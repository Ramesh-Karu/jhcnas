import { ColumnMapping, RowValidationError, TransformationType, DataType } from '../types';

export class MappingEngine {
  static applyTransformation(
    rawVal: any,
    transformation: TransformationType,
    defaultValue?: string
  ): any {
    if (rawVal === null || rawVal === undefined || String(rawVal).trim() === '') {
      return defaultValue !== undefined && defaultValue !== '' ? defaultValue : null;
    }

    const valStr = String(rawVal).trim();

    switch (transformation) {
      case 'trim':
        return valStr;

      case 'uppercase':
        return valStr.toUpperCase();

      case 'lowercase':
        return valStr.toLowerCase();

      case 'parse_date': {
        // If already in ISO or date format
        const isoMatch = valStr.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (isoMatch) {
          const y = isoMatch[1];
          const m = isoMatch[2].padStart(2, '0');
          const d = isoMatch[3].padStart(2, '0');
          return `${y}-${m}-${d}`;
        }
        // Try parsing JS Date
        const parsed = new Date(valStr);
        if (!isNaN(parsed.getTime())) {
          return parsed.toISOString().split('T')[0];
        }
        return valStr;
      }

      case 'parse_number': {
        const clean = valStr.replace(/[^\d.-]/g, '');
        const num = Number(clean);
        return isNaN(num) ? rawVal : num;
      }

      case 'yes_no_to_boolean': {
        const lower = valStr.toLowerCase();
        if (['yes', 'y', 'true', '1', 't'].includes(lower)) return true;
        if (['no', 'n', 'false', '0', 'f'].includes(lower)) return false;
        return defaultValue ? defaultValue === 'true' : null;
      }

      case 'pa_to_status': {
        const upper = valStr.toUpperCase();
        const map: Record<string, string> = {
          P: 'Present',
          PRESENT: 'Present',
          A: 'Absent',
          ABSENT: 'Absent',
          L: 'Late',
          LATE: 'Late',
          E: 'Excused',
          EXCUSED: 'Excused'
        };
        return map[upper] || valStr;
      }

      case 'normalize_phone': {
        const hasPlus = valStr.startsWith('+');
        const digits = valStr.replace(/\D/g, '');
        return hasPlus ? `+${digits}` : digits;
      }

      case 'normalize_id': {
        return valStr.replace(/\s+/g, '').toUpperCase();
      }

      default:
        return valStr;
    }
  }

  static transformAndValidate(
    rawRecord: Record<string, any>,
    colMappings: ColumnMapping[],
    worksheetName: string,
    sectionHeadingTargetCol?: string,
    seenUniqueKeys?: Set<string>
  ): { cleanRecord: Record<string, any>; errors: RowValidationError[] } {
    const rowNumber = rawRecord.__rowNumber || 0;
    const cleanRecord: Record<string, any> = {};
    const errors: RowValidationError[] = [];

    // Assign section heading if mapped (e.g. class = 'CLASS 10A')
    if (sectionHeadingTargetCol && rawRecord.__sectionHeading) {
      cleanRecord[sectionHeadingTargetCol] = rawRecord.__sectionHeading;
    }

    for (const cm of colMappings) {
      // Find raw value by column letter or header name
      let val = rawRecord[cm.excelColumn];
      if (val === undefined || val === null) {
        val = rawRecord[`header_${cm.excelHeader}`];
      }

      // 1. Transformation
      const transformed = this.applyTransformation(val, cm.transformation, cm.defaultValue);
      cleanRecord[cm.supabaseColumn] = transformed;

      // 2. Required Check
      if (cm.required && (transformed === null || transformed === undefined || String(transformed).trim() === '')) {
        errors.push({
          worksheetName,
          rowNumber,
          excelColumn: cm.excelColumn,
          columnName: cm.supabaseColumn,
          rawValue: String(val ?? ''),
          errorMessage: `Row ${rowNumber}: Required field '${cm.supabaseColumn}' is missing or empty.`,
          errorType: 'missing_required'
        });
        continue;
      }

      if (transformed === null || transformed === undefined) continue;

      // 3. Type Validation
      const strVal = String(transformed).trim();
      switch (cm.dataType) {
        case 'integer': {
          const num = Number(strVal);
          if (isNaN(num) || !Number.isInteger(num)) {
            errors.push({
              worksheetName,
              rowNumber,
              excelColumn: cm.excelColumn,
              columnName: cm.supabaseColumn,
              rawValue: strVal,
              errorMessage: `Row ${rowNumber}: Field '${cm.supabaseColumn}' value '${strVal}' is not a valid integer.`,
              errorType: 'type_mismatch'
            });
          }
          break;
        }

        case 'decimal': {
          const num = Number(strVal);
          if (isNaN(num)) {
            errors.push({
              worksheetName,
              rowNumber,
              excelColumn: cm.excelColumn,
              columnName: cm.supabaseColumn,
              rawValue: strVal,
              errorMessage: `Row ${rowNumber}: Field '${cm.supabaseColumn}' value '${strVal}' is not a valid decimal number.`,
              errorType: 'type_mismatch'
            });
          }
          break;
        }

        case 'date': {
          if (!/^\d{4}-\d{2}-\d{2}$/.test(strVal)) {
            errors.push({
              worksheetName,
              rowNumber,
              excelColumn: cm.excelColumn,
              columnName: cm.supabaseColumn,
              rawValue: strVal,
              errorMessage: `Row ${rowNumber}: Invalid date format '${strVal}'. Expected YYYY-MM-DD.`,
              errorType: 'invalid_date'
            });
          }
          break;
        }

        case 'boolean': {
          if (typeof transformed !== 'boolean' && !['true', 'false', '1', '0'].includes(strVal.toLowerCase())) {
            errors.push({
              worksheetName,
              rowNumber,
              excelColumn: cm.excelColumn,
              columnName: cm.supabaseColumn,
              rawValue: strVal,
              errorMessage: `Row ${rowNumber}: Field '${cm.supabaseColumn}' value '${strVal}' is not a valid boolean.`,
              errorType: 'type_mismatch'
            });
          }
          break;
        }
      }

      // 4. Duplicate Check within batch
      if (cm.uniqueKey && transformed !== null && seenUniqueKeys) {
        const uniqueValKey = `${cm.supabaseColumn}:${strVal}`;
        if (seenUniqueKeys.has(uniqueValKey)) {
          errors.push({
            worksheetName,
            rowNumber,
            excelColumn: cm.excelColumn,
            columnName: cm.supabaseColumn,
            rawValue: strVal,
            errorMessage: `Row ${rowNumber}: Duplicate unique key '${strVal}' for '${cm.supabaseColumn}' detected in batch.`,
            errorType: 'duplicate'
          });
        } else {
          seenUniqueKeys.add(uniqueValKey);
        }
      }
    }

    return { cleanRecord, errors };
  }
}
