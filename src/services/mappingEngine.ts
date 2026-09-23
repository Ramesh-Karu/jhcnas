import { ColumnMapping, RowValidationError, TransformationType, DataType } from '../types';

export class MappingEngine {
  static parseFlexibleDate(val: any): string | null {
    if (val === null || val === undefined || String(val).trim() === '') {
      return null;
    }

    // 1. If it's a native Date object
    if (val instanceof Date && !isNaN(val.getTime())) {
      return val.toISOString().split('T')[0];
    }

    const valStr = String(val).trim();

    // 2. If it's an Excel numeric serial date (e.g. 39448)
    const numericSerial = Number(valStr);
    if (!isNaN(numericSerial) && numericSerial > 1000 && numericSerial < 100000) {
      try {
        const dateObj = new Date(Math.round((numericSerial - 25569) * 86400 * 1000));
        if (!isNaN(dateObj.getTime())) {
          return dateObj.toISOString().split('T')[0];
        }
      } catch {}
    }

    // 3. ISO format: YYYY-MM-DD or YYYY/MM/DD
    const isoMatch = valStr.match(/^(\d{4})[-/. ](\d{1,2})[-/. ](\d{1,2})/);
    if (isoMatch) {
      const y = isoMatch[1];
      const m = isoMatch[2].padStart(2, '0');
      const d = isoMatch[3].padStart(2, '0');
      return `${y}-${m}-${d}`;
    }

    // 4. European / Commonwealth format: DD/MM/YYYY or DD-MM-YYYY
    const dmyMatch = valStr.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
    if (dmyMatch) {
      let d = parseInt(dmyMatch[1], 10);
      let m = parseInt(dmyMatch[2], 10);
      const y = dmyMatch[3];
      // Disambiguate if m > 12 (then it must be MM/DD/YYYY)
      if (m > 12 && d <= 12) {
        const tmp = d;
        d = m;
        m = tmp;
      }
      return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    }

    // 5. General JS Date parsing fallback
    const parsed = new Date(valStr);
    if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 1900) {
      return parsed.toISOString().split('T')[0];
    }

    return null;
  }

  static applyTransformation(
    rawVal: any,
    transformation: TransformationType,
    defaultValue?: string,
    targetDataType?: DataType
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
        const parsed = this.parseFlexibleDate(rawVal);
        return parsed !== null ? parsed : valStr;
      }

      case 'parse_number': {
        const clean = valStr.replace(/[$,€£₹% ]/g, '').replace(/,/g, '');
        const num = Number(clean);
        return isNaN(num) ? rawVal : num;
      }

      case 'yes_no_to_boolean': {
        const lower = valStr.toLowerCase();
        if (['yes', 'y', 'true', '1', 't', 'p'].includes(lower)) return true;
        if (['no', 'n', 'false', '0', 'f', 'a'].includes(lower)) return false;
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
        // If data type is date or integer, automatically normalize even if transformation was not explicitly selected
        if (targetDataType === 'date') {
          const parsed = this.parseFlexibleDate(rawVal);
          return parsed !== null ? parsed : valStr;
        }
        if (targetDataType === 'integer' || targetDataType === 'decimal') {
          const clean = valStr.replace(/[$,€£₹% ]/g, '').replace(/,/g, '');
          const num = Number(clean);
          return isNaN(num) ? valStr : num;
        }
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
      if (val === undefined || val === null) {
        val = rawRecord[cm.excelHeader];
      }

      // 1. Transformation (with intelligent datatype awareness)
      const transformed = this.applyTransformation(val, cm.transformation, cm.defaultValue, cm.dataType);

      // Check if value is truly empty
      const isEmpty = transformed === null || transformed === undefined || String(transformed).trim() === '';

      // 2. Required Check
      if (cm.required && isEmpty) {
        errors.push({
          worksheetName,
          rowNumber,
          excelColumn: cm.excelColumn,
          columnName: cm.supabaseColumn,
          rawValue: String(val ?? ''),
          errorMessage: `Row ${rowNumber}: Required field '${cm.supabaseColumn}' is missing or empty.`,
          errorType: 'missing_required'
        });
        cleanRecord[cm.supabaseColumn] = null;
        continue;
      }

      // If optional and empty, coerce to null and skip further validation!
      if (isEmpty) {
        cleanRecord[cm.supabaseColumn] = null;
        continue;
      }

      // 3. Type Validation & Normalization
      const strVal = String(transformed).trim();
      let coercedVal = transformed;

      switch (cm.dataType) {
        case 'integer': {
          const cleanNumStr = strVal.replace(/[$,€£₹% ]/g, '').replace(/,/g, '');
          const num = Number(cleanNumStr);
          if (isNaN(num)) {
            errors.push({
              worksheetName,
              rowNumber,
              excelColumn: cm.excelColumn,
              columnName: cm.supabaseColumn,
              rawValue: strVal,
              errorMessage: `Row ${rowNumber}: Field '${cm.supabaseColumn}' value '${strVal}' is not a valid integer.`,
              errorType: 'type_mismatch'
            });
          } else {
            coercedVal = Math.round(num);
          }
          break;
        }

        case 'decimal': {
          const cleanNumStr = strVal.replace(/[$,€£₹% ]/g, '').replace(/,/g, '');
          const num = Number(cleanNumStr);
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
          } else {
            coercedVal = num;
          }
          break;
        }

        case 'date': {
          const parsedDate = this.parseFlexibleDate(transformed);
          if (!parsedDate) {
            errors.push({
              worksheetName,
              rowNumber,
              excelColumn: cm.excelColumn,
              columnName: cm.supabaseColumn,
              rawValue: strVal,
              errorMessage: `Row ${rowNumber}: Invalid date '${strVal}' for '${cm.supabaseColumn}'. Could not parse to standard YYYY-MM-DD.`,
              errorType: 'invalid_date'
            });
          } else {
            coercedVal = parsedDate;
          }
          break;
        }

        case 'boolean': {
          if (typeof transformed === 'boolean') {
            coercedVal = transformed;
          } else {
            const lower = strVal.toLowerCase();
            if (['true', '1', 'yes', 'y', 't', 'p'].includes(lower)) {
              coercedVal = true;
            } else if (['false', '0', 'no', 'n', 'f', 'a'].includes(lower)) {
              coercedVal = false;
            } else {
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
          }
          break;
        }

        default:
          coercedVal = strVal;
          break;
      }

      cleanRecord[cm.supabaseColumn] = coercedVal;

      // 4. Duplicate Check within batch
      if (cm.uniqueKey && coercedVal !== null && seenUniqueKeys) {
        const uniqueValKey = `${cm.supabaseColumn}:${String(coercedVal).toLowerCase()}`;
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
