import React, { useState } from 'react';
import { 
  Download, 
  X, 
  FileSpreadsheet, 
  Layers, 
  Table, 
  Check, 
  Play, 
  ExternalLink,
  Sparkles,
  Building2,
  Users
} from 'lucide-react';
import { createComplexSampleWorkbook } from '../services/sampleWorkbook';
import { generateJhcInventoryWorkbook } from '../services/presetSamples';
import { ExcelAnalyzer } from '../services/excelAnalyzer';
import { WorkbookAnalysis } from '../types';

interface SampleWorkbookModalProps {
  isOpen: boolean;
  onClose: () => void;
  onLoadIntoAnalyzer: (analysis: WorkbookAnalysis) => void;
}

export const SampleWorkbookModal: React.FC<SampleWorkbookModalProps> = ({
  isOpen,
  onClose,
  onLoadIntoAnalyzer
}) => {
  const [selectedSample, setSelectedSample] = useState<'jhc_inventory' | 'students'>('jhc_inventory');

  if (!isOpen) return null;

  const handleDownload = () => {
    if (selectedSample === 'jhc_inventory') {
      const buffer = generateJhcInventoryWorkbook();
      const blob = new Blob([new Uint8Array(buffer)], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Jaffna_Hindu_College_Inventory.xlsx';
      a.click();
      URL.revokeObjectURL(url);
    } else {
      const sample = createComplexSampleWorkbook();
      const blob = new Blob([sample.binaryData.buffer as ArrayBuffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'students_complex.xlsx';
      a.click();
      URL.revokeObjectURL(url);
    }
  };

  const handleLoad = () => {
    if (selectedSample === 'jhc_inventory') {
      const buffer = generateJhcInventoryWorkbook();
      const { analysis } = ExcelAnalyzer.parseBuffer(buffer, 'Jaffna Hindu College Inventory.xlsx');
      analysis.fileHash = 'jhc_inventory_master_ledger_hash';
      analysis.detectedArchetype = 'MULTI_SHEET_LEDGER';
      analysis.archetypeTitle = '🏫 Jaffna Hindu College Multi-Book Inventory & Ledger';
      analysis.archetypeBadge = '🏫 JHC Multi-Book Inventory';
      onLoadIntoAnalyzer(analysis);
      onClose();
    } else {
      const sample = createComplexSampleWorkbook();
      const { analysis } = ExcelAnalyzer.parseBuffer(sample.binaryData, sample.filename);
      analysis.fileHash = 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';
      onLoadIntoAnalyzer(analysis);
      onClose();
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden">
        {/* Modal Header */}
        <div className="p-5 border-b border-slate-200 flex items-center justify-between bg-slate-50">
          <div className="flex items-center space-x-3">
            <div className="p-2.5 rounded-xl bg-purple-600 text-white shadow-xs">
              <FileSpreadsheet className="w-6 h-6" />
            </div>
            <div>
              <h2 className="text-base font-bold text-slate-900">Pre-Configured Workbooks & Presets</h2>
              <p className="text-xs text-slate-500">
                Load authentic school inventory and academic spreadsheets directly into the engine.
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-600 hover:bg-slate-100 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Tab Selection */}
        <div className="px-6 pt-4 flex gap-2 border-b border-slate-100 bg-white">
          <button
            type="button"
            onClick={() => setSelectedSample('jhc_inventory')}
            className={`pb-2.5 px-3 text-xs font-bold border-b-2 flex items-center space-x-1.5 transition-colors ${
              selectedSample === 'jhc_inventory'
                ? 'border-purple-600 text-purple-700'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            <Building2 className="w-3.5 h-3.5" />
            <span>JHC Inventory & Departmental Ledger (20 Sheets)</span>
          </button>
          <button
            type="button"
            onClick={() => setSelectedSample('students')}
            className={`pb-2.5 px-3 text-xs font-bold border-b-2 flex items-center space-x-1.5 transition-colors ${
              selectedSample === 'students'
                ? 'border-purple-600 text-purple-700'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            <Users className="w-3.5 h-3.5" />
            <span>Students Complex (5 Sheets)</span>
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-6 overflow-y-auto space-y-4 text-xs text-slate-600">
          {selectedSample === 'jhc_inventory' ? (
            <div className="space-y-3">
              <div className="p-3.5 rounded-xl border border-purple-200 bg-purple-50/50 space-y-1.5">
                <div className="font-bold text-purple-950 flex items-center space-x-1.5">
                  <Sparkles className="w-4 h-4 text-purple-600" />
                  <span>Jaffna Hindu College Multi-Book Inventory Preset</span>
                </div>
                <p className="text-slate-600 leading-relaxed text-[11px]">
                  Matches the authentic institutional Google Spreadsheet with 20 sheets: Master Inventory consolidation (1,741 items across 64 book columns), Responsible Person directory, and 18 departmental books.
                </p>
              </div>

              <div className="space-y-2">
                <div className="p-3 rounded-xl border border-slate-200 bg-slate-50">
                  <div className="font-semibold text-slate-900 flex items-center justify-between">
                    <span>1. Master Inventory (Consolidation Sheet)</span>
                    <span className="text-[10px] bg-purple-100 text-purple-800 px-2 py-0.5 rounded font-mono">
                      1,741 Rows • 64 Columns
                    </span>
                  </div>
                  <p className="mt-1 text-slate-500">
                    Aggregates ledger balances, physical balances on hand, and stock distribution across 18 departmental book columns.
                  </p>
                </div>

                <div className="p-3 rounded-xl border border-slate-200 bg-slate-50">
                  <div className="font-semibold text-slate-900 flex items-center justify-between">
                    <span>2. Responsible Person Directory</span>
                    <span className="text-[10px] bg-blue-100 text-blue-800 px-2 py-0.5 rounded font-mono">
                      Department Custodians
                    </span>
                  </div>
                  <p className="mt-1 text-slate-500">
                    Directory of teachers and officers accountable for each book ledger with contact numbers.
                  </p>
                </div>

                <div className="p-3 rounded-xl border border-slate-200 bg-slate-50">
                  <div className="font-semibold text-slate-900 flex items-center justify-between">
                    <span>3. 18 Departmental Books (Admin, ICT, Labs, Sports, Hostel, Tech)</span>
                    <span className="text-[10px] bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded font-mono">
                      Stock Audits
                    </span>
                  </div>
                  <p className="mt-1 text-slate-500">
                    Departmental registers with item descriptions, section references, ledger balances, actual on-hand balances, and surplus/deficiency flags.
                  </p>
                </div>
              </div>
            </div>
          ) : (
            <div className="space-y-3">
              <div className="p-3 rounded-xl border border-slate-200 bg-slate-50">
                <div className="font-semibold text-slate-900 flex items-center justify-between">
                  <span>1. Student Details (Worksheet)</span>
                  <span className="text-[10px] bg-purple-100 text-purple-800 px-2 py-0.5 rounded font-mono">
                    A1:H1 Title • A3:H3 & A13:H13 Section Headings
                  </span>
                </div>
                <p className="mt-1 text-slate-500">
                  Contains Title at row 1, Section Heading <code className="text-purple-700 font-bold">CLASS 10A</code> at row 3 (propagated to 7 student rows), and Section Heading <code className="text-purple-700 font-bold">CLASS 10B</code> at row 13 (propagated to 5 student rows).
                </p>
              </div>

              <div className="p-3 rounded-xl border border-slate-200 bg-slate-50">
                <div className="font-semibold text-slate-900 flex items-center justify-between">
                  <span>2. Attendance (Worksheet)</span>
                  <span className="text-[10px] bg-blue-100 text-blue-800 px-2 py-0.5 rounded font-mono">
                    P/A/L/E Status Codes
                  </span>
                </div>
                <p className="mt-1 text-slate-500">
                  Normalizes single-letter status codes: <code className="font-bold">P</code> → Present, <code className="font-bold">A</code> → Absent, <code className="font-bold">L</code> → Late, <code className="font-bold">E</code> → Excused.
                </p>
              </div>

              <div className="p-3 rounded-xl border border-slate-200 bg-slate-50">
                <div className="font-semibold text-slate-900 flex items-center justify-between">
                  <span>3. Sports & 4. Medical (Worksheets)</span>
                  <span className="text-[10px] bg-teal-100 text-teal-800 px-2 py-0.5 rounded font-mono">
                    Relational Foreign Keys
                  </span>
                </div>
                <p className="mt-1 text-slate-500">
                  Maps student activities, medical clearances (Yes/No boolean conversion), blood groups, and doctor emergency contacts.
                </p>
              </div>

              <div className="p-3 rounded-xl border border-slate-200 bg-slate-50">
                <div className="font-semibold text-slate-900 flex items-center justify-between">
                  <span>5. Results (Worksheet)</span>
                  <span className="text-[10px] bg-indigo-100 text-indigo-800 px-2 py-0.5 rounded font-mono">
                    Academic Scores
                  </span>
                </div>
                <p className="mt-1 text-slate-500">
                  Term exams with numeric scores, grade letters, and composite unique keys <code className="font-mono">(student_number, term)</code>.
                </p>
              </div>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="p-4 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
          <button
            onClick={handleDownload}
            className="inline-flex items-center space-x-1.5 px-4 py-2 rounded-lg border border-slate-300 text-xs font-semibold text-slate-700 bg-white hover:bg-slate-50 shadow-2xs"
          >
            <Download className="w-4 h-4 text-slate-500" />
            <span>Download .xlsx File</span>
          </button>

          <button
            onClick={handleLoad}
            className="inline-flex items-center space-x-1.5 px-4 py-2 rounded-lg bg-purple-600 hover:bg-purple-700 text-xs font-semibold text-white shadow-2xs"
          >
            <Play className="w-4 h-4" />
            <span>Load into Analyzer & Dry Run</span>
          </button>
        </div>
      </div>
    </div>
  );
};

