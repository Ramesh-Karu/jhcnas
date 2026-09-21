import React from 'react';
import { 
  Download, 
  X, 
  FileSpreadsheet, 
  Layers, 
  Table, 
  Check, 
  Play, 
  ExternalLink 
} from 'lucide-react';
import { createComplexSampleWorkbook } from '../services/sampleWorkbook';
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
  if (!isOpen) return null;

  const sample = createComplexSampleWorkbook();

  const handleDownload = () => {
    const blob = new Blob([sample.binaryData.buffer as ArrayBuffer], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'students_complex.xlsx';
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleLoad = () => {
    const { analysis } = ExcelAnalyzer.parseBuffer(sample.binaryData, sample.filename);
    analysis.fileHash = 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';
    onLoadIntoAnalyzer(analysis);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden">
        {/* Modal Header */}
        <div className="p-6 border-b border-slate-200 flex items-center justify-between bg-slate-50">
          <div className="flex items-center space-x-3">
            <div className="p-2.5 rounded-xl bg-emerald-600 text-white shadow-xs">
              <FileSpreadsheet className="w-6 h-6" />
            </div>
            <div>
              <h2 className="text-lg font-bold text-slate-900">Complex Sample Workbook: students.xlsx</h2>
              <p className="text-xs text-slate-500">
                Multi-worksheet workbook with merged cell titles and section heading propagation.
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

        {/* Modal Body */}
        <div className="p-6 overflow-y-auto space-y-4 text-xs text-slate-600">
          <p className="leading-relaxed">
            This realistic workbook is constructed according to the project specifications. It illustrates how the synchronizer safely navigates complex human-formatted spreadsheets:
          </p>

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
            className="inline-flex items-center space-x-1.5 px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-xs font-semibold text-white shadow-2xs"
          >
            <Play className="w-4 h-4" />
            <span>Load into Analyzer & Dry Run</span>
          </button>
        </div>
      </div>
    </div>
  );
};
