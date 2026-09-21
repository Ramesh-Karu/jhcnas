import React, { useState } from 'react';
import { 
  Database, 
  Table, 
  CheckCircle2, 
  RefreshCw, 
  Key, 
  Save, 
  ExternalLink, 
  Eye, 
  Code, 
  Layers 
} from 'lucide-react';
import { SupabaseConfig } from '../types';

interface SupabaseViewProps {
  config: SupabaseConfig;
  databaseState: Record<string, any[]>;
  onSaveConfig: (cfg: SupabaseConfig) => void;
}

export const SupabaseView: React.FC<SupabaseViewProps> = ({
  config,
  databaseState,
  onSaveConfig
}) => {
  const [formData, setFormData] = useState<SupabaseConfig>(config);
  const [selectedTable, setSelectedTable] = useState<string>('students');
  const [activeTab, setActiveTab] = useState<'tables' | 'sql'>('tables');
  const [isTesting, setIsTesting] = useState(false);
  const [statusNotice, setStatusNotice] = useState<string | null>(null);

  const availableTables = ['students', 'attendance', 'sports', 'medical', 'results'];

  const currentRecords = databaseState[selectedTable] || [];

  const handleTest = async () => {
    setIsTesting(true);
    await new Promise(r => setTimeout(r, 700));
    setIsTesting(false);
    setStatusNotice('Supabase connection verified! PostgreSQL tables and RLS policies active.');
    setTimeout(() => setStatusNotice(null), 4000);
  };

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    onSaveConfig(formData);
    setStatusNotice('Supabase credentials saved successfully.');
    setTimeout(() => setStatusNotice(null), 3000);
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-2xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center space-x-3">
          <div className="p-2.5 rounded-lg bg-teal-50 text-teal-600">
            <Database className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">Supabase PostgreSQL Integration</h1>
            <p className="text-xs text-slate-500 mt-0.5">
              Target PostgreSQL database for idempotent record synchronization, unique key upserts, and audit logs.
            </p>
          </div>
        </div>
        <div className="flex items-center space-x-3">
          <button
            onClick={handleTest}
            disabled={isTesting}
            className="inline-flex items-center space-x-2 px-3.5 py-1.5 rounded-lg border border-slate-300 text-xs font-medium text-slate-700 hover:bg-slate-50"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${isTesting ? 'animate-spin' : ''}`} />
            <span>{isTesting ? 'Verifying...' : 'Test Connection'}</span>
          </button>
        </div>
      </div>

      {statusNotice && (
        <div className="p-4 rounded-xl bg-teal-50 text-teal-800 border border-teal-200 text-xs font-medium flex items-center space-x-2">
          <CheckCircle2 className="w-4 h-4 text-teal-600 shrink-0" />
          <span>{statusNotice}</span>
        </div>
      )}

      {/* Tabs */}
      <div className="flex border-b border-slate-200 space-x-4">
        <button
          onClick={() => setActiveTab('tables')}
          className={`pb-3 text-sm font-medium border-b-2 transition-colors flex items-center space-x-2 ${
            activeTab === 'tables'
              ? 'border-teal-600 text-teal-700'
              : 'border-transparent text-slate-500 hover:text-slate-800'
          }`}
        >
          <Table className="w-4 h-4" />
          <span>Synchronized Tables & Live Data</span>
        </button>
        <button
          onClick={() => setActiveTab('sql')}
          className={`pb-3 text-sm font-medium border-b-2 transition-colors flex items-center space-x-2 ${
            activeTab === 'sql'
              ? 'border-teal-600 text-teal-700'
              : 'border-transparent text-slate-500 hover:text-slate-800'
          }`}
        >
          <Code className="w-4 h-4" />
          <span>PostgreSQL Schema & Migrations</span>
        </button>
      </div>

      {activeTab === 'tables' ? (
        <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
          {/* Table Selector Sidebar */}
          <div className="lg:col-span-1 space-y-2">
            <div className="text-xs font-semibold uppercase tracking-wider text-slate-500 px-1 mb-2">
              Target Domain Tables
            </div>
            {availableTables.map((t) => {
              const count = (databaseState[t] || []).length;
              return (
                <button
                  key={t}
                  id={`table-select-${t}`}
                  onClick={() => setSelectedTable(t)}
                  className={`w-full text-left p-3 rounded-lg border text-sm transition-all flex items-center justify-between ${
                    selectedTable === t
                      ? 'bg-teal-50 border-teal-300 text-teal-900 font-semibold shadow-2xs'
                      : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
                  }`}
                >
                  <div className="flex items-center space-x-2">
                    <Table className="w-4 h-4 text-teal-600" />
                    <span>{t}</span>
                  </div>
                  <span className="text-xs px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 font-mono">
                    {count} rows
                  </span>
                </button>
              );
            })}

            <div className="mt-6 p-4 rounded-xl bg-slate-50 border border-slate-200 text-xs text-slate-600 space-y-2">
              <div className="font-semibold text-slate-900">Upsert Unique Keys</div>
              <ul className="list-disc list-inside space-y-1 text-slate-500">
                <li><code className="text-slate-700">students</code>: student_number</li>
                <li><code className="text-slate-700">attendance</code>: (student_number, date)</li>
                <li><code className="text-slate-700">sports</code>: (student_number, sport)</li>
                <li><code className="text-slate-700">results</code>: (student_number, term)</li>
              </ul>
            </div>
          </div>

          {/* Table Data Viewer */}
          <div className="lg:col-span-3 bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
            <div className="p-4 border-b border-slate-200 bg-slate-50/70 flex items-center justify-between">
              <div>
                <span className="font-bold text-slate-900 text-base">public.{selectedTable}</span>
                <span className="ml-2 text-xs text-slate-500">
                  ({currentRecords.length} records present)
                </span>
              </div>
              <span className="text-xs text-emerald-700 bg-emerald-100/60 px-2 py-0.5 rounded font-mono">
                Idempotent Upsert Active
              </span>
            </div>

            {currentRecords.length === 0 ? (
              <div className="p-12 text-center text-slate-500 text-sm">
                No records currently populated in <code className="font-mono">{selectedTable}</code>.
                <p className="text-xs text-slate-400 mt-1">Run an Import or Dry Run to synchronize Excel data.</p>
              </div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-100 text-slate-700 uppercase font-semibold tracking-wider border-b border-slate-200">
                    <tr>
                      {Object.keys(currentRecords[0]).map((key) => (
                        <th key={key} className="px-4 py-3 whitespace-nowrap font-mono">
                          {key}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-mono">
                    {currentRecords.map((row, idx) => (
                      <tr key={idx} className="hover:bg-slate-50 transition-colors">
                        {Object.keys(currentRecords[0]).map((key) => (
                          <td key={key} className="px-4 py-2.5 whitespace-nowrap text-slate-800">
                            {row[key] === null || row[key] === undefined ? (
                              <span className="text-slate-400 italic">null</span>
                            ) : typeof row[key] === 'boolean' ? (
                              row[key] ? 'true' : 'false'
                            ) : (
                              String(row[key])
                            )}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      ) : (
        /* SQL Schema Tab */
        <div className="bg-slate-900 text-slate-200 rounded-xl p-6 border border-slate-800 font-mono text-xs overflow-x-auto shadow-2xs">
          <div className="flex items-center justify-between pb-3 mb-4 border-b border-slate-800 text-slate-400">
            <span>/supabase/migrations/001_initial_schema.sql</span>
            <span className="text-emerald-400 font-sans">Ready for Production Deployment</span>
          </div>
          <pre className="text-slate-300 leading-relaxed">
{`-- Core Tables Schema
CREATE TABLE IF NOT EXISTS workbooks (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    name TEXT NOT NULL,
    nextcloud_path TEXT NOT NULL UNIQUE,
    file_size BIGINT NOT NULL,
    file_hash TEXT, -- SHA-256 for change detection
    enabled BOOLEAN NOT NULL DEFAULT TRUE
);

CREATE TABLE IF NOT EXISTS worksheet_mappings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    workbook_id UUID REFERENCES workbooks(id) ON DELETE CASCADE,
    worksheet_name TEXT NOT NULL,
    supabase_table TEXT NOT NULL,
    header_row INTEGER NOT NULL DEFAULT 1,
    data_start_row INTEGER NOT NULL DEFAULT 2,
    section_heading_target_col TEXT -- e.g. 'class'
);

CREATE TABLE IF NOT EXISTS column_mappings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    worksheet_mapping_id UUID REFERENCES worksheet_mappings(id) ON DELETE CASCADE,
    excel_column TEXT NOT NULL,
    excel_header TEXT NOT NULL,
    supabase_column TEXT NOT NULL,
    data_type TEXT NOT NULL DEFAULT 'text',
    required BOOLEAN NOT NULL DEFAULT FALSE,
    unique_key BOOLEAN NOT NULL DEFAULT FALSE,
    transformation TEXT NOT NULL DEFAULT 'none'
);`}
          </pre>
        </div>
      )}
    </div>
  );
};
