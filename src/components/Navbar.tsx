import React from 'react';
import { 
  Cloud, 
  Database, 
  Cpu, 
  RefreshCw, 
  Download, 
  Sparkles, 
  ShieldCheck,
  CheckCircle2,
  AlertCircle
} from 'lucide-react';
import { NextcloudConfig, SupabaseConfig, SyncSettings } from '../types';

interface NavbarProps {
  nextcloud: NextcloudConfig;
  supabase: SupabaseConfig;
  syncSettings: SyncSettings;
  onRefreshAll: () => void;
  onOpenSampleModal: () => void;
  onOpenAiAssistant: () => void;
}

export const Navbar: React.FC<NavbarProps> = ({
  nextcloud,
  supabase,
  syncSettings,
  onRefreshAll,
  onOpenSampleModal,
  onOpenAiAssistant
}) => {
  return (
    <header className="sticky top-0 z-40 bg-white border-b border-slate-200 shadow-xs">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex items-center justify-between h-16">
          {/* Brand Logo & Name */}
          <div className="flex items-center space-x-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-700 flex items-center justify-center text-white font-bold shadow-xs">
              <span className="text-lg tracking-tighter font-mono">NX</span>
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <span className="font-semibold text-slate-900 text-lg tracking-tight">Nextcloud Excel Sync</span>
                <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200">
                  TrueNAS • Coolify • Supabase
                </span>
              </div>
              <p className="text-xs text-slate-500 hidden sm:block">
                Intelligent Excel multi-sheet & merged cell synchronizer
              </p>
            </div>
          </div>

          {/* Quick System Indicators */}
          <div className="hidden md:flex items-center space-x-4">
            {/* Nextcloud Status */}
            <div className="flex items-center space-x-1.5 px-2.5 py-1 rounded-lg bg-slate-50 border border-slate-200 text-xs">
              <Cloud className={`w-3.5 h-3.5 ${nextcloud.isConnected ? 'text-emerald-600' : 'text-amber-500'}`} />
              <span className="text-slate-600 font-medium">TrueNAS WebDAV:</span>
              <span className={nextcloud.isConnected ? 'text-emerald-700 font-semibold' : 'text-amber-600 font-medium'}>
                {nextcloud.isConnected ? 'Connected' : 'Unverified'}
              </span>
            </div>

            {/* Supabase Status */}
            <div className="flex items-center space-x-1.5 px-2.5 py-1 rounded-lg bg-slate-50 border border-slate-200 text-xs">
              <Database className={`w-3.5 h-3.5 ${
                supabase.isConnected 
                  ? 'text-teal-600' 
                  : (!supabase.url || (!supabase.anonKey && !supabase.serviceKey && !supabase.serviceRoleKey))
                    ? 'text-slate-400' 
                    : 'text-rose-500'
              }`} />
              <span className="text-slate-600 font-medium">Supabase:</span>
              <span className={
                supabase.isConnected 
                  ? 'text-teal-700 font-semibold' 
                  : (!supabase.url || (!supabase.anonKey && !supabase.serviceKey && !supabase.serviceRoleKey))
                    ? 'text-slate-500 font-medium' 
                    : 'text-rose-600 font-medium'
              }>
                {supabase.isConnected 
                  ? 'Connected' 
                  : (!supabase.url || (!supabase.anonKey && !supabase.serviceKey && !supabase.serviceRoleKey))
                    ? 'Not Configured' 
                    : 'Disconnected'}
              </span>
            </div>

            {/* Coolify Worker Status */}
            <div className="flex items-center space-x-1.5 px-2.5 py-1 rounded-lg bg-slate-50 border border-slate-200 text-xs">
              <Cpu className="w-3.5 h-3.5 text-blue-600" />
              <span className="text-slate-600 font-medium">Worker:</span>
              <span className="text-blue-700 font-semibold">{syncSettings.syncInterval}</span>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="flex items-center space-x-2">
            <button
              id="btn-sample-workbook"
              onClick={onOpenSampleModal}
              className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg border border-slate-300 bg-white text-xs font-medium text-slate-700 hover:bg-slate-50 transition-colors shadow-2xs"
              title="Download or preview students.xlsx complex sample"
            >
              <Download className="w-3.5 h-3.5 text-slate-500" />
              <span className="hidden sm:inline">Sample Workbook</span>
            </button>

            <button
              id="btn-ai-assistant"
              onClick={onOpenAiAssistant}
              className="inline-flex items-center space-x-1.5 px-3 py-1.5 rounded-lg bg-purple-50 border border-purple-200 text-xs font-medium text-purple-700 hover:bg-purple-100 transition-colors shadow-2xs"
            >
              <Sparkles className="w-3.5 h-3.5 text-purple-600" />
              <span>AI Mapping Assistant</span>
            </button>

            <button
              id="btn-refresh-all"
              onClick={onRefreshAll}
              className="p-1.5 rounded-lg border border-slate-300 text-slate-600 hover:bg-slate-100 transition-colors"
              title="Refresh all metrics"
            >
              <RefreshCw className="w-4 h-4" />
            </button>
          </div>
        </div>
      </div>
    </header>
  );
};
