import React from 'react';
import { 
  Cloud, 
  Database, 
  Cpu, 
  RefreshCw, 
  Download, 
  Sparkles, 
  Key,
  Menu,
  ShieldCheck
} from 'lucide-react';
import { NextcloudConfig, SupabaseConfig, SyncSettings, LiveSchedulerStatus } from '../types';

interface NavbarProps {
  nextcloud: NextcloudConfig;
  supabase: SupabaseConfig;
  syncSettings: SyncSettings;
  onRefreshAll: () => void;
  onOpenSampleModal: () => void;
  onOpenAiAssistant: () => void;
  onOpenSecretsVault?: () => void;
  onToggleMobileMenu?: () => void;
  allSecretsConnected?: boolean;
  schedulerStatus?: LiveSchedulerStatus;
}

export const Navbar: React.FC<NavbarProps> = ({
  nextcloud,
  supabase,
  syncSettings,
  onRefreshAll,
  onOpenSampleModal,
  onOpenAiAssistant,
  onOpenSecretsVault,
  onToggleMobileMenu,
  allSecretsConnected = true,
  schedulerStatus,
}) => {
  const isSupabaseConfigured = Boolean(
    supabase.url && (supabase.anonKey || supabase.serviceKey || supabase.serviceRoleKey)
  );

  return (
    <header className="sticky top-0 z-30 bg-white border-b border-slate-200 shadow-2xs">
      <div className="w-full px-3 sm:px-6">
        <div className="flex items-center justify-between h-16 gap-2">
          
          {/* Left Zone: Mobile Hamburger + Brand Logo & Title */}
          <div className="flex items-center space-x-2.5 sm:space-x-3 shrink-0">
            {/* Mobile Hamburger Toggle Button */}
            {onToggleMobileMenu && (
              <button
                id="btn-mobile-menu-toggle"
                onClick={onToggleMobileMenu}
                className="md:hidden p-2 rounded-lg text-slate-600 hover:text-slate-900 hover:bg-slate-100 transition-colors"
                aria-label="Toggle Navigation Menu"
              >
                <Menu className="w-5 h-5" />
              </button>
            )}

            <div className="w-9 h-9 sm:w-10 sm:h-10 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-700 flex items-center justify-center text-white font-bold shadow-xs shrink-0">
              <span className="text-sm sm:text-base tracking-tighter font-mono">NX</span>
            </div>

            <div className="min-w-0">
              <div className="flex items-center space-x-2">
                <span className="font-bold text-slate-900 text-sm sm:text-base tracking-tight truncate">
                  Nextcloud Excel Sync
                </span>
                <span className="hidden lg:inline-flex items-center px-2 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">
                  TrueNAS • Coolify • Supabase
                </span>
              </div>
              <p className="text-[11px] text-slate-500 hidden sm:block truncate">
                Multi-sheet workbook schema inference & automated database sync
              </p>
            </div>
          </div>

          {/* Center Zone: System Subsystem Health Indicators (Desktop) */}
          <div className="hidden xl:flex items-center space-x-2.5">
            {/* Nextcloud Status */}
            <div className="flex items-center space-x-1.5 px-2.5 py-1 rounded-lg bg-slate-50 border border-slate-200 text-xs">
              <Cloud className={`w-3.5 h-3.5 ${nextcloud.isConnected ? 'text-emerald-600' : 'text-amber-500'}`} />
              <span className="text-slate-600 font-medium">TrueNAS WebDAV:</span>
              <span className={nextcloud.isConnected ? 'text-emerald-700 font-bold' : 'text-amber-600 font-medium'}>
                {nextcloud.isConnected ? 'Connected' : 'Unverified'}
              </span>
            </div>

            {/* Supabase Status */}
            <div className="flex items-center space-x-1.5 px-2.5 py-1 rounded-lg bg-slate-50 border border-slate-200 text-xs">
              <Database className={`w-3.5 h-3.5 ${
                supabase.isConnected 
                  ? 'text-teal-600' 
                  : !isSupabaseConfigured
                    ? 'text-slate-400' 
                    : 'text-rose-500'
              }`} />
              <span className="text-slate-600 font-medium">Supabase:</span>
              <span className={
                supabase.isConnected 
                  ? 'text-teal-700 font-bold' 
                  : !isSupabaseConfigured
                    ? 'text-slate-500 font-medium' 
                    : 'text-rose-600 font-medium'
              }>
                {supabase.isConnected ? 'Connected' : !isSupabaseConfigured ? 'Not Configured' : 'Disconnected'}
              </span>
            </div>

            {/* Production Engine / Scheduler Status */}
            <div className="flex items-center space-x-1.5 px-2.5 py-1 rounded-lg bg-slate-50 border border-slate-200 text-xs">
              <Cpu className="w-3.5 h-3.5 text-indigo-600" />
              <span className="text-slate-600 font-medium">Engine:</span>
              {schedulerStatus?.state === 'SYNCING' ? (
                <span className="text-amber-700 font-bold font-mono inline-flex items-center space-x-1">
                  <RefreshCw className="w-3 h-3 animate-spin text-amber-600" />
                  <span>Syncing...</span>
                </span>
              ) : schedulerStatus?.enabled ? (
                <span className="text-emerald-700 font-bold font-mono">
                  Auto ({schedulerStatus.intervalLabel}
                  {schedulerStatus.secondsUntilNextRun !== null && schedulerStatus.secondsUntilNextRun < 60
                    ? ' • <1m'
                    : schedulerStatus.secondsUntilNextRun !== null
                      ? ` • in ${Math.ceil(schedulerStatus.secondsUntilNextRun / 60)}m`
                      : ''})
                </span>
              ) : (
                <span className="text-slate-600 font-medium font-mono">Manual Trigger</span>
              )}
            </div>
          </div>

          {/* Right Zone: Primary Actions */}
          <div className="flex items-center space-x-1.5 sm:space-x-2 shrink-0">
            {/* Unified Secrets Vault Button */}
            {onOpenSecretsVault && (
              <button
                id="btn-secrets-vault"
                onClick={onOpenSecretsVault}
                className={`inline-flex items-center space-x-1.5 px-2.5 sm:px-3 py-1.5 rounded-lg text-xs font-semibold transition-all shadow-2xs border ${
                  allSecretsConnected
                    ? 'bg-white border-slate-300 text-slate-700 hover:bg-slate-50'
                    : 'bg-amber-50 border-amber-300 text-amber-900 hover:bg-amber-100'
                }`}
                title="Verify and edit secrets for Nextcloud, Supabase, and Worker Array"
              >
                <Key className={`w-3.5 h-3.5 ${allSecretsConnected ? 'text-amber-500' : 'text-amber-600 animate-bounce'}`} />
                <span className="hidden sm:inline">Secrets Vault</span>
                <span className={`w-2 h-2 rounded-full ${allSecretsConnected ? 'bg-emerald-500' : 'bg-amber-500 animate-pulse'}`} />
              </button>
            )}

            {/* Sample Workbook Button */}
            <button
              id="btn-sample-workbook"
              onClick={onOpenSampleModal}
              className="inline-flex items-center space-x-1.5 px-2.5 sm:px-3 py-1.5 rounded-lg border border-slate-300 bg-white text-xs font-medium text-slate-700 hover:bg-slate-50 transition-colors shadow-2xs"
              title="Download or preview students.xlsx complex sample"
            >
              <Download className="w-3.5 h-3.5 text-slate-500" />
              <span className="hidden md:inline">Sample Workbook</span>
            </button>

            {/* AI Assistant Button */}
            <button
              id="btn-ai-assistant"
              onClick={onOpenAiAssistant}
              className="inline-flex items-center space-x-1.5 px-2.5 sm:px-3 py-1.5 rounded-lg bg-purple-50 border border-purple-200 text-xs font-semibold text-purple-700 hover:bg-purple-100 transition-colors shadow-2xs"
            >
              <Sparkles className="w-3.5 h-3.5 text-purple-600" />
              <span className="hidden sm:inline">AI Assistant</span>
            </button>

            {/* Refresh Button */}
            <button
              id="btn-refresh-all"
              onClick={onRefreshAll}
              className="p-1.5 sm:p-2 rounded-lg border border-slate-300 text-slate-600 hover:bg-slate-100 transition-colors"
              title="Refresh all metrics and files"
            >
              <RefreshCw className="w-4 h-4" />
            </button>
          </div>

        </div>
      </div>
    </header>
  );
};
