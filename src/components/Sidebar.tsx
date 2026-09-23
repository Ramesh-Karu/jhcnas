import React from 'react';
import { 
  LayoutDashboard, 
  Cloud, 
  Database, 
  FileSpreadsheet, 
  Microscope, 
  GitFork, 
  PlaySquare, 
  ArrowLeftRight,
  History, 
  AlertTriangle, 
  Terminal, 
  Settings,
  Key,
  X
} from 'lucide-react';
import { NavigationTab, LiveSchedulerStatus } from '../types';

interface SidebarProps {
  currentTab: NavigationTab;
  onSelectTab: (tab: NavigationTab) => void;
  errorCount: number;
  filesWaitingCount: number;
  conflictCount?: number;
  isOpenMobile?: boolean;
  onCloseMobile?: () => void;
  onOpenSecretsVault?: () => void;
  allSecretsConnected?: boolean;
  schedulerStatus?: LiveSchedulerStatus;
}

export const Sidebar: React.FC<SidebarProps> = ({
  currentTab,
  onSelectTab,
  errorCount,
  filesWaitingCount,
  conflictCount = 0,
  isOpenMobile = false,
  onCloseMobile,
  onOpenSecretsVault,
  allSecretsConnected = true,
  schedulerStatus,
}) => {
  const navItems = [
    { id: 'dashboard' as NavigationTab, label: 'Dashboard', icon: LayoutDashboard },
    { id: 'nextcloud' as NavigationTab, label: 'Nextcloud WebDAV', icon: Cloud },
    { id: 'supabase' as NavigationTab, label: 'Supabase PostgreSQL', icon: Database },
    { 
      id: 'files' as NavigationTab, 
      label: 'Excel Files', 
      icon: FileSpreadsheet, 
      badge: filesWaitingCount > 0 ? filesWaitingCount : undefined,
      badgeColor: 'bg-amber-100 text-amber-700'
    },
    { id: 'analyzer' as NavigationTab, label: 'Workbook Analyzer', icon: Microscope },
    { id: 'mappings' as NavigationTab, label: 'Table & Column Mappings', icon: GitFork },
    { id: 'import' as NavigationTab, label: 'Import & Dry Run', icon: PlaySquare },
    { 
      id: 'twoway' as NavigationTab, 
      label: 'Two-Way Sync', 
      icon: ArrowLeftRight,
      badge: conflictCount > 0 ? conflictCount : undefined,
      badgeColor: 'bg-amber-500 text-white animate-pulse'
    },
    { id: 'history' as NavigationTab, label: 'Audit History', icon: History },
    { 
      id: 'errors' as NavigationTab, 
      label: 'Validation Errors', 
      icon: AlertTriangle,
      badge: errorCount > 0 ? errorCount : undefined,
      badgeColor: 'bg-rose-100 text-rose-700'
    },
    { id: 'logs' as NavigationTab, label: 'Execution Logs', icon: Terminal },
    { id: 'settings' as NavigationTab, label: 'Settings', icon: Settings }
  ];

  const handleItemClick = (id: NavigationTab) => {
    onSelectTab(id);
    if (onCloseMobile) {
      onCloseMobile();
    }
  };

  const sidebarContent = (
    <div className="flex flex-col h-full bg-slate-900 text-slate-300">
      {/* Mobile Header with Close Button */}
      <div className="p-4 border-b border-slate-800 flex items-center justify-between">
        <div>
          <div className="text-[10px] uppercase font-bold tracking-wider text-slate-400">Production Pipeline</div>
          <div className="text-xs text-slate-300 font-mono mt-0.5">
            <span className="text-emerald-400 font-semibold">WebDAV</span> → <span className="text-indigo-400 font-semibold">{schedulerStatus?.engineMode === 'EXTERNAL_WORKER_DAEMON' ? 'Worker' : 'Sync Engine'}</span> → <span className="text-teal-400 font-semibold">Supabase</span>
          </div>
        </div>
        {onCloseMobile && (
          <button
            onClick={onCloseMobile}
            className="md:hidden p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
            title="Close menu"
          >
            <X className="w-5 h-5" />
          </button>
        )}
      </div>

      {/* Navigation Items */}
      <nav className="flex-1 px-3 py-3 space-y-1 overflow-y-auto">
        {navItems.map((item) => {
          const Icon = item.icon;
          const isActive = currentTab === item.id;
          return (
            <button
              key={item.id}
              id={`nav-item-${item.id}`}
              onClick={() => handleItemClick(item.id)}
              className={`w-full flex items-center justify-between px-3 py-2 rounded-lg text-xs sm:text-sm font-medium transition-all ${
                isActive
                  ? 'bg-emerald-600 text-white shadow-xs'
                  : 'text-slate-300 hover:bg-slate-800 hover:text-white'
              }`}
            >
              <div className="flex items-center space-x-3 truncate">
                <Icon className={`w-4 h-4 shrink-0 ${isActive ? 'text-white' : 'text-slate-400'}`} />
                <span className="truncate">{item.label}</span>
              </div>
              {item.badge !== undefined && (
                <span className={`text-[11px] px-2 py-0.5 rounded-full font-bold shrink-0 ml-2 ${
                  isActive ? 'bg-white/20 text-white' : item.badgeColor
                }`}>
                  {item.badge}
                </span>
              )}
            </button>
          );
        })}
      </nav>

      {/* Secrets Vault Button */}
      {onOpenSecretsVault && (
        <div className="px-3 pb-3">
          <button
            onClick={() => {
              onOpenSecretsVault();
              if (onCloseMobile) onCloseMobile();
            }}
            className="w-full flex items-center justify-between px-3 py-2 rounded-lg bg-slate-800/90 hover:bg-slate-800 border border-slate-700/80 text-xs font-semibold text-slate-200 transition-colors group"
          >
            <div className="flex items-center space-x-2.5">
              <Key className="w-3.5 h-3.5 text-amber-400 group-hover:rotate-12 transition-transform" />
              <span>Secrets & Credentials</span>
            </div>
            <span className={`w-2 h-2 rounded-full ${allSecretsConnected ? 'bg-emerald-400' : 'bg-amber-400 animate-ping'}`} />
          </button>
        </div>
      )}

      {/* Footer Info */}
      <div className="p-3 border-t border-slate-800 text-[11px] text-slate-400 bg-slate-950/40">
        <div className="flex items-center justify-between">
          <span className="font-medium text-slate-300">
            {schedulerStatus?.engineMode === 'EXTERNAL_WORKER_DAEMON' ? 'Worker Daemon' : 'In-App Sync Engine'}
          </span>
          <span className={`font-mono font-bold ${schedulerStatus?.enabled ? 'text-emerald-400' : 'text-slate-400'}`}>
            {schedulerStatus?.enabled
              ? schedulerStatus.secondsUntilNextRun !== null && schedulerStatus.secondsUntilNextRun < 60
                ? `< 1m run`
                : schedulerStatus.secondsUntilNextRun !== null
                  ? `${Math.ceil(schedulerStatus.secondsUntilNextRun / 60)}m run`
                  : schedulerStatus.intervalLabel
              : 'Manual'}
          </span>
        </div>
        <div className="mt-1 text-[10px] text-slate-400 flex items-center justify-between">
          <span className="truncate max-w-[130px]">
            {schedulerStatus?.engineMode === 'EXTERNAL_WORKER_DAEMON'
              ? schedulerStatus.workerEndpoint?.replace(/^https?:\/\//, '') || 'Remote Worker'
              : 'Node.js Production Pipeline'}
          </span>
          <span className={`font-mono font-semibold ${
            schedulerStatus?.state === 'SYNCING' ? 'text-amber-400 animate-pulse' :
            schedulerStatus?.state === 'SCHEDULED' ? 'text-emerald-400' :
            'text-slate-400'
          }`}>
            {schedulerStatus?.state === 'SYNCING' ? 'Syncing...' : schedulerStatus?.state === 'SCHEDULED' ? 'Active' : 'Standby'}
          </span>
        </div>
      </div>
    </div>
  );

  return (
    <>
      {/* Desktop Sticky Sidebar (Visible on md and up) */}
      <aside className="hidden md:flex flex-col w-64 shrink-0 bg-slate-900 border-r border-slate-800 min-h-[calc(100vh-4rem)] sticky top-16 z-20">
        {sidebarContent}
      </aside>

      {/* Mobile Drawer (Visible on <md when open) */}
      {isOpenMobile && (
        <div className="fixed inset-0 z-50 md:hidden flex animate-in fade-in duration-200">
          {/* Backdrop overlay */}
          <div 
            className="fixed inset-0 bg-slate-900/70 backdrop-blur-xs transition-opacity"
            onClick={onCloseMobile}
          />
          {/* Slide-out panel */}
          <div className="relative flex flex-col w-72 max-w-[85vw] h-full shadow-2xl z-10 animate-in slide-in-from-left duration-200">
            {sidebarContent}
          </div>
        </div>
      )}
    </>
  );
};
