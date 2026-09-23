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
  Settings 
} from 'lucide-react';
import { NavigationTab } from '../types';

interface SidebarProps {
  currentTab: NavigationTab;
  onSelectTab: (tab: NavigationTab) => void;
  errorCount: number;
  filesWaitingCount: number;
  conflictCount?: number;
}

export const Sidebar: React.FC<SidebarProps> = ({
  currentTab,
  onSelectTab,
  errorCount,
  filesWaitingCount,
  conflictCount = 0
}) => {
  const navItems = [
    { id: 'dashboard' as NavigationTab, label: 'Dashboard', icon: LayoutDashboard },
    { id: 'nextcloud' as NavigationTab, label: 'Nextcloud', icon: Cloud },
    { id: 'supabase' as NavigationTab, label: 'Supabase', icon: Database },
    { 
      id: 'files' as NavigationTab, 
      label: 'Excel Files', 
      icon: FileSpreadsheet, 
      badge: filesWaitingCount > 0 ? filesWaitingCount : undefined,
      badgeColor: 'bg-amber-100 text-amber-700'
    },
    { id: 'analyzer' as NavigationTab, label: 'Workbook Analyzer', icon: Microscope },
    { id: 'mappings' as NavigationTab, label: 'Mappings', icon: GitFork },
    { id: 'import' as NavigationTab, label: 'Import & Dry Run', icon: PlaySquare },
    { 
      id: 'twoway' as NavigationTab, 
      label: 'Two-Way Sync', 
      icon: ArrowLeftRight,
      badge: conflictCount > 0 ? conflictCount : undefined,
      badgeColor: 'bg-amber-500 text-white animate-pulse'
    },
    { id: 'history' as NavigationTab, label: 'Import History', icon: History },
    { 
      id: 'errors' as NavigationTab, 
      label: 'Errors', 
      icon: AlertTriangle,
      badge: errorCount > 0 ? errorCount : undefined,
      badgeColor: 'bg-rose-100 text-rose-700'
    },
    { id: 'logs' as NavigationTab, label: 'Logs', icon: Terminal },
    { id: 'settings' as NavigationTab, label: 'Settings', icon: Settings }
  ];

  return (
    <aside className="w-64 bg-slate-900 text-slate-300 flex flex-col shrink-0 min-h-[calc(100vh-4rem)] border-r border-slate-800">
      <div className="p-4 border-b border-slate-800">
        <div className="text-xs uppercase font-bold tracking-wider text-slate-400 mb-1">Architecture</div>
        <div className="text-xs text-slate-300 font-mono bg-slate-800/80 p-2 rounded-md leading-relaxed">
          <span className="text-emerald-400">Nextcloud</span> → <span className="text-blue-400">Coolify</span> → <span className="text-teal-400">Supabase</span>
        </div>
      </div>

      <nav className="flex-1 px-3 py-4 space-y-1">
        {navItems.map((item) => {
          const Icon = item.icon;
          const isActive = currentTab === item.id;
          return (
            <button
              key={item.id}
              id={`nav-item-${item.id}`}
              onClick={() => onSelectTab(item.id)}
              className={`w-full flex items-center justify-between px-3 py-2.5 rounded-lg text-sm font-medium transition-all ${
                isActive
                  ? 'bg-emerald-600 text-white shadow-sm'
                  : 'text-slate-300 hover:bg-slate-800 hover:text-white'
              }`}
            >
              <div className="flex items-center space-x-3">
                <Icon className={`w-4 h-4 ${isActive ? 'text-white' : 'text-slate-400'}`} />
                <span>{item.label}</span>
              </div>
              {item.badge !== undefined && (
                <span className={`text-xs px-2 py-0.5 rounded-full font-semibold ${
                  isActive ? 'bg-white/20 text-white' : item.badgeColor
                }`}>
                  {item.badge}
                </span>
              )}
            </button>
          );
        })}
      </nav>

      <div className="p-4 border-t border-slate-800 text-xs text-slate-400">
        <div className="flex items-center justify-between">
          <span>Worker Interval</span>
          <span className="text-emerald-400 font-mono">15 min</span>
        </div>
        <div className="mt-1 text-[11px] text-slate-400">
          TrueNAS SCALE • WebDAV HTTPS
        </div>
      </div>
    </aside>
  );
};
