import {
  FolderKanban,
  LogOut,
  PanelLeft,
  PanelLeftClose,
  Settings,
} from 'lucide-react';

type Page = 'projects' | 'settings';

interface SidebarProps {
  currentPage: Page;
  onPageChange: (page: Page) => void;
  user: { name: string; email: string } | null;
  onLogout: () => void;
  collapsed: boolean;
  onCollapsedChange: (collapsed: boolean) => void;
}

const navItems = [
  { id: 'projects' as const, label: 'Progetti', Icon: FolderKanban },
  { id: 'settings' as const, label: 'Impostazioni', Icon: Settings },
];

function initials(name: string): string {
  return name
    .split(' ')
    .filter(Boolean)
    .map((part) => part[0])
    .join('')
    .slice(0, 2)
    .toUpperCase();
}

export default function Sidebar({
  currentPage,
  onPageChange,
  user,
  onLogout,
  collapsed,
  onCollapsedChange,
}: SidebarProps) {
  const ToggleIcon = collapsed ? PanelLeft : PanelLeftClose;

  return (
    <aside className={`sidebar${collapsed ? ' collapsed' : ''}`}>
      <div className="sidebar-header">
        <button
          className="icon-button"
          onClick={() => onCollapsedChange(!collapsed)}
          title={collapsed ? 'Espandi sidebar' : 'Comprimi sidebar'}
        >
          <ToggleIcon size={19} />
        </button>
        {!collapsed && <span className="sidebar-title">Project Steps</span>}
      </div>

      <nav className="sidebar-nav">
        {navItems.map(({ id, label, Icon }) => (
          <button
            key={id}
            className={currentPage === id ? 'active' : ''}
            onClick={() => onPageChange(id)}
            title={collapsed ? label : undefined}
          >
            <Icon size={18} />
            {!collapsed && <span>{label}</span>}
          </button>
        ))}
      </nav>

      <div className="sidebar-user">
        <div className="sidebar-avatar">{user ? initials(user.name || user.email) : '?'}</div>
        {!collapsed && (
          <div className="sidebar-user-info">
            <div className="sidebar-user-name">{user?.name || 'Non connesso'}</div>
            <div className="sidebar-user-email">{user?.email || 'Configura e accedi'}</div>
          </div>
        )}
        {user && (
          <button className="sidebar-logout-btn" onClick={onLogout} title="Logout">
            <LogOut size={16} />
          </button>
        )}
      </div>
    </aside>
  );
}
