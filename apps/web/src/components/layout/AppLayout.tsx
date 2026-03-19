import { Outlet, NavLink, useNavigate } from 'react-router-dom'
import { Home, PlusSquare, MessageCircle, Bell, User, Shield, LogOut } from 'lucide-react'
import { useAuthStore } from '@/stores/auth'

export default function AppLayout() {
  const { user, logout } = useAuthStore()
  const navigate = useNavigate()
  const isAdmin = user?.role === 'admin' || user?.role === 'moderator'

  const handleLogout = async () => {
    await logout()
    navigate('/login')
  }

  return (
    <div className="min-h-screen flex flex-col max-w-lg mx-auto bg-bg">
      {/* Header */}
      <header className="sticky top-0 z-50 bg-surface/80 backdrop-blur-sm border-b border-border px-4 py-3 flex items-center justify-between">
        <h1 className="text-xl font-bold bg-gradient-to-r from-primary to-secondary bg-clip-text text-transparent">
          VYBE
        </h1>
        <div className="flex items-center gap-3">
          {isAdmin && (
            <NavLink to="/admin" className="text-text-muted hover:text-primary transition-colors">
              <Shield size={20} />
            </NavLink>
          )}
          <button onClick={handleLogout} className="text-text-muted hover:text-error transition-colors">
            <LogOut size={20} />
          </button>
        </div>
      </header>

      {/* Main content */}
      <main className="flex-1 overflow-y-auto">
        <Outlet />
      </main>

      {/* Bottom nav */}
      <nav className="sticky bottom-0 bg-surface/90 backdrop-blur-sm border-t border-border">
        <div className="flex items-center justify-around py-2">
          <NavLink to="/feed" className={({ isActive }) =>
            `flex flex-col items-center gap-1 px-3 py-1 transition-colors ${isActive ? 'text-primary' : 'text-text-muted hover:text-text'}`
          }>
            <Home size={22} />
            <span className="text-xs">Feed</span>
          </NavLink>
          <NavLink to="/create" className={({ isActive }) =>
            `flex flex-col items-center gap-1 px-3 py-1 transition-colors ${isActive ? 'text-primary' : 'text-text-muted hover:text-text'}`
          }>
            <PlusSquare size={22} />
            <span className="text-xs">Create</span>
          </NavLink>
          <NavLink to="/messages" className={({ isActive }) =>
            `flex flex-col items-center gap-1 px-3 py-1 transition-colors ${isActive ? 'text-primary' : 'text-text-muted hover:text-text'}`
          }>
            <MessageCircle size={22} />
            <span className="text-xs">Messages</span>
          </NavLink>
          <NavLink to="/notifications" className={({ isActive }) =>
            `flex flex-col items-center gap-1 px-3 py-1 transition-colors ${isActive ? 'text-primary' : 'text-text-muted hover:text-text'}`
          }>
            <Bell size={22} />
            <span className="text-xs">Alerts</span>
          </NavLink>
          <NavLink to="/profile" className={({ isActive }) =>
            `flex flex-col items-center gap-1 px-3 py-1 transition-colors ${isActive ? 'text-primary' : 'text-text-muted hover:text-text'}`
          }>
            <User size={22} />
            <span className="text-xs">Profile</span>
          </NavLink>
        </div>
      </nav>
    </div>
  )
}
