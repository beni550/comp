import { useState, useEffect } from 'react'
import { useNavigate } from 'react-router-dom'
import api from '@/api/client'
import { formatRelativeTime } from '@/lib/utils'
import { Heart, MessageCircle, UserPlus, Bell, Check, Loader2 } from 'lucide-react'

interface Notification {
  id: string
  type: string
  title: string
  body: string
  data: Record<string, string>
  isRead: boolean
  createdAt: string
}

const ICONS: Record<string, typeof Heart> = {
  like: Heart,
  comment: MessageCircle,
  follow: UserPlus,
  default: Bell,
}

export default function NotificationsPage() {
  const [notifications, setNotifications] = useState<Notification[]>([])
  const [loading, setLoading] = useState(true)
  const navigate = useNavigate()

  useEffect(() => {
    const fetchNotifications = async () => {
      setLoading(true)
      try {
        const { data } = await api.get('/notifications')
        setNotifications(data.data || [])
      } catch {
        setNotifications([])
      }
      setLoading(false)
    }
    fetchNotifications()
  }, [])

  const markAsRead = async (notifId: string) => {
    try {
      await api.post(`/notifications/${notifId}/read`)
      setNotifications(prev => prev.map(n => n.id === notifId ? { ...n, isRead: true } : n))
    } catch { /* ignore */ }
  }

  const markAllRead = async () => {
    try {
      await api.post('/notifications/read-all')
      setNotifications(prev => prev.map(n => ({ ...n, isRead: true })))
    } catch { /* ignore */ }
  }

  const handleClick = (notif: Notification) => {
    if (!notif.isRead) markAsRead(notif.id)
    if (notif.data?.contentId) {
      navigate(`/post/${notif.data.contentId}`)
    } else if (notif.data?.username) {
      navigate(`/profile/${notif.data.username}`)
    }
  }

  return (
    <div>
      <div className="px-4 py-3 border-b border-border flex items-center justify-between">
        <h2 className="text-lg font-bold text-text">Notifications</h2>
        {notifications.some(n => !n.isRead) && (
          <button onClick={markAllRead} className="text-primary text-sm hover:text-primary-light flex items-center gap-1">
            <Check size={16} /> Mark all read
          </button>
        )}
      </div>

      {loading ? (
        <div className="flex justify-center py-12"><Loader2 size={24} className="animate-spin text-primary" /></div>
      ) : notifications.length === 0 ? (
        <div className="text-center py-12 text-text-muted">
          <Bell size={48} className="mx-auto mb-4 opacity-50" />
          <p className="text-lg mb-1">No notifications</p>
          <p className="text-sm">You're all caught up!</p>
        </div>
      ) : (
        <div className="divide-y divide-border">
          {notifications.map(notif => {
            const Icon = ICONS[notif.type] || ICONS.default
            return (
              <button key={notif.id} onClick={() => handleClick(notif)}
                className={`w-full text-left px-4 py-3 flex items-start gap-3 transition-colors hover:bg-surface/50 ${
                  !notif.isRead ? 'bg-primary/5' : ''
                }`}>
                <div className={`w-10 h-10 rounded-full flex items-center justify-center shrink-0 ${
                  notif.type === 'like' ? 'bg-secondary/20 text-secondary' :
                  notif.type === 'follow' ? 'bg-primary/20 text-primary' :
                  notif.type === 'comment' ? 'bg-success/20 text-success' :
                  'bg-surface-light text-text-muted'
                }`}>
                  <Icon size={18} />
                </div>
                <div className="flex-1 min-w-0">
                  <p className="text-sm text-text">{notif.title}</p>
                  <p className="text-xs text-text-muted mt-0.5">{notif.body}</p>
                  <p className="text-xs text-text-muted mt-1">{formatRelativeTime(notif.createdAt)}</p>
                </div>
                {!notif.isRead && (
                  <div className="w-2 h-2 rounded-full bg-primary mt-2 shrink-0" />
                )}
              </button>
            )
          })}
        </div>
      )}
    </div>
  )
}
