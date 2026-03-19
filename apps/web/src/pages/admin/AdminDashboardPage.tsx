import { useState, useEffect } from 'react'
import { useNavigate } from 'react-router-dom'
import api from '@/api/client'
import { useAuthStore } from '@/stores/auth'
import { Users, FileText, Flag, Shield, ArrowLeft, Loader2, Ban, CheckCircle } from 'lucide-react'

interface DashboardStats {
  totalUsers: number
  totalPosts: number
  totalReports: number
  activeUsers: number
}

interface Report {
  id: string
  type: string
  reason: string
  status: string
  reporter: { username: string; displayName: string }
  createdAt: string
}

export default function AdminDashboardPage() {
  const { user } = useAuthStore()
  const navigate = useNavigate()
  const [stats, setStats] = useState<DashboardStats | null>(null)
  const [reports, setReports] = useState<Report[]>([])
  const [loading, setLoading] = useState(true)
  const [tab, setTab] = useState<'overview' | 'reports'>('overview')

  const isAdmin = user?.role === 'admin' || user?.role === 'moderator'

  useEffect(() => {
    if (!isAdmin) {
      navigate('/feed')
      return
    }
    const fetchData = async () => {
      setLoading(true)
      try {
        const [statsRes, reportsRes] = await Promise.all([
          api.get('/admin/dashboard'),
          api.get('/reports/mine'),
        ])
        setStats(statsRes.data.data)
        setReports(reportsRes.data.data || [])
      } catch {
        // Fallback if endpoints don't exist yet
        setStats({ totalUsers: 0, totalPosts: 0, totalReports: 0, activeUsers: 0 })
        setReports([])
      }
      setLoading(false)
    }
    fetchData()
  }, [isAdmin, navigate])

  const handleReportAction = async (reportId: string, action: 'resolve' | 'dismiss') => {
    try {
      await api.post(`/admin/moderation/cases/${reportId}/resolve`, { action })
      setReports(prev => prev.map(r => r.id === reportId ? { ...r, status: action === 'resolve' ? 'resolved' : 'dismissed' } : r))
    } catch { /* ignore */ }
  }

  if (loading) {
    return <div className="flex justify-center py-12"><Loader2 size={24} className="animate-spin text-primary" /></div>
  }

  return (
    <div>
      {/* Header */}
      <div className="flex items-center gap-3 px-4 py-3 border-b border-border">
        <button onClick={() => navigate(-1)} className="text-text-muted hover:text-text">
          <ArrowLeft size={20} />
        </button>
        <Shield size={20} className="text-primary" />
        <span className="font-bold text-text">Admin Dashboard</span>
      </div>

      {/* Tabs */}
      <div className="flex border-b border-border">
        <button onClick={() => setTab('overview')}
          className={`flex-1 py-3 text-sm font-medium transition-colors ${tab === 'overview' ? 'text-primary border-b-2 border-primary' : 'text-text-muted'}`}>
          Overview
        </button>
        <button onClick={() => setTab('reports')}
          className={`flex-1 py-3 text-sm font-medium transition-colors ${tab === 'reports' ? 'text-primary border-b-2 border-primary' : 'text-text-muted'}`}>
          Reports ({reports.filter(r => r.status === 'pending').length})
        </button>
      </div>

      {tab === 'overview' && stats && (
        <div className="p-4 grid grid-cols-2 gap-3">
          <div className="bg-surface rounded-xl p-4 border border-border">
            <div className="flex items-center gap-2 mb-2">
              <Users size={18} className="text-primary" />
              <span className="text-xs text-text-muted">Total Users</span>
            </div>
            <p className="text-2xl font-bold text-text">{stats.totalUsers}</p>
          </div>
          <div className="bg-surface rounded-xl p-4 border border-border">
            <div className="flex items-center gap-2 mb-2">
              <FileText size={18} className="text-success" />
              <span className="text-xs text-text-muted">Total Posts</span>
            </div>
            <p className="text-2xl font-bold text-text">{stats.totalPosts}</p>
          </div>
          <div className="bg-surface rounded-xl p-4 border border-border">
            <div className="flex items-center gap-2 mb-2">
              <Flag size={18} className="text-warning" />
              <span className="text-xs text-text-muted">Reports</span>
            </div>
            <p className="text-2xl font-bold text-text">{stats.totalReports}</p>
          </div>
          <div className="bg-surface rounded-xl p-4 border border-border">
            <div className="flex items-center gap-2 mb-2">
              <Users size={18} className="text-secondary" />
              <span className="text-xs text-text-muted">Active Users</span>
            </div>
            <p className="text-2xl font-bold text-text">{stats.activeUsers}</p>
          </div>
        </div>
      )}

      {tab === 'reports' && (
        <div className="divide-y divide-border">
          {reports.length === 0 ? (
            <div className="text-center py-12 text-text-muted">
              <Flag size={48} className="mx-auto mb-4 opacity-50" />
              <p>No reports</p>
            </div>
          ) : (
            reports.map(report => (
              <div key={report.id} className="px-4 py-3">
                <div className="flex items-start justify-between">
                  <div>
                    <p className="text-sm font-medium text-text">{report.type} report</p>
                    <p className="text-xs text-text-muted mt-0.5">By @{report.reporter.username}</p>
                    <p className="text-sm text-text mt-1">{report.reason}</p>
                  </div>
                  <span className={`text-xs px-2 py-1 rounded-full ${
                    report.status === 'pending' ? 'bg-warning/20 text-warning' :
                    report.status === 'resolved' ? 'bg-success/20 text-success' :
                    'bg-surface-light text-text-muted'
                  }`}>{report.status}</span>
                </div>
                {report.status === 'pending' && (
                  <div className="flex gap-2 mt-2">
                    <button onClick={() => handleReportAction(report.id, 'resolve')}
                      className="flex items-center gap-1 text-xs bg-success/20 text-success px-3 py-1.5 rounded-lg hover:bg-success/30">
                      <CheckCircle size={14} /> Resolve
                    </button>
                    <button onClick={() => handleReportAction(report.id, 'dismiss')}
                      className="flex items-center gap-1 text-xs bg-surface-light text-text-muted px-3 py-1.5 rounded-lg hover:bg-border">
                      <Ban size={14} /> Dismiss
                    </button>
                  </div>
                )}
              </div>
            ))
          )}
        </div>
      )}
    </div>
  )
}
