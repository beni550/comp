import { useState, useEffect } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import api from '@/api/client'
import { useAuthStore } from '@/stores/auth'
import { formatRelativeTime } from '@/lib/utils'
import { Settings, UserPlus, UserCheck, MapPin, Calendar, Link as LinkIcon, ArrowLeft, Loader2, Heart, MessageCircle } from 'lucide-react'

interface Profile {
  id: string
  username: string
  displayName: string
  bio: string | null
  avatarUrl: string | null
  location: string | null
  website: string | null
  followersCount: number
  followingCount: number
  postsCount: number
  isFollowing: boolean
  joinedAt: string
}

interface UserPost {
  id: string
  caption: string
  likesCount: number
  commentsCount: number
  createdAt: string
  media: { url: string; type: string }[]
}

export default function ProfilePage() {
  const { username } = useParams<{ username: string }>()
  const { user } = useAuthStore()
  const navigate = useNavigate()
  const [profile, setProfile] = useState<Profile | null>(null)
  const [posts, setPosts] = useState<UserPost[]>([])
  const [loading, setLoading] = useState(true)
  const [editing, setEditing] = useState(false)
  const [editForm, setEditForm] = useState({ bio: '', displayName: '' })

  const isOwnProfile = !username || username === user?.username

  useEffect(() => {
    const fetchProfile = async () => {
      setLoading(true)
      try {
        const endpoint = isOwnProfile ? '/users/me' : `/users/${username}`
        const { data } = await api.get(endpoint)
        setProfile(data.data)
        setEditForm({ bio: data.data.bio || '', displayName: data.data.displayName || '' })

        // Fetch user's posts
        const postsEndpoint = isOwnProfile ? '/users/me/posts' : `/users/${username}/posts`
        try {
          const { data: postsData } = await api.get(postsEndpoint)
          setPosts(postsData.data || [])
        } catch {
          setPosts([])
        }
      } catch {
        setProfile(null)
      }
      setLoading(false)
    }
    fetchProfile()
  }, [username, isOwnProfile])

  const handleFollow = async () => {
    if (!profile) return
    try {
      if (profile.isFollowing) {
        await api.delete(`/social/follow/${profile.id}`)
      } else {
        await api.post(`/social/follow/${profile.id}`)
      }
      setProfile(prev => prev ? {
        ...prev,
        isFollowing: !prev.isFollowing,
        followersCount: prev.isFollowing ? prev.followersCount - 1 : prev.followersCount + 1,
      } : null)
    } catch { /* ignore */ }
  }

  const handleSaveProfile = async () => {
    try {
      await api.patch('/users/me', editForm)
      setProfile(prev => prev ? { ...prev, ...editForm } : null)
      setEditing(false)
    } catch { /* ignore */ }
  }

  if (loading) {
    return <div className="flex justify-center py-12"><Loader2 size={24} className="animate-spin text-primary" /></div>
  }

  if (!profile) {
    return (
      <div className="text-center py-12 text-text-muted">
        <p>User not found</p>
        <button onClick={() => navigate('/feed')} className="text-primary mt-2">Go back</button>
      </div>
    )
  }

  return (
    <div>
      {/* Header */}
      {!isOwnProfile && (
        <div className="flex items-center gap-3 px-4 py-3 border-b border-border">
          <button onClick={() => navigate(-1)} className="text-text-muted hover:text-text"><ArrowLeft size={20} /></button>
          <span className="font-medium">{profile.displayName}</span>
        </div>
      )}

      {/* Profile header */}
      <div className="px-4 py-6">
        <div className="flex items-start gap-4">
          <div className="w-20 h-20 rounded-full bg-primary/20 flex items-center justify-center text-primary text-2xl font-bold shrink-0">
            {profile.displayName[0]?.toUpperCase()}
          </div>
          <div className="flex-1 min-w-0">
            <h2 className="text-xl font-bold text-text">{profile.displayName}</h2>
            <p className="text-sm text-text-muted">@{profile.username}</p>
          </div>
          {isOwnProfile ? (
            <button onClick={() => setEditing(!editing)}
              className="border border-border text-text-muted hover:text-text rounded-full p-2 transition-colors">
              <Settings size={18} />
            </button>
          ) : (
            <button onClick={handleFollow}
              className={`flex items-center gap-1.5 rounded-full px-4 py-2 text-sm font-medium transition-colors ${
                profile.isFollowing
                  ? 'bg-surface border border-border text-text hover:border-error hover:text-error'
                  : 'bg-primary hover:bg-primary-dark text-white'
              }`}>
              {profile.isFollowing ? <><UserCheck size={16} /> Following</> : <><UserPlus size={16} /> Follow</>}
            </button>
          )}
        </div>

        {/* Edit form */}
        {editing && (
          <div className="mt-4 space-y-3 p-4 bg-surface rounded-lg border border-border">
            <div>
              <label className="block text-xs text-text-muted mb-1">Display Name</label>
              <input value={editForm.displayName} onChange={e => setEditForm(prev => ({ ...prev, displayName: e.target.value }))}
                className="w-full bg-bg border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary" />
            </div>
            <div>
              <label className="block text-xs text-text-muted mb-1">Bio</label>
              <textarea value={editForm.bio} onChange={e => setEditForm(prev => ({ ...prev, bio: e.target.value }))}
                className="w-full bg-bg border border-border rounded-lg px-3 py-2 text-sm text-text focus:outline-none focus:border-primary resize-none h-20" maxLength={300} />
            </div>
            <div className="flex gap-2">
              <button onClick={handleSaveProfile} className="bg-primary hover:bg-primary-dark text-white text-sm rounded-lg px-4 py-2">Save</button>
              <button onClick={() => setEditing(false)} className="text-text-muted text-sm hover:text-text">Cancel</button>
            </div>
          </div>
        )}

        {/* Bio */}
        {profile.bio && <p className="text-sm text-text mt-3">{profile.bio}</p>}

        {/* Meta info */}
        <div className="flex flex-wrap gap-4 mt-3 text-xs text-text-muted">
          {profile.location && <span className="flex items-center gap-1"><MapPin size={14} />{profile.location}</span>}
          {profile.website && <a href={profile.website} target="_blank" rel="noopener noreferrer" className="flex items-center gap-1 text-primary hover:text-primary-light"><LinkIcon size={14} />{profile.website}</a>}
          <span className="flex items-center gap-1"><Calendar size={14} />Joined {new Date(profile.joinedAt).toLocaleDateString()}</span>
        </div>

        {/* Stats */}
        <div className="flex gap-6 mt-4">
          <div><span className="font-bold text-text">{profile.postsCount}</span> <span className="text-text-muted text-sm">posts</span></div>
          <div><span className="font-bold text-text">{profile.followersCount}</span> <span className="text-text-muted text-sm">followers</span></div>
          <div><span className="font-bold text-text">{profile.followingCount}</span> <span className="text-text-muted text-sm">following</span></div>
        </div>
      </div>

      {/* Posts grid */}
      <div className="border-t border-border">
        <h3 className="px-4 py-3 text-sm font-medium text-text-muted">Posts</h3>
        {posts.length === 0 ? (
          <p className="text-center py-8 text-text-muted text-sm">No posts yet</p>
        ) : (
          <div className="divide-y divide-border">
            {posts.map(post => (
              <button key={post.id} onClick={() => navigate(`/post/${post.id}`)} className="w-full text-left px-4 py-3 hover:bg-surface/50 transition-colors">
                <p className="text-sm text-text line-clamp-2">{post.caption}</p>
                <div className="flex items-center gap-4 mt-2 text-xs text-text-muted">
                  <span className="flex items-center gap-1"><Heart size={12} />{post.likesCount}</span>
                  <span className="flex items-center gap-1"><MessageCircle size={12} />{post.commentsCount}</span>
                  <span>{formatRelativeTime(post.createdAt)}</span>
                </div>
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  )
}
