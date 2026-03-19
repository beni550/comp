import { useState, useEffect } from 'react'
import { Link } from 'react-router-dom'
import api from '@/api/client'
import { formatRelativeTime } from '@/lib/utils'
import { Heart, MessageCircle, Share2, Bookmark, MoreHorizontal, Loader2 } from 'lucide-react'

interface ContentItem {
  id: string
  author: { id: string; username: string; displayName: string; avatarUrl: string | null }
  type: string
  caption: string
  media: { id: string; url: string; type: string }[]
  likesCount: number
  commentsCount: number
  sharesCount: number
  isLiked: boolean
  isSaved: boolean
  createdAt: string
  publishedAt: string | null
}

export default function FeedPage() {
  const [tab, setTab] = useState<'my-vybe' | 'for-you'>('my-vybe')
  const [items, setItems] = useState<ContentItem[]>([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    let cancelled = false
    const fetchFeed = async () => {
      setLoading(true)
      try {
        const { data } = await api.get(`/feed/${tab}`)
        if (!cancelled) setItems(data.data || [])
      } catch {
        if (!cancelled) setItems([])
      }
      if (!cancelled) setLoading(false)
    }
    fetchFeed()
    return () => { cancelled = true }
  }, [tab])

  const handleReact = async (contentId: string, index: number) => {
    try {
      const { data } = await api.post(`/content/${contentId}/react`, { type: 'like' })
      setItems(prev => prev.map((item, i) => i === index ? {
        ...item,
        isLiked: data.data.action === 'added',
        likesCount: data.data.action === 'added' ? item.likesCount + 1 : item.likesCount - 1,
      } : item))
    } catch { /* ignore */ }
  }

  const handleSave = async (contentId: string, index: number) => {
    try {
      await api.post(`/content/${contentId}/save`)
      setItems(prev => prev.map((item, i) => i === index ? {
        ...item,
        isSaved: !item.isSaved,
        savesCount: undefined as never,
      } : item))
    } catch { /* ignore */ }
  }

  return (
    <div className="pb-4">
      {/* Feed tabs */}
      <div className="flex border-b border-border sticky top-14 bg-bg z-10">
        <button onClick={() => setTab('my-vybe')}
          className={`flex-1 py-3 text-sm font-medium transition-colors ${tab === 'my-vybe' ? 'text-primary border-b-2 border-primary' : 'text-text-muted'}`}>
          My VYBE
        </button>
        <button onClick={() => setTab('for-you')}
          className={`flex-1 py-3 text-sm font-medium transition-colors ${tab === 'for-you' ? 'text-primary border-b-2 border-primary' : 'text-text-muted'}`}>
          For You
        </button>
      </div>

      {loading ? (
        <div className="flex justify-center py-12">
          <Loader2 size={24} className="animate-spin text-primary" />
        </div>
      ) : items.length === 0 ? (
        <div className="text-center py-12 text-text-muted">
          <p className="text-lg mb-2">No posts yet</p>
          <p className="text-sm">Follow people or create your own posts!</p>
        </div>
      ) : (
        <div className="divide-y divide-border">
          {items.map((item, index) => (
            <article key={item.id} className="px-4 py-4">
              {/* Author row */}
              <div className="flex items-center gap-3 mb-3">
                <Link to={`/profile/${item.author.username}`}>
                  <div className="w-10 h-10 rounded-full bg-primary/20 flex items-center justify-center text-primary font-medium">
                    {item.author.displayName[0]?.toUpperCase()}
                  </div>
                </Link>
                <div className="flex-1 min-w-0">
                  <Link to={`/profile/${item.author.username}`} className="font-medium text-sm text-text hover:text-primary">
                    {item.author.displayName}
                  </Link>
                  <p className="text-xs text-text-muted">@{item.author.username} · {formatRelativeTime(item.publishedAt || item.createdAt)}</p>
                </div>
                <button className="text-text-muted hover:text-text"><MoreHorizontal size={18} /></button>
              </div>

              {/* Caption */}
              <Link to={`/post/${item.id}`}>
                <p className="text-sm text-text mb-3 whitespace-pre-wrap">{item.caption}</p>
              </Link>

              {/* Media */}
              {item.media && item.media.length > 0 && (
                <div className="rounded-xl overflow-hidden mb-3 bg-surface">
                  <img src={item.media[0].url} alt="" className="w-full object-cover max-h-80" />
                </div>
              )}

              {/* Actions */}
              <div className="flex items-center justify-between text-text-muted">
                <button onClick={() => handleReact(item.id, index)}
                  className={`flex items-center gap-1.5 text-sm transition-colors ${item.isLiked ? 'text-secondary' : 'hover:text-secondary'}`}>
                  <Heart size={18} fill={item.isLiked ? 'currentColor' : 'none'} />
                  <span>{item.likesCount || ''}</span>
                </button>
                <Link to={`/post/${item.id}`} className="flex items-center gap-1.5 text-sm hover:text-primary transition-colors">
                  <MessageCircle size={18} />
                  <span>{item.commentsCount || ''}</span>
                </Link>
                <button className="flex items-center gap-1.5 text-sm hover:text-primary transition-colors">
                  <Share2 size={18} />
                  <span>{item.sharesCount || ''}</span>
                </button>
                <button onClick={() => handleSave(item.id, index)}
                  className={`transition-colors ${item.isSaved ? 'text-warning' : 'hover:text-warning'}`}>
                  <Bookmark size={18} fill={item.isSaved ? 'currentColor' : 'none'} />
                </button>
              </div>
            </article>
          ))}
        </div>
      )}
    </div>
  )
}
