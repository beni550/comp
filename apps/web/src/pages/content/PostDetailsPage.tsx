import { useState, useEffect } from 'react'
import { useParams, Link, useNavigate } from 'react-router-dom'
import api from '@/api/client'
import { formatRelativeTime } from '@/lib/utils'
import { Heart, MessageCircle, Share2, Bookmark, ArrowLeft, Send, MoreHorizontal, Loader2 } from 'lucide-react'

interface Author {
  id: string
  username: string
  displayName: string
  avatarUrl: string | null
}

interface Comment {
  id: string
  author: Author
  body: string
  createdAt: string
  likesCount: number
}

interface Post {
  id: string
  author: Author
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

export default function PostDetailsPage() {
  const { id } = useParams<{ id: string }>()
  const navigate = useNavigate()
  const [post, setPost] = useState<Post | null>(null)
  const [comments, setComments] = useState<Comment[]>([])
  const [newComment, setNewComment] = useState('')
  const [loading, setLoading] = useState(true)
  const [submitting, setSubmitting] = useState(false)

  useEffect(() => {
    const fetchPost = async () => {
      setLoading(true)
      try {
        const [postRes, commentsRes] = await Promise.all([
          api.get(`/content/${id}`),
          api.get(`/content/${id}/comments`),
        ])
        setPost(postRes.data.data)
        setComments(commentsRes.data.data || [])
      } catch {
        // Post not found
      }
      setLoading(false)
    }
    if (id) fetchPost()
  }, [id])

  const handleReact = async () => {
    if (!post) return
    try {
      const { data } = await api.post(`/content/${post.id}/react`, { type: 'like' })
      setPost(prev => prev ? {
        ...prev,
        isLiked: data.data.action === 'added',
        likesCount: data.data.action === 'added' ? prev.likesCount + 1 : prev.likesCount - 1,
      } : null)
    } catch { /* ignore */ }
  }

  const handleComment = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!newComment.trim() || !post) return
    setSubmitting(true)
    try {
      const { data } = await api.post(`/content/${post.id}/comments`, { body: newComment })
      setComments(prev => [data.data, ...prev])
      setNewComment('')
      setPost(prev => prev ? { ...prev, commentsCount: prev.commentsCount + 1 } : null)
    } catch { /* ignore */ }
    setSubmitting(false)
  }

  const handleSave = async () => {
    if (!post) return
    try {
      await api.post(`/content/${post.id}/save`)
      setPost(prev => prev ? { ...prev, isSaved: !prev.isSaved } : null)
    } catch { /* ignore */ }
  }

  if (loading) {
    return (
      <div className="flex justify-center py-12">
        <Loader2 size={24} className="animate-spin text-primary" />
      </div>
    )
  }

  if (!post) {
    return (
      <div className="text-center py-12 text-text-muted">
        <p>Post not found</p>
        <button onClick={() => navigate('/feed')} className="text-primary mt-2">Go back to feed</button>
      </div>
    )
  }

  return (
    <div className="flex flex-col h-full">
      {/* Header */}
      <div className="flex items-center gap-3 px-4 py-3 border-b border-border">
        <button onClick={() => navigate(-1)} className="text-text-muted hover:text-text">
          <ArrowLeft size={20} />
        </button>
        <span className="font-medium">Post</span>
      </div>

      {/* Post content */}
      <div className="flex-1 overflow-y-auto">
        <div className="px-4 py-4">
          {/* Author */}
          <div className="flex items-center gap-3 mb-3">
            <Link to={`/profile/${post.author.username}`}>
              <div className="w-12 h-12 rounded-full bg-primary/20 flex items-center justify-center text-primary font-medium text-lg">
                {post.author.displayName[0]?.toUpperCase()}
              </div>
            </Link>
            <div className="flex-1">
              <Link to={`/profile/${post.author.username}`} className="font-medium text-text hover:text-primary">
                {post.author.displayName}
              </Link>
              <p className="text-xs text-text-muted">@{post.author.username} · {formatRelativeTime(post.publishedAt || post.createdAt)}</p>
            </div>
            <button className="text-text-muted hover:text-text"><MoreHorizontal size={18} /></button>
          </div>

          {/* Caption */}
          <p className="text-text mb-4 whitespace-pre-wrap">{post.caption}</p>

          {/* Media */}
          {post.media && post.media.length > 0 && (
            <div className="rounded-xl overflow-hidden mb-4 bg-surface">
              <img src={post.media[0].url} alt="" className="w-full object-cover" />
            </div>
          )}

          {/* Actions */}
          <div className="flex items-center justify-between text-text-muted py-3 border-t border-b border-border">
            <button onClick={handleReact}
              className={`flex items-center gap-1.5 text-sm transition-colors ${post.isLiked ? 'text-secondary' : 'hover:text-secondary'}`}>
              <Heart size={20} fill={post.isLiked ? 'currentColor' : 'none'} />
              <span>{post.likesCount}</span>
            </button>
            <div className="flex items-center gap-1.5 text-sm">
              <MessageCircle size={20} />
              <span>{post.commentsCount}</span>
            </div>
            <button className="flex items-center gap-1.5 text-sm hover:text-primary transition-colors">
              <Share2 size={20} />
              <span>{post.sharesCount}</span>
            </button>
            <button onClick={handleSave}
              className={`transition-colors ${post.isSaved ? 'text-warning' : 'hover:text-warning'}`}>
              <Bookmark size={20} fill={post.isSaved ? 'currentColor' : 'none'} />
            </button>
          </div>
        </div>

        {/* Comments */}
        <div className="px-4">
          <h3 className="font-medium text-sm text-text-muted mb-3">Comments ({post.commentsCount})</h3>
          {comments.length === 0 ? (
            <p className="text-text-muted text-sm py-4">No comments yet. Be the first!</p>
          ) : (
            <div className="space-y-4">
              {comments.map(comment => (
                <div key={comment.id} className="flex gap-3">
                  <div className="w-8 h-8 rounded-full bg-surface-light flex items-center justify-center text-text-muted text-xs font-medium shrink-0">
                    {comment.author.displayName[0]?.toUpperCase()}
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2">
                      <Link to={`/profile/${comment.author.username}`} className="text-sm font-medium text-text hover:text-primary">
                        {comment.author.displayName}
                      </Link>
                      <span className="text-xs text-text-muted">{formatRelativeTime(comment.createdAt)}</span>
                    </div>
                    <p className="text-sm text-text mt-1">{comment.body}</p>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* Comment input */}
      <form onSubmit={handleComment} className="px-4 py-3 border-t border-border flex gap-2">
        <input
          value={newComment}
          onChange={e => setNewComment(e.target.value)}
          className="flex-1 bg-surface border border-border rounded-full px-4 py-2 text-sm text-text placeholder-text-muted focus:outline-none focus:border-primary"
          placeholder="Add a comment..."
        />
        <button type="submit" disabled={submitting || !newComment.trim()}
          className="bg-primary hover:bg-primary-dark text-white rounded-full p-2 transition-colors disabled:opacity-50">
          {submitting ? <Loader2 size={18} className="animate-spin" /> : <Send size={18} />}
        </button>
      </form>
    </div>
  )
}
