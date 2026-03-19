import { useState, useEffect, useRef, useCallback } from 'react'
import api from '@/api/client'
import { useAuthStore } from '@/stores/auth'
import { formatRelativeTime } from '@/lib/utils'
import { ArrowLeft, Send, Loader2, MessageCircle } from 'lucide-react'
import { io, type Socket } from 'socket.io-client'

interface Conversation {
  id: string
  participants: { id: string; username: string; displayName: string; avatarUrl: string | null }[]
  lastMessage: { body: string; createdAt: string; senderId: string } | null
  unreadCount: number
}

interface Message {
  id: string
  senderId: string
  body: string
  createdAt: string
  status: string
}

export default function MessagesPage() {
  const { user } = useAuthStore()
  const [conversations, setConversations] = useState<Conversation[]>([])
  const [activeConvo, setActiveConvo] = useState<Conversation | null>(null)
  const [messages, setMessages] = useState<Message[]>([])
  const [newMessage, setNewMessage] = useState('')
  const [loading, setLoading] = useState(true)
  const [sending, setSending] = useState(false)
  const messagesEndRef = useRef<HTMLDivElement>(null)
  const socketRef = useRef<Socket | null>(null)

  useEffect(() => {
    const fetchConversations = async () => {
      setLoading(true)
      try {
        const { data } = await api.get('/conversations')
        setConversations(data.data || [])
      } catch {
        setConversations([])
      }
      setLoading(false)
    }
    fetchConversations()
  }, [])

  const scrollToBottom = useCallback(() => {
    setTimeout(() => messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' }), 100)
  }, [])

  // WebSocket connection
  useEffect(() => {
    const token = localStorage.getItem('accessToken')
    if (!token) return

    const socket = io(window.location.origin, {
      path: '/socket.io',
      auth: { token },
      transports: ['websocket'],
    })

    socket.on('new_message', (msg: Message) => {
      setMessages(prev => [...prev, msg])
      scrollToBottom()
    })

    socketRef.current = socket
    return () => { socket.disconnect() }
  }, [scrollToBottom])

  const openConversation = async (convo: Conversation) => {
    setActiveConvo(convo)
    try {
      const { data } = await api.get(`/conversations/${convo.id}/messages`)
      setMessages(data.data || [])
      scrollToBottom()

      // Mark as read
      await api.post(`/conversations/${convo.id}/read`)
      setConversations(prev => prev.map(c => c.id === convo.id ? { ...c, unreadCount: 0 } : c))
    } catch {
      setMessages([])
    }
  }

  const handleSend = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!newMessage.trim() || !activeConvo) return
    setSending(true)
    try {
      const { data } = await api.post(`/conversations/${activeConvo.id}/messages`, {
        body: newMessage,
      })
      setMessages(prev => [...prev, data.data])
      setNewMessage('')
      scrollToBottom()
    } catch { /* ignore */ }
    setSending(false)
  }

  const getOtherParticipant = (convo: Conversation) => {
    return convo.participants.find(p => p.id !== user?.id) || convo.participants[0]
  }

  // Conversation list view
  if (!activeConvo) {
    return (
      <div>
        <div className="px-4 py-3 border-b border-border">
          <h2 className="text-lg font-bold text-text">Messages</h2>
        </div>

        {loading ? (
          <div className="flex justify-center py-12"><Loader2 size={24} className="animate-spin text-primary" /></div>
        ) : conversations.length === 0 ? (
          <div className="text-center py-12 text-text-muted">
            <MessageCircle size={48} className="mx-auto mb-4 opacity-50" />
            <p className="text-lg mb-1">No messages yet</p>
            <p className="text-sm">Start a conversation by visiting someone's profile</p>
          </div>
        ) : (
          <div className="divide-y divide-border">
            {conversations.map(convo => {
              const other = getOtherParticipant(convo)
              return (
                <button key={convo.id} onClick={() => openConversation(convo)}
                  className="w-full text-left px-4 py-3 hover:bg-surface/50 transition-colors flex items-center gap-3">
                  <div className="w-12 h-12 rounded-full bg-primary/20 flex items-center justify-center text-primary font-medium shrink-0">
                    {other.displayName[0]?.toUpperCase()}
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center justify-between">
                      <span className="font-medium text-sm text-text">{other.displayName}</span>
                      {convo.lastMessage && (
                        <span className="text-xs text-text-muted">{formatRelativeTime(convo.lastMessage.createdAt)}</span>
                      )}
                    </div>
                    <p className="text-sm text-text-muted truncate">
                      {convo.lastMessage?.body || 'No messages yet'}
                    </p>
                  </div>
                  {convo.unreadCount > 0 && (
                    <span className="bg-primary text-white text-xs rounded-full w-5 h-5 flex items-center justify-center">
                      {convo.unreadCount}
                    </span>
                  )}
                </button>
              )
            })}
          </div>
        )}
      </div>
    )
  }

  // Active conversation view
  const other = getOtherParticipant(activeConvo)
  return (
    <div className="flex flex-col h-full">
      {/* Chat header */}
      <div className="flex items-center gap-3 px-4 py-3 border-b border-border">
        <button onClick={() => setActiveConvo(null)} className="text-text-muted hover:text-text">
          <ArrowLeft size={20} />
        </button>
        <div className="w-8 h-8 rounded-full bg-primary/20 flex items-center justify-center text-primary text-sm font-medium">
          {other.displayName[0]?.toUpperCase()}
        </div>
        <div>
          <p className="font-medium text-sm text-text">{other.displayName}</p>
          <p className="text-xs text-text-muted">@{other.username}</p>
        </div>
      </div>

      {/* Messages */}
      <div className="flex-1 overflow-y-auto px-4 py-4 space-y-3">
        {messages.map(msg => {
          const isMine = msg.senderId === user?.id
          return (
            <div key={msg.id} className={`flex ${isMine ? 'justify-end' : 'justify-start'}`}>
              <div className={`max-w-xs px-4 py-2 rounded-2xl text-sm ${
                isMine
                  ? 'bg-primary text-white rounded-br-sm'
                  : 'bg-surface text-text rounded-bl-sm'
              }`}>
                <p>{msg.body}</p>
                <p className={`text-xs mt-1 ${isMine ? 'text-white/60' : 'text-text-muted'}`}>
                  {formatRelativeTime(msg.createdAt)}
                </p>
              </div>
            </div>
          )
        })}
        <div ref={messagesEndRef} />
      </div>

      {/* Message input */}
      <form onSubmit={handleSend} className="px-4 py-3 border-t border-border flex gap-2">
        <input
          value={newMessage}
          onChange={e => setNewMessage(e.target.value)}
          className="flex-1 bg-surface border border-border rounded-full px-4 py-2 text-sm text-text placeholder-text-muted focus:outline-none focus:border-primary"
          placeholder="Type a message..."
        />
        <button type="submit" disabled={sending || !newMessage.trim()}
          className="bg-primary hover:bg-primary-dark text-white rounded-full p-2 transition-colors disabled:opacity-50">
          {sending ? <Loader2 size={18} className="animate-spin" /> : <Send size={18} />}
        </button>
      </form>
    </div>
  )
}
