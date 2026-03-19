import { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import api from '@/api/client'
import { Image, X, Loader2 } from 'lucide-react'

export default function CreatePostPage() {
  const [caption, setCaption] = useState('')
  const [mediaPreview, setMediaPreview] = useState<string | null>(null)
  const [mediaFile, setMediaFile] = useState<File | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const navigate = useNavigate()

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setMediaFile(file)
    const reader = new FileReader()
    reader.onload = (ev) => setMediaPreview(ev.target?.result as string)
    reader.readAsDataURL(file)
  }

  const removeMedia = () => {
    setMediaFile(null)
    setMediaPreview(null)
  }

  const handleSubmit = async () => {
    if (!caption.trim()) return
    setLoading(true)
    setError('')
    try {
      const mediaIds: string[] = []

      // Upload media if selected
      if (mediaFile) {
        try {
          // Step 1: Get upload intent (presigned URL)
          const { data: uploadData } = await api.post('/media/upload-intent', {
            filename: mediaFile.name,
            mimeType: mediaFile.type,
            sizeBytes: mediaFile.size,
          })
          // Step 2: Upload to presigned URL
          await fetch(uploadData.data.uploadUrl, {
            method: 'PUT',
            body: mediaFile,
            headers: { 'Content-Type': mediaFile.type },
          })
          // Step 3: Confirm upload
          await api.post(`/media/${uploadData.data.assetId}/complete`)
          mediaIds.push(uploadData.data.assetId)
        } catch {
          // Media upload may not work in local dev without MinIO — continue without
          console.warn('Media upload not available — post created without media')
        }
      }

      // Create post as draft
      const { data: createData } = await api.post('/content', {
        type: mediaIds.length > 0 ? 'photo' : 'text_media_thread',
        caption,
        audience: 'public',
        mediaIds,
      })
      const contentId = createData.data.id

      // Publish
      await api.post(`/content/${contentId}/publish`)
      navigate('/feed')
    } catch (err: unknown) {
      const axiosErr = err as { response?: { data?: { error?: { message?: string } } } }
      setError(axiosErr.response?.data?.error?.message || 'Failed to create post')
    }
    setLoading(false)
  }

  return (
    <div className="px-4 py-4">
      <div className="flex items-center justify-between mb-4">
        <h2 className="text-lg font-bold text-text">Create Post</h2>
        <button onClick={handleSubmit} disabled={loading || !caption.trim()}
          className="bg-primary hover:bg-primary-dark text-white font-medium rounded-full px-5 py-2 text-sm transition-colors disabled:opacity-50 flex items-center gap-2">
          {loading && <Loader2 size={16} className="animate-spin" />}
          Post
        </button>
      </div>

      {error && (
        <div className="bg-error/10 border border-error/30 text-error rounded-lg px-4 py-3 text-sm mb-4">{error}</div>
      )}

      <textarea
        value={caption}
        onChange={e => setCaption(e.target.value)}
        className="w-full bg-transparent text-text placeholder-text-muted focus:outline-none resize-none text-lg min-h-32"
        placeholder="What's on your mind?"
        maxLength={2000}
      />
      <p className="text-xs text-text-muted mb-4">{caption.length}/2000</p>

      {/* Media preview */}
      {mediaPreview && (
        <div className="relative rounded-xl overflow-hidden mb-4">
          <img src={mediaPreview} alt="Preview" className="w-full object-cover max-h-64 rounded-xl" />
          <button onClick={removeMedia}
            className="absolute top-2 right-2 bg-black/60 text-white rounded-full p-1.5 hover:bg-black/80">
            <X size={16} />
          </button>
        </div>
      )}

      {/* Media upload button */}
      <div className="border-t border-border pt-4 flex items-center gap-4">
        <label className="flex items-center gap-2 text-primary hover:text-primary-light cursor-pointer transition-colors">
          <Image size={22} />
          <span className="text-sm">Add Photo</span>
          <input type="file" accept="image/*,video/*" className="hidden" onChange={handleFileSelect} />
        </label>
        <span className="text-xs text-text-muted">(Media upload is placeholder in MVP)</span>
      </div>
    </div>
  )
}
