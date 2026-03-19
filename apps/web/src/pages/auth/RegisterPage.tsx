import { useState } from 'react'
import { Link, useNavigate } from 'react-router-dom'
import { useAuthStore } from '@/stores/auth'
import { Loader2 } from 'lucide-react'

export default function RegisterPage() {
  const [form, setForm] = useState({ email: '', username: '', displayName: '', password: '', birthDate: '' })
  const [error, setError] = useState('')
  const { register, isLoading } = useAuthStore()
  const navigate = useNavigate()

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setError('')
    try {
      await register(form)
      navigate('/onboarding')
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Registration failed'
      if (msg.includes('REGISTRATION_PENDING')) {
        setError('Registration requires OTP verification. For MVP demo, use seeded test accounts (e.g. alice@test.com / Password123!).')
      } else {
        const axiosErr = err as { response?: { data?: { error?: { message?: string } } } }
        setError(axiosErr.response?.data?.error?.message || msg)
      }
    }
  }

  const update = (field: string, value: string) => setForm(prev => ({ ...prev, [field]: value }))

  return (
    <div className="min-h-screen flex items-center justify-center px-4 bg-bg">
      <div className="w-full max-w-sm">
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold bg-gradient-to-r from-primary to-secondary bg-clip-text text-transparent">VYBE</h1>
          <p className="text-text-muted mt-2">Create your account</p>
        </div>

        <form onSubmit={handleSubmit} className="space-y-4">
          {error && (
            <div className="bg-error/10 border border-error/30 text-error rounded-lg px-4 py-3 text-sm">{error}</div>
          )}
          <div>
            <label className="block text-sm text-text-muted mb-1">Email</label>
            <input type="email" value={form.email} onChange={e => update('email', e.target.value)}
              className="w-full bg-surface border border-border rounded-lg px-4 py-3 text-text placeholder-text-muted focus:outline-none focus:border-primary" placeholder="you@example.com" required />
          </div>
          <div>
            <label className="block text-sm text-text-muted mb-1">Username</label>
            <input type="text" value={form.username} onChange={e => update('username', e.target.value)}
              className="w-full bg-surface border border-border rounded-lg px-4 py-3 text-text placeholder-text-muted focus:outline-none focus:border-primary" placeholder="Choose a username" required />
          </div>
          <div>
            <label className="block text-sm text-text-muted mb-1">Display Name</label>
            <input type="text" value={form.displayName} onChange={e => update('displayName', e.target.value)}
              className="w-full bg-surface border border-border rounded-lg px-4 py-3 text-text placeholder-text-muted focus:outline-none focus:border-primary" placeholder="Your display name" required />
          </div>
          <div>
            <label className="block text-sm text-text-muted mb-1">Birth Date</label>
            <input type="date" value={form.birthDate} onChange={e => update('birthDate', e.target.value)}
              className="w-full bg-surface border border-border rounded-lg px-4 py-3 text-text placeholder-text-muted focus:outline-none focus:border-primary" required />
          </div>
          <div>
            <label className="block text-sm text-text-muted mb-1">Password</label>
            <input type="password" value={form.password} onChange={e => update('password', e.target.value)}
              className="w-full bg-surface border border-border rounded-lg px-4 py-3 text-text placeholder-text-muted focus:outline-none focus:border-primary" placeholder="Min 8 characters" required minLength={8} />
          </div>
          <button type="submit" disabled={isLoading}
            className="w-full bg-primary hover:bg-primary-dark text-white font-medium rounded-lg px-4 py-3 transition-colors disabled:opacity-50 flex items-center justify-center gap-2">
            {isLoading && <Loader2 size={18} className="animate-spin" />}
            Create Account
          </button>
        </form>

        <p className="text-center text-text-muted text-sm mt-6">
          Already have an account? <Link to="/login" className="text-primary hover:text-primary-light">Sign In</Link>
        </p>
      </div>
    </div>
  )
}
