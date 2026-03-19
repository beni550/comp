import { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import api from '@/api/client'
import { Loader2 } from 'lucide-react'

const INTERESTS = [
  'Music', 'Art', 'Photography', 'Gaming', 'Sports',
  'Fashion', 'Food', 'Travel', 'Tech', 'Fitness',
]

export default function OnboardingPage() {
  const [step, setStep] = useState(1)
  const [bio, setBio] = useState('')
  const [selected, setSelected] = useState<string[]>([])
  const [loading, setLoading] = useState(false)
  const navigate = useNavigate()

  const toggle = (interest: string) => {
    setSelected(prev =>
      prev.includes(interest) ? prev.filter(i => i !== interest) : [...prev, interest]
    )
  }

  const handleComplete = async () => {
    setLoading(true)
    try {
      if (bio) {
        await api.patch('/users/me', { bio })
      }
      if (selected.length > 0) {
        await api.post('/users/me/interests', { interests: selected })
      }
      await api.post('/users/me/onboarding/complete', {})
    } catch {
      // Non-critical — continue anyway
    }
    setLoading(false)
    navigate('/feed')
  }

  return (
    <div className="min-h-screen flex items-center justify-center px-4 bg-bg">
      <div className="w-full max-w-sm">
        <div className="text-center mb-8">
          <h1 className="text-2xl font-bold text-text">Set up your profile</h1>
          <p className="text-text-muted mt-1">Step {step} of 2</p>
          <div className="flex gap-2 mt-4 justify-center">
            <div className={`h-1 w-16 rounded ${step >= 1 ? 'bg-primary' : 'bg-border'}`} />
            <div className={`h-1 w-16 rounded ${step >= 2 ? 'bg-primary' : 'bg-border'}`} />
          </div>
        </div>

        {step === 1 && (
          <div className="space-y-4">
            <div>
              <label className="block text-sm text-text-muted mb-1">Bio</label>
              <textarea
                value={bio}
                onChange={e => setBio(e.target.value)}
                className="w-full bg-surface border border-border rounded-lg px-4 py-3 text-text placeholder-text-muted focus:outline-none focus:border-primary resize-none h-24"
                placeholder="Tell people about yourself..."
                maxLength={300}
              />
              <p className="text-xs text-text-muted mt-1">{bio.length}/300</p>
            </div>
            <button onClick={() => setStep(2)}
              className="w-full bg-primary hover:bg-primary-dark text-white font-medium rounded-lg px-4 py-3 transition-colors">
              Next
            </button>
            <button onClick={() => navigate('/feed')}
              className="w-full text-text-muted hover:text-text text-sm py-2 transition-colors">
              Skip for now
            </button>
          </div>
        )}

        {step === 2 && (
          <div className="space-y-4">
            <p className="text-sm text-text-muted">Select your interests</p>
            <div className="flex flex-wrap gap-2">
              {INTERESTS.map(interest => (
                <button key={interest} onClick={() => toggle(interest)}
                  className={`px-4 py-2 rounded-full text-sm transition-colors ${
                    selected.includes(interest)
                      ? 'bg-primary text-white'
                      : 'bg-surface border border-border text-text-muted hover:border-primary'
                  }`}>
                  {interest}
                </button>
              ))}
            </div>
            <button onClick={handleComplete} disabled={loading}
              className="w-full bg-primary hover:bg-primary-dark text-white font-medium rounded-lg px-4 py-3 transition-colors disabled:opacity-50 flex items-center justify-center gap-2">
              {loading && <Loader2 size={18} className="animate-spin" />}
              Get Started
            </button>
            <button onClick={() => setStep(1)}
              className="w-full text-text-muted hover:text-text text-sm py-2 transition-colors">
              Back
            </button>
          </div>
        )}
      </div>
    </div>
  )
}
