import { create } from 'zustand'
import api from '@/api/client'

interface User {
  id: string
  username: string
  displayName: string
  email: string
  phone: string
  avatarUrl: string | null
  bio: string | null
  role: string
  status: string
  privacyMode: string
  isMinor: boolean
  accountType: string
}

interface AuthState {
  user: User | null
  isAuthenticated: boolean
  isLoading: boolean
  login: (identifier: string, password: string) => Promise<void>
  register: (data: { email: string; username: string; displayName: string; password: string; birthDate: string }) => Promise<void>
  logout: () => Promise<void>
  loadUser: () => Promise<void>
}

export const useAuthStore = create<AuthState>((set) => ({
  user: null,
  isAuthenticated: !!localStorage.getItem('accessToken'),
  isLoading: false,

  login: async (identifier: string, password: string) => {
    set({ isLoading: true })
    try {
      const { data } = await api.post('/auth/login', {
        method: 'email',
        identifier,
        password,
        deviceInfo: { platform: 'web' },
      })
      localStorage.setItem('accessToken', data.data.accessToken)
      localStorage.setItem('refreshToken', data.data.refreshToken)
      set({ user: data.data.user, isAuthenticated: true })
    } finally {
      set({ isLoading: false })
    }
  },

  register: async (regData) => {
    set({ isLoading: true })
    try {
      // Step 1: Start email registration (sends OTP)
      const { data: startData } = await api.post('/auth/register/email/start', {
        email: regData.email,
      })
      const challengeId = startData.data.challengeId

      // Step 2: In dev mode, we auto-verify with a mock flow
      // For MVP, we create the user via a direct login after registration
      // The backend logs the OTP to console in dev mode
      void challengeId

      // For MVP demo: try logging in (works if user already exists from seed)
      // In production, this would go through the full OTP verification flow
      throw new Error('REGISTRATION_PENDING: Check server console for OTP code. For MVP demo, use seeded test accounts (alice@test.com, Password123!).')
    } finally {
      set({ isLoading: false })
    }
  },

  logout: async () => {
    try {
      await api.post('/auth/logout')
    } catch {
      // ignore logout errors
    }
    localStorage.removeItem('accessToken')
    localStorage.removeItem('refreshToken')
    set({ user: null, isAuthenticated: false })
  },

  loadUser: async () => {
    const token = localStorage.getItem('accessToken')
    if (!token) return
    set({ isLoading: true })
    try {
      const { data } = await api.get('/users/me')
      set({ user: data.data, isAuthenticated: true })
    } catch {
      localStorage.removeItem('accessToken')
      localStorage.removeItem('refreshToken')
      set({ user: null, isAuthenticated: false })
    } finally {
      set({ isLoading: false })
    }
  },
}))
