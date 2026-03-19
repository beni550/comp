import { z } from 'zod';

// ─── Username ───────────────────────────────────────────────
export const usernameSchema = z
  .string()
  .min(3, 'Username must be at least 3 characters')
  .max(24, 'Username must be at most 24 characters')
  .regex(
    /^(?!\.)(?!.*\.\.)(?!.*\.$)[a-zA-Z0-9_.]+$/,
    'Username can only contain letters, numbers, underscores, and dots. Cannot start/end with dot or have consecutive dots.'
  );

// ─── Display Name ───────────────────────────────────────────
export const displayNameSchema = z
  .string()
  .min(1, 'Display name is required')
  .max(60, 'Display name must be at most 60 characters');

// ─── Phone ──────────────────────────────────────────────────
export const phoneSchema = z
  .string()
  .regex(/^\+[1-9]\d{1,14}$/, 'Phone must be in E.164 format (e.g., +1234567890)');

// ─── Email ──────────────────────────────────────────────────
export const emailSchema = z
  .string()
  .email('Invalid email address')
  .transform((val) => val.toLowerCase().trim());

// ─── Birth Date ─────────────────────────────────────────────
export const birthDateSchema = z
  .string()
  .regex(/^\d{4}-\d{2}-\d{2}$/, 'Birth date must be in YYYY-MM-DD format')
  .refine((val) => {
    const date = new Date(val);
    return !isNaN(date.getTime()) && date < new Date();
  }, 'Birth date must be a valid past date');

// ─── Bio ────────────────────────────────────────────────────
export const bioSchema = z
  .string()
  .max(300, 'Bio must be at most 300 characters')
  .default('');

// ─── Comment ────────────────────────────────────────────────
export const commentBodySchema = z
  .string()
  .min(1, 'Comment cannot be empty')
  .max(2000, 'Comment must be at most 2000 characters');

// ─── Auth Schemas ───────────────────────────────────────────
export const registerPhoneStartSchema = z.object({
  phone: phoneSchema,
});

export const registerPhoneVerifySchema = z.object({
  challengeId: z.string().uuid(),
  code: z.string().length(6, 'OTP code must be 6 digits'),
  username: usernameSchema,
  displayName: displayNameSchema,
  birthDate: birthDateSchema,
});

export const registerEmailStartSchema = z.object({
  email: emailSchema,
});

export const loginSchema = z.object({
  method: z.enum(['phone', 'email', 'google', 'apple']),
  identifier: z.string().optional(),
  code: z.string().optional(),
  password: z.string().optional(),
  providerToken: z.string().optional(),
  deviceInfo: z.object({
    platform: z.enum(['ios', 'android', 'web']),
    appVersion: z.string().optional(),
  }),
});

// ─── Profile Schemas ────────────────────────────────────────
export const updateProfileSchema = z.object({
  bio: bioSchema.optional(),
  avatarMediaId: z.string().uuid().optional().nullable(),
  theme: z.enum(['light', 'dark', 'system']).optional(),
  mood: z.string().max(50).optional().nullable(),
  moodEmoji: z.string().max(10).optional().nullable(),
  links: z.array(z.string().url()).max(5).optional(),
  profileMode: z.enum(['standard', 'creator', 'portfolio']).optional(),
  featuredContentIds: z.array(z.string().uuid()).max(5).optional(),
});

export const updateUserSchema = z.object({
  displayName: displayNameSchema.optional(),
  username: usernameSchema.optional(),
  email: emailSchema.optional(),
  phone: phoneSchema.optional(),
});

export const updateSettingsSchema = z.object({
  defaultAudience: z.enum(['public', 'followers', 'friends', 'circle', 'private']).optional(),
  dmPermission: z.enum(['everyone', 'followers', 'friends', 'nobody']).optional(),
  readReceipts: z.boolean().optional(),
  screenshotAlert: z.boolean().optional(),
  downloadPermission: z.enum(['everyone', 'followers', 'nobody']).optional(),
  discoverableByPhone: z.boolean().optional(),
  discoverableByEmail: z.boolean().optional(),
  sensitiveContentLevel: z.enum(['standard', 'less', 'more']).optional(),
  autoplay: z.boolean().optional(),
  textScale: z.number().min(0.5).max(2.0).optional(),
  reduceMotion: z.boolean().optional(),
  dailyUsageLimit: z.number().int().min(1).optional().nullable(),
  breakReminders: z.boolean().optional(),
  quietHoursStart: z.string().optional().nullable(),
  quietHoursEnd: z.string().optional().nullable(),
  locale: z.string().max(10).optional(),
  pushNotifications: z.record(z.boolean()).optional(),
  emailNotifications: z.record(z.boolean()).optional(),
});

// ─── Content Schemas ────────────────────────────────────────
export const createContentSchema = z.object({
  type: z.enum(['video', 'photo', 'carousel', 'text_media_thread']),
  caption: z.string().max(5000).optional().default(''),
  audience: z.enum(['public', 'followers', 'friends', 'circle', 'private']),
  audienceCircleId: z.string().uuid().optional(),
  mediaIds: z.array(z.string().uuid()).min(0).max(10),
  allowComments: z.boolean().optional().default(true),
  allowSharing: z.boolean().optional().default(true),
  allowDownload: z.boolean().optional().default(false),
  publishNow: z.boolean().optional().default(false),
});

export const updateContentSchema = z.object({
  caption: z.string().max(5000).optional(),
  audience: z.enum(['public', 'followers', 'friends', 'circle', 'private']).optional(),
  mediaIds: z.array(z.string().uuid()).max(10).optional(),
});

export const uploadIntentSchema = z.object({
  filename: z.string().min(1),
  mimeType: z.string().min(1),
  sizeBytes: z.number().int().positive().max(1073741824), // 1GB max
});

// ─── Social Schemas ─────────────────────────────────────────
export const createCircleSchema = z.object({
  name: z.string().min(1).max(100),
  description: z.string().max(500).optional().default(''),
  type: z.enum(['general', 'interest', 'close_friends']).optional().default('general'),
  privacyMode: z.enum(['private', 'invite_only']).optional().default('private'),
});

export const updateCircleSchema = z.object({
  name: z.string().min(1).max(100).optional(),
  description: z.string().max(500).optional(),
});

// ─── Messaging Schemas ──────────────────────────────────────
export const createConversationSchema = z.object({
  type: z.enum(['direct', 'circle_group']),
  participantIds: z.array(z.string().uuid()).optional(),
  circleId: z.string().uuid().optional(),
  title: z.string().max(100).optional(),
});

export const sendMessageSchema = z.object({
  kind: z.enum(['text', 'image', 'video', 'reply']),
  body: z.string().max(5000).optional(),
  mediaIds: z.array(z.string().uuid()).optional(),
  replyToId: z.string().uuid().optional(),
});

// ─── Report Schema ──────────────────────────────────────────
export const createReportSchema = z.object({
  targetType: z.enum(['user', 'content', 'comment', 'message', 'circle']),
  targetId: z.string().uuid(),
  reason: z.enum(['spam', 'harassment', 'hate_speech', 'violence', 'nudity', 'misinformation', 'other']),
  description: z.string().max(1000).optional(),
});

// ─── Admin Schemas ──────────────────────────────────────────
export const createFeatureFlagSchema = z.object({
  key: z.string().min(1).max(100),
  description: z.string().max(500).optional().default(''),
  isEnabled: z.boolean().default(false),
  rolloutPercentage: z.number().int().min(0).max(100).optional().default(0),
  environment: z.enum(['development', 'staging', 'production']).optional().default('development'),
});

export const updateFeatureFlagSchema = z.object({
  isEnabled: z.boolean().optional(),
  rolloutPercentage: z.number().int().min(0).max(100).optional(),
  description: z.string().max(500).optional(),
});

// ─── Analytics Schema ───────────────────────────────────────
export const trackEventSchema = z.object({
  eventType: z.string().min(1).max(100),
  attributes: z.record(z.unknown()),
  sessionId: z.string().optional(),
});

// ─── Story Schema ───────────────────────────────────────────
export const createStorySchema = z.object({
  mediaId: z.string().uuid(),
  audience: z.enum(['public', 'followers', 'friends', 'circle']),
  interactiveType: z.enum(['poll', 'countdown', 'sticker']).optional(),
  interactiveData: z.record(z.unknown()).optional(),
  saveToArchive: z.boolean().optional().default(false),
});

export const createHighlightSchema = z.object({
  title: z.string().min(1).max(100),
  storyIds: z.array(z.string().uuid()).min(1),
  coverMediaId: z.string().uuid().optional(),
});
