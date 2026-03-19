export const APP_CONFIG = {
  // Auth
  ACCESS_TOKEN_TTL: '15m',
  REFRESH_TOKEN_TTL_DAYS: 30,
  OTP_TTL_MINUTES: 10,
  OTP_MAX_ATTEMPTS: 5,
  OTP_CODE_LENGTH: 6,
  USERNAME_CHANGE_COOLDOWN_DAYS: 30,
  MIN_AGE: 13,

  // Content
  MAX_CAROUSEL_ITEMS: 10,
  MAX_CAPTION_LENGTH: 5000,
  MAX_BIO_LENGTH: 300,
  MAX_LINKS: 5,
  MAX_FEATURED_ITEMS: 5,
  MAX_UPLOAD_SIZE_BYTES: 1073741824, // 1GB
  STORY_TTL_HOURS: 24,
  MAX_VIDEO_DURATION_SECONDS: 600, // 10 minutes
  MIN_VIDEO_DURATION_SECONDS: 5,

  // Social
  MAX_CIRCLE_MEMBERS: 25,
  MAX_CIRCLE_OWNERS: 2,

  // Messaging
  DELETE_FOR_ALL_WINDOW_MINUTES: 10,
  MAX_MESSAGE_LENGTH: 5000,

  // Notifications
  DEDUP_WINDOW_MINUTES: 5,

  // Moderation
  REPORT_SLA_HOURS: 24,

  // Feed
  DEFAULT_PAGE_SIZE: 20,
  MAX_PAGE_SIZE: 50,

  // Rate Limiting
  RATE_LIMIT_WINDOW_MS: 60000,
  RATE_LIMIT_MAX_REQUESTS: 100,
  AUTH_RATE_LIMIT_MAX: 10,
  OTP_RATE_LIMIT_MAX: 3,

  // Password
  BCRYPT_ROUNDS: 12,
} as const;

export const RESERVED_USERNAMES = [
  'admin', 'administrator', 'mod', 'moderator', 'support',
  'help', 'vybe', 'official', 'system', 'null', 'undefined',
  'api', 'www', 'app', 'mail', 'email', 'root', 'superuser',
  'test', 'demo', 'staff', 'team', 'bot', 'webhook',
] as const;

export const ONBOARDING_STEPS = [
  { name: 'interests', skippable: false, order: 1 },
  { name: 'avatar', skippable: true, order: 2 },
  { name: 'contacts', skippable: true, order: 3 },
  { name: 'follow_suggestions', skippable: true, order: 4 },
  { name: 'first_post', skippable: true, order: 5 },
] as const;

export const ALLOWED_MIME_TYPES = {
  image: ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/heic'],
  video: ['video/mp4', 'video/quicktime', 'video/webm'],
  audio: ['audio/mpeg', 'audio/wav', 'audio/ogg', 'audio/aac'],
} as const;

export const NOTIFICATION_TYPES = {
  FOLLOW_ACCEPTED: 'follow_accepted',
  FOLLOW_REQUEST: 'follow_request',
  FRIEND_REQUEST: 'friend_request',
  FRIEND_ACCEPTED: 'friend_accepted',
  NEW_DM: 'new_dm',
  STORY_REPLY: 'story_reply',
  CONTENT_LIKE: 'content_like',
  CONTENT_COMMENT: 'content_comment',
  SECURITY_LOGIN: 'security_login',
  MODERATION_DECISION: 'moderation_decision',
  CIRCLE_INVITE: 'circle_invite',
} as const;

export const ANALYTICS_EVENTS = {
  SIGNUP_COMPLETED: 'signup_completed',
  ONBOARDING_STEP_COMPLETED: 'onboarding_step_completed',
  CONTENT_PUBLISHED: 'content_published',
  FEED_IMPRESSION: 'feed_impression',
  CONTENT_ENGAGEMENT: 'content_engagement',
  DM_SENT: 'dm_sent',
  REPORT_SUBMITTED: 'report_submitted',
} as const;
