// ─── User & Auth ────────────────────────────────────────────
export enum AccountType {
  STANDARD = 'standard',
  CREATOR = 'creator',
  PORTFOLIO = 'portfolio',
}

export enum UserStatus {
  ACTIVE = 'active',
  SUSPENDED = 'suspended',
  BANNED = 'banned',
  DEACTIVATED = 'deactivated',
  REJECTED_PENDING_DELETE = 'rejected_pending_delete',
}

export enum PrivacyMode {
  PUBLIC = 'public',
  FOLLOWERS = 'followers',
  FRIENDS_ONLY = 'friends_only',
  HIDDEN = 'hidden',
}

export enum AuthProvider {
  PHONE = 'phone',
  EMAIL = 'email',
  GOOGLE = 'google',
  APPLE = 'apple',
}

export enum OtpChannel {
  SMS = 'sms',
  EMAIL = 'email',
}

export enum OtpPurpose {
  REGISTRATION = 'registration',
  LOGIN = 'login',
  RECOVERY = 'recovery',
  VERIFICATION = 'verification',
}

export enum DevicePlatform {
  IOS = 'ios',
  ANDROID = 'android',
  WEB = 'web',
}

export enum VerificationRequestType {
  CREATOR = 'creator',
  IDENTITY = 'identity',
}

export enum VerificationRequestStatus {
  PENDING = 'pending',
  APPROVED = 'approved',
  REJECTED = 'rejected',
}

// ─── Profile ────────────────────────────────────────────────
export enum Theme {
  LIGHT = 'light',
  DARK = 'dark',
  SYSTEM = 'system',
}

export enum ProfileMode {
  STANDARD = 'standard',
  CREATOR = 'creator',
  PORTFOLIO = 'portfolio',
}

// ─── Social ─────────────────────────────────────────────────
export enum FollowStatus {
  ACTIVE = 'active',
  PENDING = 'pending',
}

export enum FriendStatus {
  PENDING = 'pending',
  ACTIVE = 'active',
  DECLINED = 'declined',
}

export enum CircleType {
  GENERAL = 'general',
  INTEREST = 'interest',
  CLOSE_FRIENDS = 'close_friends',
}

export enum CirclePrivacyMode {
  PRIVATE = 'private',
  INVITE_ONLY = 'invite_only',
}

export enum CircleMemberRole {
  OWNER = 'owner',
  ADMIN = 'admin',
  MEMBER = 'member',
}

export enum BlockScope {
  FULL = 'full',
  CONTENT_ONLY = 'content_only',
}

export enum MuteTargetType {
  USER = 'user',
  CONTENT = 'content',
  CIRCLE = 'circle',
}

// ─── Content ────────────────────────────────────────────────
export enum ContentType {
  VIDEO = 'video',
  PHOTO = 'photo',
  CAROUSEL = 'carousel',
  TEXT_MEDIA_THREAD = 'text_media_thread',
}

export enum ContentStatus {
  DRAFT = 'draft',
  PROCESSING = 'processing',
  PUBLISHED = 'published',
  HIDDEN_REVIEW = 'hidden_review',
  REMOVED = 'removed',
  DRAFT_FAILED = 'draft_failed',
}

export enum Audience {
  PUBLIC = 'public',
  FOLLOWERS = 'followers',
  FRIENDS = 'friends',
  CIRCLE = 'circle',
  PRIVATE = 'private',
}

export enum MediaKind {
  IMAGE = 'image',
  VIDEO = 'video',
  AUDIO = 'audio',
}

export enum MediaStatus {
  PENDING = 'pending',
  PROCESSING = 'processing',
  READY = 'ready',
  FAILED = 'failed',
}

export enum ModerationStatus {
  PENDING = 'pending',
  APPROVED = 'approved',
  FLAGGED = 'flagged',
  REJECTED = 'rejected',
}

export enum ContentMediaRole {
  PRIMARY = 'primary',
  THUMBNAIL = 'thumbnail',
  ATTACHMENT = 'attachment',
}

export enum InteractiveType {
  POLL = 'poll',
  COUNTDOWN = 'countdown',
  STICKER = 'sticker',
}

export enum CommentStatus {
  ACTIVE = 'active',
  HIDDEN = 'hidden',
  REMOVED = 'removed',
}

export enum ReactionTargetType {
  CONTENT = 'content',
  COMMENT = 'comment',
  STORY = 'story',
  MESSAGE = 'message',
}

// ─── Messaging ──────────────────────────────────────────────
export enum ConversationType {
  DIRECT = 'direct',
  CIRCLE_GROUP = 'circle_group',
}

export enum MessageKind {
  TEXT = 'text',
  IMAGE = 'image',
  VIDEO = 'video',
  SYSTEM = 'system',
  REPLY = 'reply',
}

export enum MessageStatus {
  SENT = 'sent',
  DELIVERED = 'delivered',
  READ = 'read',
  DELETED_FOR_SENDER = 'deleted_for_sender',
  DELETED_FOR_ALL = 'deleted_for_all',
}

export enum DmPermission {
  EVERYONE = 'everyone',
  FOLLOWERS = 'followers',
  FRIENDS = 'friends',
  NOBODY = 'nobody',
}

// ─── Notifications ──────────────────────────────────────────
export enum NotificationChannel {
  IN_APP = 'in_app',
  PUSH = 'push',
  EMAIL = 'email',
}

export enum NotificationOutboxStatus {
  PENDING = 'pending',
  SENT = 'sent',
  FAILED = 'failed',
  SKIPPED = 'skipped',
}

// ─── Moderation ─────────────────────────────────────────────
export enum ReportReason {
  SPAM = 'spam',
  HARASSMENT = 'harassment',
  HATE_SPEECH = 'hate_speech',
  VIOLENCE = 'violence',
  NUDITY = 'nudity',
  MISINFORMATION = 'misinformation',
  OTHER = 'other',
}

export enum ReportTargetType {
  USER = 'user',
  CONTENT = 'content',
  COMMENT = 'comment',
  MESSAGE = 'message',
  CIRCLE = 'circle',
}

export enum ReportStatus {
  OPEN = 'open',
  IN_REVIEW = 'in_review',
  RESOLVED = 'resolved',
  ESCALATED = 'escalated',
}

export enum CasePriority {
  LOW = 'low',
  NORMAL = 'normal',
  HIGH = 'high',
  CRITICAL = 'critical',
}

export enum CaseResolution {
  NO_ACTION = 'no_action',
  WARN = 'warn',
  HIDE_CONTENT = 'hide_content',
  SUSPEND_USER = 'suspend_user',
  BAN_USER = 'ban_user',
}

export enum ModerationActionType {
  WARN = 'warn',
  HIDE = 'hide',
  REMOVE = 'remove',
  SUSPEND = 'suspend',
  BAN = 'ban',
  ESCALATE = 'escalate',
  DISMISS = 'dismiss',
}

// ─── Admin ──────────────────────────────────────────────────
export enum FeatureFlagEnvironment {
  DEVELOPMENT = 'development',
  STAGING = 'staging',
  PRODUCTION = 'production',
}

// ─── Feed ───────────────────────────────────────────────────
export enum FeedType {
  MY_VYBE = 'my_vybe',
  FOR_YOU = 'for_you',
  CIRCLE = 'circle',
  PROFILE = 'profile',
}

export enum TrendingTimeframe {
  ONE_HOUR = '1h',
  TWENTY_FOUR_HOURS = '24h',
  SEVEN_DAYS = '7d',
}

// ─── User Roles ─────────────────────────────────────────────
export enum UserRole {
  GUEST = 'guest',
  USER = 'user',
  MINOR = 'minor',
  CREATOR = 'creator',
  MODERATOR = 'moderator',
  SUPPORT = 'support',
  ADMIN = 'admin',
}

// ─── Wallet ─────────────────────────────────────────────────
export enum WalletTransactionType {
  INVITE_BONUS = 'invite_bonus',
  DAILY_REWARD = 'daily_reward',
  SPENT = 'spent',
}

// ─── Download Permission ────────────────────────────────────
export enum DownloadPermission {
  EVERYONE = 'everyone',
  FOLLOWERS = 'followers',
  NOBODY = 'nobody',
}

export enum SensitiveContentLevel {
  STANDARD = 'standard',
  LESS = 'less',
  MORE = 'more',
}

// ─── Safety ─────────────────────────────────────────────────
export enum SafetyRuleType {
  KEYWORD = 'keyword',
  PATTERN = 'pattern',
  POLICY = 'policy',
}

export enum SafetyRuleAction {
  FLAG = 'flag',
  BLOCK = 'block',
  ESCALATE = 'escalate',
}

export enum BlockedWordSeverity {
  LOW = 'low',
  MEDIUM = 'medium',
  HIGH = 'high',
}
