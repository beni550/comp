import {
  AccountType, UserStatus, PrivacyMode, AuthProvider, Theme, ProfileMode,
  FollowStatus, FriendStatus, CircleType, CirclePrivacyMode, CircleMemberRole,
  ContentType, ContentStatus, Audience, MediaKind, MediaStatus, ContentMediaRole,
  ConversationType, MessageKind, MessageStatus, DmPermission,
  ReportReason, ReportTargetType, ReportStatus, CasePriority, CaseResolution,
  ModerationActionType, FeatureFlagEnvironment, UserRole,
  DownloadPermission, SensitiveContentLevel,
} from './enums';

// ─── Auth DTOs ──────────────────────────────────────────────
export interface RegisterPhoneStartDTO {
  phone: string;
}

export interface RegisterPhoneVerifyDTO {
  challengeId: string;
  code: string;
  username: string;
  displayName: string;
  birthDate: string;
}

export interface RegisterEmailStartDTO {
  email: string;
}

export interface LoginDTO {
  method: AuthProvider;
  identifier?: string;
  code?: string;
  password?: string;
  providerToken?: string;
  deviceInfo: {
    platform: string;
    appVersion?: string;
  };
}

export interface AuthTokensDTO {
  accessToken: string;
  refreshToken: string;
}

export interface SessionDTO {
  id: string;
  deviceId: string;
  platform: string;
  ip: string;
  userAgent: string;
  createdAt: string;
  isCurrent: boolean;
}

// ─── User DTOs ──────────────────────────────────────────────
export interface UserDTO {
  id: string;
  username: string;
  displayName: string;
  email?: string;
  phone?: string;
  birthDate: string;
  accountType: AccountType;
  status: UserStatus;
  privacyMode: PrivacyMode;
  isMinor: boolean;
  role: UserRole;
  createdAt: string;
}

export interface PublicUserDTO {
  id: string;
  username: string;
  displayName: string;
  avatarUrl?: string;
  bio?: string;
  profileMode: ProfileMode;
  privacyMode: PrivacyMode;
  isVerified: boolean;
  followersCount?: number;
  followingCount?: number;
  postsCount?: number;
}

export interface ProfileDTO {
  bio: string;
  avatarMediaId?: string;
  avatarUrl?: string;
  theme: Theme;
  mood?: string;
  moodEmoji?: string;
  links: string[];
  profileMode: ProfileMode;
  featuredContentIds: string[];
  followersCount: number;
  followingCount: number;
  friendsCount: number;
  postsCount: number;
}

export interface SettingsDTO {
  defaultAudience: Audience;
  dmPermission: DmPermission;
  readReceipts: boolean;
  screenshotAlert: boolean;
  downloadPermission: DownloadPermission;
  discoverableByPhone: boolean;
  discoverableByEmail: boolean;
  sensitiveContentLevel: SensitiveContentLevel;
  autoplay: boolean;
  textScale: number;
  reduceMotion: boolean;
  dailyUsageLimit?: number;
  breakReminders: boolean;
  quietHoursStart?: string;
  quietHoursEnd?: string;
  locale: string;
  pushNotifications: Record<string, boolean>;
  emailNotifications: Record<string, boolean>;
}

// ─── Social DTOs ────────────────────────────────────────────
export interface FollowEdgeDTO {
  id: string;
  followerId: string;
  followeeId: string;
  status: FollowStatus;
  createdAt: string;
}

export interface FollowRequestDTO {
  id: string;
  follower: PublicUserDTO;
  createdAt: string;
}

export interface FriendEdgeDTO {
  id: string;
  user: PublicUserDTO;
  status: FriendStatus;
  createdAt: string;
}

export interface CircleDTO {
  id: string;
  ownerId: string;
  name: string;
  description: string;
  type: CircleType;
  privacyMode: CirclePrivacyMode;
  memberLimit: number;
  memberCount: number;
  avatarUrl?: string;
  createdAt: string;
}

export interface CircleMemberDTO {
  id: string;
  userId: string;
  username: string;
  displayName: string;
  avatarUrl?: string;
  role: CircleMemberRole;
  joinedAt: string;
}

// ─── Content DTOs ───────────────────────────────────────────
export interface MediaAssetDTO {
  id: string;
  kind: MediaKind;
  mimeType: string;
  width?: number;
  height?: number;
  durationMs?: number;
  url: string;
  thumbnailUrl?: string;
  status: MediaStatus;
}

export interface ContentItemDTO {
  id: string;
  authorId: string;
  author: PublicUserDTO;
  type: ContentType;
  status: ContentStatus;
  audience: Audience;
  caption: string;
  hashtags: string[];
  location?: string;
  allowComments: boolean;
  allowSharing: boolean;
  allowDownload: boolean;
  media: MediaAssetDTO[];
  likesCount: number;
  commentsCount: number;
  sharesCount: number;
  savesCount: number;
  isLiked?: boolean;
  isSaved?: boolean;
  publishedAt?: string;
  createdAt: string;
}

export interface CreateContentDTO {
  type: ContentType;
  caption?: string;
  audience: Audience;
  audienceCircleId?: string;
  mediaIds: string[];
  allowComments?: boolean;
  allowSharing?: boolean;
  allowDownload?: boolean;
  publishNow?: boolean;
}

export interface CommentDTO {
  id: string;
  contentId: string;
  authorId: string;
  author: PublicUserDTO;
  parentId?: string;
  body: string;
  likesCount: number;
  isLiked?: boolean;
  createdAt: string;
  replies?: CommentDTO[];
}

export interface ReactionDTO {
  id: string;
  actorId: string;
  reactionType: string;
  createdAt: string;
}

export interface StoryDTO {
  id: string;
  authorId: string;
  author: PublicUserDTO;
  media: MediaAssetDTO;
  audience: Audience;
  interactiveType?: string;
  interactiveData?: Record<string, unknown>;
  viewsCount: number;
  expiresAt: string;
  createdAt: string;
}

export interface StoryGroupDTO {
  author: PublicUserDTO;
  stories: StoryDTO[];
  hasUnviewed: boolean;
}

export interface HighlightDTO {
  id: string;
  ownerId: string;
  title: string;
  coverUrl?: string;
  storyCount: number;
  createdAt: string;
}

// ─── Feed DTOs ──────────────────────────────────────────────
export interface FeedItemDTO {
  content: ContentItemDTO;
  feedScore?: number;
  explainTopReason?: string;
}

// ─── Messaging DTOs ─────────────────────────────────────────
export interface ConversationDTO {
  id: string;
  type: ConversationType;
  title?: string;
  participants: PublicUserDTO[];
  lastMessage?: MessageDTO;
  unreadCount: number;
  createdAt: string;
}

export interface ConversationMemberDTO {
  userId: string;
  user: PublicUserDTO;
  role: CircleMemberRole;
  lastReadMessageId?: string;
  joinedAt: string;
}

export interface MessageDTO {
  id: string;
  conversationId: string;
  senderId: string;
  sender: PublicUserDTO;
  kind: MessageKind;
  body?: string;
  media?: MediaAssetDTO[];
  replyTo?: MessageDTO;
  reactions: MessageReactionDTO[];
  status: MessageStatus;
  createdAt: string;
}

export interface MessageReactionDTO {
  actorId: string;
  emoji: string;
  createdAt: string;
}

// ─── Notification DTOs ──────────────────────────────────────
export interface NotificationDTO {
  id: string;
  type: string;
  title: string;
  body: string;
  data: Record<string, unknown>;
  isRead: boolean;
  createdAt: string;
}

// ─── Moderation DTOs ────────────────────────────────────────
export interface ReportDTO {
  id: string;
  targetType: ReportTargetType;
  targetId: string;
  reason: ReportReason;
  description?: string;
  status: ReportStatus;
  createdAt: string;
}

export interface ModerationCaseDTO {
  id: string;
  reportId: string;
  report: ReportDTO;
  assigneeId?: string;
  status: ReportStatus;
  priority: CasePriority;
  resolution?: CaseResolution;
  notes?: string;
  slaDeadline: string;
  createdAt: string;
  resolvedAt?: string;
}

export interface ModerationActionDTO {
  id: string;
  caseId: string;
  actorId: string;
  actionType: ModerationActionType;
  targetType: ReportTargetType;
  targetId: string;
  reason: string;
  createdAt: string;
}

// ─── Admin DTOs ─────────────────────────────────────────────
export interface FeatureFlagDTO {
  id: string;
  key: string;
  description: string;
  isEnabled: boolean;
  rolloutPercentage: number;
  environment: FeatureFlagEnvironment;
  createdAt: string;
  updatedAt: string;
}

export interface AuditLogDTO {
  id: string;
  actorId?: string;
  action: string;
  targetType?: string;
  targetId?: string;
  metadata: Record<string, unknown>;
  ip?: string;
  createdAt: string;
}

export interface DashboardStatsDTO {
  signups24h: number;
  dau: number;
  moderationQueueSize: number;
  storageUsageBytes: number;
  activeUsers: number;
  totalContent: number;
  totalReports: number;
}

// ─── Analytics DTOs ─────────────────────────────────────────
export interface AnalyticsEventDTO {
  eventType: string;
  attributes: Record<string, unknown>;
  sessionId?: string;
}

// ─── Common ─────────────────────────────────────────────────
export interface PaginationMeta {
  cursor?: string;
  hasMore: boolean;
  total?: number;
}

export interface UploadIntentDTO {
  assetId: string;
  uploadUrl: string;
  expiresAt: string;
}
