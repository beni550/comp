# VYBE MVP - API Contracts

## Base URL
```
/api/v1
```

## Common Response Envelope

### Success
```json
{
  "data": { ... },
  "meta": {
    "requestId": "uuid",
    "timestamp": "ISO8601"
  }
}
```

### Error
```json
{
  "error": {
    "code": "ERROR_CODE",
    "message": "Human-readable message",
    "details": { ... },
    "requestId": "uuid"
  }
}
```

### Paginated Response
```json
{
  "data": [ ... ],
  "meta": {
    "requestId": "uuid",
    "cursor": "next-cursor-value",
    "hasMore": true,
    "total": 100
  }
}
```

---

## Module: Auth

### POST /auth/register/phone/start
Start phone registration - sends OTP.
```
Request: { phone: string (E.164) }
Response: { challengeId: string, expiresAt: string }
Rate limit: 3/min per phone
```

### POST /auth/register/phone/verify
Verify OTP and complete registration.
```
Request: { challengeId: string, code: string, username: string, displayName: string, birthDate: string }
Response: { user: UserDTO, accessToken: string, refreshToken: string }
```

### POST /auth/register/email/start
Start email registration - sends verification code.
```
Request: { email: string }
Response: { challengeId: string, expiresAt: string }
```

### POST /auth/login
Login with credentials or provider token.
```
Request: { 
  method: "phone" | "email" | "google" | "apple",
  identifier?: string,    // phone or email
  code?: string,           // OTP code
  password?: string,       // if password auth
  providerToken?: string,  // social login token
  deviceInfo: { platform: string, appVersion?: string }
}
Response: { user: UserDTO, accessToken: string, refreshToken: string, requiresTwoFactor?: boolean }
```

### POST /auth/refresh
Refresh access token.
```
Request: { refreshToken: string }
Response: { accessToken: string, refreshToken: string }
```

### POST /auth/logout
Logout current session.
```
Headers: Authorization: Bearer <token>
Response: { success: true }
```

### POST /auth/logout-all
Logout all sessions except current.
```
Headers: Authorization: Bearer <token>
Response: { revokedCount: number }
```

### GET /auth/sessions
List active sessions.
```
Headers: Authorization: Bearer <token>
Response: { sessions: SessionDTO[] }
```

### DELETE /auth/sessions/:sessionId
Revoke a specific session.
```
Headers: Authorization: Bearer <token>
Response: { success: true }
```

### POST /auth/2fa/challenge
Send 2FA challenge for new device login.
```
Request: { method: "sms" | "email" }
Response: { challengeId: string }
```

### POST /auth/recovery/start
Start account recovery flow.
```
Request: { identifier: string }
Response: { challengeId: string, method: string }
```

### GET /auth/username/check/:username
Check username availability.
```
Response: { available: boolean, reason?: string }
```

---

## Module: Onboarding

### GET /onboarding/progress
Get current onboarding state.
```
Headers: Authorization: Bearer <token>
Response: { currentStep: string, completedSteps: string[], isComplete: boolean }
```

### POST /onboarding/step/:stepName
Complete an onboarding step.
```
Headers: Authorization: Bearer <token>
Request: { data: object }  // Step-specific data (interests, avatar, etc.)
Response: { nextStep: string | null, isComplete: boolean }
```

### POST /onboarding/skip/:stepName
Skip a skippable onboarding step.
```
Headers: Authorization: Bearer <token>
Response: { nextStep: string | null }
Error: 400 if step is not skippable
```

---

## Module: Users / Profile

### GET /users/me
Get current user profile.
```
Headers: Authorization: Bearer <token>
Response: { user: UserDTO, profile: ProfileDTO, settings: SettingsDTO }
```

### PATCH /users/me
Update current user basic info.
```
Headers: Authorization: Bearer <token>
Request: { displayName?: string, username?: string, email?: string, phone?: string }
Response: { user: UserDTO }
Constraints: username change max once per 30 days
```

### GET /users/:userId
Get user public profile.
```
Headers: Authorization: Bearer <token> (optional for public profiles)
Response: { user: PublicUserDTO, profile: PublicProfileDTO }
Note: Counters visibility based on privacy settings
```

### PATCH /users/me/profile
Update profile details.
```
Headers: Authorization: Bearer <token>
Request: { bio?: string, avatarMediaId?: string, theme?: string, mood?: string, moodEmoji?: string, links?: string[], profileMode?: string, featuredContentIds?: string[] }
Response: { profile: ProfileDTO }
```

### PATCH /users/me/settings
Update user settings.
```
Headers: Authorization: Bearer <token>
Request: Partial<SettingsDTO>
Response: { settings: SettingsDTO }
```

### GET /users/search
Search users by username/displayName.
```
Query: { q: string, limit?: number, cursor?: string }
Response: { users: PublicUserDTO[], meta: PaginationMeta }
Note: pg_trgm for typo tolerance
```

### DELETE /users/me
Soft-delete (deactivate) account.
```
Headers: Authorization: Bearer <token>
Request: { confirmation: "DELETE" }
Response: { success: true }
```

---

## Module: Social

### POST /social/follow/:userId
Follow a user (or create follow request for private accounts).
```
Headers: Authorization: Bearer <token>
Response: { status: "following" | "requested" }
```

### DELETE /social/follow/:userId
Unfollow or cancel follow request.
```
Headers: Authorization: Bearer <token>
Response: { success: true }
```

### GET /social/followers
List current user's followers.
```
Query: { limit?, cursor?, search? }
Response: { users: PublicUserDTO[], meta }
```

### GET /social/following
List users the current user follows.
```
Query: { limit?, cursor?, search? }
Response: { users: PublicUserDTO[], meta }
```

### POST /social/follow-requests/:requestId/respond
Accept or decline a follow request.
```
Request: { action: "accept" | "decline" }
Response: { success: true }
```

### GET /social/follow-requests
List pending follow requests.
```
Query: { limit?, cursor? }
Response: { requests: FollowRequestDTO[], meta }
```

### POST /social/friend-request/:userId
Send friend request.
```
Headers: Authorization: Bearer <token>
Response: { friendEdge: FriendEdgeDTO }
```

### POST /social/friend-request/:id/respond
Accept or decline friend request.
```
Request: { action: "accept" | "decline" }
Response: { success: true }
```

### GET /social/friends
List friends.
```
Query: { limit?, cursor?, search? }
Response: { users: PublicUserDTO[], meta }
```

### POST /social/block/:userId
Block a user (removes follow/friend).
```
Headers: Authorization: Bearer <token>
Response: { success: true }
```

### DELETE /social/block/:userId
Unblock a user.
```
Headers: Authorization: Bearer <token>
Response: { success: true }
```

### GET /social/blocked
List blocked users.
```
Query: { limit?, cursor? }
Response: { users: PublicUserDTO[], meta }
```

### POST /social/mute
Mute a target.
```
Request: { targetId: string, targetType: "user" | "content" | "circle" }
Response: { success: true }
```

### DELETE /social/mute/:targetId
Unmute.
```
Response: { success: true }
```

### POST /circles
Create a circle.
```
Request: { name: string, description?: string, type?: string, privacyMode?: string }
Response: { circle: CircleDTO }
Constraint: memberLimit defaults to 25
```

### GET /circles
List user's circles.
```
Response: { circles: CircleDTO[] }
```

### GET /circles/:id
Get circle details.
```
Response: { circle: CircleDTO, members: CircleMemberDTO[] }
```

### PATCH /circles/:id
Update circle.
```
Request: { name?: string, description?: string }
Response: { circle: CircleDTO }
```

### DELETE /circles/:id
Delete circle.
```
Response: { success: true }
```

### POST /circles/:id/members
Add member to circle.
```
Request: { userId: string, role?: string }
Response: { member: CircleMemberDTO }
```

### DELETE /circles/:id/members/:userId
Remove member from circle.
```
Response: { success: true }
```

---

## Module: Content

### POST /media/intents
Create upload intent - returns signed URL.
```
Headers: Authorization: Bearer <token>
Request: { filename: string, mimeType: string, sizeBytes: number }
Response: { assetId: string, uploadUrl: string, expiresAt: string }
Validation: mime type whitelist, size limits
```

### POST /media/:id/complete
Confirm upload completion.
```
Headers: Authorization: Bearer <token>
Response: { asset: MediaAssetDTO }
Triggers: background processing (scan, metadata, thumbnail)
```

### POST /content
Create content item (draft or published).
```
Headers: Authorization: Bearer <token>
Request: {
  type: "video" | "photo" | "carousel" | "text_media_thread",
  caption?: string,
  audience: "public" | "followers" | "friends" | "circle" | "private",
  audienceCircleId?: string,
  mediaIds: string[],    // ordered media asset IDs
  allowComments?: boolean,
  allowSharing?: boolean,
  allowDownload?: boolean,
  publishNow?: boolean
}
Response: { content: ContentItemDTO }
```

### PATCH /content/:id
Edit content (caption, audience, media order).
```
Request: { caption?: string, audience?: string, mediaIds?: string[] }
Response: { content: ContentItemDTO }
Note: Does not reset counters; records editHistory
```

### POST /content/:id/publish
Publish a draft content item.
```
Response: { content: ContentItemDTO }
Error: 400 if any media not in 'ready' status
```

### DELETE /content/:id
Soft-delete content.
```
Response: { success: true }
Note: Existing shares show "content unavailable"
```

### GET /content/:id
Get content item.
```
Response: { content: ContentItemDTO, media: MediaAssetDTO[] }
Auth: respects audience permissions
```

### POST /content/:id/comments
Add comment.
```
Request: { body: string, parentId?: string }
Response: { comment: CommentDTO }
```

### GET /content/:id/comments
List comments.
```
Query: { limit?, cursor?, sort?: "newest" | "oldest" | "top" }
Response: { comments: CommentDTO[], meta }
```

### DELETE /comments/:id
Delete own comment.
```
Response: { success: true }
```

### POST /content/:id/reactions
Add/change reaction.
```
Request: { reactionType?: string }  // default: "like"
Response: { reaction: ReactionDTO }
```

### DELETE /content/:id/reactions
Remove reaction.
```
Response: { success: true }
```

### POST /content/:id/save
Save content.
```
Response: { saveItem: SaveItemDTO }
```

### DELETE /content/:id/save
Unsave content.
```
Response: { success: true }
```

### POST /stories
Create story.
```
Request: { mediaId: string, audience: string, interactiveType?: string, interactiveData?: object, saveToArchive?: boolean }
Response: { story: StoryDTO }
```

### GET /stories/feed
Get stories from followed users.
```
Response: { stories: StoryGroupDTO[] }  // Grouped by author
```

### POST /highlights
Create highlight collection.
```
Request: { title: string, storyIds: string[], coverMediaId?: string }
Response: { highlight: HighlightDTO }
```

---

## Module: Feed

### GET /feed/my
Chronological feed from followed users, friends, circles.
```
Query: { limit?, cursor? }
Response: { items: FeedItemDTO[], meta }
```

### GET /feed/for-you
Ranked discovery feed.
```
Query: { limit?, cursor? }
Response: { items: FeedItemDTO[], meta }
Each item includes explainTopReason
```

### GET /feed/circle/:circleId
Circle-specific feed.
```
Query: { limit?, cursor? }
Response: { items: FeedItemDTO[], meta }
```

### GET /users/:userId/feed
User profile feed.
```
Query: { limit?, cursor? }
Response: { items: FeedItemDTO[], meta }
```

### POST /feedback/not-interested
Negative feedback signal for feed.
```
Request: { contentId: string, reason?: string }
Response: { success: true }
```

### GET /discover/trending
Trending content.
```
Query: { timeframe?: "1h" | "24h" | "7d", category?: string }
Response: { items: FeedItemDTO[], meta }
```

---

## Module: Messaging

### GET /conversations
List user's conversations.
```
Query: { limit?, cursor? }
Response: { conversations: ConversationDTO[], meta }
```

### POST /conversations
Create new conversation.
```
Request: { type: "direct" | "circle_group", participantIds?: string[], circleId?: string, title?: string }
Response: { conversation: ConversationDTO }
DM constraint: checks target's dmPermission setting
```

### GET /conversations/:id
Get conversation details.
```
Response: { conversation: ConversationDTO, members: ConversationMemberDTO[] }
```

### GET /conversations/:id/messages
Get messages in conversation.
```
Query: { limit?, cursor?, before? }
Response: { messages: MessageDTO[], meta }
```

### POST /conversations/:id/messages
Send message.
```
Request: { kind: "text" | "image" | "video" | "reply", body?: string, mediaIds?: string[], replyToId?: string }
Response: { message: MessageDTO }
```

### DELETE /messages/:id
Delete message.
```
Query: { forEveryone?: boolean }
Response: { success: true }
Note: forEveryone only within 10-minute window
```

### POST /messages/:id/reactions
React to message.
```
Request: { emoji: string }
Response: { reaction: MessageReactionDTO }
```

### POST /conversations/:id/read
Mark conversation as read up to a message.
```
Request: { messageId: string }
Response: { success: true }
```

---

## Module: Notifications

### GET /notifications
List notifications.
```
Query: { limit?, cursor?, unreadOnly?: boolean }
Response: { notifications: NotificationDTO[], meta, unreadCount: number }
```

### POST /notifications/:id/read
Mark notification as read.
```
Response: { success: true }
```

### POST /notifications/read-all
Mark all as read.
```
Response: { updatedCount: number }
```

### PATCH /notifications/settings
Update notification preferences.
```
Request: { [eventType]: { push: boolean, email: boolean, inApp: boolean } }
Response: { settings: NotificationSettingsDTO }
```

---

## Module: Moderation

### POST /reports
Submit a report.
```
Headers: Authorization: Bearer <token>
Request: {
  targetType: "user" | "content" | "comment" | "message" | "circle",
  targetId: string,
  reason: "spam" | "harassment" | "hate_speech" | "violence" | "nudity" | "misinformation" | "other",
  description?: string
}
Response: { report: ReportDTO, caseId: string }
Auto: evidence snapshot captured, SLA 24h
```

### GET /reports/mine
My submitted reports with status.
```
Response: { reports: ReportStatusDTO[] }
```

---

## Module: Admin (requires admin/moderator role)

### GET /admin/dashboard
Dashboard stats.
```
Response: { signups24h, dau, moderationQueueSize, storageUsageBytes, activeUsers }
```

### GET /admin/users
User lookup.
```
Query: { q?, status?, limit?, cursor? }
Response: { users: AdminUserDTO[], meta }
```

### GET /admin/users/:id
User detail with privacy-safe summary.
```
Response: { user: AdminUserDetailDTO, sessions: SessionDTO[], recentActions: AuditLogDTO[] }
```

### POST /admin/users/:id/suspend
Suspend user.
```
Request: { reason: string, durationDays?: number }
Response: { success: true }
```

### POST /admin/users/:id/ban
Ban user.
```
Request: { reason: string }
Response: { success: true }
```

### GET /admin/moderation/queue
Moderation case queue.
```
Query: { status?, priority?, limit?, cursor? }
Response: { cases: ModerationCaseDTO[], meta }
```

### PATCH /admin/moderation/cases/:id
Update moderation case.
```
Request: { status?, assigneeId?, resolution?, notes? }
Response: { case: ModerationCaseDTO }
```

### POST /admin/moderation/cases/:id/actions
Take moderation action.
```
Request: { actionType: string, reason: string }
Response: { action: ModerationActionDTO }
```

### GET /admin/feature-flags
List feature flags.
```
Response: { flags: FeatureFlagDTO[] }
```

### POST /admin/feature-flags
Create feature flag.
```
Request: { key: string, description?: string, isEnabled: boolean, rolloutPercentage?: number, environment?: string }
Response: { flag: FeatureFlagDTO }
```

### PATCH /admin/feature-flags/:id
Update feature flag.
```
Request: { isEnabled?: boolean, rolloutPercentage?: number }
Response: { flag: FeatureFlagDTO }
```

### GET /admin/audit-logs
View audit logs.
```
Query: { actorId?, action?, targetType?, from?, to?, limit?, cursor? }
Response: { logs: AuditLogDTO[], meta }
```

---

## Module: Analytics

### POST /analytics/events
Track analytics event.
```
Headers: Authorization: Bearer <token>
Request: { eventType: string, attributes: object, sessionId?: string }
Response: { success: true }
```

### GET /admin/analytics/summary
Analytics summary for admin dashboard.
```
Query: { from?: string, to?: string }
Response: { signups: number, activeUsers: number, contentCreated: number, engagementRate: number, ... }
```

---

## WebSocket Events (Socket.io)

### Client → Server
| Event | Payload | Notes |
|-------|---------|-------|
| `conversation:join` | `{ conversationId }` | Join chat room after auth |
| `message:send` | `{ conversationId, kind, body, mediaIds?, replyToId? }` | Send message |
| `message:read` | `{ conversationId, messageId }` | Update read cursor |
| `typing:start` | `{ conversationId }` | Typing indicator |
| `typing:stop` | `{ conversationId }` | |
| `presence:update` | `{ status: "online" | "away" | "offline" }` | |

### Server → Client
| Event | Payload | Notes |
|-------|---------|-------|
| `message:new` | `MessageDTO` | New message in conversation |
| `message:deleted` | `{ messageId, conversationId }` | Message deleted |
| `message:reaction` | `{ messageId, actorId, emoji }` | Reaction added |
| `typing:update` | `{ conversationId, userId, isTyping }` | |
| `presence:update` | `{ userId, status }` | Ephemeral |
| `notification:new` | `NotificationDTO` | Real-time notification |
| `conversation:updated` | `ConversationDTO` | Conversation metadata change |
