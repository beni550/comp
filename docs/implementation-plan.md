# VYBE MVP - Implementation Plan

## Milestones (aligned with spec section 21)

### M1 - Foundations
**Goal**: Working dev environment with auth, users, settings, CI, admin shell

1. Bootstrap monorepo structure
   - Root package.json with workspaces
   - TypeScript configuration (base + per-app)
   - ESLint + Prettier configuration
   - apps/api, apps/web, packages/types, packages/validation, packages/config

2. Database schema + Prisma setup
   - Complete Prisma schema with all entities
   - Initial migration
   - Seed data: interests, safety rules, blocked words, system templates, feature flags, feed config

3. API server scaffold
   - Express + TypeScript with middleware stack
   - Error handling, request ID, CORS, rate limiting
   - Module registration pattern
   - Health check endpoint

4. Auth module
   - Phone/email registration with OTP
   - Login (credentials + social stubs)
   - JWT access/refresh token management
   - Session + device management
   - 2FA challenge flow
   - Username validation + availability check
   - Age verification
   - Account recovery flow

5. User module
   - Profile CRUD
   - Settings CRUD
   - Username change with 30-day limit
   - Account deletion (soft)

6. Onboarding module
   - Step-based onboarding flow
   - Interest selection
   - Avatar upload
   - Skip/resume support

7. Admin basics
   - Admin auth middleware
   - User lookup
   - Feature flags CRUD
   - Audit log viewer
   - Dashboard stats

**Deliverable**: Login flow works end-to-end, admin can look up users

### M2 - Social Core
**Goal**: Active social graph

8. Social module - Follow
   - Follow/unfollow
   - Follow requests for private accounts
   - Accept/decline requests
   - Followers/following lists with search

9. Social module - Friends
   - Friend request send/respond
   - Friends list
   - Mutual friends detection

10. Social module - Circles
    - Create/edit/delete circles
    - Add/remove members
    - Circle detail view
    - Member limit enforcement (25)

11. Social module - Block & Mute
    - Block/unblock (cascading follow/friend removal)
    - Mute/unmute
    - Blocked users list management

12. User search
    - PostgreSQL trigram search on username/displayName
    - Privacy-respecting results

**Deliverable**: Users can follow, friend, create circles, block

### M3 - Content Core
**Goal**: Full content creation and consumption

13. Media pipeline
    - Upload intent + signed URL generation
    - Upload completion + background processing
    - Thumbnail generation
    - Asset status tracking

14. Content module
    - Create/edit/delete content items
    - Publish flow (only when all media ready)
    - Audience permissions enforcement
    - Edit history tracking

15. Comments & reactions
    - Hierarchical comments
    - Like/reaction system
    - Save/bookmark

16. Stories
    - Create/view stories
    - 24h expiry
    - Story feed (grouped by author)
    - Highlights

**Deliverable**: Users can create posts, comment, like, share stories

### M4 - Feed & Chat
**Goal**: Daily active loop

17. Feed module
    - My VYBE (chronological from follows)
    - For You (ranked with configurable weights)
    - Circle feed
    - Profile feed
    - Negative feedback signals
    - Trending

18. Messaging module
    - 1:1 conversations
    - Circle group chats
    - Message types: text, image, reply, system
    - Read receipts (optional per settings)
    - Delete for self / delete for everyone (10-min window)
    - WebSocket integration

19. Notifications module
    - In-app notifications
    - Deduplication (5-min window)
    - Notification preferences
    - WebSocket real-time delivery
    - Notification outbox for tracking

**Deliverable**: Users can browse feed, chat, get notifications

### M5 - Safety & Beta Readiness
**Goal**: Beta candidate

20. Moderation module
    - Report submission with evidence snapshot
    - Moderation case queue
    - Case assignment + resolution
    - Moderation actions (warn, hide, remove, suspend, ban)
    - SLA tracking (24h)

21. Analytics module
    - Event tracking endpoint
    - Required v1 events implementation
    - Admin analytics summary

22. Wallet basics
    - Points ledger (invite bonus, etc.)
    - Balance queries

23. Frontend implementation
    - Auth flows (login, register, onboarding)
    - Profile view/edit
    - Social (follow, friends, circles)
    - Content creation + media upload
    - Feed views
    - Chat/messaging
    - Notifications
    - Admin panel basics

24. Testing
    - Unit tests for services + validators
    - Integration tests for auth, publish, report flows
    - API contract tests

25. Documentation & cleanup
    - Module READMEs
    - Local dev setup guide
    - Seed data verification

## Implementation Order

```
Phase 1 (Foundation):
  1. Monorepo + tooling
  2. Prisma schema + migrations + seed
  3. API scaffold + middleware
  4. Auth module
  5. User/profile module
  6. Onboarding module
  7. Admin basics

Phase 2 (Social):
  8. Follow/unfollow
  9. Friends
  10. Circles
  11. Block/mute
  12. User search

Phase 3 (Content):
  13. Media upload pipeline
  14. Content CRUD + publish
  15. Comments + reactions
  16. Stories + highlights

Phase 4 (Engagement):
  17. Feed (my, for-you, circle, profile, trending)
  18. Messaging + WebSocket
  19. Notifications

Phase 5 (Safety & Polish):
  20. Moderation + reports
  21. Analytics events
  22. Wallet points
  23. Frontend
  24. Tests
  25. Documentation
```

## Assumptions Documented

1. Web React app serves as primary MVP client (mobile React Native is future)
2. MinIO used as S3-compatible storage for local development
3. OTP codes logged to console in dev mode (no SMS gateway)
4. No E2EE in v1; encryption at rest + transport only
5. No real payment processing; wallet is points ledger only
6. Single-region deployment
7. PostgreSQL FTS with pg_trgm for search; no external search engine
8. Feed ranking uses configurable weights in DB, not ML model
9. Push notifications interface defined but delivery deferred (in-app only for v1)
10. Contact sync interface defined but actual import deferred
11. Video transcoding stubbed - stores original file, thumbnail generated
12. Story interactive types (poll, countdown) are basic JSON storage only

## Remaining Gaps After MVP

- React Native mobile app
- E2EE for messaging
- Real SMS/email delivery (SendGrid, Twilio integration)
- Real push notifications (FCM/APNs)
- Video transcoding pipeline
- ML-based feed ranking
- AI Curator
- AR features
- Marketplace / creator monetization
- Multi-region deployment
- Advanced moderation (ML classifiers)
- Contact sync implementation
- WCAG accessibility audit
- Performance load testing
