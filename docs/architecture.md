# VYBE MVP - System Architecture

## 1. Architecture Overview

VYBE v1 follows a **Modular Monolith** pattern — all domain modules live in a single deployable API application but maintain strict boundaries: each module owns its service layer, repository, DTOs, events, and tests.

```
┌─────────────────────────────────────────────────────────────┐
│                        Clients                              │
│   ┌──────────┐  ┌──────────┐  ┌──────────┐                │
│   │ Mobile   │  │ Web App  │  │ Admin    │                │
│   │ (future) │  │ (React)  │  │ (React)  │                │
│   └────┬─────┘  └────┬─────┘  └────┬─────┘                │
└────────┼─────────────┼─────────────┼────────────────────────┘
         │             │             │
         ▼             ▼             ▼
┌─────────────────────────────────────────────────────────────┐
│                    API Gateway Layer                         │
│  ┌─────────────────────────────────────────────────────┐    │
│  │  Express + TypeScript    /api/v1/*                  │    │
│  │  ┌──────┐ ┌──────┐ ┌────────┐ ┌────────────┐      │    │
│  │  │ CORS │ │ Rate │ │ Auth   │ │ Request ID │      │    │
│  │  │      │ │Limit │ │ Guard  │ │ Middleware  │      │    │
│  │  └──────┘ └──────┘ └────────┘ └────────────┘      │    │
│  └─────────────────────────────────────────────────────┘    │
│                                                             │
│  ┌─────────────────────────────────────────────────────┐    │
│  │              Domain Modules                         │    │
│  │  ┌────────┐ ┌────────┐ ┌─────────┐ ┌──────────┐   │    │
│  │  │  Auth  │ │ Users  │ │ Social  │ │ Content  │   │    │
│  │  └────────┘ └────────┘ └─────────┘ └──────────┘   │    │
│  │  ┌────────┐ ┌────────┐ ┌─────────┐ ┌──────────┐   │    │
│  │  │  Feed  │ │Messag- │ │ Notifi- │ │Moderat-  │   │    │
│  │  │        │ │  ing   │ │ cations │ │  ion     │   │    │
│  │  └────────┘ └────────┘ └─────────┘ └──────────┘   │    │
│  │  ┌────────┐ ┌────────┐                             │    │
│  │  │ Admin  │ │Analyt- │                             │    │
│  │  │        │ │  ics   │                             │    │
│  │  └────────┘ └────────┘                             │    │
│  └─────────────────────────────────────────────────────┘    │
│                                                             │
│  ┌──────────────────┐  ┌──────────────────┐                │
│  │  WebSocket GW    │  │  Background Jobs │                │
│  │  (Socket.io)     │  │  (BullMQ/Redis)  │                │
│  └──────────────────┘  └──────────────────┘                │
└─────────────────────────────────────────────────────────────┘
         │             │             │
         ▼             ▼             ▼
┌─────────────────────────────────────────────────────────────┐
│                    Data Layer                                │
│  ┌──────────┐  ┌──────────┐  ┌──────────────┐             │
│  │PostgreSQL│  │  Redis   │  │ S3-compatible│             │
│  │ (Prisma) │  │ Cache +  │  │ Object Store │             │
│  │          │  │ Queues   │  │ (Media)      │             │
│  └──────────┘  └──────────┘  └──────────────┘             │
└─────────────────────────────────────────────────────────────┘
```

## 2. Monorepo Structure

```
vybe/
├── apps/
│   ├── api/                  # Express + TypeScript API server
│   │   ├── src/
│   │   │   ├── modules/      # Domain modules
│   │   │   │   ├── auth/
│   │   │   │   ├── users/
│   │   │   │   ├── social/
│   │   │   │   ├── content/
│   │   │   │   ├── feed/
│   │   │   │   ├── messaging/
│   │   │   │   ├── notifications/
│   │   │   │   ├── moderation/
│   │   │   │   ├── admin/
│   │   │   │   └── analytics/
│   │   │   ├── middleware/
│   │   │   ├── common/
│   │   │   └── main.ts
│   │   ├── prisma/
│   │   │   ├── schema.prisma
│   │   │   ├── migrations/
│   │   │   └── seed.ts
│   │   └── package.json
│   └── web/                  # React + Vite + Tailwind frontend
│       ├── src/
│       │   ├── pages/
│       │   ├── components/
│       │   ├── hooks/
│       │   ├── stores/
│       │   ├── api/
│       │   └── App.tsx
│       └── package.json
├── packages/
│   ├── types/                # Shared DTOs, enums, interfaces
│   ├── validation/           # Zod schemas shared client+server
│   └── config/               # Shared constants, feature flags
├── package.json              # Root workspace
├── tsconfig.base.json
└── docs/
```

## 3. Technology Stack

| Layer | Technology | Rationale |
|-------|-----------|-----------|
| Language | TypeScript end-to-end | Spec requirement; shared types |
| API Framework | Express.js + TypeScript | Mature, modular, spec-compatible |
| ORM | Prisma | Type-safe queries, migrations, seeding |
| Database | PostgreSQL | Spec requirement; FTS + trigram for search |
| Cache/Queue | Redis + BullMQ | Rate limiting, sessions, job queues, presence |
| Realtime | Socket.io | Chat, typing, presence, notification fan-out |
| Auth | JWT (access + refresh tokens) | Stateless access, revocable refresh |
| Media Storage | S3-compatible (MinIO for dev) | Signed uploads per spec |
| Frontend | React + Vite + Tailwind + Zustand | Modern, fast, spec-aligned |
| Validation | Zod | Shared client/server validation |
| Testing | Vitest + Supertest | Fast, TypeScript-native |

## 4. Module Boundaries

Each module follows this internal structure:
```
module/
├── module.router.ts      # Express routes
├── module.service.ts     # Business logic
├── module.repository.ts  # Database access (Prisma)
├── module.dto.ts         # Request/Response DTOs
├── module.events.ts      # Domain events emitted
├── module.middleware.ts   # Module-specific middleware
└── module.test.ts        # Tests
```

### Module Dependency Rules
- Modules communicate via service interfaces, not direct repository access
- Auth module is a dependency of all other modules (middleware)
- Analytics module listens to events from all modules
- No circular dependencies between domain modules

## 5. Cross-Cutting Concerns

### Authentication & Authorization
- JWT access tokens (15min TTL) + refresh tokens (30 day TTL)
- Device-bound sessions stored in DB
- Role-based access: guest, user, minor, creator, moderator, support, admin
- Auth middleware validates JWT and attaches user context

### Error Handling
- Consistent error envelope: `{ code, message, details, requestId }`
- HTTP status codes follow REST conventions
- Domain errors mapped to appropriate HTTP codes

### Feature Flags
- Stored in DB with environment segregation
- Checked via middleware or service calls
- Admin UI for management with audit trail

### Audit Logging
- Sensitive actions logged: login, password change, block, delete, moderator actions
- Stored in `audit_log` table with actor, action, target, metadata, timestamp

### Rate Limiting
- Redis-backed rate limiter per endpoint category
- Configurable limits per user role
- OTP: 5 attempts per challenge, cooldown escalation

## 6. Data Flow Patterns

### Media Upload Flow
1. Client requests upload intent → server returns signed URL + asset draft ID
2. Client uploads directly to object storage
3. Client calls complete endpoint → server marks upload complete
4. Background worker: virus scan, metadata extraction, thumbnail generation
5. Asset status transitions: `pending → processing → ready / failed`

### Feed Generation
1. Candidate generation from follow graph + interests + locale
2. Scoring with configurable weights (stored in DB/config, not hardcoded)
3. Fairness rules: exposure floor, diversity, duplicate suppression
4. Cursor-based pagination for infinite scroll

### Notification Pipeline
1. Domain event emitted (e.g., `follow.accepted`)
2. Notification service creates notification record
3. Deduplication check (5-minute window)
4. Fan-out: in-app storage + push queue + email queue
5. Delivery tracking via `notification_outbox`

## 7. Security Considerations

- TLS everywhere in production
- Passwords hashed with bcrypt (cost 12)
- OTP codes hashed before storage
- No secrets in code; all via environment variables
- CORS configured for known origins
- Rate limiting on auth endpoints
- Input validation on all endpoints via Zod
- SQL injection prevention via Prisma parameterized queries
- XSS prevention via React's built-in escaping

## 8. Assumptions

1. **MVP is web-first**: React web app serves as the primary client; React Native mobile is future work
2. **Single-region deployment**: No multi-region for v1
3. **MinIO for local dev**: S3-compatible local object storage for development
4. **Redis required**: Used for cache, rate limiting, sessions, queues, presence
5. **No E2EE in v1**: Encryption at rest + transport only; E2EE interface defined for beta
6. **No real payments**: Wallet/points ledger only; no payment gateway integration
7. **No SMS gateway in dev**: OTP codes logged to console in development mode
8. **PostgreSQL FTS**: Full-text search via PostgreSQL; no external search engine in v1

## 9. Ambiguities & Resolutions

| Ambiguity | Resolution |
|-----------|-----------|
| Tables in spec show entity fields but not all column types | Inferred from context; documented in data-model.md |
| Spec mentions React Native mobile but MVP scope unclear | Building React web app as primary client for MVP demo |
| WebSocket gateway specifics not fully defined | Using Socket.io with room-based architecture for chat + notifications |
| "Signed uploads" implementation details | Using pre-signed S3 PUT URLs with expiry |
| Feed ranking weights storage | Stored in `feed_config` DB table, editable via admin |
| Notification push delivery | In-app notifications only for v1; push interface defined but not connected |
| Contact sync / phone hashing | Interface defined; actual contact import deferred |
