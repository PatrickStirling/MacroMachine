# Macro Machine Free vs Pro Strategy Notes

## Goal

Capture the current thinking around:

- whether Macro Machine should become a free + paid product
- how to split functionality in a way that still feels valuable for free users
- whether to ship one licensed app or two builds from one codebase
- what backend/private update infrastructure would be needed for a paid version

This is a planning note, not a final decision.

## Current Context

Macro Machine currently ships as one Electron app.

- Packaging lives in `FusionMacroReorderer-electron/`
- The app currently uses `electron-updater`
- Releases are currently published through public GitHub releases

That means any future paid/private path will require update/distribution changes even if the feature split is delayed.

## Product Split Philosophy

The free version should feel like a genuinely useful macro utility, not a broken demo.

The paid version should feel clearly more powerful for serious authors, product builders, and automation-heavy users.

Important principle:

- Do not make free worse at parsing/importing/exporting normal macros just to force an upgrade.
- Gate advanced creation, automation, and scale features instead.

## Proposed Free Tier

Free should likely include the core publishing-controls workflow:

- import/export `.setting`
- basic DRFX-related workflow if already part of the normal import/export experience
- nodes panel browsing
- search/filtering
- publish/unpublish controls
- reorder published controls
- rename controls
- page assignment
- basic label/button/separator support
- solid round-trip fidelity
- compatibility/parsing support for normal macros

This keeps Free useful and good-will-building.

## Proposed Premium Tier

Premium should focus on the high-leverage authoring and automation features:

- presets engine
- CSV / live data linking
- reload data workflows
- macro-root / group-layer authoring tools
- advanced control creation and editing
- combo-control authoring
- header image tools
- advanced expression / on-change helpers
- quick publish sets / templates
- multi-tab / batch / higher-scale workflows
- future advanced automation features

This makes Premium feel like the serious production version, while Free remains worthwhile.

## Option A: One Licensed App

### Model

Ship one app to everyone.

- all users download the same build
- free users get limited mode
- paid users unlock premium features with a license key or sign-in

### Advantages

- simplest packaging model
- one installer
- one app ID
- one update channel
- easiest free-to-paid upgrade UX
- fastest path to launch

### Disadvantages

- premium code ships inside the free app
- in Electron/JS, determined users can inspect or patch app code
- licensing acts more as a commercial barrier than a hard technical wall
- product messaging can feel blurrier over time

### Best For

- fastest launch
- lowest engineering overhead
- validating pricing/market fit before building more infrastructure

## Option B: Two Builds From One Codebase

### Model

Keep one source codebase, but produce two packaged apps:

- Macro Machine Free
- Macro Machine Pro

Feature gating happens at build time instead of only runtime.

### Advantages

- cleaner product split
- premium-only UI/features can be excluded from the free build
- stronger separation than runtime-only licensing
- easier to keep Pro updates private while leaving Free public
- clearer customer messaging

### Disadvantages

- more release complexity
- two installers to build/test/sign
- likely two update feeds/channels
- possibly two app IDs
- upgrade flow is slightly less seamless than a single licensed app

### Best For

- cleaner long-term product structure
- stronger separation of free vs paid
- better private-update strategy for paid users

## Important Clarification

Two builds from one codebase is not the same as two separate codebases.

Recommended if split happens:

- one codebase
- shared core logic
- build flags / environment flags / feature manifests decide Free vs Pro packaging

Not recommended:

- maintaining two independent apps/repositories with duplicated logic

That would create unnecessary maintenance overhead and drift.

## Current Recommendation

If the goal is speed:

- start with one licensed app

If the goal is the better long-term architecture:

- ship two builds from one codebase

Current lean:

- two builds from one codebase is the better long-term answer
- one licensed app is the easier short-term answer

This decision can be delayed until the product/feature set stabilizes more.

## Private Paid Updates

If any paid version exists, public GitHub releases are not a good long-term home for paid update delivery.

Even if Free and Pro remain one app, paid/private updates will still need a more private distribution path.

### Current Public Update Setup

Today the app updates from public GitHub releases via `electron-updater`.

That is fine for a public build.

It is not ideal for a paid/private product.

## Backend Pieces Needed For Paid / Private Distribution

At minimum:

- purchase system
- license database
- activation endpoint
- entitlement/token issuance
- private update feed or artifact server

### Purchase System

Possible providers:

- Stripe
- Lemon Squeezy
- Paddle
- Gumroad

This handles checkout and payment.

### Licensing

Basic flow:

1. user buys Pro
2. webhook creates a license record
3. user receives a license key or account access
4. app sends license + machine info to your server
5. server returns a signed entitlement token
6. app caches that token locally with an offline grace period
7. premium features unlock

### Private Update Delivery

Better than public GitHub releases for paid builds:

- S3 / Cloudflare R2 / Backblaze / similar object storage
- small authenticated update endpoint
- signed URLs for downloads
- Electron `generic` provider or equivalent private feed

This is usually cleaner than trying to force private GitHub release auth into the desktop app.

## If You Use One Licensed App

Backend implications:

- one app build
- one app ID
- one runtime license system
- likely one update feed with entitlement-aware access

Pros:

- simpler operations

Cons:

- premium code still ships in the app users already have

## If You Use Two Builds

Backend implications:

- two release outputs
- usually two app IDs
- separate update feeds/channels
- Pro updates can stay private
- Free updates can remain public if desired

Suggested app-identity pattern:

- `com.stirlingsupply.macromachine`
- `com.stirlingsupply.macromachinepro`

This avoids updater crossover and keeps channel separation cleaner.

## Security Reality

No Electron-only solution is perfect protection.

Important reality:

- if premium logic ships in the free app, it can be inspected
- licensing in Electron is deterrence/convenience, not absolute security

That is why two builds is stronger than one licensed binary, even though both still need real backend/license enforcement.

## Suggested Phased Path

### Phase 1

- keep one codebase
- define a clear Free vs Pro feature matrix
- decide whether to launch with one licensed app or two builds
- move paid/private updates off public GitHub releases

### Phase 2

- implement licensing / entitlements
- implement private update infrastructure
- refine the upgrade path from Free to Pro

### Phase 3

- revisit the split after real customer usage
- adjust the feature boundary based on actual value, not guesswork

## Recommended Feature Boundary for MM

Best current split:

Free:

- core macro publishing workflow
- basic creation/editing
- reliable import/export
- good browsing and organization

Pro:

- automation
- presets
- linking
- macro-root advanced tooling
- advanced authoring tools
- higher-scale workflows

That keeps Free useful and Pro clearly worth paying for.

## Next Time We Revisit This

Questions to answer:

1. Is speed-to-market more important than stronger feature separation?
2. Do we want Free and Pro to have separate installers/app identities?
3. Which premium features are truly essential vs just “nice to have”?
4. Which payment/license provider do we want?
5. Do we want public Free updates and private Pro updates?

## Short Conclusion

The best long-term product architecture is probably:

- one codebase
- two packaged builds
- Free centered on core publishing
- Pro centered on automation, presets, data linking, and advanced macro authoring

The fastest launch architecture is:

- one app
- license unlock for premium features

Either way, paid/private updates will require moving beyond the current public GitHub release update flow.
