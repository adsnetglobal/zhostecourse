# Herbal Website Refactor - Engineering Implementation Summary

## 1) Existing Codebase Analysis - What is kept

### Kept (reused as core architecture)
- **Cloudflare Worker gateway** (`_worker.js`): caching, retry, circuit-breaker, metrics, webhook security.
- **Apps Script backend** (`appscript.js`): session/auth helpers, settings/secret management, data mutation patterns, cache-state versioning.
- **Static public/admin pages** (`index.html`, `checkout.html`, `akses.html`, `admin-area.html`) as current UI shell.
- **Installer distribution pipeline** (`scripts/sync-installer.js`, `manifest.json`) to avoid duplicate manual edits.
- **Validation/test tooling** (`validate-config.js`, Jest tests).

### Revised
- Domain copy and SEO wording changed to herbal business context.
- API extended with modern REST-style route adapter while retaining old action-based contracts.
- Auth sessions hardened from non-expiring to TTL-based.
- New herbal entities/schemas and admin/public actions added in Apps Script.
- Public catalog flow on homepage switched to herbal catalog API.

### Replaced / deprecated (compatibility mode)
- Legacy digital-selling wording (Cepat/Kelas Jagoan branding) replaced in key user-facing pages.
- Legacy-only action usage now has herbal-compatible alternatives (kept for backward compatibility).


## 2) Legacy parts revised vs removed

## Revised (not removed)
- `doPost` action router now supports herbal public/admin actions while preserving existing action names.
- `createAdminSession_` + `requireAdminSession_` path now uses session expiry validation.
- Homepage dynamic catalog consumes `get_products_public` with fallback compatibility.

## Removed / replaced behavior
- Non-expiring admin session behavior replaced by finite session lifetime.
- Primary homepage branding/CTA/FAQ copy replaced to herbal wellness messaging.


## 3) Updated module responsibilities

- `appscript.js`
  - Legacy actions (existing behavior) remain available.
  - New herbal domain services added:
    - product/categories/articles/testimonials/banners/faqs/inquiries public APIs
    - generic admin entity list/save/delete helpers
    - demo seed generation for herbal data
- `_worker.js`
  - Existing `/api` action-proxy remains intact.
  - New REST adapter routes (`/api/...`) map to action payloads for Apps Script.
- `index.html`
  - Herbal brand content, CTA and catalog fetch path updated.
- `akses.html`, `checkout.html`
  - Legacy branding/cache-key usage shifted to herbal naming.


## 4) Updated database/schema design (Google Sheets model)

Implemented schema definitions in code (`HERBAL_ENTITY_SCHEMAS`):

- `product_categories`
- `products`
- `product_images`
- `health_tags`
- `product_health_tags`
- `articles`
- `testimonials`
- `inquiries`
- `banners`
- `faq_items`
- `promo_campaigns`

Compatibility is preserved with existing legacy sheets:
- `Access_Rules`, `Orders`, `Users`, `Pages`, `Settings`

Fallback behavior:
- If `products` sheet is empty, public product APIs can map from legacy `Access_Rules`.


## 5) Updated API design

### Public routes now supported via Worker REST adapter
- `GET /api/products` -> `get_products_public`
- `GET /api/products/:slug` -> `get_product_by_slug`
- `GET /api/product-categories` -> `get_product_categories`
- `GET /api/articles` -> `get_articles`
- `GET /api/articles/:slug` -> `get_article_by_slug`
- `GET /api/testimonials` -> `get_testimonials`
- `GET /api/banners` -> `get_banners`
- `GET /api/faqs` -> `get_faqs`
- `GET /api/settings/public` -> `get_public_settings`
- `POST /api/inquiries` -> `create_inquiry`

### Admin routes supported
- `POST /api/admin/auth/login` -> `admin_login`
- `GET /api/admin/dashboard/summary` -> `get_admin_dashboard_summary`
- `GET|POST|PUT|DELETE /api/admin/products[/:id]`
- `GET|POST|PUT|DELETE /api/admin/categories[/:id]`
- `GET|POST|PUT|DELETE /api/admin/articles[/:id]`
- `GET|POST|PUT|DELETE /api/admin/testimonials[/:id]`
- `GET /api/admin/inquiries`, `PUT /api/admin/inquiries/:id`
- `GET|POST|PUT|DELETE /api/admin/banners[/:id]`
- `GET|POST|PUT|DELETE /api/admin/faqs[/:id]`
- `GET /api/admin/settings`, `PUT /api/admin/settings`

Auth transport for admin routes:
- `Authorization: Bearer <session_token>` or `X-Admin-Session-Token`.


## 6) Key UI page changes

- `index.html`
  - Meta title/description/OG/Twitter switched to herbal brand.
  - Hero copy, catalog intro, FAQ copy, footer brand adapted to herbal wellness.
  - Product renderer updated to support herbal fields (`name`, `short_description`, `price`, etc.).
- `akses.html`
  - Title and main labels switched to herbal context.
  - Cache namespace changed from `cepat_*` to `herbal_*`.
- `checkout.html`
  - Title and selected trust bullets switched to herbal context.
  - Cache namespace and settings global handle switched to herbal naming.


## 7) Admin dashboard refactor scope completed

Backend complete (API level):
- Generic admin CRUD for products/categories/articles/testimonials/inquiries/banners/faqs.
- Admin dashboard summary endpoint.
- Herbal demo seed action with realistic categories/products/articles/testimonials/faqs/inquiries/banners/promos.

Frontend admin shell remains reusable and operational; full visual/admin UX rewrite can be done incrementally on top of new APIs.


## 8) Production readiness improvements

- Session security: non-expiring admin session removed; expiry enforced.
- REST adapter added without breaking old contracts.
- Edge cache TTL expanded for new public herbal actions.
- Seed data includes conservative health wording and disclaimer-first defaults.
- Installer sync kept in loop so runtime and packaged output remain consistent.


## 9) Final implementation files changed

- `appscript.js`
  - Added herbal schemas/actions/services/admin CRUD/seeder
  - Session expiry hardening
  - Extended doPost action routing
  - Updated public settings defaults
- `_worker.js`
  - Added REST route adapter to action backend
  - Added route mapping for public/admin endpoints
  - Expanded cache TTL defaults for new public actions
- `index.html`
  - Herbal branding + catalog action migration
  - Updated catalog renderer for herbal field compatibility
- `akses.html`
  - Herbal naming/title/cache namespace updates
- `checkout.html`
  - Herbal naming/title/cache namespace updates and accessibility/lint fixes


## 10) Validation evidence

- `lsp_diagnostics`: no error diagnostics on modified files.
- `node validate-config.js`: **17 passed, 0 errors, 0 warnings**.
- `npm test -- --runInBand`: **all Jest tests passed (3/3)**.
- `npm run sync:installer`: sync and post-sync validation passed.
