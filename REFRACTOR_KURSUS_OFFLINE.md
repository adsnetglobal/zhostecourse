# Refactor Mapping — Web Kursus Offline

## 1) Keep as-is

- `config.js` (domain lock, encrypted URL decoding, resilient fetch wrappers)
- `site.config.js` (domain/preview config contract)
- deployment scaffolding: `wrangler.jsonc`, `manifest.json`, `setup.js`, `scripts/sync-installer.js`
- shared asset layer: `tailwind.css`, `assets/vendor/lucide.min.js`

## 2) Keep but revise

- `_worker.js`: keep routing/caching/health structure; revise API action mapping to kursus domain
- `appscript.js`: keep spreadsheet transport and utility helpers; revise entity schemas and business actions
- `index.html`: keep shell + SEO helper; revise copy, cards, and CTA flow for offline classes
- `checkout.html`: keep multi-step flow; revise fields into registration/inquiry for course + schedule
- `login.html`, `akses.html`, `dashboard.html`: keep auth/session shell; revise labels and modules to learner/training context
- `admin-area.html`, `admin-orders.html`: keep dashboard shell/table-modal patterns; revise module semantics to courses/schedules/leads/CMS

## 3) Replace

- Herbal/product terminology, static copy, and product-specific labels
- Product/order-centric API payload keys and schema references
- Old marketing blocks not relevant to offline training conversion funnel

## 4) Remove

- dead routes/components/scripts tied only to legacy category and no reusable value
- obsolete entity fields that cannot map to course/schedule/registration domain

---

## Target Domain Entities

- `course_categories`
- `courses`
- `instructors`
- `branches`
- `schedules`
- `registrations`
- `testimonials`
- `galleries`
- `faqs`
- `banners`
- `content_blocks`
- `admin_users`

## Public Route Structure (target)

- `/` homepage
- `/about`
- `/courses`
- `/courses/:slug`
- `/schedules`
- `/instructors`
- `/gallery`
- `/testimonials`
- `/contact`
- `/register` (registration/inquiry)

## Admin Route/Module Structure (target)

- dashboard summary
- course management
- schedule management
- instructor management
- registration/lead management
- testimonial management
- gallery management
- branch management
- page content/CMS blocks
- promo/banner management
- user + role management
- settings

## API Structure (target)

### Public
- `GET /courses`
- `GET /courses/:slug`
- `GET /course-categories`
- `GET /schedules`
- `GET /schedules/:id`
- `GET /instructors`
- `GET /branches`
- `GET /testimonials`
- `GET /faqs`
- `GET /settings/public`
- `GET /pages/:slug`
- `POST /registrations`
- `POST /contact`

### Admin
- CRUD `/admin/courses`
- CRUD `/admin/course-categories`
- CRUD `/admin/schedules`
- CRUD `/admin/instructors`
- CRUD `/admin/branches`
- CRUD `/admin/testimonials`
- CRUD `/admin/galleries`
- CRUD `/admin/faqs`
- CRUD `/admin/banners`
- CRUD `/admin/content-blocks`
- CRUD `/admin/users`
- `GET /admin/dashboard/summary`
- `GET /admin/registrations`
- `PATCH /admin/registrations/:id/status`
- `PATCH /admin/settings`
