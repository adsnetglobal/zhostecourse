const ss = SpreadsheetApp.getActiveSpreadsheet();

/* =========================
   CONFIG.JS INTEGRATION
   (Server-side Configuration)
========================= */
const SCRIPT_CONFIG = {
  // SCRIPT_URL sengaja tidak di-hardcode untuk menghindari exposure endpoint di source.
  // Set via Script Properties: APP_SCRIPT_URL jika memang diperlukan.
  SCRIPT_URL: "",
  ENV: "production"
};

const ADMIN_SESSION_CACHE_TTL_SECONDS = 6 * 60 * 60;
const ADMIN_SESSION_CACHE_PREFIX = "admin_session_cache_";
const ADMIN_SESSION_PROPERTY_PREFIX = "admin_session_";
const ADMIN_SESSION_DURATION_MS = 12 * 60 * 60 * 1000;
const PUBLIC_CACHE_STATE_PROPERTY = "public_cache_state_v1";
const PUBLIC_CACHE_SCOPES = ["settings", "catalog", "pages", "dashboard"];
const PRODUCT_DESC_MAX_LENGTH = 280;

const HERBAL_ENTITY_SCHEMAS = {
  product_categories: ["id", "name", "slug", "description", "icon", "sort_order", "is_active", "created_at", "updated_at"],
  products: [
    "id", "category_id", "name", "slug", "short_description", "description", "ingredients", "benefits", "usage_instructions",
    "dosage", "caution_notes", "form_type", "packaging_size", "certification_info", "sku", "price", "discount_price", "stock",
    "weight", "is_featured", "is_best_seller", "is_active", "seo_title", "seo_description", "created_at", "updated_at"
  ],
  product_images: ["id", "product_id", "image_url", "alt_text", "sort_order", "created_at"],
  health_tags: ["id", "name", "slug", "description", "created_at", "updated_at"],
  product_health_tags: ["id", "product_id", "health_tag_id"],
  articles: ["id", "title", "slug", "excerpt", "content", "featured_image", "author_name", "category", "tags", "is_published", "published_at", "seo_title", "seo_description", "created_at", "updated_at"],
  testimonials: ["id", "customer_name", "city", "rating", "review_text", "related_product_id", "is_featured", "is_active", "created_at", "updated_at"],
  inquiries: ["id", "name", "phone", "email", "subject", "message", "inquiry_type", "related_product_id", "status", "created_at", "updated_at"],
  banners: ["id", "title", "subtitle", "image_url", "button_label", "button_link", "page_key", "sort_order", "is_active", "created_at", "updated_at"],
  faq_items: ["id", "question", "answer", "sort_order", "is_active", "created_at", "updated_at"],
  promo_campaigns: ["id", "title", "slug", "description", "discount_type", "discount_value", "start_date", "end_date", "is_active", "created_at", "updated_at"],
  course_categories: ["id", "name", "slug", "description", "icon", "sort_order", "is_active", "created_at", "updated_at"],
  courses: [
    "id", "category_id", "title", "slug", "short_description", "full_description", "learning_outcomes", "benefits",
    "target_participants", "level", "duration_text", "total_sessions", "format_type", "requirements", "facilities_included",
    "certificate_available", "price", "promo_price", "thumbnail", "featured", "published", "seo_title", "seo_description",
    "created_at", "updated_at"
  ],
  instructors: ["id", "full_name", "slug", "photo", "specialization", "bio", "experience_years", "certifications", "featured", "is_active", "created_at", "updated_at"],
  branches: ["id", "name", "city", "address", "maps_url", "phone", "email", "operating_hours", "is_active", "created_at", "updated_at"],
  schedules: ["id", "course_id", "branch_id", "instructor_id", "batch_code", "start_date", "end_date", "class_days", "class_time", "quota", "booked_seats", "available_seats", "registration_deadline", "status", "notes", "created_at", "updated_at"],
  registrations: ["id", "course_id", "schedule_id", "full_name", "phone", "email", "age", "occupation_or_background", "preferred_branch_id", "notes", "lead_source", "status", "assigned_to", "follow_up_notes", "created_at", "updated_at"],
  galleries: ["id", "title", "media_type", "media_url", "category", "related_course_id", "related_branch_id", "is_featured", "sort_order", "is_published", "created_at", "updated_at"],
  content_blocks: ["id", "page_slug", "block_key", "title", "subtitle", "content", "image_url", "cta_label", "cta_url", "sort_order", "is_published", "seo_title", "seo_description", "created_at", "updated_at"],
  admin_users: ["id", "name", "email", "password_hash", "role", "branch_id", "is_active", "last_login_at", "created_at", "updated_at"]
};

const HERBAL_ACTION_ALIASES = {
  get_product_categories: "get_product_categories",
  get_products_public: "get_products_public",
  get_product_by_slug: "get_product_by_slug",
  get_articles: "get_articles",
  get_article_by_slug: "get_article_by_slug",
  get_testimonials: "get_testimonials",
  get_banners: "get_banners",
  get_faqs: "get_faqs",
  get_public_settings: "get_public_settings",
  create_inquiry: "create_inquiry",
  get_admin_dashboard_summary: "get_admin_dashboard_summary",
  get_admin_products_v2: "get_admin_products_v2",
  save_admin_product_v2: "save_admin_product_v2",
  delete_admin_product_v2: "delete_admin_product_v2",
  get_admin_categories: "get_admin_categories",
  save_admin_category: "save_admin_category",
  delete_admin_category: "delete_admin_category",
  get_admin_articles_v2: "get_admin_articles_v2",
  save_admin_article_v2: "save_admin_article_v2",
  delete_admin_article_v2: "delete_admin_article_v2",
  get_admin_testimonials_v2: "get_admin_testimonials_v2",
  save_admin_testimonial_v2: "save_admin_testimonial_v2",
  delete_admin_testimonial_v2: "delete_admin_testimonial_v2",
  get_admin_inquiries_v2: "get_admin_inquiries_v2",
  update_admin_inquiry_v2: "update_admin_inquiry_v2",
  get_admin_banners_v2: "get_admin_banners_v2",
  save_admin_banner_v2: "save_admin_banner_v2",
  delete_admin_banner_v2: "delete_admin_banner_v2",
  get_admin_faqs_v2: "get_admin_faqs_v2",
  save_admin_faq_v2: "save_admin_faq_v2",
  delete_admin_faq_v2: "delete_admin_faq_v2",
  seed_herbal_demo_data: "seed_herbal_demo_data",

  // Kursus Offline aliases (public)
  get_course_categories: "get_course_categories",
  get_courses_public: "get_courses_public",
  get_courses: "get_courses_public",
  get_course_public: "get_course_public",
  get_course_by_slug: "get_course_by_slug",
  get_schedules_public: "get_schedules_public",
  get_schedules: "get_schedules_public",
  get_schedule_public: "get_schedule_public",
  get_instructors_public: "get_instructors_public",
  get_instructors: "get_instructors_public",
  get_branches_public: "get_branches_public",
  get_branches: "get_branches_public",
  get_testimonials_public: "get_testimonials_public",
  get_galleries_public: "get_galleries_public",
  get_faqs_public: "get_faqs_public",
  get_settings_public: "get_settings_public",
  create_registration: "create_registration",

  // Kursus Offline aliases (admin)
  get_admin_courses_v2: "get_admin_courses_v2",
  save_admin_course_v2: "save_admin_course_v2",
  delete_admin_course_v2: "delete_admin_course_v2",
  get_admin_course_categories_v2: "get_admin_course_categories_v2",
  save_admin_course_category_v2: "save_admin_course_category_v2",
  delete_admin_course_category_v2: "delete_admin_course_category_v2",
  get_admin_schedules_v2: "get_admin_schedules_v2",
  save_admin_schedule_v2: "save_admin_schedule_v2",
  delete_admin_schedule_v2: "delete_admin_schedule_v2",
  get_admin_instructors_v2: "get_admin_instructors_v2",
  save_admin_instructor_v2: "save_admin_instructor_v2",
  delete_admin_instructor_v2: "delete_admin_instructor_v2",
  get_admin_branches_v2: "get_admin_branches_v2",
  save_admin_branch_v2: "save_admin_branch_v2",
  delete_admin_branch_v2: "delete_admin_branch_v2",
  get_admin_registrations_v2: "get_admin_registrations_v2",
  save_admin_registration_v2: "save_admin_registration_v2",
  patch_admin_registration_status_v2: "patch_admin_registration_status_v2",
  get_admin_galleries_v2: "get_admin_galleries_v2",
  save_admin_gallery_v2: "save_admin_gallery_v2",
  delete_admin_gallery_v2: "delete_admin_gallery_v2",
  get_admin_content_blocks_v2: "get_admin_content_blocks_v2",
  save_admin_content_block_v2: "save_admin_content_block_v2",
  delete_admin_content_block_v2: "delete_admin_content_block_v2",
  get_admin_users_v2: "get_admin_users_v2",
  save_admin_user_v2: "save_admin_user_v2",
  delete_admin_user_v2: "delete_admin_user_v2",
  seed_course_demo_data: "seed_course_demo_data"
};

function getScriptConfig(key) {
  try {
    const p = PropertiesService.getScriptProperties();
    const v = p.getProperty(String(key || ""));
    if (v !== null && v !== undefined && String(v) !== "") return String(v);
  } catch (e) {}
  return SCRIPT_CONFIG[key] || "";
}

function testConfiguration() {
  const url = getScriptConfig("SCRIPT_URL");
  return { status: "success", script_url_configured: !!url };
}

/* =========================
   UTIL / HARDENING HELPERS
========================= */
function jsonRes(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  return ContentService.createTextOutput("System API Ready!")
    .setMimeType(ContentService.MimeType.TEXT);
}

function createDefaultPublicCacheState_() {
  const now = Date.now();
  return {
    settings: now,
    catalog: now,
    pages: now,
    dashboard: now,
    last_updated: now
  };
}

function normalizePublicCacheState_(state) {
  const source = state && typeof state === "object" ? state : {};
  const fallback = createDefaultPublicCacheState_();
  const normalized = {};
  PUBLIC_CACHE_SCOPES.forEach(function(scope) {
    const value = Number(source[scope] || 0);
    normalized[scope] = value > 0 ? value : fallback[scope];
  });
  normalized.last_updated = Math.max.apply(null, PUBLIC_CACHE_SCOPES.map(function(scope) {
    return Number(normalized[scope] || 0);
  }));
  return normalized;
}

function readPublicCacheState_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(PUBLIC_CACHE_STATE_PROPERTY);
    if (!raw) {
      const seeded = normalizePublicCacheState_(null);
      props.setProperty(PUBLIC_CACHE_STATE_PROPERTY, JSON.stringify(seeded));
      return seeded;
    }
    return normalizePublicCacheState_(JSON.parse(raw));
  } catch (e) {
    const fallback = normalizePublicCacheState_(null);
    try {
      PropertiesService.getScriptProperties().setProperty(PUBLIC_CACHE_STATE_PROPERTY, JSON.stringify(fallback));
    } catch (err) {}
    return fallback;
  }
}

function writePublicCacheState_(state) {
  const normalized = normalizePublicCacheState_(state);
  PropertiesService.getScriptProperties().setProperty(PUBLIC_CACHE_STATE_PROPERTY, JSON.stringify(normalized));
  return normalized;
}

function bumpPublicCacheState_(scopes) {
  const validScopes = Array.isArray(scopes)
    ? scopes.map(function(scope) { return String(scope || "").trim().toLowerCase(); }).filter(function(scope, index, arr) {
        return PUBLIC_CACHE_SCOPES.indexOf(scope) !== -1 && arr.indexOf(scope) === index;
      })
    : [];
  if (!validScopes.length) return readPublicCacheState_();

  const next = readPublicCacheState_();
  let seed = Date.now();
  validScopes.forEach(function(scope, index) {
    const previous = Number(next[scope] || 0);
    next[scope] = Math.max(seed + index, previous + 1);
  });
  next.last_updated = Math.max.apply(null, PUBLIC_CACHE_SCOPES.map(function(scope) {
    return Number(next[scope] || 0);
  }));
  return writePublicCacheState_(next);
}

function publicCacheVersionToken_(scope, state) {
  const key = String(scope || "").trim().toLowerCase();
  const source = state && typeof state === "object" ? normalizePublicCacheState_(state) : readPublicCacheState_();
  if (PUBLIC_CACHE_SCOPES.indexOf(key) === -1) return "0";
  return String(Number(source[key] || 0));
}

function withPublicCacheVersion_(payload, scope) {
  const target = payload && typeof payload === "object" ? payload : {};
  target.cache_version = publicCacheVersionToken_(scope);
  return target;
}

function withPublicCacheState_(payload, state) {
  const target = payload && typeof payload === "object" ? payload : {};
  target.cache_state = state && typeof state === "object" ? normalizePublicCacheState_(state) : readPublicCacheState_();
  return target;
}

function getPublicCacheState() {
  return {
    status: "success",
    data: readPublicCacheState_()
  };
}

// CACHING WRAPPER
function getCachedData_(key, fetcherFn, expirationInSeconds = 600) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached) {
    return JSON.parse(cached);
  }
  const data = fetcherFn();
  if (data) {
    try {
      cache.put(key, JSON.stringify(data), expirationInSeconds);
    } catch (e) {
      // Data might be too large for cache (100KB limit)
      console.error("Cache Put Error for " + key + ": " + e.toString());
    }
  }
  return data;
}

function getSettingsMap_() {
  return getCachedData_("settings_map", () => {
    const s = ss.getSheetByName("Settings");
    if (!s) return {};
    const d = s.getDataRange().getValues();
    const map = {};
    for (let i = 1; i < d.length; i++) {
      const k = String(d[i][0] || "").trim();
      if (k) map[k] = d[i][1];
    }
    return map;
  }, 1800); // Cache for 30 minutes
}
function getCfgFrom_(cfg, name) {
  return (cfg && cfg[name] !== undefined && cfg[name] !== null) ? cfg[name] : "";
}
function mustSheet_(name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet "${name}" tidak ditemukan`);
  return sh;
}
function toNumberSafe_(v) {
  const n = Number(String(v ?? "").replace(/[^\d]/g, ""));
  return isFinite(n) ? n : 0;
}
function toISODate_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function normalizePlainText_(value) {
  return String(value === null || value === undefined ? "" : value)
    .replace(/\r\n?/g, "\n")
    .replace(/\u00A0/g, " ")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function containsHtmlMarkup_(value) {
  return /<\s*\/?\s*[a-z][^>]*>/i.test(String(value === null || value === undefined ? "" : value));
}

function normalizeProductDescription_(value) {
  return normalizePlainText_(value);
}

function validateProductDescription_(value) {
  const raw = String(value === null || value === undefined ? "" : value);
  const normalized = normalizeProductDescription_(raw);
  const errors = [];
  if (containsHtmlMarkup_(raw)) {
    errors.push("Deskripsi singkat produk tidak boleh mengandung tag HTML.");
  }
  if (normalized.length > PRODUCT_DESC_MAX_LENGTH) {
    errors.push("Deskripsi singkat produk maksimal " + PRODUCT_DESC_MAX_LENGTH + " karakter.");
  }
  return {
    value: normalized,
    errors: errors
  };
}

function normalizeProductRow_(row) {
  const next = Array.isArray(row) ? row.slice() : [];
  if (next.length > 1) next[1] = normalizePlainText_(next[1]);
  if (next.length > 2) next[2] = normalizeProductDescription_(next[2]);
  return next;
}

function nowIsoDateTime_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function toSlug_(value) {
  const base = String(value || "").trim().toLowerCase();
  if (!base) return "";
  return base
    .replace(/[^a-z0-9\s-]/g, " ")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "");
}

function parseBooleanLike_(value, fallback) {
  if (value === true || value === false) return value;
  const normalized = String(value === null || value === undefined ? "" : value).trim().toLowerCase();
  if (!normalized) return !!fallback;
  return normalized === "1" || normalized === "true" || normalized === "yes" || normalized === "active" || normalized === "published";
}

function parseNumberSafe_(value, fallback) {
  if (value === null || value === undefined || value === "") return Number(fallback || 0);
  const n = Number(value);
  return isFinite(n) ? n : Number(fallback || 0);
}

function parseCsvList_(value) {
  if (Array.isArray(value)) {
    return value.map(function(item) { return normalizePlainText_(item); }).filter(Boolean);
  }
  return String(value || "")
    .split(",")
    .map(function(item) { return normalizePlainText_(item); })
    .filter(Boolean);
}

function ensureSheetWithHeaders_(sheetName, headers) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    return sh;
  }

  const range = sh.getDataRange();
  if (range.getNumRows() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    return sh;
  }

  const existingHeaders = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  let appended = false;
  headers.forEach(function(header) {
    if (existingHeaders.indexOf(header) === -1) {
      sh.getRange(1, existingHeaders.length + 1).setValue(header);
      existingHeaders.push(header);
      appended = true;
    }
  });
  if (appended) sh.setFrozenRows(1);
  return sh;
}

function getSheetHeaderIndexMap_(sh) {
  const headerRow = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  const map = {};
  headerRow.forEach(function(header, index) {
    const key = String(header || "").trim();
    if (key) map[key] = index;
  });
  return map;
}

function readSheetAsObjects_(sheetName, headers) {
  const sh = ensureSheetWithHeaders_(sheetName, headers);
  const data = sh.getDataRange().getValues();
  const indexMap = getSheetHeaderIndexMap_(sh);
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach(function(header) {
      const idx = indexMap[header];
      row[header] = idx !== undefined ? data[i][idx] : "";
    });
    rows.push(row);
  }
  return rows;
}

function upsertSheetObject_(sheetName, headers, idKey, payload) {
  const sh = ensureSheetWithHeaders_(sheetName, headers);
  const indexMap = getSheetHeaderIndexMap_(sh);
  const all = sh.getDataRange().getValues();
  const now = nowIsoDateTime_();
  const next = Object.assign({}, payload || {});
  if (!next[idKey]) next[idKey] = "ID-" + Date.now();
  if (headers.indexOf("created_at") !== -1 && !next.created_at) next.created_at = now;
  if (headers.indexOf("updated_at") !== -1) next.updated_at = now;

  const rowValues = headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(next, header) ? next[header] : "";
  });

  let updated = false;
  const idColumnIndex = indexMap[idKey];
  if (idColumnIndex !== undefined) {
    for (let i = 1; i < all.length; i++) {
      if (String(all[i][idColumnIndex]).trim() === String(next[idKey]).trim()) {
        sh.getRange(i + 1, 1, 1, headers.length).setValues([rowValues]);
        updated = true;
        break;
      }
    }
  }

  if (!updated) sh.appendRow(rowValues);
  return next;
}

function deleteSheetObjectById_(sheetName, idKey, idValue) {
  const headers = HERBAL_ENTITY_SCHEMAS[sheetName];
  const sh = ensureSheetWithHeaders_(sheetName, headers || [idKey]);
  const all = sh.getDataRange().getValues();
  const indexMap = getSheetHeaderIndexMap_(sh);
  const idColumnIndex = indexMap[idKey];
  if (idColumnIndex === undefined) return false;
  const target = String(idValue || "").trim();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][idColumnIndex]).trim() === target) {
      sh.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function normalizeActionAlias_(actionName) {
  const normalized = String(actionName || "").trim();
  return HERBAL_ACTION_ALIASES[normalized] || normalized;
}

function mapLegacyRuleToHerbalProduct_(row) {
  const id = String(row[0] || "").trim();
  const name = normalizePlainText_(row[1] || "");
  const slug = toSlug_(name || id || ("produk-" + Date.now()));
  return {
    id: id,
    category_id: "",
    name: name,
    slug: slug,
    short_description: normalizeProductDescription_(row[2] || ""),
    description: normalizeProductDescription_(row[2] || ""),
    ingredients: "",
    benefits: "",
    usage_instructions: "Ikuti petunjuk penggunaan pada label produk.",
    dosage: "Sesuai rekomendasi pada kemasan.",
    caution_notes: "Bukan pengganti saran medis profesional.",
    form_type: "",
    packaging_size: "",
    certification_info: "",
    sku: id,
    price: parseNumberSafe_(row[4], 0),
    discount_price: "",
    stock: 0,
    weight: "",
    is_featured: false,
    is_best_seller: false,
    is_active: String(row[5] || "").trim().toLowerCase() === "active",
    seo_title: name,
    seo_description: normalizeProductDescription_(row[2] || ""),
    created_at: "",
    updated_at: "",
    image_url: String(row[7] || ""),
    legacy_lp_url: String(row[6] || ""),
    legacy_access_url: String(row[3] || "")
  };
}

function readHerbalProducts_(cfg) {
  const headers = HERBAL_ENTITY_SCHEMAS.products;
  const sheetRows = readSheetAsObjects_("products", headers);
  if (sheetRows.length) {
    return sheetRows.map(function(row) {
      return {
        id: String(row.id || "").trim(),
        category_id: String(row.category_id || "").trim(),
        name: normalizePlainText_(row.name || ""),
        slug: toSlug_(row.slug || row.name || row.id),
        short_description: normalizeProductDescription_(row.short_description || ""),
        description: normalizeProductDescription_(row.description || ""),
        ingredients: String(row.ingredients || ""),
        benefits: String(row.benefits || ""),
        usage_instructions: String(row.usage_instructions || ""),
        dosage: String(row.dosage || ""),
        caution_notes: String(row.caution_notes || ""),
        form_type: String(row.form_type || ""),
        packaging_size: String(row.packaging_size || ""),
        certification_info: String(row.certification_info || ""),
        sku: String(row.sku || ""),
        price: parseNumberSafe_(row.price, 0),
        discount_price: row.discount_price === "" ? "" : parseNumberSafe_(row.discount_price, 0),
        stock: parseNumberSafe_(row.stock, 0),
        weight: String(row.weight || ""),
        is_featured: parseBooleanLike_(row.is_featured, false),
        is_best_seller: parseBooleanLike_(row.is_best_seller, false),
        is_active: parseBooleanLike_(row.is_active, true),
        seo_title: String(row.seo_title || ""),
        seo_description: String(row.seo_description || ""),
        created_at: String(row.created_at || ""),
        updated_at: String(row.updated_at || "")
      };
    });
  }

  const rules = mustSheet_("Access_Rules").getDataRange().getValues();
  const mapped = [];
  for (let i = 1; i < rules.length; i++) mapped.push(mapLegacyRuleToHerbalProduct_(rules[i]));
  return mapped;
}

function readHerbalProductImages_() {
  const headers = HERBAL_ENTITY_SCHEMAS.product_images;
  return readSheetAsObjects_("product_images", headers).map(function(row) {
    return {
      id: String(row.id || "").trim(),
      product_id: String(row.product_id || "").trim(),
      image_url: String(row.image_url || "").trim(),
      alt_text: String(row.alt_text || "").trim(),
      sort_order: parseNumberSafe_(row.sort_order, 0)
    };
  });
}

function readHerbalCategories_() {
  const headers = HERBAL_ENTITY_SCHEMAS.product_categories;
  return readSheetAsObjects_("product_categories", headers).map(function(row) {
    return {
      id: String(row.id || "").trim(),
      name: String(row.name || "").trim(),
      slug: toSlug_(row.slug || row.name),
      description: String(row.description || "").trim(),
      icon: String(row.icon || "").trim(),
      sort_order: parseNumberSafe_(row.sort_order, 0),
      is_active: parseBooleanLike_(row.is_active, true)
    };
  });
}

function readHerbalHealthTags_() {
  const headers = HERBAL_ENTITY_SCHEMAS.health_tags;
  return readSheetAsObjects_("health_tags", headers).map(function(row) {
    return {
      id: String(row.id || "").trim(),
      name: String(row.name || "").trim(),
      slug: toSlug_(row.slug || row.name),
      description: String(row.description || "").trim()
    };
  });
}

function readProductHealthTagLinks_() {
  const headers = HERBAL_ENTITY_SCHEMAS.product_health_tags;
  return readSheetAsObjects_("product_health_tags", headers).map(function(row) {
    return {
      id: String(row.id || "").trim(),
      product_id: String(row.product_id || "").trim(),
      health_tag_id: String(row.health_tag_id || "").trim()
    };
  });
}

function hydrateHerbalProduct_(product, categories, images, tags, links) {
  const category = categories.find(function(item) { return item.id === product.category_id; }) || null;
  const productImages = images
    .filter(function(item) { return item.product_id === product.id && item.image_url; })
    .sort(function(a, b) { return a.sort_order - b.sort_order; });
  const linkedTagIds = links.filter(function(item) { return item.product_id === product.id; }).map(function(item) { return item.health_tag_id; });
  const linkedTags = tags.filter(function(tag) { return linkedTagIds.indexOf(tag.id) !== -1; });
  return Object.assign({}, product, {
    category: category,
    health_tags: linkedTags,
    health_tag_slugs: linkedTags.map(function(tag) { return tag.slug; }),
    images: productImages,
    image_url: productImages.length ? productImages[0].image_url : ""
  });
}

function getHerbalProductsPublic(d, cfg) {
  cfg = cfg || getSettingsMap_();
  const categories = readHerbalCategories_();
  const images = readHerbalProductImages_();
  const tags = readHerbalHealthTags_();
  const links = readProductHealthTagLinks_();
  const products = readHerbalProducts_(cfg).map(function(product) {
    return hydrateHerbalProduct_(product, categories, images, tags, links);
  });

  const query = String(d.query || d.search || "").trim().toLowerCase();
  const categorySlug = String(d.category || "").trim().toLowerCase();
  const healthSlug = String(d.health_tag || d.benefit || "").trim().toLowerCase();
  const formType = String(d.form_type || "").trim().toLowerCase();
  const inStock = String(d.in_stock || "").trim().toLowerCase();
  const minPrice = parseNumberSafe_(d.min_price, 0);
  const maxPriceRaw = String(d.max_price || "").trim();
  const maxPrice = maxPriceRaw ? parseNumberSafe_(d.max_price, 0) : null;

  let filtered = products.filter(function(item) {
    if (!item.is_active) return false;
    if (query) {
      const hay = (item.name + " " + item.short_description + " " + item.description).toLowerCase();
      if (hay.indexOf(query) === -1) return false;
    }
    if (categorySlug) {
      const cSlug = item.category && item.category.slug ? String(item.category.slug).toLowerCase() : "";
      if (cSlug !== categorySlug) return false;
    }
    if (healthSlug && item.health_tag_slugs.indexOf(healthSlug) === -1) return false;
    if (formType && String(item.form_type || "").toLowerCase() !== formType) return false;
    if (item.price < minPrice) return false;
    if (maxPrice !== null && item.price > maxPrice) return false;
    if (inStock === "true" || inStock === "1") {
      if (parseNumberSafe_(item.stock, 0) <= 0) return false;
    }
    return true;
  });

  const sort = String(d.sort || "newest").trim().toLowerCase();
  if (sort === "price") filtered.sort(function(a, b) { return a.price - b.price; });
  else if (sort === "price_desc") filtered.sort(function(a, b) { return b.price - a.price; });
  else if (sort === "best_selling") filtered.sort(function(a, b) { return Number(b.is_best_seller) - Number(a.is_best_seller); });
  else filtered.sort(function(a, b) { return String(b.updated_at || "").localeCompare(String(a.updated_at || "")); });

  return withPublicCacheVersion_({ status: "success", data: filtered, total: filtered.length }, "catalog");
}

function getHerbalProductBySlug(d, cfg) {
  cfg = cfg || getSettingsMap_();
  const slug = toSlug_(d.slug || "");
  if (!slug) return { status: "error", message: "Slug produk wajib diisi." };
  const productsResult = getHerbalProductsPublic({}, cfg);
  const list = (productsResult && productsResult.data) ? productsResult.data : [];
  const found = list.find(function(item) { return item.slug === slug; });
  if (!found) return { status: "error", message: "Produk herbal tidak ditemukan." };

  const faqs = getHerbalFaqs({ page_key: "product" }, cfg);
  const related = list.filter(function(item) {
    return item.id !== found.id && item.category_id && found.category_id && item.category_id === found.category_id;
  }).slice(0, 4);

  return withPublicCacheVersion_({
    status: "success",
    data: Object.assign({}, found, {
      related_products: related,
      faq: (faqs && faqs.data) ? faqs.data.slice(0, 6) : []
    })
  }, "catalog");
}

function getHerbalCategories(d, cfg) {
  cfg = cfg || getSettingsMap_();
  const categories = readHerbalCategories_().filter(function(item) { return item.is_active; })
    .sort(function(a, b) { return a.sort_order - b.sort_order; });
  return withPublicCacheVersion_({ status: "success", data: categories, total: categories.length }, "catalog");
}

function getHerbalArticles(d, cfg) {
  cfg = cfg || getSettingsMap_();
  const headers = HERBAL_ENTITY_SCHEMAS.articles;
  const rows = readSheetAsObjects_("articles", headers)
    .filter(function(row) { return parseBooleanLike_(row.is_published, true); })
    .map(function(row) {
      return {
        id: String(row.id || "").trim(),
        title: String(row.title || "").trim(),
        slug: toSlug_(row.slug || row.title),
        excerpt: String(row.excerpt || "").trim(),
        content: String(row.content || "").trim(),
        featured_image: String(row.featured_image || "").trim(),
        author_name: String(row.author_name || "Admin"),
        category: String(row.category || "Edukasi Herbal"),
        tags: parseCsvList_(row.tags),
        published_at: String(row.published_at || ""),
        seo_title: String(row.seo_title || ""),
        seo_description: String(row.seo_description || "")
      };
    })
    .sort(function(a, b) { return String(b.published_at || "").localeCompare(String(a.published_at || "")); });
  return withPublicCacheVersion_({ status: "success", data: rows, total: rows.length }, "pages");
}

function getHerbalArticleBySlug(d, cfg) {
  cfg = cfg || getSettingsMap_();
  const slug = toSlug_(d.slug || "");
  if (!slug) return { status: "error", message: "Slug artikel wajib diisi." };
  const articles = getHerbalArticles({}, cfg).data || [];
  const found = articles.find(function(item) { return item.slug === slug; });
  if (!found) return { status: "error", message: "Artikel tidak ditemukan." };
  return withPublicCacheVersion_({ status: "success", data: found }, "pages");
}

function getHerbalTestimonials(d, cfg) {
  cfg = cfg || getSettingsMap_();
  const headers = HERBAL_ENTITY_SCHEMAS.testimonials;
  const rows = readSheetAsObjects_("testimonials", headers)
    .filter(function(row) { return parseBooleanLike_(row.is_active, true); })
    .map(function(row) {
      return {
        id: String(row.id || ""),
        customer_name: String(row.customer_name || "Pelanggan"),
        city: String(row.city || ""),
        rating: parseNumberSafe_(row.rating, 5),
        review_text: String(row.review_text || ""),
        related_product_id: String(row.related_product_id || ""),
        is_featured: parseBooleanLike_(row.is_featured, false)
      };
    })
    .sort(function(a, b) { return Number(b.is_featured) - Number(a.is_featured); });
  return withPublicCacheVersion_({ status: "success", data: rows, total: rows.length }, "pages");
}

function getHerbalBanners(d, cfg) {
  cfg = cfg || getSettingsMap_();
  const pageKey = String((d && d.page_key) || "home").trim().toLowerCase();
  const headers = HERBAL_ENTITY_SCHEMAS.banners;
  const rows = readSheetAsObjects_("banners", headers)
    .filter(function(row) {
      if (!parseBooleanLike_(row.is_active, true)) return false;
      const target = String(row.page_key || "home").trim().toLowerCase();
      return target === pageKey;
    })
    .map(function(row) {
      return {
        id: String(row.id || ""),
        title: String(row.title || ""),
        subtitle: String(row.subtitle || ""),
        image_url: String(row.image_url || ""),
        button_label: String(row.button_label || ""),
        button_link: String(row.button_link || ""),
        page_key: String(row.page_key || "home"),
        sort_order: parseNumberSafe_(row.sort_order, 0)
      };
    })
    .sort(function(a, b) { return a.sort_order - b.sort_order; });
  return withPublicCacheVersion_({ status: "success", data: rows, total: rows.length }, "pages");
}

function getHerbalFaqs(d, cfg) {
  cfg = cfg || getSettingsMap_();
  const headers = HERBAL_ENTITY_SCHEMAS.faq_items;
  const rows = readSheetAsObjects_("faq_items", headers)
    .filter(function(row) { return parseBooleanLike_(row.is_active, true); })
    .map(function(row) {
      return {
        id: String(row.id || ""),
        question: String(row.question || ""),
        answer: String(row.answer || ""),
        sort_order: parseNumberSafe_(row.sort_order, 0)
      };
    })
    .sort(function(a, b) { return a.sort_order - b.sort_order; });
  return withPublicCacheVersion_({ status: "success", data: rows, total: rows.length }, "pages");
}

function getHerbalPublicSettings(cfg) {
  cfg = cfg || getSettingsMap_();
  return withPublicCacheVersion_({
    status: "success",
    data: {
      site_name: getCfgFrom_(cfg, "site_name") || "NaturaHerb Wellness",
      site_tagline: getCfgFrom_(cfg, "site_tagline") || "Herbal alami untuk dukungan gaya hidup sehat harian.",
      contact_email: getCfgFrom_(cfg, "contact_email") || "",
      wa_admin: getCfgFrom_(cfg, "wa_admin") || "",
      legal_disclaimer: getCfgFrom_(cfg, "legal_disclaimer") || "Produk herbal ini bukan pengganti diagnosis atau terapi medis profesional.",
      certifications: parseCsvList_(getCfgFrom_(cfg, "certifications") || "BPOM, Halal, GMP"),
      social_links: {
        instagram: getCfgFrom_(cfg, "social_instagram") || "",
        facebook: getCfgFrom_(cfg, "social_facebook") || "",
        tiktok: getCfgFrom_(cfg, "social_tiktok") || ""
      }
    }
  }, "settings");
}

function createInquiry(d, cfg) {
  cfg = cfg || getSettingsMap_();
  const name = normalizePlainText_(d.name || "");
  const phone = normalizePlainText_(d.phone || "");
  const email = normalizePlainText_(d.email || "");
  const message = normalizePlainText_(d.message || "");
  const subject = normalizePlainText_(d.subject || "Informasi Produk Herbal");
  const inquiryType = normalizePlainText_(d.inquiry_type || "general");
  if (!name || !message) return { status: "error", message: "Nama dan pesan wajib diisi." };

  const payload = {
    id: "INQ-" + Date.now(),
    name: name,
    phone: phone,
    email: email,
    subject: subject,
    message: message,
    inquiry_type: inquiryType,
    related_product_id: normalizePlainText_(d.related_product_id || ""),
    status: "new"
  };
  upsertSheetObject_("inquiries", HERBAL_ENTITY_SCHEMAS.inquiries, "id", payload);
  return withPublicCacheState_({ status: "success", message: "Pertanyaan Anda telah diterima. Tim kami akan segera menghubungi." }, bumpPublicCacheState_(["dashboard", "pages"]));
}

function ensureCourseDomainSheets_() {
  [
    "course_categories",
    "courses",
    "instructors",
    "branches",
    "schedules",
    "registrations",
    "galleries",
    "content_blocks",
    "admin_users"
  ].forEach(function(entity) {
    ensureSheetWithHeaders_(entity, HERBAL_ENTITY_SCHEMAS[entity]);
  });
}

function mapLegacyProductToCourse_(item) {
  const now = nowIsoDateTime_();
  return {
    id: String(item.id || ""),
    category_id: String(item.category_id || ""),
    title: String(item.name || item.title || "Program Kursus Offline"),
    slug: toSlug_(item.slug || item.name || item.title || "program-kursus"),
    short_description: String(item.short_description || item.desc || ""),
    full_description: String(item.description || ""),
    learning_outcomes: String(item.benefits || ""),
    benefits: String(item.benefits || ""),
    target_participants: "Siswa, mahasiswa, fresh graduate, karyawan, profesional",
    level: String(item.level || "Pemula"),
    duration_text: String(item.duration_text || "8 sesi"),
    total_sessions: parseNumberSafe_(item.total_sessions, 8),
    format_type: "offline",
    requirements: String(item.requirements || "Komitmen hadir sesuai jadwal"),
    facilities_included: String(item.facilities_included || "Modul, ruang kelas, praktik langsung"),
    certificate_available: true,
    price: parseNumberSafe_(item.price, 0),
    promo_price: item.discount_price === "" ? "" : parseNumberSafe_(item.discount_price, 0),
    thumbnail: String(item.image_url || ""),
    featured: parseBooleanLike_(item.is_featured, false),
    published: parseBooleanLike_(item.is_active, true),
    seo_title: String(item.seo_title || item.name || ""),
    seo_description: String(item.seo_description || item.short_description || ""),
    created_at: String(item.created_at || now),
    updated_at: String(item.updated_at || now),
    category_name: item.category && item.category.name ? String(item.category.name) : "",
    category_slug: item.category && item.category.slug ? String(item.category.slug) : "",
    branch_city: String(item.branch_city || item.city || "")
  };
}

function buildCourseCategoryMap_() {
  const rows = readSheetAsObjects_("course_categories", HERBAL_ENTITY_SCHEMAS.course_categories);
  const map = {};
  rows.forEach(function(item) {
    const id = String(item.id || "").trim();
    if (!id) return;
    map[id] = {
      id: id,
      name: String(item.name || ""),
      slug: toSlug_(item.slug || item.name),
      description: String(item.description || ""),
      sort_order: parseNumberSafe_(item.sort_order, 0),
      is_active: parseBooleanLike_(item.is_active, true)
    };
  });
  return map;
}

function getCourseCategories(d, cfg) {
  cfg = cfg || getSettingsMap_();
  ensureCourseDomainSheets_();

  const rows = readSheetAsObjects_("course_categories", HERBAL_ENTITY_SCHEMAS.course_categories)
    .map(function(item) {
      return {
        id: String(item.id || ""),
        name: String(item.name || ""),
        slug: toSlug_(item.slug || item.name),
        description: String(item.description || ""),
        icon: String(item.icon || ""),
        sort_order: parseNumberSafe_(item.sort_order, 0),
        is_active: parseBooleanLike_(item.is_active, true)
      };
    })
    .filter(function(item) { return item.is_active; })
    .sort(function(a, b) { return a.sort_order - b.sort_order; });

  if (rows.length > 0) {
    return withPublicCacheVersion_({ status: "success", data: rows, total: rows.length }, "catalog");
  }

  const fallback = readHerbalCategories_().map(function(item) {
    return {
      id: item.id,
      name: item.name,
      slug: item.slug,
      description: item.description,
      icon: item.icon,
      sort_order: item.sort_order,
      is_active: item.is_active
    };
  });

  return withPublicCacheVersion_({ status: "success", data: fallback, total: fallback.length }, "catalog");
}

function getCoursesPublic(d, cfg) {
  cfg = cfg || getSettingsMap_();
  ensureCourseDomainSheets_();

  const categoryMap = buildCourseCategoryMap_();
  const rows = readSheetAsObjects_("courses", HERBAL_ENTITY_SCHEMAS.courses);
  let courses = rows.map(function(item) {
    const category = categoryMap[String(item.category_id || "")] || null;
    return {
      id: String(item.id || ""),
      category_id: String(item.category_id || ""),
      category_name: category ? String(category.name || "") : "",
      category_slug: category ? String(category.slug || "") : "",
      title: String(item.title || ""),
      slug: toSlug_(item.slug || item.title),
      short_description: String(item.short_description || ""),
      full_description: String(item.full_description || ""),
      learning_outcomes: String(item.learning_outcomes || ""),
      benefits: String(item.benefits || ""),
      target_participants: String(item.target_participants || ""),
      level: String(item.level || "Pemula"),
      duration_text: String(item.duration_text || ""),
      total_sessions: parseNumberSafe_(item.total_sessions, 0),
      format_type: String(item.format_type || "offline"),
      requirements: String(item.requirements || ""),
      facilities_included: String(item.facilities_included || ""),
      certificate_available: parseBooleanLike_(item.certificate_available, true),
      price: parseNumberSafe_(item.price, 0),
      promo_price: item.promo_price === "" ? "" : parseNumberSafe_(item.promo_price, 0),
      image_url: String(item.thumbnail || ""),
      featured: parseBooleanLike_(item.featured, false),
      published: parseBooleanLike_(item.published, true),
      seo_title: String(item.seo_title || ""),
      seo_description: String(item.seo_description || ""),
      created_at: String(item.created_at || ""),
      updated_at: String(item.updated_at || "")
    };
  });

  if (!courses.length) {
    const categories = readHerbalCategories_();
    const images = readHerbalProductImages_();
    const tags = readHerbalHealthTags_();
    const links = readProductHealthTagLinks_();
    const legacy = readHerbalProducts_(cfg).map(function(product) {
      return hydrateHerbalProduct_(product, categories, images, tags, links);
    });
    courses = legacy.map(mapLegacyProductToCourse_);
  }

  const query = String(d.query || d.search || "").trim().toLowerCase();
  const category = String(d.category || "").trim().toLowerCase();
  const level = String(d.level || "").trim().toLowerCase();
  const city = String(d.city || d.branch || "").trim().toLowerCase();
  const minPrice = parseNumberSafe_(d.min_price, 0);
  const maxPriceRaw = String(d.max_price || "").trim();
  const maxPrice = maxPriceRaw ? parseNumberSafe_(d.max_price, 0) : null;

  let filtered = courses.filter(function(item) {
    if (!parseBooleanLike_(item.published, true)) return false;
    if (query) {
      const hay = (String(item.title || "") + " " + String(item.short_description || "") + " " + String(item.full_description || "")).toLowerCase();
      if (hay.indexOf(query) === -1) return false;
    }
    if (category) {
      const cSlug = String(item.category_slug || "").toLowerCase();
      const cName = String(item.category_name || "").toLowerCase();
      if (cSlug !== category && cName !== category) return false;
    }
    if (level && String(item.level || "").toLowerCase() !== level) return false;
    if (city && String(item.branch_city || "").toLowerCase().indexOf(city) === -1) return false;
    if (item.price < minPrice) return false;
    if (maxPrice !== null && item.price > maxPrice) return false;
    return true;
  });

  const sort = String(d.sort || "newest").trim().toLowerCase();
  if (sort === "price") filtered.sort(function(a, b) { return a.price - b.price; });
  else if (sort === "price_desc") filtered.sort(function(a, b) { return b.price - a.price; });
  else if (sort === "popular") filtered.sort(function(a, b) { return Number(b.featured) - Number(a.featured); });
  else filtered.sort(function(a, b) { return String(b.updated_at || "").localeCompare(String(a.updated_at || "")); });

  return withPublicCacheVersion_({ status: "success", data: filtered, total: filtered.length }, "catalog");
}

function getCourseBySlug(d, cfg) {
  cfg = cfg || getSettingsMap_();
  const slug = toSlug_(d.slug || "");
  if (!slug) return { status: "error", message: "Slug kursus wajib diisi." };
  const list = getCoursesPublic({}, cfg).data || [];
  const found = list.find(function(item) { return item.slug === slug; });
  if (!found) return { status: "error", message: "Kursus tidak ditemukan." };
  return withPublicCacheVersion_({ status: "success", data: found }, "catalog");
}

function getCoursePublic(d, cfg) {
  cfg = cfg || getSettingsMap_();
  if (d.slug) return getCourseBySlug(d, cfg);

  const id = String(d.id || d.course_id || "").trim();
  if (!id) return { status: "error", message: "ID kursus wajib diisi." };

  const list = getCoursesPublic({}, cfg).data || [];
  const found = list.find(function(item) { return String(item.id || "") === id; });
  if (found) return withPublicCacheVersion_({ status: "success", data: found }, "catalog");

  const legacy = getProductDetail({ id: id }, cfg);
  if (legacy && legacy.status === "success") {
    const mapped = mapLegacyProductToCourse_(legacy.data || {});
    return withPublicCacheVersion_({ status: "success", data: mapped }, "catalog");
  }
  return { status: "error", message: "Kursus tidak ditemukan." };
}

function getSchedulesPublic(d, cfg) {
  cfg = cfg || getSettingsMap_();
  ensureCourseDomainSheets_();

  const rows = readSheetAsObjects_("schedules", HERBAL_ENTITY_SCHEMAS.schedules);
  const courses = getCoursesPublic({}, cfg).data || [];
  const courseMap = {};
  courses.forEach(function(item) { courseMap[String(item.id || "")] = item; });

  let schedules = rows.map(function(item) {
    const quota = parseNumberSafe_(item.quota, 0);
    const booked = parseNumberSafe_(item.booked_seats, 0);
    const available = item.available_seats === "" ? Math.max(quota - booked, 0) : parseNumberSafe_(item.available_seats, 0);
    const course = courseMap[String(item.course_id || "")] || null;
    return {
      id: String(item.id || ""),
      course_id: String(item.course_id || ""),
      course_name: course ? String(course.title || "") : "",
      category: course ? String(course.category_name || "") : "",
      level: course ? String(course.level || "") : "",
      branch_id: String(item.branch_id || ""),
      location: String(item.branch_id || ""),
      instructor_id: String(item.instructor_id || ""),
      batch_code: String(item.batch_code || ""),
      start_date: String(item.start_date || ""),
      end_date: String(item.end_date || ""),
      class_days: String(item.class_days || ""),
      class_time: String(item.class_time || ""),
      quota: quota,
      booked_seats: booked,
      available_seats: available,
      registration_deadline: String(item.registration_deadline || ""),
      status: String(item.status || "open").toLowerCase(),
      notes: String(item.notes || "")
    };
  });

  if (!schedules.length) {
    schedules = courses.slice(0, 12).map(function(item, idx) {
      return {
        id: "SCH-AUTO-" + (idx + 1),
        course_id: item.id,
        course_name: item.title,
        category: item.category_name,
        level: item.level,
        branch_id: "",
        location: "Cabang tersedia",
        instructor_id: "",
        batch_code: "BATCH-" + String(idx + 1).padStart(2, "0"),
        start_date: "",
        end_date: "",
        class_days: idx % 2 === 0 ? "Weekday" : "Weekend",
        class_time: idx % 2 === 0 ? "19:00 - 21:00" : "09:00 - 12:00",
        quota: 20,
        booked_seats: 0,
        available_seats: 20,
        registration_deadline: "",
        status: "open",
        notes: ""
      };
    });
  }

  const categoryFilter = String(d.category || "").trim().toLowerCase();
  const levelFilter = String(d.level || "").trim().toLowerCase();
  const seatOnly = String(d.seat_available || d.available_only || "").trim().toLowerCase();
  const mode = String(d.mode || d.class_mode || "").trim().toLowerCase();

  let filtered = schedules.filter(function(item) {
    if (categoryFilter && String(item.category || "").toLowerCase() !== categoryFilter) return false;
    if (levelFilter && String(item.level || "").toLowerCase() !== levelFilter) return false;
    if ((seatOnly === "1" || seatOnly === "true") && parseNumberSafe_(item.available_seats, 0) <= 0) return false;
    if (mode === "weekend" && String(item.class_days || "").toLowerCase().indexOf("weekend") === -1) return false;
    if (mode === "weekday" && String(item.class_days || "").toLowerCase().indexOf("weekday") === -1) return false;
    return true;
  });

  filtered.sort(function(a, b) { return String(a.start_date || "").localeCompare(String(b.start_date || "")); });
  return withPublicCacheVersion_({ status: "success", data: filtered, total: filtered.length }, "catalog");
}

function getSchedulePublic(d, cfg) {
  cfg = cfg || getSettingsMap_();
  const id = String(d.id || d.schedule_id || "").trim();
  if (!id) return { status: "error", message: "ID jadwal wajib diisi." };
  const list = getSchedulesPublic({}, cfg).data || [];
  const found = list.find(function(item) { return String(item.id || "") === id; });
  if (!found) return { status: "error", message: "Jadwal tidak ditemukan." };
  return withPublicCacheVersion_({ status: "success", data: found }, "catalog");
}

function getInstructorsPublic(d, cfg) {
  cfg = cfg || getSettingsMap_();
  ensureCourseDomainSheets_();
  const rows = readSheetAsObjects_("instructors", HERBAL_ENTITY_SCHEMAS.instructors)
    .filter(function(item) { return parseBooleanLike_(item.is_active, true); })
    .map(function(item) {
      return {
        id: String(item.id || ""),
        full_name: String(item.full_name || ""),
        slug: toSlug_(item.slug || item.full_name),
        photo: String(item.photo || ""),
        specialization: String(item.specialization || ""),
        experience_years: parseNumberSafe_(item.experience_years, 0),
        certifications: parseCsvList_(item.certifications),
        bio: String(item.bio || ""),
        featured: parseBooleanLike_(item.featured, false)
      };
    });
  return withPublicCacheVersion_({ status: "success", data: rows, total: rows.length }, "pages");
}

function getBranchesPublic(d, cfg) {
  cfg = cfg || getSettingsMap_();
  ensureCourseDomainSheets_();
  let rows = readSheetAsObjects_("branches", HERBAL_ENTITY_SCHEMAS.branches)
    .filter(function(item) { return parseBooleanLike_(item.is_active, true); })
    .map(function(item) {
      return {
        id: String(item.id || ""),
        name: String(item.name || ""),
        city: String(item.city || ""),
        address: String(item.address || ""),
        maps_url: String(item.maps_url || ""),
        phone: String(item.phone || ""),
        email: String(item.email || ""),
        operating_hours: String(item.operating_hours || "")
      };
    });

  if (!rows.length) {
    rows = [
      { id: "BR-JKT", name: "Jakarta Selatan", city: "Jakarta", address: "Jakarta Selatan", maps_url: "", phone: "", email: "", operating_hours: "09:00 - 20:00" },
      { id: "BR-BDG", name: "Bandung", city: "Bandung", address: "Bandung", maps_url: "", phone: "", email: "", operating_hours: "09:00 - 20:00" },
      { id: "BR-SBY", name: "Surabaya", city: "Surabaya", address: "Surabaya", maps_url: "", phone: "", email: "", operating_hours: "09:00 - 20:00" }
    ];
  }
  return withPublicCacheVersion_({ status: "success", data: rows, total: rows.length }, "pages");
}

function getTestimonialsPublic(d, cfg) {
  cfg = cfg || getSettingsMap_();
  const legacy = getHerbalTestimonials(d || {}, cfg);
  const list = (legacy && legacy.data ? legacy.data : []).map(function(item) {
    return {
      id: item.id,
      student_name: item.customer_name || "Peserta",
      city: item.city || "",
      rating: parseNumberSafe_(item.rating, 5),
      review: item.review_text || "",
      joined_course: item.related_product_id || "",
      is_featured: parseBooleanLike_(item.is_featured, false)
    };
  });
  return withPublicCacheVersion_({ status: "success", data: list, total: list.length }, "pages");
}

function getGalleriesPublic(d, cfg) {
  cfg = cfg || getSettingsMap_();
  ensureCourseDomainSheets_();
  const categoryFilter = String(d.category || "").trim().toLowerCase();

  let rows = readSheetAsObjects_("galleries", HERBAL_ENTITY_SCHEMAS.galleries)
    .filter(function(item) { return parseBooleanLike_(item.is_published, true); })
    .map(function(item) {
      return {
        id: String(item.id || ""),
        title: String(item.title || ""),
        media_type: String(item.media_type || "image"),
        media_url: String(item.media_url || ""),
        category: String(item.category || "general"),
        related_course_id: String(item.related_course_id || ""),
        related_branch_id: String(item.related_branch_id || ""),
        is_featured: parseBooleanLike_(item.is_featured, false),
        sort_order: parseNumberSafe_(item.sort_order, 0)
      };
    });

  if (categoryFilter) {
    rows = rows.filter(function(item) {
      return String(item.category || "").toLowerCase() === categoryFilter;
    });
  }

  rows.sort(function(a, b) { return a.sort_order - b.sort_order; });
  return withPublicCacheVersion_({ status: "success", data: rows, total: rows.length }, "pages");
}

function getFaqsPublic(d, cfg) {
  return getHerbalFaqs(d || {}, cfg || getSettingsMap_());
}

function getSettingsPublic(cfg) {
  cfg = cfg || getSettingsMap_();
  const base = getHerbalPublicSettings(cfg);
  const payload = base && base.data ? base.data : {};
  payload.site_name = payload.site_name || "Web Kursus Offline";
  payload.site_tagline = payload.site_tagline || "Pelatihan tatap muka untuk skill praktis dan siap kerja.";
  payload.primary_cta = "Daftar Sekarang";
  payload.secondary_cta = "Lihat Jadwal";
  return withPublicCacheVersion_({ status: "success", data: payload }, "settings");
}

function createRegistration(d, cfg) {
  cfg = cfg || getSettingsMap_();
  ensureCourseDomainSheets_();

  const fullName = normalizePlainText_(d.full_name || d.nama || "");
  const phone = normalizePlainText_(d.phone || d.whatsapp || "");
  const email = normalizePlainText_(d.email || "").toLowerCase();
  const courseId = normalizePlainText_(d.course_id || d.id_produk || "");
  const scheduleId = normalizePlainText_(d.schedule_id || d.selected_schedule || "");
  const background = normalizePlainText_(d.occupation_or_background || d.background || "");
  const preferredBranch = normalizePlainText_(d.preferred_branch_id || d.preferred_branch || "");
  const leadSource = normalizePlainText_(d.lead_source || "");

  if (!fullName || !phone || !email || !scheduleId || !background || !preferredBranch || !leadSource) {
    return { status: "error", message: "Data registrasi wajib belum lengkap." };
  }

  const now = nowIsoDateTime_();
  const regId = "REG-" + Date.now();
  const payload = {
    id: regId,
    course_id: courseId,
    schedule_id: scheduleId,
    full_name: fullName,
    phone: phone,
    email: email,
    age: normalizePlainText_(d.age || ""),
    occupation_or_background: background,
    preferred_branch_id: preferredBranch,
    notes: normalizePlainText_(d.notes || ""),
    lead_source: leadSource,
    status: normalizePlainText_(d.status || "new") || "new",
    assigned_to: normalizePlainText_(d.assigned_to || ""),
    follow_up_notes: normalizePlainText_(d.follow_up_notes || ""),
    created_at: now,
    updated_at: now
  };

  upsertSheetObject_("registrations", HERBAL_ENTITY_SCHEMAS.registrations, "id", payload);

  const estimatedPrice = parseNumberSafe_(d.harga || d.price, 0);
  const waAdmin = getCfgFrom_(cfg, "wa_admin");
  if (waAdmin) {
    const msg = [
      "📥 REGISTRASI BARU",
      "Kode: " + regId,
      "Nama: " + fullName,
      "Program: " + normalizePlainText_(d.nama_produk || d.course_title || "-"),
      "Jadwal: " + scheduleId,
      "Cabang: " + preferredBranch,
      "No. HP: " + phone,
      "Sumber: " + leadSource
    ].join("\n");
    sendWA(waAdmin, msg, cfg);
  }

  return withPublicCacheState_({
    status: "success",
    message: "Registrasi berhasil dikirim.",
    invoice: regId,
    registration_code: regId,
    tagihan: estimatedPrice,
    follow_up_status: "new"
  }, bumpPublicCacheState_(["dashboard", "catalog", "pages"]));
}

function requireAdminAndGetEntityRows_(d, entityName) {
  requireAdminSession_(d, { actionName: "admin_" + entityName });
  const headers = HERBAL_ENTITY_SCHEMAS[entityName];
  if (!headers) throw new Error("Entity tidak didukung: " + entityName);
  return readSheetAsObjects_(entityName, headers);
}

function adminListEntity_(d, entityName) {
  const rows = requireAdminAndGetEntityRows_(d, entityName);
  return { status: "success", data: rows, total: rows.length };
}

function adminSaveEntity_(d, entityName, idKey) {
  requireAdminSession_(d, { actionName: "save_" + entityName });
  const headers = HERBAL_ENTITY_SCHEMAS[entityName];
  if (!headers) return { status: "error", message: "Entity tidak didukung." };
  const payload = Object.assign({}, d && d.payload && typeof d.payload === "object" ? d.payload : d || {});
  const sanitized = {};
  headers.forEach(function(header) {
    sanitized[header] = Object.prototype.hasOwnProperty.call(payload, header) ? payload[header] : "";
  });
  if (!sanitized[idKey]) sanitized[idKey] = String(entityName).substring(0, 3).toUpperCase() + "-" + Date.now();
  const saved = upsertSheetObject_(entityName, headers, idKey, sanitized);
  return withPublicCacheState_({ status: "success", data: saved }, bumpPublicCacheState_(["catalog", "pages", "dashboard", "settings"]));
}

function adminDeleteEntity_(d, entityName, idKey) {
  requireAdminSession_(d, { actionName: "delete_" + entityName });
  const id = String((d && (d.id || d[idKey])) || "").trim();
  if (!id) return { status: "error", message: "ID wajib diisi." };
  const ok = deleteSheetObjectById_(entityName, idKey, id);
  if (!ok) return { status: "error", message: "Data tidak ditemukan." };
  return withPublicCacheState_({ status: "success", message: "Data berhasil dihapus." }, bumpPublicCacheState_(["catalog", "pages", "dashboard", "settings"]));
}

function getAdminDashboardSummary(d, cfg) {
  requireAdminSession_(d, { actionName: "get_admin_dashboard_summary" });
  cfg = cfg || getSettingsMap_();

  ensureCourseDomainSheets_();

  const courses = readSheetAsObjects_("courses", HERBAL_ENTITY_SCHEMAS.courses);
  const schedules = readSheetAsObjects_("schedules", HERBAL_ENTITY_SCHEMAS.schedules);
  const registrations = readSheetAsObjects_("registrations", HERBAL_ENTITY_SCHEMAS.registrations);
  const branches = readSheetAsObjects_("branches", HERBAL_ENTITY_SCHEMAS.branches);

  const legacyProducts = readHerbalProducts_(cfg);
  const legacyCategories = readHerbalCategories_();
  const legacyArticles = readSheetAsObjects_("articles", HERBAL_ENTITY_SCHEMAS.articles);
  const legacyInquiries = readSheetAsObjects_("inquiries", HERBAL_ENTITY_SCHEMAS.inquiries);
  const lowStock = legacyProducts.filter(function(item) { return parseNumberSafe_(item.stock, 0) > 0 && parseNumberSafe_(item.stock, 0) <= 5; });

  const activeSchedules = schedules.filter(function(item) {
    const status = String(item.status || "").toLowerCase();
    return status === "open" || status === "limited" || status === "active";
  });

  const latestRegistrations = registrations
    .slice()
    .sort(function(a, b) { return String(b.created_at || "").localeCompare(String(a.created_at || "")); })
    .slice(0, 10);

  return {
    status: "success",
    data: {
      // new dashboard metrics
      total_courses: courses.length,
      active_schedules: activeSchedules.length,
      total_inquiries: registrations.length,
      confirmed_participants: registrations.filter(function(item) {
        const status = String(item.status || "").toLowerCase();
        return status === "confirmed" || status === "paid" || status === "enrolled";
      }).length,
      branch_count: branches.length,
      latest_registrations: latestRegistrations,

      // backward-compatible legacy metrics
      product_count: legacyProducts.length,
      category_count: legacyCategories.length,
      article_count: legacyArticles.length,
      inquiry_count: legacyInquiries.length,
      low_stock_warning: lowStock,
      latest_inquiries: legacyInquiries.slice(-10).reverse()
    }
  };
}

function seedHerbalDemoData(d) {
  requireAdminSession_(d, { actionName: "seed_herbal_demo_data" });

  ensureSheetWithHeaders_("product_categories", HERBAL_ENTITY_SCHEMAS.product_categories);
  ensureSheetWithHeaders_("products", HERBAL_ENTITY_SCHEMAS.products);
  ensureSheetWithHeaders_("product_images", HERBAL_ENTITY_SCHEMAS.product_images);
  ensureSheetWithHeaders_("health_tags", HERBAL_ENTITY_SCHEMAS.health_tags);
  ensureSheetWithHeaders_("product_health_tags", HERBAL_ENTITY_SCHEMAS.product_health_tags);
  ensureSheetWithHeaders_("articles", HERBAL_ENTITY_SCHEMAS.articles);
  ensureSheetWithHeaders_("testimonials", HERBAL_ENTITY_SCHEMAS.testimonials);
  ensureSheetWithHeaders_("inquiries", HERBAL_ENTITY_SCHEMAS.inquiries);
  ensureSheetWithHeaders_("banners", HERBAL_ENTITY_SCHEMAS.banners);
  ensureSheetWithHeaders_("faq_items", HERBAL_ENTITY_SCHEMAS.faq_items);
  ensureSheetWithHeaders_("promo_campaigns", HERBAL_ENTITY_SCHEMAS.promo_campaigns);

  const hasProducts = readSheetAsObjects_("products", HERBAL_ENTITY_SCHEMAS.products).length > 0;
  if (hasProducts) {
    return { status: "success", message: "Data herbal sudah tersedia. Seed dilewati untuk mencegah duplikasi." };
  }

  const now = nowIsoDateTime_();
  const categories = [
    ["CAT-IMMUNE", "Immune Support", "immune-support", "Produk herbal untuk dukungan daya tahan tubuh.", "shield", 1, true, now, now],
    ["CAT-DIGEST", "Digestion", "digestion", "Formulasi herbal untuk kenyamanan pencernaan.", "leaf", 2, true, now, now],
    ["CAT-DETOX", "Detox", "detox", "Produk pendukung proses detoksifikasi alami tubuh.", "drop", 3, true, now, now],
    ["CAT-ENERGY", "Energy and Stamina", "energy-stamina", "Herbal pendukung energi dan vitalitas.", "bolt", 4, true, now, now],
    ["CAT-SLEEP", "Sleep and Relaxation", "sleep-relaxation", "Dukungan relaksasi dan kualitas tidur.", "moon", 5, true, now, now],
    ["CAT-WOMEN", "Women Wellness", "women-wellness", "Herbal dukungan kesehatan wanita.", "flower", 6, true, now, now],
    ["CAT-MEN", "Men Wellness", "men-wellness", "Herbal dukungan vitalitas pria.", "mountain", 7, true, now, now],
    ["CAT-BEAUTY", "Skin and Beauty", "skin-beauty", "Perawatan herbal kulit dan kecantikan.", "sparkles", 8, true, now, now],
    ["CAT-EXTERNAL", "External Care", "external-care", "Minyak, balm, dan perawatan luar herbal.", "hand", 9, true, now, now],
    ["CAT-DRINK", "Herbal Drinks", "herbal-drinks", "Minuman herbal siap seduh.", "cup", 10, true, now, now],
    ["CAT-BUNDLE", "Bundles and Promo", "bundles-promo", "Paket bundling herbal hemat.", "gift", 11, true, now, now]
  ];
  const categorySheet = ensureSheetWithHeaders_("product_categories", HERBAL_ENTITY_SCHEMAS.product_categories);
  categorySheet.getRange(2, 1, categories.length, HERBAL_ENTITY_SCHEMAS.product_categories.length).setValues(categories);

  const products = [
    ["PRD-001","CAT-IMMUNE","Echinacea Immune Guard","echinacea-immune-guard","Ekstrak echinacea dan meniran untuk dukungan imun harian.","Formula kombinasi echinacea, meniran, dan vitamin C herbal-grade untuk menjaga daya tahan tubuh.","Echinacea, Meniran, Jahe","Membantu menjaga daya tahan tubuh","2 kapsul per hari sesudah makan","2 kapsul/hari","Tidak dianjurkan untuk ibu hamil tanpa konsultasi","capsule","60 kapsul","BPOM TR", "SKU-IMM-001",189000,159000,42,"120g",true,true,true,"Echinacea Immune Guard | NaturaHerb","Suplemen herbal untuk dukungan daya tahan tubuh.",now,now],
    ["PRD-002","CAT-DIGEST","Ginger Digest Balance","ginger-digest-balance","Dukungan pencernaan dengan jahe merah dan temulawak.","Membantu kenyamanan lambung dan pencernaan setelah makan.","Jahe merah, Temulawak, Peppermint","Membantu pencernaan lebih nyaman","1 kapsul sebelum makan besar","1-2 kapsul/hari","Hentikan penggunaan bila ada reaksi alergi","capsule","30 kapsul","BPOM TR", "SKU-DIG-002",145000,129000,58,"90g",false,false,true,"Ginger Digest Balance","Herbal pencernaan harian berbahan jahe merah.",now,now],
    ["PRD-003","CAT-SLEEP","Calm Sleep Herbal Tea","calm-sleep-herbal-tea","Teh herbal relaksasi untuk kualitas tidur.","Perpaduan chamomile, lavender, dan daun mint untuk membantu relaksasi malam.","Chamomile, Lavender, Peppermint","Membantu relaksasi sebelum tidur","Seduh 1 sachet 10 menit sebelum tidur","1 sachet/hari","Tidak untuk anak di bawah 5 tahun","drink","20 sachet","PIRT", "SKU-SLP-003",99000,89000,76,"200g",true,false,true,"Calm Sleep Herbal Tea","Teh herbal relaksasi untuk tidur lebih nyaman.",now,now],
    ["PRD-004","CAT-ENERGY","Royal Ginseng Vital","royal-ginseng-vital","Dukungan stamina dengan ginseng dan maca.","Kombinasi adaptogen untuk membantu stamina dan fokus aktivitas.","Ginseng, Maca, Madu bubuk","Membantu menjaga stamina","1 kapsul pagi hari","1 kapsul/hari","Hindari konsumsi malam hari","capsule","30 kapsul","BPOM TR", "SKU-ENG-004",210000,189000,33,"95g",true,true,true,"Royal Ginseng Vital","Suplemen herbal untuk energi dan fokus.",now,now],
    ["PRD-005","CAT-WOMEN","Women Harmony Blend","women-harmony-blend","Herbal dukungan kebugaran wanita.","Mengandung kunyit asam dan daun raspberry untuk dukungan kebugaran wanita.","Kunyit, Asam jawa, Daun raspberry","Membantu menjaga kenyamanan tubuh wanita","2 kapsul per hari","2 kapsul/hari","Konsultasikan saat hamil/menyusui","capsule","60 kapsul","BPOM TR", "SKU-WMN-005",175000,155000,27,"110g",false,false,true,"Women Harmony Blend","Herbal support untuk kebugaran wanita.",now,now],
    ["PRD-006","CAT-MEN","Men Vital Force","men-vital-force","Dukungan vitalitas pria berbahan herbal.","Kombinasi pasak bumi, tribulus, dan jahe hitam.","Pasak bumi, Tribulus, Jahe hitam","Membantu menunjang vitalitas pria","1 kapsul pagi dan malam","2 kapsul/hari","Tidak untuk penderita hipertensi tanpa saran dokter","capsule","60 kapsul","BPOM TR", "SKU-MEN-006",225000,199000,21,"120g",true,true,true,"Men Vital Force","Suplemen herbal pendukung vitalitas pria.",now,now],
    ["PRD-007","CAT-DETOX","Detox Fiber Cleanse","detox-fiber-cleanse","Serat herbal untuk dukungan detoks harian.","Serat larut, psyllium husk, dan daun senna untuk dukungan detoksifikasi.","Psyllium husk, Senna, Lemon peel","Membantu proses detoks alami","Larutkan 1 sendok takar pada 250ml air","1 kali/hari","Pastikan asupan cairan cukup","powder","250g","BPOM TR", "SKU-DTX-007",130000,115000,64,"250g",false,false,true,"Detox Fiber Cleanse","Serat herbal untuk pencernaan dan detoks.",now,now],
    ["PRD-008","CAT-BEAUTY","Collagen Herbal Glow","collagen-herbal-glow","Minuman kolagen dengan herbal antioksidan.","Kolagen laut dengan ekstrak rosella dan delima.","Kolagen, Rosella, Delima","Membantu menjaga elastisitas kulit","1 sachet dilarutkan air dingin","1 sachet/hari","Simpan di tempat sejuk","drink","15 sachet","BPOM MD", "SKU-BEA-008",199000,179000,48,"180g",true,false,true,"Collagen Herbal Glow","Kolagen + herbal antioksidan untuk kulit.",now,now],
    ["PRD-009","CAT-EXTERNAL","Herbal Muscle Balm","herbal-muscle-balm","Balm herbal untuk perawatan area pegal.","Mengandung minyak atsiri herbal untuk sensasi hangat menenangkan.","Minyak cengkeh, Kayu putih, Menthol","Membantu relaksasi otot","Oleskan tipis pada area yang dibutuhkan","Sesuai kebutuhan","Hanya untuk pemakaian luar","balm","30g","BPOM NA", "SKU-EXT-009",75000,65000,88,"30g",false,false,true,"Herbal Muscle Balm","Balm herbal untuk area pegal.",now,now],
    ["PRD-010","CAT-DRINK","Turmeric Wellness Drink","turmeric-wellness-drink","Minuman kunyit siap seduh.","Kunyit, jahe, dan kayu manis untuk minuman harian.","Kunyit, Jahe, Kayu manis","Membantu menjaga kebugaran","Seduh 1 sachet dengan air hangat","1-2 sachet/hari","Tidak dianjurkan berlebihan","drink","20 sachet","PIRT", "SKU-DRK-010",85000,79000,95,"220g",true,false,true,"Turmeric Wellness Drink","Minuman herbal kunyit-jahe harian.",now,now],
    ["PRD-011","CAT-BUNDLE","Immune Family Bundle","immune-family-bundle","Paket hemat dukungan imun keluarga.","Bundling produk immune support untuk 1 bulan.","Variatif","Dukungan imun praktis dan hemat","Ikuti petunjuk masing-masing produk","Sesuai petunjuk","Lihat detail tiap produk","bundle","1 paket","-", "SKU-BND-011",399000,329000,14,"1.2kg",true,true,true,"Immune Family Bundle","Paket hemat produk dukungan imun keluarga.",now,now],
    ["PRD-012","CAT-DIGEST","Probiotic Herbal Mix","probiotic-herbal-mix","Probiotik dengan ekstrak herbal pencernaan.","Membantu keseimbangan mikrobiota usus.","Probiotik, Adas, Jahe","Membantu kesehatan pencernaan","1 sachet setelah makan","1 sachet/hari","Simpan di tempat kering","powder","14 sachet","BPOM MD", "SKU-DIG-012",165000,149000,39,"160g",false,false,true,"Probiotic Herbal Mix","Probiotik dan herbal untuk pencernaan.",now,now],
    ["PRD-013","CAT-SLEEP","Relax Balm Roll On","relax-balm-roll-on","Roll on aromaterapi relaksasi.","Aroma lavender dan chamomile untuk kenyamanan istirahat.","Lavender oil, Chamomile oil","Membantu relaksasi","Oleskan pada pelipis/pergelangan","Sesuai kebutuhan","Hindari area mata","oil","10ml","-", "SKU-SLP-013",69000,59000,71,"10ml",false,false,true,"Relax Balm Roll On","Aromaterapi roll on untuk relaksasi.",now,now],
    ["PRD-014","CAT-ENERGY","Moringa Daily Boost","moringa-daily-boost","Kapsul daun kelor untuk energi harian.","Kaya nutrisi mikro untuk dukungan aktivitas.","Daun kelor","Membantu menjaga energi harian","2 kapsul setelah sarapan","2 kapsul/hari","Konsultasikan jika memiliki kondisi medis khusus","capsule","60 kapsul","BPOM TR", "SKU-ENG-014",120000,105000,53,"100g",false,false,true,"Moringa Daily Boost","Kapsul kelor dukungan energi alami.",now,now],
    ["PRD-015","CAT-DETOX","Liver Care Herbal","liver-care-herbal","Dukungan fungsi hati dengan herbal terpilih.","Formulasi milk thistle dan temulawak.","Milk thistle, Temulawak","Membantu mendukung fungsi hati","1 kapsul setelah makan malam","1 kapsul/hari","Tidak untuk anak-anak","capsule","30 kapsul","BPOM TR", "SKU-DTX-015",178000,159000,26,"95g",false,false,true,"Liver Care Herbal","Herbal dukungan fungsi hati.",now,now],
    ["PRD-016","CAT-WOMEN","Herbal Cranberry Care","herbal-cranberry-care","Dukungan kenyamanan area kewanitaan.","Kombinasi cranberry dan daun sirih.","Cranberry, Daun sirih","Membantu menjaga kebersihan area intim","1 kapsul per hari","1 kapsul/hari","Hentikan bila terjadi ketidaknyamanan","capsule","30 kapsul","BPOM TR", "SKU-WMN-016",168000,149000,34,"90g",false,false,true,"Herbal Cranberry Care","Herbal support kenyamanan wanita.",now,now],
    ["PRD-017","CAT-MEN","Tribulus Stamina Drink","tribulus-stamina-drink","Minuman stamina pria dengan tribulus.","Formula cair siap minum untuk aktivitas tinggi.","Tribulus, Madu, Jahe","Membantu menjaga stamina pria","Minum 1 botol sebelum aktivitas","1 botol/hari","Batasi konsumsi sesuai anjuran","drink","10 botol","BPOM MD", "SKU-MEN-017",240000,219000,19,"500ml",true,false,true,"Tribulus Stamina Drink","Minuman herbal dukungan stamina pria.",now,now],
    ["PRD-018","CAT-BEAUTY","Aloe Skin Soothing Gel","aloe-skin-soothing-gel","Gel aloe untuk perawatan kulit harian.","Menenangkan kulit kering/terpapar matahari.","Aloe vera, Green tea extract","Membantu menenangkan kulit","Oleskan tipis pada kulit bersih","2-3 kali/hari","Untuk pemakaian luar","gel","100ml","BPOM NA", "SKU-BEA-018",89000,79000,81,"100ml",false,false,true,"Aloe Skin Soothing Gel","Gel aloe vera untuk perawatan kulit.",now,now]
  ];
  const productSheet = ensureSheetWithHeaders_("products", HERBAL_ENTITY_SCHEMAS.products);
  productSheet.getRange(2, 1, products.length, HERBAL_ENTITY_SCHEMAS.products.length).setValues(products);

  const images = products.map(function(item, idx) {
    return ["IMG-" + (idx + 1), item[0], "https://images.unsplash.com/photo-1514996937319-344454492b37?auto=format&fit=crop&w=1200&q=80", item[2], 1, now];
  });
  const imageSheet = ensureSheetWithHeaders_("product_images", HERBAL_ENTITY_SCHEMAS.product_images);
  imageSheet.getRange(2, 1, images.length, HERBAL_ENTITY_SCHEMAS.product_images.length).setValues(images);

  const tags = [
    ["TAG-IMM", "Immune Support", "immune-support", "Dukungan daya tahan tubuh", now, now],
    ["TAG-DIG", "Digestion", "digestion", "Kesehatan pencernaan", now, now],
    ["TAG-SLEEP", "Sleep", "sleep", "Relaksasi dan tidur", now, now],
    ["TAG-ENERGY", "Energy", "energy", "Energi dan stamina", now, now],
    ["TAG-DETOX", "Detox", "detox", "Detoksifikasi", now, now],
    ["TAG-BEAUTY", "Beauty", "beauty", "Kecantikan dan kulit", now, now]
  ];
  const tagSheet = ensureSheetWithHeaders_("health_tags", HERBAL_ENTITY_SCHEMAS.health_tags);
  tagSheet.getRange(2, 1, tags.length, HERBAL_ENTITY_SCHEMAS.health_tags.length).setValues(tags);

  const links = products.map(function(item, idx) {
    const tag = tags[idx % tags.length];
    return ["PHT-" + (idx + 1), item[0], tag[0]];
  });
  const linkSheet = ensureSheetWithHeaders_("product_health_tags", HERBAL_ENTITY_SCHEMAS.product_health_tags);
  linkSheet.getRange(2, 1, links.length, HERBAL_ENTITY_SCHEMAS.product_health_tags.length).setValues(links);

  const articles = [
    ["ART-001", "Panduan Memilih Herbal Berdasarkan Kebutuhan Harian", "panduan-memilih-herbal", "Cara memilih produk herbal berdasarkan tujuan kesehatan.", "Konten edukasi herbal tentang cara membaca komposisi, dosis, dan keamanan penggunaan harian.", "", "Tim NaturaHerb", "Edukasi Herbal", "herbal,edukasi,panduan", true, now, "Panduan Memilih Herbal", "Panduan praktis memilih produk herbal aman.", now, now],
    ["ART-002", "Perbedaan Herbal Capsule, Powder, dan Drink", "perbedaan-form-herbal", "Memahami bentuk sediaan herbal.", "Penjelasan kelebihan dan kekurangan capsule, powder, dan minuman herbal.", "", "Tim NaturaHerb", "Tips Produk", "capsule,powder,drink", true, now, "Perbedaan Bentuk Herbal", "Edukasi bentuk sediaan herbal.", now, now],
    ["ART-003", "Mitos vs Fakta Produk Herbal", "mitos-vs-fakta-herbal", "Membedakan klaim yang realistis dan berlebihan.", "Konten untuk membantu pengguna memahami batasan klaim produk herbal.", "", "Tim NaturaHerb", "FAQ", "mitos,fakta,herbal", true, now, "Mitos vs Fakta Herbal", "Pahami klaim herbal dengan bijak.", now, now]
  ];
  const articleSheet = ensureSheetWithHeaders_("articles", HERBAL_ENTITY_SCHEMAS.articles);
  articleSheet.getRange(2, 1, articles.length, HERBAL_ENTITY_SCHEMAS.articles.length).setValues(articles);

  const testimonials = [
    ["TS-001", "Rina", "Bandung", 5, "Produknya membantu saya menjaga stamina harian.", "PRD-001", true, true, now, now],
    ["TS-002", "Andi", "Jakarta", 5, "Pencernaan jadi lebih nyaman setelah rutin konsumsi.", "PRD-002", true, true, now, now],
    ["TS-003", "Maya", "Surabaya", 4, "Teh relaksasi membantu kualitas tidur saya.", "PRD-003", false, true, now, now]
  ];
  const testimonialSheet = ensureSheetWithHeaders_("testimonials", HERBAL_ENTITY_SCHEMAS.testimonials);
  testimonialSheet.getRange(2, 1, testimonials.length, HERBAL_ENTITY_SCHEMAS.testimonials.length).setValues(testimonials);

  const banners = [
    ["BNR-001", "Dukungan Herbal untuk Keseharian Anda", "Produk herbal alami dengan kualitas terjaga", "https://images.unsplash.com/photo-1502741338009-cac2772e18bc?auto=format&fit=crop&w=1600&q=80", "Belanja Sekarang", "/produk", "home", 1, true, now, now],
    ["BNR-002", "Promo Bundling Immune Support", "Hemat hingga 20% untuk paket keluarga", "https://images.unsplash.com/photo-1471193945509-9ad0617afabf?auto=format&fit=crop&w=1600&q=80", "Lihat Promo", "/promo", "home", 2, true, now, now]
  ];
  const bannerSheet = ensureSheetWithHeaders_("banners", HERBAL_ENTITY_SCHEMAS.banners);
  bannerSheet.getRange(2, 1, banners.length, HERBAL_ENTITY_SCHEMAS.banners.length).setValues(banners);

  const faqs = [
    ["FAQ-001", "Apakah produk herbal aman dikonsumsi setiap hari?", "Gunakan sesuai petunjuk label dan konsultasikan jika memiliki kondisi medis khusus.", 1, true, now, now],
    ["FAQ-002", "Apakah produk herbal ini bisa menggantikan obat dokter?", "Tidak. Produk herbal adalah dukungan kesehatan, bukan pengganti diagnosis atau terapi medis.", 2, true, now, now],
    ["FAQ-003", "Berapa lama hasil konsumsi herbal dapat dirasakan?", "Respons tubuh berbeda tiap individu. Konsumsi rutin sesuai anjuran biasanya membantu hasil yang lebih optimal.", 3, true, now, now]
  ];
  const faqSheet = ensureSheetWithHeaders_("faq_items", HERBAL_ENTITY_SCHEMAS.faq_items);
  faqSheet.getRange(2, 1, faqs.length, HERBAL_ENTITY_SCHEMAS.faq_items.length).setValues(faqs);

  const inquiries = [
    ["INQ-001", "Siti", "081234567890", "siti@email.com", "Konsultasi produk", "Produk mana yang cocok untuk dukungan tidur?", "consultation", "PRD-003", "new", now, now],
    ["INQ-002", "Budi", "082233445566", "budi@email.com", "Informasi reseller", "Apakah tersedia program reseller untuk kota Semarang?", "reseller", "", "processed", now, now]
  ];
  const inquirySheet = ensureSheetWithHeaders_("inquiries", HERBAL_ENTITY_SCHEMAS.inquiries);
  inquirySheet.getRange(2, 1, inquiries.length, HERBAL_ENTITY_SCHEMAS.inquiries.length).setValues(inquiries);

  const promos = [
    ["PRM-001", "Promo Bundling Imunitas", "promo-bundling-imunitas", "Diskon khusus paket imunitas.", "percentage", 20, toISODate_(), "2099-12-31", true, now, now]
  ];
  const promoSheet = ensureSheetWithHeaders_("promo_campaigns", HERBAL_ENTITY_SCHEMAS.promo_campaigns);
  promoSheet.getRange(2, 1, promos.length, HERBAL_ENTITY_SCHEMAS.promo_campaigns.length).setValues(promos);

  updateSettings({
    auth_session_token: getAdminSessionToken_(d),
    payload: {
      site_name: "NaturaHerb Wellness",
      site_tagline: "Herbal alami untuk dukungan kesehatan harian Anda",
      legal_disclaimer: "Produk herbal membantu menjaga kesehatan dan bukan pengganti diagnosis maupun terapi medis profesional.",
      certifications: "BPOM,Halal,GMP"
    }
  });

  return withPublicCacheState_({ status: "success", message: "Seed data herbal berhasil dibuat.", products_seeded: products.length }, bumpPublicCacheState_(["catalog", "pages", "settings", "dashboard"]));
}

function seedCourseDemoData(d) {
  requireAdminSession_(d, { actionName: "seed_course_demo_data" });
  ensureCourseDomainSheets_();

  const existing = readSheetAsObjects_("courses", HERBAL_ENTITY_SCHEMAS.courses);
  if (existing.length > 0) {
    return { status: "success", message: "Data kursus sudah tersedia. Seed dilewati untuk mencegah duplikasi." };
  }

  const now = nowIsoDateTime_();

  const categories = [
    ["CC-BHS", "Bahasa", "bahasa", "Program bahasa tatap muka", "languages", 1, true, now, now],
    ["CC-KMP", "Komputer", "komputer", "Pelatihan komputer dan office", "monitor", 2, true, now, now],
    ["CC-DSN", "Desain", "desain", "Program desain kreatif", "palette", 3, true, now, now],
    ["CC-DGM", "Digital Marketing", "digital-marketing", "Strategi pemasaran digital", "megaphone", 4, true, now, now],
    ["CC-PSP", "Public Speaking", "public-speaking", "Pelatihan komunikasi publik", "mic", 5, true, now, now],
    ["CC-PRG", "Programming", "programming", "Kelas coding untuk pemula-menengah", "code", 6, true, now, now],
    ["CC-OFC", "Office Skills", "office-skills", "Kelas skill administrasi kantor", "briefcase", 7, true, now, now],
    ["CC-KNG", "Keuangan", "keuangan", "Pelatihan akuntansi dan finansial", "wallet", 8, true, now, now],
    ["CC-KLN", "Kuliner", "kuliner", "Kelas praktik usaha kuliner", "chef-hat", 9, true, now, now],
    ["CC-ANR", "Anak dan Remaja", "anak-remaja", "Program skill untuk pelajar", "users", 10, true, now, now]
  ];
  ensureSheetWithHeaders_("course_categories", HERBAL_ENTITY_SCHEMAS.course_categories)
    .getRange(2, 1, categories.length, HERBAL_ENTITY_SCHEMAS.course_categories.length)
    .setValues(categories);

  const courses = [
    ["CRS-001", "CC-BHS", "Kursus Bahasa Inggris Intensif", "kursus-bahasa-inggris-intensif", "Program speaking & grammar intensif", "Program tatap muka fokus komunikasi aktif untuk akademik dan kerja.", "Percaya diri speaking, grammar aplikatif", "Praktik percakapan langsung", "SMA, mahasiswa, pekerja", "Pemula", "8 minggu", 16, "offline", "Membawa alat tulis", "Modul, simulasi interview", true, 1800000, 1500000, "https://images.unsplash.com/photo-1456513080510-7bf3a84b82f8?w=1200", true, true, "Kursus Bahasa Inggris Intensif", "Kursus bahasa Inggris tatap muka untuk pelajar dan profesional", now, now],
    ["CRS-002", "CC-DSN", "Kelas Desain Grafis Tatap Muka", "kelas-desain-grafis-tatap-muka", "Belajar desain dari dasar sampai project", "Pelatihan Adobe & Canva berbasis proyek untuk kebutuhan konten profesional.", "Mahir desain sosial media dan branding", "Studi kasus bisnis nyata", "Pemula, UMKM, content creator", "Pemula", "6 minggu", 12, "offline", "Laptop pribadi", "Template, feedback mentor", true, 2200000, 1900000, "https://images.unsplash.com/photo-1507238691740-187a5b1d37b8?w=1200", true, true, "Kelas Desain Grafis Tatap Muka", "Kelas desain grafis offline berbasis praktik", now, now],
    ["CRS-003", "CC-KMP", "Pelatihan Microsoft Excel untuk Kerja", "pelatihan-microsoft-excel-untuk-kerja", "Excel praktis untuk kebutuhan kantor", "Fokus formula, pivot table, dashboard laporan dan automasi dasar.", "Meningkatkan produktivitas kerja", "Latihan file kasus real", "Karyawan, fresh graduate", "Pemula", "4 minggu", 8, "offline", "Laptop dengan Excel", "File latihan & template", true, 1450000, 1200000, "https://images.unsplash.com/photo-1461749280684-dccba630e2f6?w=1200", false, true, "Pelatihan Microsoft Excel untuk Kerja", "Kursus Excel tatap muka untuk profesional", now, now],
    ["CRS-004", "CC-DGM", "Bootcamp Digital Marketing Offline", "bootcamp-digital-marketing-offline", "Bootcamp performa pemasaran digital", "Mencakup strategi konten, ads, funnel, dan optimasi conversion.", "Siap eksekusi campaign digital", "Praktik ads dan audit funnel", "UMKM, marketer, owner bisnis", "Menengah", "10 minggu", 20, "offline", "Laptop dan akun bisnis", "Template campaign + mentoring", true, 3500000, 2990000, "https://images.unsplash.com/photo-1432888622747-4eb9a8efeb07?w=1200", true, true, "Bootcamp Digital Marketing Offline", "Bootcamp marketing offline untuk pertumbuhan bisnis", now, now],
    ["CRS-005", "CC-PSP", "Kelas Public Speaking", "kelas-public-speaking", "Latihan komunikasi percaya diri", "Teknik vokal, struktur presentasi, handling nervous, dan body language.", "Presentasi lebih efektif", "Live speaking simulation", "Pelajar, mahasiswa, profesional", "Pemula", "5 minggu", 10, "offline", "Komitmen latihan", "Video feedback personal", true, 1600000, 1350000, "https://images.unsplash.com/photo-1475721027785-f74eccf877e2?w=1200", false, true, "Kelas Public Speaking", "Pelatihan public speaking tatap muka", now, now],
    ["CRS-006", "CC-PRG", "Kursus Programming Dasar", "kursus-programming-dasar", "Belajar logika coding dari nol", "Dasar pemrograman, algoritma, dan mini project aplikasi sederhana.", "Memahami fundamental coding", "Project mingguan", "Pelajar, fresh graduate", "Pemula", "8 minggu", 16, "offline", "Laptop", "Mentoring dan code review", true, 2400000, 2100000, "https://images.unsplash.com/photo-1515879218367-8466d910aaa4?w=1200", true, true, "Kursus Programming Dasar", "Kelas coding dasar tatap muka untuk pemula", now, now],
    ["CRS-007", "CC-KNG", "Workshop Akuntansi UMKM", "workshop-akuntansi-umkm", "Workshop laporan keuangan praktis", "Pencatatan kas, laba rugi, arus kas, dan kontrol biaya untuk usaha kecil.", "Mampu menyusun laporan keuangan", "Template pembukuan siap pakai", "Pemilik UMKM", "Pemula", "3 minggu", 6, "offline", "Membawa data usaha", "Workbook akuntansi", true, 1300000, 1100000, "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=1200", false, true, "Workshop Akuntansi UMKM", "Pelatihan akuntansi praktis untuk UMKM", now, now],
    ["CRS-008", "CC-KLN", "Kelas Barista Pemula", "kelas-barista-pemula", "Pelatihan dasar barista & brew", "Teori kopi, grinding, espresso, milk steaming, dan service dasar kafe.", "Siap kerja entry-level barista", "Praktik alat espresso", "Pemula, pencari kerja", "Pemula", "4 minggu", 8, "offline", "Fisik sehat", "Sertifikat penyelesaian", true, 2800000, 2500000, "https://images.unsplash.com/photo-1495474472287-4d71bcdd2085?w=1200", true, true, "Kelas Barista Pemula", "Pelatihan barista tatap muka untuk karier", now, now],
    ["CRS-009", "CC-ANR", "Kelas Make Up Basic", "kelas-make-up-basic", "Kelas makeup dasar untuk pemula", "Pengenalan tools, teknik complexion, eye makeup, dan look natural/party.", "Menguasai teknik makeup dasar", "Praktik langsung dengan model", "Remaja, dewasa", "Pemula", "4 minggu", 8, "offline", "Membawa kit pribadi", "Face chart dan modul", true, 2100000, 1800000, "https://images.unsplash.com/photo-1522335789203-aabd1fc54bc9?w=1200", false, true, "Kelas Make Up Basic", "Kursus makeup tatap muka untuk pemula", now, now],
    ["CRS-010", "CC-OFC", "Kursus Komputer Administrasi Perkantoran", "kursus-komputer-administrasi-perkantoran", "Program siap kerja administrasi", "Word, Excel, PowerPoint, email bisnis, dan etika administrasi kantor.", "Meningkatkan kesiapan kerja kantor", "Simulasi tugas administrasi", "SMA/SMK, fresh graduate", "Pemula", "6 minggu", 12, "offline", "Laptop", "Template dokumen kantor", true, 1700000, 1450000, "https://images.unsplash.com/photo-1520607162513-77705c0f0d4a?w=1200", true, true, "Kursus Komputer Administrasi Perkantoran", "Pelatihan admin office tatap muka", now, now]
  ];
  ensureSheetWithHeaders_("courses", HERBAL_ENTITY_SCHEMAS.courses)
    .getRange(2, 1, courses.length, HERBAL_ENTITY_SCHEMAS.courses.length)
    .setValues(courses);

  const branches = [
    ["BR-001", "Jakarta Selatan", "Jakarta", "Jl. TB Simatupang No. 88", "https://maps.google.com", "081200000001", "jaksel@kursusoffline.id", "Senin - Sabtu 09:00-20:00", true, now, now],
    ["BR-002", "Bandung", "Bandung", "Jl. Buah Batu No. 120", "https://maps.google.com", "081200000002", "bandung@kursusoffline.id", "Senin - Sabtu 09:00-20:00", true, now, now],
    ["BR-003", "Surabaya", "Surabaya", "Jl. Darmo No. 45", "https://maps.google.com", "081200000003", "surabaya@kursusoffline.id", "Senin - Sabtu 09:00-20:00", true, now, now],
    ["BR-004", "Yogyakarta", "Yogyakarta", "Jl. Kaliurang KM 5", "https://maps.google.com", "081200000004", "yogya@kursusoffline.id", "Senin - Sabtu 09:00-20:00", true, now, now],
    ["BR-005", "Bekasi", "Bekasi", "Jl. Ahmad Yani No. 21", "https://maps.google.com", "081200000005", "bekasi@kursusoffline.id", "Senin - Sabtu 09:00-20:00", true, now, now]
  ];
  ensureSheetWithHeaders_("branches", HERBAL_ENTITY_SCHEMAS.branches)
    .getRange(2, 1, branches.length, HERBAL_ENTITY_SCHEMAS.branches.length)
    .setValues(branches);

  const instructors = [
    ["INS-001", "Dwi Pratama", "dwi-pratama", "", "English Communication", "Trainer bahasa Inggris korporat", 8, "TOEFL ITP, CELTA", true, true, now, now],
    ["INS-002", "Rina Aulia", "rina-aulia", "", "Graphic Design", "Praktisi desain brand UMKM", 7, "Adobe Certified", true, true, now, now],
    ["INS-003", "Arif Nugroho", "arif-nugroho", "", "Data & Excel", "Konsultan data reporting", 10, "MOS Expert", true, true, now, now],
    ["INS-004", "Fajar Ramadhan", "fajar-ramadhan", "", "Digital Marketing", "Performance marketer lintas industri", 9, "Meta Blueprint", true, true, now, now],
    ["INS-005", "Nadia Putri", "nadia-putri", "", "Public Speaking", "Coach komunikasi profesional", 6, "BNSP Trainer", true, true, now, now],
    ["INS-006", "Bagas Saputra", "bagas-saputra", "", "Programming", "Engineer & coding mentor", 8, "AWS CCP", true, true, now, now],
    ["INS-007", "Mila Anggraini", "mila-anggraini", "", "Barista & F&B", "Trainer coffee shop operations", 5, "SCA Foundation", true, true, now, now]
  ];
  ensureSheetWithHeaders_("instructors", HERBAL_ENTITY_SCHEMAS.instructors)
    .getRange(2, 1, instructors.length, HERBAL_ENTITY_SCHEMAS.instructors.length)
    .setValues(instructors);

  const schedules = [];
  for (let i = 0; i < 18; i++) {
    const c = courses[i % courses.length];
    const b = branches[i % branches.length];
    const ins = instructors[i % instructors.length];
    const quota = 20 + (i % 3) * 5;
    const booked = 5 + (i % 7);
    schedules.push([
      "SCH-" + String(i + 1).padStart(3, "0"),
      c[0],
      b[0],
      ins[0],
      "BATCH-" + String(i + 1).padStart(2, "0"),
      "2026-04-" + String((i % 20) + 5).padStart(2, "0"),
      "2026-05-" + String((i % 20) + 5).padStart(2, "0"),
      i % 2 === 0 ? "Weekday" : "Weekend",
      i % 2 === 0 ? "19:00-21:00" : "09:00-12:00",
      quota,
      booked,
      quota - booked,
      "2026-04-" + String((i % 20) + 1).padStart(2, "0"),
      quota - booked <= 2 ? "limited" : "open",
      "Jadwal reguler",
      now,
      now
    ]);
  }
  ensureSheetWithHeaders_("schedules", HERBAL_ENTITY_SCHEMAS.schedules)
    .getRange(2, 1, schedules.length, HERBAL_ENTITY_SCHEMAS.schedules.length)
    .setValues(schedules);

  const registrations = [];
  for (let j = 0; j < 12; j++) {
    registrations.push([
      "REG-" + String(j + 1).padStart(3, "0"),
      courses[j % courses.length][0],
      schedules[j % schedules.length][0],
      ["Andi", "Siti", "Budi", "Rina", "Fajar", "Nisa", "Kevin", "Aulia", "Rizki", "Dina", "Farhan", "Maya"][j],
      "08123" + String(100000 + j),
      "lead" + (j + 1) + "@mail.com",
      19 + (j % 12),
      ["Mahasiswa", "Karyawan", "Fresh Graduate", "UMKM Owner"][j % 4],
      branches[j % branches.length][0],
      "Butuh kelas dengan jadwal fleksibel",
      ["instagram", "google", "friend", "tiktok", "corporate"][j % 5],
      ["new", "contacted", "follow_up", "confirmed", "paid", "enrolled"][j % 6],
      "staff-" + ((j % 3) + 1),
      "Follow-up awal",
      now,
      now
    ]);
  }
  ensureSheetWithHeaders_("registrations", HERBAL_ENTITY_SCHEMAS.registrations)
    .getRange(2, 1, registrations.length, HERBAL_ENTITY_SCHEMAS.registrations.length)
    .setValues(registrations);

  const gallery = [];
  for (let k = 0; k < 12; k++) {
    gallery.push([
      "GAL-" + String(k + 1).padStart(3, "0"),
      "Dokumentasi Kelas " + (k + 1),
      "image",
      "https://images.unsplash.com/photo-1523240795612-9a054b0db644?w=1200",
      ["classroom", "activity", "facility", "certificate"][k % 4],
      courses[k % courses.length][0],
      branches[k % branches.length][0],
      k < 4,
      k + 1,
      true,
      now,
      now
    ]);
  }
  ensureSheetWithHeaders_("galleries", HERBAL_ENTITY_SCHEMAS.galleries)
    .getRange(2, 1, gallery.length, HERBAL_ENTITY_SCHEMAS.galleries.length)
    .setValues(gallery);

  const contentBlocks = [
    ["CB-001", "home", "hero", "Belajar Skill Praktis Lewat Kelas Tatap Muka", "Program siap kerja untuk individu & corporate", "Konten hero utama", "", "Daftar Sekarang", "/register", 1, true, "Web Kursus Offline", "Platform pelatihan tatap muka profesional", now, now],
    ["CB-002", "home", "featured_programs", "Program Unggulan", "Pilih kelas paling populer", "Konten featured programs", "", "Lihat Program", "/courses", 2, true, "Program Unggulan Kursus", "Program pelatihan paling diminati", now, now],
    ["CB-003", "about", "institution_profile", "Profil Lembaga", "Komitmen mutu pelatihan", "Konten profil lembaga", "", "Hubungi Admin", "/contact", 1, true, "Tentang Lembaga Kursus", "Profil institusi pelatihan", now, now]
  ];
  ensureSheetWithHeaders_("content_blocks", HERBAL_ENTITY_SCHEMAS.content_blocks)
    .getRange(2, 1, contentBlocks.length, HERBAL_ENTITY_SCHEMAS.content_blocks.length)
    .setValues(contentBlocks);

  return withPublicCacheState_({
    status: "success",
    message: "Seed data kursus offline berhasil dibuat.",
    courses_seeded: courses.length,
    schedules_seeded: schedules.length,
    registrations_seeded: registrations.length
  }, bumpPublicCacheState_(["catalog", "pages", "settings", "dashboard"]));
}

function getSecret_(name, cfg) {
  const k = String(name || "").trim();
  if (!k) return "";
  try {
    const p = PropertiesService.getScriptProperties();
    const v = p.getProperty(k);
    if (v !== null && v !== undefined && String(v).trim() !== "") return String(v).trim();
  } catch (e) {}
  return String(getCfgFrom_(cfg || getSettingsMap_(), k) || "").trim();
}

function isDebugAllowed_() {
  try {
    const p = PropertiesService.getScriptProperties();
    return String(p.getProperty("DEBUG_MODE") || "false").toLowerCase() === "true";
  } catch (e) {
    return false;
  }
}

function hashPassword_(plain) {
  const input = String(plain || "");
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  const hex = digest.map(function(b){
    const v = (b < 0 ? b + 256 : b);
    return ("0" + v.toString(16)).slice(-2);
  }).join("");
  return "sha256$" + hex;
}

function verifyPassword_(input, stored) {
  const inStr = String(input || "");
  const st = String(stored || "").trim();
  if (!st) return false;
  if (st.indexOf("sha256$") === 0) return hashPassword_(inStr) === st;
  return inStr === st;
}

function getAdminSessionToken_(data) {
  const source = data || {};
  return String(
    source.auth_session_token ||
    source.admin_session_token ||
    source.session_token ||
    ""
  ).trim();
}

function getAdminSessionPropertyStore_() {
  return PropertiesService.getScriptProperties();
}

function persistAdminSession_(token, session) {
  const key = String(token || "").trim();
  if (!key || !session || typeof session !== "object") return;
  const serialized = JSON.stringify(session);
  getAdminSessionPropertyStore_().setProperty(ADMIN_SESSION_PROPERTY_PREFIX + key, serialized);
  try {
    CacheService.getScriptCache().put(ADMIN_SESSION_CACHE_PREFIX + key, serialized, ADMIN_SESSION_CACHE_TTL_SECONDS);
  } catch (e) {}
}

function revokeAdminSession_(token) {
  const key = String(token || "").trim();
  if (!key) return;
  getAdminSessionPropertyStore_().deleteProperty(ADMIN_SESSION_PROPERTY_PREFIX + key);
  try {
    CacheService.getScriptCache().remove(ADMIN_SESSION_CACHE_PREFIX + key);
  } catch (e) {}
}

function createAdminSession_(sessionData) {
  const issuedAt = Date.now();
  const token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  const expiresAt = issuedAt + ADMIN_SESSION_DURATION_MS;
  const session = Object.assign({
    id: "",
    email: "",
    name: "Admin",
    role: "admin",
    issued_at: issuedAt,
    expires_at: expiresAt
  }, sessionData || {});
  persistAdminSession_(token, session);
  return {
    token: token,
    expires_at: session.expires_at,
    session: session
  };
}

function getAdminSession_(token) {
  const key = String(token || "").trim();
  if (!key) return null;
  let cached = null;
  try {
    cached = CacheService.getScriptCache().get(ADMIN_SESSION_CACHE_PREFIX + key);
  } catch (e) {}
  try {
    const parsed = cached ? JSON.parse(cached) : null;
    if (parsed && typeof parsed === "object") {
      const propKey = ADMIN_SESSION_PROPERTY_PREFIX + key;
      if (!getAdminSessionPropertyStore_().getProperty(propKey)) {
        try {
          getAdminSessionPropertyStore_().setProperty(propKey, cached);
        } catch (e) {}
      }
      return parsed;
    }
  } catch (e) {}

  const stored = getAdminSessionPropertyStore_().getProperty(ADMIN_SESSION_PROPERTY_PREFIX + key);
  if (!stored) return null;
  try {
    const parsed = JSON.parse(stored);
    if (parsed && typeof parsed === "object") {
      try {
        CacheService.getScriptCache().put(ADMIN_SESSION_CACHE_PREFIX + key, stored, ADMIN_SESSION_CACHE_TTL_SECONDS);
      } catch (e) {}
      return parsed;
    }
  } catch (e) {
    revokeAdminSession_(key);
  }
  return null;
}

function validateAdminSessionAccess_(session, options) {
  const opts = options || {};
  const actionName = String(opts.actionName || "aksi admin").trim();
  const allowedRoles = Array.isArray(opts.allowedRoles) && opts.allowedRoles.length
    ? opts.allowedRoles.map(function(role) { return String(role || "").trim().toLowerCase(); })
    : ["admin"];
  if (!session || typeof session !== "object") {
    throw new Error("Sesi admin tidak valid. Silakan login ulang.");
  }
  const expiresAt = Number(session.expires_at || 0);
  if (!expiresAt || Date.now() >= expiresAt) {
    throw new Error("Sesi admin sudah kedaluwarsa. Silakan login ulang.");
  }
  const role = String(session.role || "").trim().toLowerCase();
  if (!role || allowedRoles.indexOf(role) === -1) {
    throw new Error("Akses admin ditolak untuk aksi " + actionName + ".");
  }
  return session;
}

function requireAdminSession_(data, options) {
  const token = getAdminSessionToken_(data);
  if (!token) throw new Error("Sesi admin tidak ditemukan. Silakan login ulang.");
  const session = getAdminSession_(token);
  if (!session) throw new Error("Sesi admin tidak valid. Silakan login ulang.");
  return validateAdminSessionAccess_(session, options);
}

function adminLogout(d) {
  const token = getAdminSessionToken_(d);
  if (token) revokeAdminSession_(token);
  return { status: "success", message: "Sesi admin berhasil ditutup." };
}

function sanitizeAssetUrl_(raw) {
  const value = String(raw || "").trim();
  if (!value) return "";
  if (/^data:image\//i.test(value)) return value;
  if (value.charAt(0) === "/") return value;
  if (!/^https?:\/\//i.test(value)) return "";

  const match = value.match(/^https?:\/\/([^\/?#]+)/i);
  const host = match && match[1] ? String(match[1]).toLowerCase() : "";
  if (!host) return "";

  if (
    host === "example.com" ||
    host === "example.org" ||
    host === "example.net" ||
    /(^|\.)example\.(com|org|net)$/i.test(host)
  ) {
    return "";
  }

  return value;
}

function getCurrentWebAppUrl_() {
  const fromConfig = String(getScriptConfig("APP_SCRIPT_URL") || getScriptConfig("SCRIPT_URL") || "").trim();
  if (fromConfig) return fromConfig;
  try {
    const url = ScriptApp.getService().getUrl();
    return String(url || "").trim();
  } catch (e) {
    return "";
  }
}

function normalizeMootaUrl_(raw) {
  const value = String(raw || "").trim();
  if (!value) return "";
  const match = value.match(/^(https:\/\/[^?#]+?)(?:[?#].*)?$/i);
  return match && match[1] ? String(match[1]).trim() : value;
}

function getMootaUrlHost_(raw) {
  const value = normalizeMootaUrl_(raw);
  const match = value.match(/^https:\/\/([^\/?#]+)/i);
  return match && match[1] ? String(match[1]).toLowerCase() : "";
}

function isDirectAppsScriptUrl_(raw) {
  const host = getMootaUrlHost_(raw);
  return host === "script.google.com" || host === "script.googleusercontent.com";
}

function isValidMootaUrl_(value) {
  const url = normalizeMootaUrl_(value);
  if (!url) return false;
  return /^https:\/\/[^\s?#]+$/i.test(url);
}

function isValidMootaToken_(value) {
  return /^[A-Za-z0-9]{8,200}$/.test(String(value || "").trim());
}

function resolveMootaConfig_(data, cfg) {
  const payload = data || {};
  const fallbackUrl = getCfgFrom_(cfg, "moota_gas_url") || getCurrentWebAppUrl_();
  const storedToken = getSecret_("moota_token", cfg);
  const legacySecret = getSecret_("moota_secret", cfg);
  const nextToken = payload.moota_token !== undefined
    ? payload.moota_token
    : (payload.moota_secret !== undefined ? payload.moota_secret : (storedToken || legacySecret));
  return {
    gasUrl: normalizeMootaUrl_(payload.moota_gas_url !== undefined ? payload.moota_gas_url : fallbackUrl),
    token: String(nextToken || "").trim()
  };
}

function validateMootaConfigFormat_(mootaCfg, opts) {
  const options = opts || {};
  const errors = [];
  const requireUrl = options.requireUrl !== false;
  const requireToken = options.requireToken !== false;

  if (requireUrl && !mootaCfg.gasUrl) errors.push("Link webhook Moota wajib diisi.");
  if (requireUrl && mootaCfg.gasUrl && !isValidMootaUrl_(mootaCfg.gasUrl)) {
    errors.push("Format link webhook Moota tidak valid. Gunakan URL HTTPS tanpa query string.");
  }
  if (requireUrl && mootaCfg.gasUrl && isDirectAppsScriptUrl_(mootaCfg.gasUrl)) {
    errors.push("Link webhook Moota tidak boleh langsung ke Google Apps Script. Gunakan endpoint Cloudflare Worker atau proxy publik agar header Signature bisa diteruskan.");
  }

  if (requireToken && !mootaCfg.token) errors.push("Secret Token Moota wajib diisi.");
  if (requireToken && mootaCfg.token && !isValidMootaToken_(mootaCfg.token)) {
    errors.push("Format Secret Token Moota tidak valid. Gunakan minimal 8 karakter alphanumeric tanpa spasi.");
  }

  return errors;
}

function normalizeMootaSignature_(raw) {
  let value = String(raw || "").trim();
  if (!value) return "";
  value = value.replace(/^sha256=/i, "").trim();
  return value.replace(/[^a-f0-9]/ig, "").toLowerCase();
}

function computeMootaSignatureHex_(payloadString, secretToken) {
  const computed = Utilities.computeHmacSha256Signature(String(payloadString || ""), String(secretToken || ""));
  return computed.map(function(chr) {
    const value = chr < 0 ? chr + 256 : chr;
    return ("0" + value.toString(16)).slice(-2);
  }).join("").toLowerCase();
}

function verifyMootaSignature_(payloadString, secretToken, rawSignature) {
  const secret = String(secretToken || "").trim();
  const received = normalizeMootaSignature_(rawSignature);
  if (!secret) {
    return { ok: false, code: "missing_secret", received: received, expected: "" };
  }
  if (!received) {
    return { ok: false, code: "missing_signature", received: "", expected: "" };
  }
  const expected = computeMootaSignatureHex_(payloadString, secret);
  return {
    ok: received === expected,
    code: received === expected ? "ok" : "invalid_signature",
    received: received,
    expected: expected
  };
}

function maskMootaSignatureForLog_(value) {
  const sig = String(value || "");
  if (!sig) return "";
  if (sig.length <= 12) return sig;
  return sig.substring(0, 8) + "..." + sig.substring(sig.length - 4);
}

function extractMootaSignatureMeta_(e) {
  const params = (e && e.parameter) || {};
  const rawSignature = params.moota_signature !== undefined
    ? params.moota_signature
    : (params.signature !== undefined ? params.signature : "");
  const signatureSource = params.moota_signature !== undefined
    ? "query:moota_signature"
    : (params.signature !== undefined ? "query:signature" : "missing");
  return {
    raw: String(rawSignature || "").trim(),
    normalized: normalizeMootaSignature_(rawSignature),
    source: signatureSource,
    forwardedByWorker: String(params.moota_forwarded || "").trim() === "1",
    workerSawSignature: String(params.moota_sig_present || "").trim() === "1",
    workerVerifiedSignature: String(params.moota_sig_verified || "").trim() === "1",
    workerVerificationSource: String(params.moota_sig_verified_by || "").trim(),
    userAgent: String(params.moota_user_agent || "").trim(),
    mootaUser: String(params.moota_user || "").trim(),
    mootaWebhook: String(params.moota_webhook || "").trim(),
    paramKeys: Object.keys(params)
  };
}

function logMootaSignatureEvent_(type, meta, extra) {
  try {
    const payload = Object.assign({
      source: meta && meta.source ? meta.source : "missing",
      forwarded_by_worker: !!(meta && meta.forwardedByWorker),
      worker_saw_signature: !!(meta && meta.workerSawSignature),
      worker_verified_signature: !!(meta && meta.workerVerifiedSignature),
      worker_verification_source: meta && meta.workerVerificationSource ? meta.workerVerificationSource : "",
      received_signature: maskMootaSignatureForLog_(meta && meta.normalized),
      signature_length: meta && meta.normalized ? meta.normalized.length : 0,
      moota_user: meta && meta.mootaUser ? String(meta.mootaUser).substring(0, 40) : "",
      moota_webhook: meta && meta.mootaWebhook ? String(meta.mootaWebhook).substring(0, 40) : "",
      user_agent: meta && meta.userAgent ? String(meta.userAgent).substring(0, 120) : "",
      param_keys: meta && meta.paramKeys ? meta.paramKeys.slice(0, 10).join(",") : ""
    }, extra || {});
    logMoota_(type, JSON.stringify(payload));
  } catch (err) {
    Logger.log("logMootaSignatureEvent_ error: " + err);
  }
}

function classifyMootaSignatureMissing_(mootaCfg, meta) {
  if (isDirectAppsScriptUrl_(mootaCfg && mootaCfg.gasUrl)) {
    return {
      code: "direct_apps_script_url",
      message: "ERROR: Missing Signature. Webhook Moota masih diarahkan langsung ke Google Apps Script. Gunakan endpoint Cloudflare Worker/proxy publik sebagai URL webhook di dashboard Moota."
    };
  }
  if (!(meta && meta.forwardedByWorker)) {
    return {
      code: "worker_not_detected",
      message: "ERROR: Missing Signature. Request tidak terlihat datang dari Worker/proxy. Pastikan URL webhook Moota mengarah ke endpoint Worker/proxy yang aktif dan terbaru."
    };
  }
  if (meta.forwardedByWorker && !meta.workerSawSignature) {
    return {
      code: "worker_missing_signature_header",
      message: "ERROR: Missing Signature. Worker/proxy menerima request tetapi header Signature dari Moota tidak ditemukan. Periksa pengaturan webhook Moota dan deploy Worker versi terbaru."
    };
  }
  return {
    code: "signature_not_forwarded",
    message: "ERROR: Missing Signature. Header Signature dari Moota tidak berhasil diteruskan ke Apps Script. Pastikan Worker meneruskan query param `moota_signature` atau `signature`."
  };
}

function appendQueryParams_(url, params) {
  const base = String(url || "").trim();
  if (!base) return "";
  const entries = [];
  const source = params || {};
  for (let key in source) {
    if (!Object.prototype.hasOwnProperty.call(source, key)) continue;
    const value = source[key];
    if (value === undefined || value === null || String(value) === "") continue;
    entries.push(encodeURIComponent(String(key)) + "=" + encodeURIComponent(String(value)));
  }
  if (!entries.length) return base;
  return base + (base.indexOf("?") === -1 ? "?" : "&") + entries.join("&");
}

function normalizeImageKitEndpoint_(raw) {
  const value = String(raw || "").trim().replace(/\/+$/, "");
  if (!value) return "";
  return value;
}

function isValidImageKitPublicKey_(value) {
  return /^public_[A-Za-z0-9+/=._-]+$/.test(String(value || "").trim());
}

function isValidImageKitPrivateKey_(value) {
  return /^private_[A-Za-z0-9+/=._-]+$/.test(String(value || "").trim());
}

function isValidImageKitEndpoint_(value) {
  const endpoint = normalizeImageKitEndpoint_(value);
  if (!endpoint) return false;
  if (!/^https:\/\/[^\s/$.?#].[^\s]*$/i.test(endpoint)) return false;
  if (/[?#]/.test(endpoint)) return false;
  return true;
}

function resolveImageKitConfig_(data, cfg) {
  const payload = data || {};
  return {
    publicKey: String((payload.ik_public_key !== undefined ? payload.ik_public_key : getCfgFrom_(cfg, "ik_public_key")) || "").trim(),
    endpoint: normalizeImageKitEndpoint_(payload.ik_endpoint !== undefined ? payload.ik_endpoint : getCfgFrom_(cfg, "ik_endpoint")),
    privateKey: String((payload.ik_private_key !== undefined ? payload.ik_private_key : getSecret_("ik_private_key", cfg)) || "").trim()
  };
}

function validateImageKitConfigFormat_(ikCfg, opts) {
  const options = opts || {};
  const errors = [];
  const requirePublic = options.requirePublic !== false;
  const requireEndpoint = options.requireEndpoint !== false;
  const requirePrivate = options.requirePrivate !== false;

  if (requirePublic && !ikCfg.publicKey) errors.push("ImageKit public key wajib diisi.");
  if (requirePublic && ikCfg.publicKey && !isValidImageKitPublicKey_(ikCfg.publicKey)) {
    errors.push("Format ImageKit public key tidak valid. Harus diawali dengan 'public_'.");
  }

  if (requireEndpoint && !ikCfg.endpoint) errors.push("ImageKit URL endpoint wajib diisi.");
  if (requireEndpoint && ikCfg.endpoint && !isValidImageKitEndpoint_(ikCfg.endpoint)) {
    errors.push("Format ImageKit URL endpoint tidak valid. Gunakan URL HTTPS seperti https://ik.imagekit.io/nama-endpoint");
  }

  if (requirePrivate && !ikCfg.privateKey) errors.push("ImageKit private key wajib diisi.");
  if (requirePrivate && ikCfg.privateKey && !isValidImageKitPrivateKey_(ikCfg.privateKey)) {
    errors.push("Format ImageKit private key tidak valid. Harus diawali dengan 'private_'.");
  }

  return errors;
}

function inferImageKitEndpointFromUrl_(fileUrl) {
  const value = String(fileUrl || "").trim();
  if (!value) return "";
  const match = value.match(/^https:\/\/([^\/?#]+)(\/[^?#]*)?/i);
  if (!match) return "";
  const host = String(match[1] || "").toLowerCase();
  const path = String(match[2] || "");
  if (!host) return "";
  if (host === "ik.imagekit.io") {
    const firstSegment = path.split("/").filter(Boolean)[0] || "";
    if (firstSegment) return "https://ik.imagekit.io/" + firstSegment;
  }
  return "https://" + host;
}

function fetchImageKitFiles_(privateKey, limit) {
  try {
    const authHeader = "Basic " + Utilities.base64Encode(String(privateKey || "").trim() + ":");
    const url = "https://api.imagekit.io/v1/files?sort=DESC_CREATED&limit=" + Number(limit || 20);
    const res = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "Authorization": authHeader },
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const text = res.getContentText();
    let data = null;
    try { data = JSON.parse(text); } catch (e) {}

    if (code >= 200 && code < 300 && Array.isArray(data)) {
      return { ok: true, files: data };
    }

    let message = "Gagal terhubung ke ImageKit.";
    if (code === 401) {
      message = "Autentikasi ImageKit gagal. Periksa private key Anda.";
    } else if (data && data.message) {
      message = "ImageKit error: " + data.message;
    } else if (text) {
      message = "ImageKit error HTTP " + code + ": " + String(text).substring(0, 200);
    }

    return { ok: false, code: code, message: message };
  } catch (e) {
    return { ok: false, code: 0, message: "Koneksi ke ImageKit gagal: " + e.toString() };
  }
}

function assertPrivilegedAction_(data, cfg) {
  if (isDebugAllowed_()) return true;
  const supplied = String((data && data.admin_token) || "").trim();
  const expected = getSecret_("ADMIN_API_TOKEN", cfg || getSettingsMap_());
  if (expected && supplied === expected) return true;
  throw new Error("Unauthorized diagnostic action");
}

/* =========================
   LEGACY getCfg (kept)
   (masih bisa dipakai, tapi lebih lambat)
========================= */
function getCfg(name) {
  try {
    const s = ss.getSheetByName("Settings");
    const d = s.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][0]).trim() === name) return d[i][1];
    }
  } catch (e) { return ""; }
  return "";
}



/* =========================
   WEBHOOK ENTRYPOINT
========================= */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonRes({ status: "error", message: "No data" });
    }

    const cfg = getSettingsMap_();



    const payloadString = e.postData.contents;
    let data = null;
    try {
       data = JSON.parse(payloadString);
    } catch(err) {
       // Ignore JSON parse error, maybe it was not JSON but handled above or invalid
       return jsonRes({ status: "error", message: "Invalid JSON format" });
    }

    // ====================================================================
    // 🚀 RADAR MOOTA: DETEKSI WEBHOOK MASUK + VERIFIKASI SIGNATURE
    // ====================================================================
    if (Array.isArray(data) && data.length > 0 && data[0].amount !== undefined) {
      const mootaCfg = resolveMootaConfig_({}, cfg);
      const signatureMeta = extractMootaSignatureMeta_(e);
      if (!mootaCfg.gasUrl) {
        logMootaSignatureEvent_("SIGNATURE_CONFIG_ERROR", signatureMeta, {
          reason: "missing_webhook_url",
          payload_bytes: payloadString.length
        });
        return ContentService.createTextOutput("ERROR: Link webhook Moota belum dikonfigurasi.")
          .setMimeType(ContentService.MimeType.TEXT);
      }
      if (!mootaCfg.token) {
        logMootaSignatureEvent_("SIGNATURE_CONFIG_ERROR", signatureMeta, {
          reason: "missing_secret_token",
          payload_bytes: payloadString.length,
          webhook_host: getMootaUrlHost_(mootaCfg.gasUrl)
        });
        return ContentService.createTextOutput("ERROR: Secret Token Moota belum dikonfigurasi.")
          .setMimeType(ContentService.MimeType.TEXT);
      }

      // Apps Script tidak menerima custom header mentah dari webhook,
      // jadi proxy/Worker perlu meneruskan header Signature ke query param ini.
      if (!signatureMeta.normalized) {
        const missingSignature = classifyMootaSignatureMissing_(mootaCfg, signatureMeta);
        logMootaSignatureEvent_("SIGNATURE_MISSING", signatureMeta, {
          reason: missingSignature.code,
          payload_bytes: payloadString.length,
          webhook_host: getMootaUrlHost_(mootaCfg.gasUrl),
          likely_direct_apps_script: isDirectAppsScriptUrl_(mootaCfg.gasUrl),
          troubleshooting_hint: isDirectAppsScriptUrl_(mootaCfg.gasUrl)
            ? "Webhook Moota diarahkan langsung ke Google Apps Script. Gunakan endpoint Worker/proxy publik."
            : "Pastikan header Signature dari Moota diteruskan ke query param moota_signature atau signature."
        });
        return ContentService.createTextOutput(missingSignature.message)
          .setMimeType(ContentService.MimeType.TEXT);
      }

      const signatureCheck = verifyMootaSignature_(payloadString, mootaCfg.token, signatureMeta.raw);
      if (!signatureCheck.ok) {
        const invalidSignatureMessage = signatureMeta.workerVerifiedSignature
          ? "ERROR: Invalid Signature. Signature sudah lolos verifikasi di Worker, jadi kemungkinan Secret Token di Apps Script berbeda dengan Secret Token di Worker/Moota."
          : "ERROR: Invalid Signature. Periksa Secret Token dan pastikan payload tidak diubah sebelum diverifikasi.";
        logMootaSignatureEvent_("SIGNATURE_INVALID", signatureMeta, {
          payload_bytes: payloadString.length,
          webhook_host: getMootaUrlHost_(mootaCfg.gasUrl),
          expected_signature: maskMootaSignatureForLog_(signatureCheck.expected),
          received_signature: maskMootaSignatureForLog_(signatureCheck.received),
          validation_code: signatureCheck.code,
          worker_verified_signature: signatureMeta.workerVerifiedSignature
        });
        return ContentService.createTextOutput(invalidSignatureMessage)
          .setMimeType(ContentService.MimeType.TEXT);
      }

      logMootaSignatureEvent_("SIGNATURE_OK", signatureMeta, {
        payload_bytes: payloadString.length,
        webhook_host: getMootaUrlHost_(mootaCfg.gasUrl)
      });

      const isMootaTest = String((e.parameter && e.parameter.test_mode) || "").trim() === "1"
        || data.some(function(item) {
          return item && (item.is_test === true || String(item.description || "").toUpperCase() === "MOOTA TEST");
        });
      if (isMootaTest) {
        return jsonRes({
          status: "success",
          message: "Test webhook Moota berhasil.",
          secret_token_configured: true,
          signature_verified: true,
          signature_source: signatureMeta.source,
          forwarded_by_worker: signatureMeta.forwardedByWorker,
          mutations_received: data.length
        });
      }

      return handleMootaWebhook(data, cfg);
    }

    // ====================================================================
    // JIKA BUKAN DARI MOOTA, JALANKAN PERINTAH DARI WEBSITE (FRONTEND)
    // ====================================================================
    const action = normalizeActionAlias_(data.action);
    switch (action) {
      case "get_global_settings": return jsonRes(getGlobalSettings(cfg));
      case "get_product": return jsonRes(getProductDetail(data, cfg));
      case "get_products": return jsonRes(getProducts(data, cfg));
      case "get_product_categories": return jsonRes(getHerbalCategories(data, cfg));
      case "get_products_public": return jsonRes(getHerbalProductsPublic(data, cfg));
      case "get_course_categories": return jsonRes(getCourseCategories(data, cfg));
      case "get_courses_public": return jsonRes(getCoursesPublic(data, cfg));
      case "get_course_by_slug": return jsonRes(getCourseBySlug(data, cfg));
      case "get_course_public": return jsonRes(getCoursePublic(data, cfg));
      case "get_schedules_public": return jsonRes(getSchedulesPublic(data, cfg));
      case "get_schedule_public": return jsonRes(getSchedulePublic(data, cfg));
      case "get_instructors_public": return jsonRes(getInstructorsPublic(data, cfg));
      case "get_branches_public": return jsonRes(getBranchesPublic(data, cfg));
      case "get_testimonials_public": return jsonRes(getTestimonialsPublic(data, cfg));
      case "get_galleries_public": return jsonRes(getGalleriesPublic(data, cfg));
      case "get_faqs_public": return jsonRes(getFaqsPublic(data, cfg));
      case "get_settings_public": return jsonRes(getSettingsPublic(cfg));
      case "get_product_by_slug": return jsonRes(getHerbalProductBySlug(data, cfg));
      case "get_articles": return jsonRes(getHerbalArticles(data, cfg));
      case "get_article_by_slug": return jsonRes(getHerbalArticleBySlug(data, cfg));
      case "get_testimonials": return jsonRes(getHerbalTestimonials(data, cfg));
      case "get_banners": return jsonRes(getHerbalBanners(data, cfg));
      case "get_faqs": return jsonRes(getHerbalFaqs(data, cfg));
      case "get_public_settings": return jsonRes(getHerbalPublicSettings(cfg));
      case "create_inquiry": return jsonRes(createInquiry(data, cfg));
      case "create_registration": return jsonRes(createRegistration(data, cfg));
      case "create_order": return jsonRes(createOrder(data, cfg));
      case "update_order_status": return jsonRes(updateOrderStatus(data, cfg));
      case "login": return jsonRes(loginUser(data));
      case "login_and_dashboard": return jsonRes(loginAndDashboard(data));
      case "get_page_content": return jsonRes(getPageContent(data));
      case "get_pages": return jsonRes(getAllPages(data));
      case "get_public_cache_state": return jsonRes(getPublicCacheState());
      case "admin_login": return jsonRes(adminLogin(data));
      case "admin_logout": return jsonRes(adminLogout(data));
      case "get_admin_data": return jsonRes(getAdminData(data, cfg));
      case "get_admin_dashboard_summary": return jsonRes(getAdminDashboardSummary(data, cfg));
      case "get_admin_products_v2": return jsonRes(adminListEntity_(data, "products"));
      case "save_admin_product_v2": return jsonRes(adminSaveEntity_(data, "products", "id"));
      case "delete_admin_product_v2": return jsonRes(adminDeleteEntity_(data, "products", "id"));
      case "get_admin_courses_v2": return jsonRes(adminListEntity_(data, "courses"));
      case "save_admin_course_v2": return jsonRes(adminSaveEntity_(data, "courses", "id"));
      case "delete_admin_course_v2": return jsonRes(adminDeleteEntity_(data, "courses", "id"));
      case "get_admin_categories": return jsonRes(adminListEntity_(data, "product_categories"));
      case "save_admin_category": return jsonRes(adminSaveEntity_(data, "product_categories", "id"));
      case "delete_admin_category": return jsonRes(adminDeleteEntity_(data, "product_categories", "id"));
      case "get_admin_course_categories_v2": return jsonRes(adminListEntity_(data, "course_categories"));
      case "save_admin_course_category_v2": return jsonRes(adminSaveEntity_(data, "course_categories", "id"));
      case "delete_admin_course_category_v2": return jsonRes(adminDeleteEntity_(data, "course_categories", "id"));
      case "get_admin_schedules_v2": return jsonRes(adminListEntity_(data, "schedules"));
      case "save_admin_schedule_v2": return jsonRes(adminSaveEntity_(data, "schedules", "id"));
      case "delete_admin_schedule_v2": return jsonRes(adminDeleteEntity_(data, "schedules", "id"));
      case "get_admin_instructors_v2": return jsonRes(adminListEntity_(data, "instructors"));
      case "save_admin_instructor_v2": return jsonRes(adminSaveEntity_(data, "instructors", "id"));
      case "delete_admin_instructor_v2": return jsonRes(adminDeleteEntity_(data, "instructors", "id"));
      case "get_admin_branches_v2": return jsonRes(adminListEntity_(data, "branches"));
      case "save_admin_branch_v2": return jsonRes(adminSaveEntity_(data, "branches", "id"));
      case "delete_admin_branch_v2": return jsonRes(adminDeleteEntity_(data, "branches", "id"));
      case "get_admin_registrations_v2": return jsonRes(adminListEntity_(data, "registrations"));
      case "save_admin_registration_v2": return jsonRes(adminSaveEntity_(data, "registrations", "id"));
      case "patch_admin_registration_status_v2": return jsonRes(adminSaveEntity_(data, "registrations", "id"));
      case "get_admin_galleries_v2": return jsonRes(adminListEntity_(data, "galleries"));
      case "save_admin_gallery_v2": return jsonRes(adminSaveEntity_(data, "galleries", "id"));
      case "delete_admin_gallery_v2": return jsonRes(adminDeleteEntity_(data, "galleries", "id"));
      case "get_admin_content_blocks_v2": return jsonRes(adminListEntity_(data, "content_blocks"));
      case "save_admin_content_block_v2": return jsonRes(adminSaveEntity_(data, "content_blocks", "id"));
      case "delete_admin_content_block_v2": return jsonRes(adminDeleteEntity_(data, "content_blocks", "id"));
      case "get_admin_users_v2": return jsonRes(adminListEntity_(data, "admin_users"));
      case "save_admin_user_v2": return jsonRes(adminSaveEntity_(data, "admin_users", "id"));
      case "delete_admin_user_v2": return jsonRes(adminDeleteEntity_(data, "admin_users", "id"));
      case "get_admin_articles_v2": return jsonRes(adminListEntity_(data, "articles"));
      case "save_admin_article_v2": return jsonRes(adminSaveEntity_(data, "articles", "id"));
      case "delete_admin_article_v2": return jsonRes(adminDeleteEntity_(data, "articles", "id"));
      case "get_admin_testimonials_v2": return jsonRes(adminListEntity_(data, "testimonials"));
      case "save_admin_testimonial_v2": return jsonRes(adminSaveEntity_(data, "testimonials", "id"));
      case "delete_admin_testimonial_v2": return jsonRes(adminDeleteEntity_(data, "testimonials", "id"));
      case "get_admin_inquiries_v2": return jsonRes(adminListEntity_(data, "inquiries"));
      case "update_admin_inquiry_v2": return jsonRes(adminSaveEntity_(data, "inquiries", "id"));
      case "get_admin_banners_v2": return jsonRes(adminListEntity_(data, "banners"));
      case "save_admin_banner_v2": return jsonRes(adminSaveEntity_(data, "banners", "id"));
      case "delete_admin_banner_v2": return jsonRes(adminDeleteEntity_(data, "banners", "id"));
      case "get_admin_faqs_v2": return jsonRes(adminListEntity_(data, "faq_items"));
      case "save_admin_faq_v2": return jsonRes(adminSaveEntity_(data, "faq_items", "id"));
      case "delete_admin_faq_v2": return jsonRes(adminDeleteEntity_(data, "faq_items", "id"));
      case "seed_herbal_demo_data": return jsonRes(seedHerbalDemoData(data));
      case "seed_course_demo_data": return jsonRes(seedCourseDemoData(data));
      case "save_product": return jsonRes(saveProduct(data));
      case "save_page": return jsonRes(savePage(data));
      case "update_settings": return jsonRes(updateSettings(data));
      case "update_moota_gateway": return jsonRes(updateMootaGatewaySettings(data));
      case "update_imagekit_media": return jsonRes(updateImageKitMediaSettings(data));
      case "import_moota_config": return jsonRes(importMootaConfig(data));
      case "get_ik_auth": return jsonRes(getImageKitAuth(data, cfg));
      case "get_media_files": return jsonRes(getIkFiles(data, cfg));
      case "test_ik_config": return jsonRes(testImageKitConfig(data, cfg));
      case "test_moota_config": return jsonRes(testMootaConfig(data, cfg));
      case "purge_cf_cache": return jsonRes(purgeCFCache(data, cfg));
      case "change_password": return jsonRes(changeUserPassword(data));
      case "update_profile": return jsonRes(updateUserProfile(data));
      case "forgot_password": return jsonRes(forgotPassword(data));
      case "get_dashboard_data": return jsonRes(getDashboardData(data));
      case "delete_product": return jsonRes(deleteProduct(data));
      case "delete_page": return jsonRes(deletePage(data));
      case "check_slug": return jsonRes(checkSlug(data));
      case "save_affiliate_pixel": return jsonRes(saveAffiliatePixel(data));
      case "get_admin_orders": return jsonRes(getAdminOrders(data));
      case "get_admin_users": return jsonRes(getAdminUsers(data));

      // DIAGNOSTIC & MONITORING ACTIONS
      case "get_email_logs":
      case "get_moota_logs":
      case "get_wa_logs":
      case "test_email":
      case "test_wa":
      case "test_lunas_notification":
      case "get_system_health":
      case "get_email_quota":
      case "debug_login":
      case "test_auth":
      case "test_moota_validation":
      case "test_moota_signature":
      case "purge_sync_logs":
      case "audit_sync_logs_cleanup":
        assertPrivilegedAction_(data, cfg);
        if (action === "get_email_logs") return jsonRes(getEmailLogs_());
        if (action === "get_moota_logs") return jsonRes(getMootaLogs_());
        if (action === "get_wa_logs") return jsonRes(getWALogs_());
        if (action === "test_email") return jsonRes(testEmailDelivery(data));
        if (action === "test_wa") return jsonRes(testWADelivery(data));
        if (action === "test_lunas_notification") return jsonRes(testLunasNotification(data));
        if (action === "get_system_health") return jsonRes(getSystemHealth());
        if (action === "get_email_quota") return jsonRes(getEmailQuotaStatus());
        if (action === "debug_login") return jsonRes(debugLogin(data));
        if (action === "test_auth") return jsonRes(runAuthTests());
        if (action === "test_moota_validation") return jsonRes(runMootaValidationTests());
        if (action === "test_moota_signature") return jsonRes(runMootaSignatureTests());
        if (action === "purge_sync_logs") return jsonRes(purgeSyncLogsArtifacts_(false, data));
        if (action === "audit_sync_logs_cleanup") return jsonRes(purgeSyncLogsArtifacts_(true, data));
        return jsonRes({ status: "error", message: "Unsupported privileged action" });

      default: return jsonRes({ status: "error", message: "Aksi tidak terdaftar: " + (action || "unknown") });
    }
  } catch (err) {
    return jsonRes({ status: "error", message: err.toString() });
  }
}



/* =========================
   WHITE-LABEL GLOBAL SETTINGS
========================= */
function getGlobalSettings(cfg) {
  cfg = cfg || getSettingsMap_();
  return withPublicCacheVersion_({
    status: "success",
    data: {
      site_name: getCfgFrom_(cfg, "site_name") || "NaturaHerb Wellness",
      site_tagline: getCfgFrom_(cfg, "site_tagline") || "Herbal alami untuk dukungan kesehatan harian.",
      site_favicon: sanitizeAssetUrl_(getCfgFrom_(cfg, "site_favicon") || ""),
      site_logo: sanitizeAssetUrl_(getCfgFrom_(cfg, "site_logo") || ""),
      contact_email: getCfgFrom_(cfg, "contact_email") || "",
      wa_admin: getCfgFrom_(cfg, "wa_admin") || ""
    }
  }, "settings");
}

/* =========================
   CLOUDFLARE PURGE
========================= */
function purgeCFCache(d, cfg) {
  try {
    requireAdminSession_(d, { actionName: "purge_cf_cache" });
    cfg = cfg || getSettingsMap_();
    const zoneId = getSecret_("cf_zone_id", cfg);
    const token = getSecret_("cf_api_token", cfg);
    if (!zoneId || !token) return { status: "error", message: "Konfigurasi Cloudflare belum disetting!" };

    const options = {
      method: "post",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify({ purge_everything: true }),
      muteHttpExceptions: true
    };

    const res = UrlFetchApp.fetch(`https://api.cloudflare.com/client/v4/zones/${zoneId}/purge_cache`, options);
    const body = JSON.parse(res.getContentText());

    if (body && body.success) {
      return { status: "success", message: "🚀 Cache Berhasil Dibersihkan!" };
    }
    const msg = (body && body.errors && body.errors.length) ? JSON.stringify(body.errors) : "Cloudflare Error";
    return { status: "error", message: msg };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function testMootaConfig(d, cfg) {
  requireAdminSession_(d, { allowDemo: false, actionName: "test_moota_config" });
  cfg = cfg || getSettingsMap_();
  const mootaCfg = resolveMootaConfig_(d, cfg);
  const errors = validateMootaConfigFormat_(mootaCfg);
  if (errors.length) {
    return { status: "error", message: errors[0], errors: errors };
  }

  if (isDirectAppsScriptUrl_(mootaCfg.gasUrl)) {
    logMoota_("CONFIG_TEST_BLOCKED", JSON.stringify({
      reason: "direct_apps_script_url",
      webhook_host: getMootaUrlHost_(mootaCfg.gasUrl),
      troubleshooting_hint: "Gunakan endpoint Cloudflare Worker atau proxy publik, bukan URL Google Apps Script langsung."
    }));
    return {
      status: "error",
      message: "Link webhook Moota tidak boleh langsung ke Google Apps Script. Gunakan endpoint Cloudflare Worker atau proxy publik agar header Signature bisa diteruskan.",
      code: "direct_apps_script_url"
    };
  }

  const payloadText = JSON.stringify([{
    amount: 1,
    type: "CR",
    description: "MOOTA TEST",
    is_test: true,
    created_at: new Date().toISOString()
  }]);

  const signature = computeMootaSignatureHex_(payloadText, mootaCfg.token);

  const targetUrl = appendQueryParams_(mootaCfg.gasUrl, {
    test_mode: "1",
    moota_signature: signature
  });

  try {
    const res = UrlFetchApp.fetch(targetUrl, {
      method: "post",
      contentType: "application/json",
      headers: {
        "Signature": signature,
        "X-MOOTA-USER": "test-user",
        "X-MOOTA-WEBHOOK": "test-webhook",
        "User-Agent": "MootaBot/1.5"
      },
      payload: payloadText,
      muteHttpExceptions: true,
      followRedirects: true
    });
    const code = res.getResponseCode();
    const text = res.getContentText();
    let data = null;
    try { data = JSON.parse(text); } catch (e) {}

    if (code >= 200 && code < 300 && data && data.status === "success") {
      return {
        status: "success",
        message: "Koneksi webhook Moota berhasil diuji.",
        gas_url: mootaCfg.gasUrl,
        secret_token_configured: !!mootaCfg.token,
        signature_preview: maskMootaSignatureForLog_(signature),
        response: data
      };
    }

    let message = "Test koneksi Moota gagal.";
    if (data && data.message) {
      message = String(data.message);
    } else if (text) {
      message = "Webhook Moota error HTTP " + code + ": " + String(text).substring(0, 200);
    }

    return {
      status: "error",
      message: message,
      http_code: code
    };
  } catch (e) {
    return {
      status: "error",
      message: "Gagal menghubungi webhook Moota: " + e.toString()
    };
  }
}

function getIkFiles(d, cfg) {
  requireAdminSession_(d, { actionName: "get_media_files" });
  cfg = cfg || getSettingsMap_();
  const ikCfg = resolveImageKitConfig_({}, cfg);
  const errors = validateImageKitConfigFormat_(ikCfg, { requirePublic: false, requireEndpoint: false, requirePrivate: true });
  if (errors.length) return { status: "error", message: errors[0] };

  const result = fetchImageKitFiles_(ikCfg.privateKey, 20);
  if (!result.ok) return { status: "error", message: result.message };

  const files = result.files.map(function(f) {
    return {
      name: f.name,
      url: f.url,
      thumbnail: f.thumbnailUrl || f.url,
      fileId: f.fileId,
      type: f.fileType
    };
  });
  return { status: "success", files: files };
}

/* =========================
   LOGGING HELPERS
========================= */
function logEmail_(status, to, subject, detail) {
  try {
    let s = ss.getSheetByName("Email_Logs");
    if (!s) {
      s = ss.insertSheet("Email_Logs");
      s.appendRow(["Timestamp", "Status", "To", "Subject", "Detail"]);
      s.setFrozenRows(1);
    }
    s.appendRow([new Date(), status, to, subject, String(detail).substring(0, 500)]);
    // Auto-trim: keep max 500 rows
    if (s.getLastRow() > 500) s.deleteRows(2, s.getLastRow() - 500);
  } catch (e) {
    Logger.log("logEmail_ error: " + e);
  }
}

function logMoota_(type, detail) {
  try {
    let s = ss.getSheetByName("Moota_Logs");
    if (!s) {
      s = ss.insertSheet("Moota_Logs");
      s.appendRow(["Timestamp", "Type", "Detail"]);
      s.setFrozenRows(1);
    }
    s.appendRow([new Date(), type, String(detail).substring(0, 1000)]);
    // Auto-trim: keep max 500 rows
    if (s.getLastRow() > 500) s.deleteRows(2, s.getLastRow() - 500);
  } catch (e) {
    Logger.log("logMoota_ error: " + e);
  }
}

function logWA_(status, target, detail) {
  try {
    let s = ss.getSheetByName("WA_Logs");
    if (!s) {
      s = ss.insertSheet("WA_Logs");
      s.appendRow(["Timestamp", "Status", "Target", "Detail"]);
      s.setFrozenRows(1);
    }
    s.appendRow([new Date(), status, target, String(detail).substring(0, 500)]);
    if (s.getLastRow() > 500) s.deleteRows(2, s.getLastRow() - 500);
  } catch (e) {
    Logger.log("logWA_ error: " + e);
  }
}

function invalidateCaches_(keys) {
  try {
    const cache = CacheService.getScriptCache();
    (keys || []).forEach(k => {
      try { cache.remove(String(k)); } catch (e) { }
    });
  } catch (e) { }
}

function referencesSyncLogs_(text) {
  return /(^|[^a-z0-9])sync[_\s]?logs([^a-z0-9]|$)/i.test(String(text || ""));
}

function normalizeEmailSafe_(value) {
  return String(value || "").trim().toLowerCase();
}

function buildSyncLogsBackup_(sheet, report) {
  const backupName = "Sync_Logs_Backup_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const backupSs = SpreadsheetApp.create(backupName);
  const manifest = backupSs.getSheetByName("Sheet1") || backupSs.getSheets()[0];
  manifest.setName("Manifest");
  manifest.clear();
  manifest.appendRow(["Section", "Key", "Value"]);
  manifest.appendRow(["summary", "source_spreadsheet_id", ss.getId()]);
  manifest.appendRow(["summary", "source_spreadsheet_name", ss.getName()]);
  manifest.appendRow(["summary", "generated_at", new Date().toISOString()]);
  manifest.appendRow(["summary", "sheet_found", String(!!sheet)]);
  manifest.appendRow(["summary", "formulas_detected", String(report.formulas_detected || 0)]);
  manifest.appendRow(["summary", "protections_detected", String(report.protections_removed || 0)]);
  manifest.appendRow(["summary", "metadata_detected", String(report.metadata_removed || 0)]);

  if (sheet) {
    const copied = sheet.copyTo(backupSs);
    copied.setName("Sync_Logs");
  }

  const rows = [];
  (report.formula_locations || []).forEach(function (item) {
    rows.push(["formula", item.sheet + "!" + item.cell, item.formula]);
  });
  (report.named_ranges_removed || []).forEach(function (name) {
    rows.push(["named_range", name, "removed"]);
  });
  (report.triggers_removed || []).forEach(function (item) {
    rows.push(["trigger", item.handler, item.event_type]);
  });
  (report.script_properties_removed || []).forEach(function (name) {
    rows.push(["script_property", name, "removed"]);
  });
  (report.permission_snapshot || []).forEach(function (item) {
    rows.push(["permission", item.role, item.email]);
  });
  (report.notes || []).forEach(function (note, idx) {
    rows.push(["note", String(idx + 1), note]);
  });

  if (rows.length > 0) {
    manifest.getRange(2 + 7, 1, rows.length, 3).setValues(rows);
  }

  const sheets = backupSs.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === "Sheet1" && sheets.length > 1) {
      backupSs.deleteSheet(sheets[i]);
      break;
    }
  }

  return {
    id: backupSs.getId(),
    url: backupSs.getUrl(),
    name: backupSs.getName()
  };
}

function captureFilePermissions_() {
  const snapshot = [];
  try {
    const file = DriveApp.getFileById(ss.getId());
    const owner = file.getOwner();
    if (owner) snapshot.push({ role: "owner", email: normalizeEmailSafe_(owner.getEmail()) });
    file.getEditors().forEach(function (user) {
      snapshot.push({ role: "editor", email: normalizeEmailSafe_(user.getEmail()) });
    });
    file.getViewers().forEach(function (user) {
      snapshot.push({ role: "viewer", email: normalizeEmailSafe_(user.getEmail()) });
    });
  } catch (e) { }
  return snapshot.filter(function (item) { return !!item.email; });
}

function revokeFilePermissions_(options, report, dryRun) {
  const cfg = options || {};
  const shouldRevoke = !!cfg.revoke_file_access;
  report.permission_snapshot = captureFilePermissions_();
  if (!shouldRevoke) {
    report.notes.push("Spreadsheet-wide Drive sharing tidak diubah otomatis. Set revoke_file_access=true dan kirim revoke_access_emails jika memang ingin mencabut akses file.");
    return;
  }

  const revokeList = Array.isArray(cfg.revoke_access_emails) ? cfg.revoke_access_emails.map(normalizeEmailSafe_).filter(Boolean) : [];
  const keepList = Array.isArray(cfg.keep_access_emails) ? cfg.keep_access_emails.map(normalizeEmailSafe_).filter(Boolean) : [];
  if (revokeList.length === 0) {
    report.notes.push("revoke_file_access=true tapi revoke_access_emails kosong, jadi tidak ada akses file yang dicabut.");
    return;
  }

  try {
    const file = DriveApp.getFileById(ss.getId());
    const ownerEmail = normalizeEmailSafe_(file.getOwner() && file.getOwner().getEmail());
    revokeList.forEach(function (email) {
      if (!email || email === ownerEmail || keepList.indexOf(email) !== -1) return;
      report.permissions_revoked.push(email);
      if (dryRun) return;
      try { file.removeEditor(email); } catch (e) { }
      try { file.removeViewer(email); } catch (e) { }
    });
  } catch (e) {
    report.notes.push("Gagal memproses revokasi akses file: " + String(e));
  }
}

function purgeSyncLogsArtifacts_(dryRun, options) {
  try {
    const cfg = options || {};
    const runMode = dryRun ? "dry_run" : "delete";
    const report = {
      status: "success",
      mode: runMode,
      sheet_found: false,
      sheet_deleted: false,
      formulas_replaced: 0,
      formulas_detected: 0,
      formula_locations: [],
      named_ranges_removed: [],
      protections_removed: 0,
      triggers_removed: [],
      script_properties_removed: [],
      metadata_removed: 0,
      permissions_revoked: [],
      permission_snapshot: [],
      backup_created: false,
      backup_id: "",
      backup_url: "",
      notes: []
    };

    const sheet = ss.getSheetByName("Sync_Logs");
    report.sheet_found = !!sheet;

    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const sh = sheets[i];
      const range = sh.getDataRange();
      if (!range) continue;
      const formulas = range.getFormulas();
      const values = range.getValues();
      for (let r = 0; r < formulas.length; r++) {
        for (let c = 0; c < formulas[r].length; c++) {
          const f = String(formulas[r][c] || "").trim();
          if (!f || !referencesSyncLogs_(f)) continue;
          report.formulas_detected++;
          if (report.formula_locations.length < 100) {
            report.formula_locations.push({
              sheet: sh.getName(),
              cell: range.getCell(r + 1, c + 1).getA1Notation(),
              formula: f.substring(0, 200)
            });
          }
          if (!dryRun) {
            range.getCell(r + 1, c + 1).setValue(values[r][c]);
            report.formulas_replaced++;
          }
        }
      }
    }

    const namedRanges = ss.getNamedRanges();
    for (let i = 0; i < namedRanges.length; i++) {
      const nr = namedRanges[i];
      let targetSheet = "";
      try { targetSheet = nr.getRange().getSheet().getName(); } catch (e) { }
      const matched = referencesSyncLogs_(nr.getName()) || referencesSyncLogs_(targetSheet);
      if (!matched) continue;
      report.named_ranges_removed.push(nr.getName());
      if (!dryRun) nr.remove();
    }

    const metadataItems = ss.getDeveloperMetadata();
    for (let i = 0; i < metadataItems.length; i++) {
      const md = metadataItems[i];
      const mk = String(md.getKey() || "");
      const mv = String(md.getValue() || "");
      if (!referencesSyncLogs_(mk) && !referencesSyncLogs_(mv)) continue;
      if (!dryRun) md.remove();
      report.metadata_removed++;
    }

    if (sheet) {
      const sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (let i = 0; i < sheetProtections.length; i++) {
        if (!dryRun) sheetProtections[i].remove();
        report.protections_removed++;
      }
      const rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (let i = 0; i < rangeProtections.length; i++) {
        if (!dryRun) rangeProtections[i].remove();
        report.protections_removed++;
      }
    }

    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      const t = triggers[i];
      const handler = String(t.getHandlerFunction() || "");
      if (!/(sync[_\s-]?logs?|sync[_\s-]?state|cepat[_\s-]?sync)/i.test(handler)) continue;
      report.triggers_removed.push({
        handler: handler,
        event_type: String(t.getEventType())
      });
      if (!dryRun) ScriptApp.deleteTrigger(t);
    }

    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    Object.keys(allProps).forEach(function (k) {
      const key = String(k || "");
      const val = String(allProps[k] || "");
      if (!referencesSyncLogs_(key) && !referencesSyncLogs_(val) && !/sync_state|cepat_sync/i.test(key)) return;
      report.script_properties_removed.push(key);
      if (!dryRun) props.deleteProperty(key);
    });

    revokeFilePermissions_(cfg, report, dryRun);

    if (!dryRun && cfg.create_backup !== false) {
      const backup = buildSyncLogsBackup_(sheet, report);
      report.backup_created = true;
      report.backup_id = backup.id;
      report.backup_url = backup.url;
      report.notes.push("Backup rollback dibuat di spreadsheet terpisah: " + backup.name);
    }

    if (sheet && !dryRun) {
      if (ss.getSheets().length === 1) {
        ss.insertSheet("System_Main");
        report.notes.push("Sync_Logs adalah sheet terakhir, dibuat sheet pengganti 'System_Main' sebelum delete.");
      }
      ss.deleteSheet(sheet);
      report.sheet_deleted = true;
    }

    if (!sheet) report.notes.push("Sheet Sync_Logs tidak ditemukan.");
    if (dryRun) report.notes.push("Dry run aktif: tidak ada perubahan yang ditulis.");
    return report;
  } catch (e) {
    return { status: "error", message: String(e) };
  }
}

/* =========================
   NOTIFICATIONS
========================= */

/**
 * Normalize Indonesian phone number for Fonnte API.
 * Strips non-digits, handles +62/62/0 prefix variations.
 * Returns clean number like "81234567890" (without country code prefix).
 */
function normalizePhone_(raw) {
  if (!raw) return "";
  // Remove all non-digit characters (+, -, spaces, parens, etc)
  let num = String(raw).replace(/[^0-9]/g, "");
  // Handle country code prefix
  if (num.startsWith("620")) num = num.substring(3); // 6208xxx → 8xxx
  else if (num.startsWith("62")) num = num.substring(2); // 628xxx → 8xxx
  // Remove leading 0 if present
  if (num.startsWith("0")) num = num.substring(1); // 08xxx → 8xxx
  return num;
}

function sendWA(target, message, cfg) {
  if (!target) {
    logWA_("SKIP", "(empty)", "No target number provided");
    return { success: false, reason: "no_target" };
  }
  cfg = cfg || getSettingsMap_();
  const token = getSecret_("fonnte_token", cfg) || getCfg("fonnte_token");
  if (!token) {
    logWA_("NO_TOKEN", target, "fonnte_token not configured in Settings");
    return { success: false, reason: "no_fonnte_token" };
  }

  // Normalize phone number: strip all non-digits, handle prefix
  const cleanTarget = normalizePhone_(target);
  if (!cleanTarget || cleanTarget.length < 9) {
    logWA_("INVALID_NUMBER", String(target), "After normalization: '" + cleanTarget + "' (too short or empty)");
    return { success: false, reason: "invalid_phone_number" };
  }

  const MAX_RETRIES = 2;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      const res = UrlFetchApp.fetch("https://api.fonnte.com/send", {
        method: "post",
        headers: { "Authorization": token },
        payload: {
          target: cleanTarget,
          message: message,
          countryCode: "62"
        },
        muteHttpExceptions: true
      });

      const httpCode = res.getResponseCode();
      const resText = res.getContentText();

      // Validate Fonnte API response
      if (httpCode >= 200 && httpCode < 300) {
        try {
          const resJson = JSON.parse(resText);
          if (resJson.status === true || resJson.status === "true") {
            logWA_("SENT", cleanTarget, "OK (attempt " + attempt + ") | Detail: " + String(resJson.detail || resJson.message || "").substring(0, 100));
            return { success: true };
          } else {
            // Fonnte returned 200 but status=false (invalid number, quota, etc)
            const reason = String(resJson.reason || resJson.detail || resJson.message || "Unknown").substring(0, 200);
            if (attempt >= MAX_RETRIES) {
              logWA_("REJECTED", cleanTarget, "Fonnte rejected: " + reason + " | Raw response: " + resText.substring(0, 200));
              return { success: false, reason: reason };
            }
          }
        } catch (parseErr) {
          // Non-JSON response but HTTP 200 - treat as success
          logWA_("SENT_UNVERIFIED", cleanTarget, "HTTP " + httpCode + " but non-JSON response (attempt " + attempt + ")");
          return { success: true };
        }
      } else {
        // HTTP error (401, 403, 500, etc)
        if (attempt >= MAX_RETRIES) {
          logWA_("HTTP_ERROR", cleanTarget, "HTTP " + httpCode + ": " + resText.substring(0, 200));
          return { success: false, reason: "HTTP " + httpCode };
        }
      }

      // Wait before retry
      if (attempt < MAX_RETRIES) Utilities.sleep(1000);

    } catch (e) {
      if (attempt >= MAX_RETRIES) {
        logWA_("EXCEPTION", cleanTarget, e.toString());
        return { success: false, reason: e.toString() };
      }
      Utilities.sleep(1000);
    }
  }
  return { success: false, reason: "exhausted_retries" };
}

function sendEmail(target, subject, body, cfg) {
  if (!target) return { success: false, reason: "no_target" };
  cfg = cfg || getSettingsMap_();

  // Check daily quota first
  const remaining = MailApp.getRemainingDailyQuota();
  if (remaining <= 0) {
    logEmail_("QUOTA_EXCEEDED", target, subject, "Daily email quota exceeded (remaining: " + remaining + ")");
    // Fallback: alert admin via WA
    const adminWA = getCfgFrom_(cfg, "wa_admin");
    if (adminWA) {
      sendWA(adminWA, "⚠️ *EMAIL QUOTA HABIS!*\n\nEmail ke " + target + " GAGAL terkirim karena quota harian habis.\nSubject: " + subject, cfg);
    }
    return { success: false, reason: "quota_exceeded" };
  }

  const senderName = getCfgFrom_(cfg, "site_name") || "Admin Sistem";
  const MAX_RETRIES = 3;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      MailApp.sendEmail({ to: target, subject: subject, htmlBody: body, name: senderName });
      logEmail_("SENT", target, subject, "OK (attempt " + attempt + ", quota left: " + (remaining - 1) + ")");
      return { success: true };
    } catch (e) {
      Logger.log("sendEmail attempt " + attempt + " failed: " + e);
      if (attempt < MAX_RETRIES) {
        Utilities.sleep(1000 * attempt); // Exponential backoff: 1s, 2s
      } else {
        logEmail_("FAILED", target, subject, e.toString());
        // Fallback: alert admin via WA
        const adminWA = getCfgFrom_(cfg, "wa_admin");
        if (adminWA) {
          sendWA(adminWA, "❌ *EMAIL GAGAL TERKIRIM!*\n\nKe: " + target + "\nSubject: " + subject + "\nError: " + String(e).substring(0, 200), cfg);
        }
        return { success: false, reason: e.toString() };
      }
    }
  }
}

function getEmailQuotaStatus() {
  const remaining = MailApp.getRemainingDailyQuota();
  return { status: "success", remaining: remaining, limit: 100, warning: remaining < 10 };
}

/* =========================
   CREATE ORDER (ANGKA UNIK + WHITE-LABEL + AFFILIATE)
========================= */
function createOrder(d, cfg) {
  try {
    if (getAdminSessionToken_(d)) {
      requireAdminSession_(d, { actionName: "create_order" });
    }
    cfg = cfg || getSettingsMap_();

    const oS = mustSheet_("Orders");
    const uS = mustSheet_("Users");

    const inv = "INV-" + Math.floor(10000 + Math.random() * 90000);
    const email = String(d.email || "").trim().toLowerCase();
    if (!email) return { status: "error", message: "Email wajib diisi" };

    // Normalize WhatsApp number at storage time
    const waRaw = String(d.whatsapp || "").trim();
    const waNormalized = normalizePhone_(waRaw);
    if (waRaw && !waNormalized) {
      Logger.log("WARNING: WA number normalization failed for: " + waRaw);
    }

    const siteName = getCfgFrom_(cfg, "site_name") || "Sistem Premium";
    const siteUrl = String(getCfgFrom_(cfg, "site_url") || "").trim();
    const loginUrl = siteUrl ? (siteUrl + "/login.html") : "Link Login Belum Disetting";

    const bankName = getCfgFrom_(cfg, "bank_name") || "-";
    const bankNorek = getCfgFrom_(cfg, "bank_norek") || "-";
    const bankOwner = getCfgFrom_(cfg, "bank_owner") || "-";

    const aff = (d.affiliate && String(d.affiliate).trim() !== "") ? String(d.affiliate).trim() : "-";

    const hargaDasar = toNumberSafe_(d.harga);
    
    // MODIFIED: Allow 0 price (Free Product)
    const isZeroPrice = hargaDasar === 0;
    if (!isZeroPrice && hargaDasar <= 0) return { status: "error", message: "Harga tidak valid" };

    let komisiNominal = 0;
    
    // Lookup Product Commission
    const pId = String(d.id_produk || "").trim();
    if (pId && aff !== "-") {
        const rules = mustSheet_("Access_Rules").getDataRange().getValues();
        for (let i = 1; i < rules.length; i++) {
            if (String(rules[i][0]) === pId) {
                // Commission is in column 12 (index 11)
                komisiNominal = Number(rules[i][11] || 0);
                break;
            }
        }
    }

    const kodeUnik = isZeroPrice ? 0 : (Math.floor(Math.random() * 900) + 100);
    const hargaTotalUnik = hargaDasar + kodeUnik;

    // Cek atau Buat User Baru
    let isNew = true;
    let pass = Math.random().toString(36).slice(-6);

    const uData = uS.getDataRange().getValues();
    for (let j = 1; j < uData.length; j++) {
      if (String(uData[j][1]).toLowerCase() === email) {
        isNew = false;
        pass = String(uData[j][2]);
        break;
      }
    }
    if (isNew) {
      // Generate Friendly Unique ID (u-XXXXXX)
      let newUserId = "u-" + Math.floor(100000 + Math.random() * 900000);
      let unique = false;
      while(!unique) {
          unique = true;
          for(let k=1; k<uData.length; k++) {
              if(String(uData[k][0]) === newUserId) {
                  unique = false;
                  newUserId = "u-" + Math.floor(100000 + Math.random() * 900000);
                  break;
              }
          }
      }
      uS.appendRow([newUserId, email, hashPassword_(pass), d.nama, "member", "Active", toISODate_(), "-"]);
    }

    const orderStatus = isZeroPrice ? "Lunas" : "Pending";

    // Simpan order (struktur kolom sama dengan script lu)
    // Store WA number as text (prefix with apostrophe prevents Google Sheets from converting to Number)
    const waForSheet = waNormalized || waRaw;
    oS.appendRow([
      inv,
      email,
      d.nama,
      "'" + waForSheet,
      d.id_produk,
      d.nama_produk,
      hargaTotalUnik,
      orderStatus,
      toISODate_(),
      aff,
      komisiNominal
    ]);

    // ==========================================
    // NOTIFIKASI (LOGIC CABANG: GRATIS vs BAYAR)
    // ==========================================
    
    const adminWA = getCfgFrom_(cfg, "wa_admin");

    if (isZeroPrice) {
       // --- SKENARIO PRODUK GRATIS (AUTO LUNAS) ---
       
       // 1. Ambil Link Akses
       let accessUrl = "";
       const pS = mustSheet_("Access_Rules");
       const pData = pS.getDataRange().getValues();
       for (let k = 1; k < pData.length; k++) {
         if (String(pData[k][0]) === String(d.id_produk)) { accessUrl = pData[k][3]; break; }
       }
       
       // 2. WA ke User (use normalized number)
       const waText = `Halo ${d.nama}, selamat datang di ${siteName}! 🎉\n\nSukses! Akses Anda untuk produk *${d.nama_produk}* telah aktif (GRATIS).\n\n🚀 *Klik link berikut untuk akses materi:*\n${accessUrl}\n\n🔐 *AKUN MEMBER AREA*\n🌐 Link: ${loginUrl}\n✉️ Email: ${email}\n🔑 Password: ${pass}\n\nTerima kasih!\n*Tim ${siteName}*`;
       sendWA(waForSheet, waText, cfg);

       // 3. Email ke User
       const emailHtml = `
       <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; color: #334155; border: 1px solid #e2e8f0; border-radius: 10px;">
          <h2 style="color: #10b981;">Akses Produk Gratis Dibuka! 🎁</h2>
          <p>Halo <b>${d.nama}</b>,</p>
          <p>Selamat! Anda telah berhasil mendapatkan akses ke produk <b>${d.nama_produk}</b> secara GRATIS.</p>
          
          <div style="text-align: center; margin: 30px 0;">
              <a href="${accessUrl}" style="background-color: #4f46e5; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold;">Akses Materi Sekarang</a>
          </div>

          <h3 style="color: #0f172a;">🔐 Akun Member Area</h3>
          <p><b>Link:</b> <a href="${loginUrl}">${loginUrl}</a><br>
          <b>Email:</b> ${email}<br>
          <b>Password:</b> <code>${pass}</code></p>
          
          <p>Salam hangat,<br><b>Tim ${siteName}</b></p>
       </div>`;
       sendEmail(email, `Akses Gratis! Produk ${d.nama_produk}`, emailHtml, cfg);

       // 4. Notif Admin
       sendWA(adminWA, `🎁 *ORDER GRATIS BARU!* 🎁\n\n📌 *Invoice:* #${inv}\n📦 *Produk:* ${d.nama_produk}\n👤 *User:* ${d.nama}\n\nStatus: Lunas (Auto)`, cfg);

    } else {
       // --- SKENARIO BERBAYAR (PENDING) ---

       // --> NOTIFIKASI PEMBELI (WHATSAPP)
    const waBuyerText =
`Halo *${d.nama}*, salam hangat dari ${siteName}! 👋

Terima kasih telah melakukan pemesanan. Berikut rincian pesanan Anda:

📦 *Produk:* ${d.nama_produk}
🔖 *Invoice:* #${inv}
💰 *Total Tagihan:* Rp ${Number(hargaTotalUnik).toLocaleString('id-ID')}

⚠️ _(Penting: Transfer *TEPAT* hingga 3 digit terakhir agar sistem dapat memvalidasi otomatis)_

Silakan selesaikan pembayaran ke rekening berikut:

🏦 *Bank:* ${bankName}
💳 *No. Rek:* ${bankNorek}
👤 *A.n:* ${bankOwner}

*(Mohon kirimkan bukti transfer ke sini agar pesanan segera diproses)*

---

🔐 *INFORMASI AKUN MEMBER*
🌐 *Link Login:* ${loginUrl}
✉️ *Email:* ${email}
🔑 *Password:* ${pass}

*(Akses materi otomatis terbuka di akun ini setelah pembayaran divalidasi)*.

Jika ada pertanyaan, silakan balas pesan ini. Terima kasih! 🙏`;
    sendWA(waForSheet, waBuyerText, cfg);

    // --> NOTIFIKASI PEMBELI (EMAIL) (template asli lu)
    const emailBuyerHtml = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; color: #334155; border: 1px solid #e2e8f0; border-radius: 10px;">
        <h2 style="color: #4f46e5; margin-bottom: 5px;">Menunggu Pembayaran Anda ⏳</h2>
        <p style="font-size: 16px; margin-top: 0;">Halo <b>${d.nama}</b>,</p>
        <p>Terima kasih atas pesanan Anda di <b>${siteName}</b>. Berikut adalah detail tagihan yang harus dibayarkan:</p>

        <div style="background-color: #f8fafc; padding: 15px 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #4f46e5;">
            <p style="margin: 0 0 5px 0;"><b>Produk:</b> ${d.nama_produk}</p>
            <p style="margin: 0 0 5px 0;"><b>Invoice:</b> #${inv}</p>
            <p style="margin: 0; font-size: 20px; color: #0f172a;"><b>Total Tagihan: Rp ${Number(hargaTotalUnik).toLocaleString('id-ID')}</b></p>
            <p style="margin: 5px 0 0 0; font-size: 12px; color: #ef4444; font-weight: bold;">*Wajib transfer TEPAT hingga 3 digit angka terakhir.</p>
        </div>

        <p>Silakan selesaikan pembayaran ke rekening berikut:</p>

        <div style="background-color: #f1f5f9; padding: 15px 20px; border-radius: 8px; margin: 20px 0; text-align: center;">
            <p style="margin: 0 0 5px 0; color: #64748b; text-transform: uppercase; font-size: 12px; font-weight: bold;">Transfer Ke Bank ${bankName}</p>
            <p style="margin: 0 0 5px 0; font-size: 22px; color: #4f46e5; font-family: monospace; font-weight: bold; letter-spacing: 2px;">${bankNorek}</p>
            <p style="margin: 0; font-size: 14px;"><b>A.n:</b> ${bankOwner}</p>
        </div>

        <p>Setelah transfer, konfirmasi melalui WhatsApp Admin agar produk segera kami aktifkan.</p>

        <hr style="border: none; border-top: 1px dashed #cbd5e1; margin: 30px 0;">

        <h3 style="color: #0f172a; margin-bottom: 10px;">🔐 Detail Akun Member Anda</h3>

        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr>
                <td style="padding: 10px; border-bottom: 1px solid #e2e8f0; width: 100px;"><b>Link Login</b></td>
                <td style="padding: 10px; border-bottom: 1px solid #e2e8f0;"><a href="${loginUrl}" style="color: #4f46e5; text-decoration: none;">${loginUrl}</a></td>
            </tr>
            <tr>
                <td style="padding: 10px; border-bottom: 1px solid #e2e8f0;"><b>Email</b></td>
                <td style="padding: 10px; border-bottom: 1px solid #e2e8f0;">${email}</td>
            </tr>
            <tr>
                <td style="padding: 10px; border-bottom: 1px solid #e2e8f0;"><b>Password</b></td>
                <td style="padding: 10px; border-bottom: 1px solid #e2e8f0;"><code style="background: #f1f5f9; padding: 3px 6px; border-radius: 4px;">${pass}</code></td>
            </tr>
        </table>

        <br>
        <p>Salam hangat,<br><b>Tim ${siteName}</b></p>
    </div>
    `;
    sendEmail(email, `Menunggu Pembayaran: Pesanan #${inv} - ${siteName}`, emailBuyerHtml, cfg);

    // --> NOTIFIKASI ADMIN
    const affMsg = aff !== "-" ? `\n🤝 *Affiliate:* ${aff}\n💸 *Potensi Komisi:* Rp ${Number(komisiNominal).toLocaleString('id-ID')}` : "";
    sendWA(adminWA, `🚨 *PESANAN BARU MASUK!* 🚨\n\n📌 *Invoice:* #${inv}\n📦 *Produk:* ${d.nama_produk}\n👤 *Customer:* ${d.nama}\n💳 *Nilai Unik:* Rp ${Number(hargaTotalUnik).toLocaleString('id-ID')}${affMsg}\n\nSilakan pantau pembayaran dari customer ini.`, cfg);
    } // End of Else (Paid)

    return { status: "success", invoice: inv, tagihan: hargaTotalUnik, is_new_user: isNew };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/* =========================
   UPDATE ORDER STATUS (MANUAL)
========================= */
function updateOrderStatus(d, cfg) {
  try {
    requireAdminSession_(d, { actionName: "update_order_status" });
    cfg = cfg || getSettingsMap_();
    const s = mustSheet_("Orders");
    const uS = mustSheet_("Users"); // kept for compatibility (even if not used)
    const pS = mustSheet_("Access_Rules");
    const r = s.getDataRange().getValues();
    const siteName = getCfgFrom_(cfg, "site_name") || "Sistem Premium";

    let orderFound = false, uEmail = "", uName = "", pId = "", pName = "", uWA = "";
    const newStatus = d.status || "Lunas";
    const isLunas = String(newStatus).trim().toLowerCase() === "lunas";

    // Trace ID for debugging this specific request
    const traceId = "UOS-" + Date.now();
    Logger.log(traceId + " updateOrderStatus called with id=" + d.id + " status=" + newStatus + " isLunas=" + isLunas);

    for (let i = 1; i < r.length; i++) {
      if (String(r[i][0]) === String(d.id)) {
        s.getRange(i + 1, 8).setValue(isLunas ? "Lunas" : newStatus);
        uEmail = r[i][1];
        uName = r[i][2];
        uWA = r[i][3];
        pId = r[i][4];
        pName = r[i][5];
        orderFound = true;
        Logger.log(traceId + " Order FOUND: row=" + (i+1) + " uWA=" + JSON.stringify(uWA) + " type=" + typeof uWA + " uEmail=" + uEmail);
        break;
      }
    }

    if (orderFound) {
      const cacheState = bumpPublicCacheState_(["dashboard"]);
      if (!isLunas) {
        Logger.log(traceId + " Not Lunas, returning early. newStatus=" + newStatus);
        return withPublicCacheState_({ status: "success", message: "Status berhasil diubah menjadi " + newStatus }, cacheState);
      }

      Logger.log(traceId + " Status=Lunas, proceeding with notifications...");

      let accessUrl = "";
      const pData = pS.getDataRange().getValues();
      for (let k = 1; k < pData.length; k++) {
        if (String(pData[k][0]) === String(pId)) { accessUrl = pData[k][3]; break; }
      }
      Logger.log(traceId + " accessUrl=" + accessUrl);

      // LOG: Debug notification target data before sending
      const waDebug = "uWA raw=" + JSON.stringify(uWA) + " type=" + typeof uWA + " normalized=" + normalizePhone_(uWA);
      logWA_("DEBUG_LUNAS", String(uWA), traceId + " | " + waDebug + " | Inv=" + d.id + " uEmail=" + uEmail);

      // STEP 1: Send WA to customer
      Logger.log(traceId + " Sending WA to: " + uWA);
      const waResult = sendWA(uWA, `🎉 *PEMBAYARAN TERVERIFIKASI!* 🎉\n\nHalo *${uName}*, kabar baik!\n\nPembayaran Anda untuk produk *${pName}* telah kami terima dan akses Anda kini *Telah Aktif*.\n\n🚀 *Klik link berikut untuk mengakses materi Anda:*\n${accessUrl}\n\nAnda juga bisa mengakses seluruh produk Anda melalui Member Area kami.\n\nTerima kasih atas kepercayaannya!\n*Tim ${siteName}*`, cfg);
      Logger.log(traceId + " WA Result: " + JSON.stringify(waResult));

      // STEP 2: Send Email to customer
      const emailActivationHtml = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; color: #334155; border: 1px solid #e2e8f0; border-radius: 10px;">
          <div style="text-align: center; margin-bottom: 20px;">
              <h1 style="color: #10b981; margin-bottom: 5px;">Akses Telah Dibuka! 🎉</h1>
          </div>
          <p style="font-size: 16px;">Halo <b>${uName}</b>,</p>
          <p>Terima kasih! Pembayaran Anda telah berhasil kami verifikasi. Akses penuh untuk produk <b>${pName}</b> sekarang sudah aktif dan dapat Anda gunakan.</p>

          <div style="text-align: center; margin: 30px 0;">
              <a href="${accessUrl}" style="background-color: #4f46e5; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 16px; display: inline-block;">Akses Materi Sekarang</a>
          </div>

          <p>Sebagai alternatif, Anda selalu bisa menemukan semua produk yang Anda miliki dengan masuk ke Member Area menggunakan akun yang telah kami kirimkan sebelumnya.</p>

          <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 30px 0;">
          <p style="font-size: 14px; color: #64748b; margin-bottom: 0;">Salam Sukses,<br><b>Tim ${siteName}</b></p>
      </div>
      `;
      Logger.log(traceId + " Sending Email to: " + uEmail);
      const emailResult = sendEmail(uEmail, `Akses Terbuka! Produk ${pName} - ${siteName}`, emailActivationHtml, cfg);
      Logger.log(traceId + " Email Result: " + JSON.stringify(emailResult));

      return withPublicCacheState_({ status: "success", trace: traceId, notifications: { wa: waResult, email: emailResult } }, cacheState);
    }

    return { status: "error", message: "Order tidak ditemukan" };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/* =========================
   HELPER: GET AFFILIATE PIXEL
========================= */
function getAffiliatePixel_(userId, productId) {
  const s = ss.getSheetByName("Affiliate_Pixels");
  if (!s) return null;
  
  const d = s.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]) === String(userId) && String(d[i][1]) === String(productId)) {
      return {
        pixel_id: String(d[i][2]),
        pixel_token: String(d[i][3]),
        pixel_test_code: String(d[i][4])
      };
    }
  }
  return null;
}

/* =========================
   PRODUCT DETAIL
========================= */
function getProductDetail(d, cfg) {
  try {
    cfg = cfg || getSettingsMap_();
    const rules = mustSheet_("Access_Rules").getDataRange().getValues();
    const pId = String(d.id).trim();
    let productData = null;

    for (let i = 1; i < rules.length; i++) {
      if (String(rules[i][0]) === pId && String(rules[i][5]).trim() === "Active") {
        productData = { 
            id: pId, 
            title: normalizePlainText_(rules[i][1]), 
            desc: normalizeProductDescription_(rules[i][2]), 
            harga: rules[i][4],
            pixel_id: rules[i][8] || "",
            pixel_token: rules[i][9] || "",
            pixel_test_code: rules[i][10] || "",
            commission: rules[i][11] || 0
        };
        break;
      }
    }
    if (!productData) return { status: "error", message: "Produk tidak ditemukan" };

    // --> CHECK AFFILIATE PIXEL OVERRIDE
    const affRef = d.ref || d.aff_id;
    if (affRef) {
        const affPixel = getAffiliatePixel_(affRef, pId);
        if (affPixel && affPixel.pixel_id) {
            productData.pixel_id = affPixel.pixel_id;
            productData.pixel_token = affPixel.pixel_token;
            productData.pixel_test_code = affPixel.pixel_test_code;
            productData.is_affiliate_pixel = true;
        }
    }

    const paymentInfo = {
      bank_name: getCfgFrom_(cfg, "bank_name"),
      bank_norek: getCfgFrom_(cfg, "bank_norek"),
      bank_owner: getCfgFrom_(cfg, "bank_owner"),
      wa_admin: getCfgFrom_(cfg, "wa_admin"),

      pixel_id: productData.pixel_id, // Pass pixel_id (possibly overridden)
      pixel_token: productData.pixel_token,
      pixel_test_code: productData.pixel_test_code
    };

    let affName = "";
    if (d.aff_id && d.aff_id !== "GUEST" && d.aff_id !== "-") {
      const users = mustSheet_("Users").getDataRange().getValues();
      for (let j = 1; j < users.length; j++) {
        if (String(users[j][0]) === String(d.aff_id)) { affName = String(users[j][3]); break; }
      }
    }

    return withPublicCacheVersion_({ status: "success", data: productData, payment: paymentInfo, aff_name: affName }, "catalog");
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/* =========================
   GET PRODUCTS + KOMISI AFFILIATE
========================= */
function getProducts(d, cfg, cachedOrders) {
  cfg = cfg || getSettingsMap_();
  
  // OPTIMIZATION: Only fetch sheets if needed, reuse cached if passed
  const rules = getCachedData_("access_rules", () => {
     return mustSheet_("Access_Rules").getDataRange().getValues();
  }, 3600); // 1 hour cache for rules

  const orders = cachedOrders || mustSheet_("Orders").getDataRange().getValues();
  const users = mustSheet_("Users").getDataRange().getValues(); // Often changes, might need real-time
  
  let email = String(d.email || "").trim().toLowerCase();
  let targetMode = false;

  // Support fetching products for a specific user (Bio Page)
  if (d.target_user_id) {
      targetMode = true;
      const tUid = String(d.target_user_id).trim();
      for (let j = 1; j < users.length; j++) {
          if (String(users[j][0]) === tUid) {
              email = String(users[j][1]).trim().toLowerCase();
              break;
          }
      }
  }

  let lunasIds = [], totalKomisi = 0, uId = "";
  let partners = [];

  if (email) {
    for (let j = 1; j < users.length; j++) {
      if (String(users[j][1]).toLowerCase() === email) { uId = String(users[j][0]); break; }
    }
    for (let x = 1; x < orders.length; x++) {
      const r = orders[x];
      if (String(r[1]).toLowerCase() === email && String(r[7]) === "Lunas") lunasIds.push(String(r[4]));
      
      // Check for Partners (Referrals) - Only calculate if not in target mode (optional, but keeps it clean)
      if (!targetMode && String(r[9]) === uId) {
          if (String(r[7]) === "Lunas") totalKomisi += Number(r[10] || 0);
          
          partners.push({
              invoice: r[0],
              name: r[2],
              product: r[5],
              status: r[7],
              date: r[8] ? String(r[8]).substring(0, 10) : "-",
              commission: r[10] || 0
          });
      }
    }
  }

  let owned = [], available = [];
  for (let i = 1; i < rules.length; i++) {
    if (String(rules[i][5]).trim() === "Active") {
      const pId = String(rules[i][0]);
      const hasAccess = lunasIds.includes(pId);
      const pObj = {
        id: pId,
        title: normalizePlainText_(rules[i][1]),
        desc: normalizeProductDescription_(rules[i][2]),
        url: hasAccess ? rules[i][3] : "#",
        harga: rules[i][4],
        access: hasAccess,
        lp_url: rules[i][6] || "",
        image_url: rules[i][7] || "",
        commission: rules[i][11] || 0
      };
      
      if (targetMode) {
          // In Bio Page mode, we show what the user OWNS as the "Available Catalog" for visitors
          if (hasAccess) available.push(pObj);
      } else {
          // Normal Dashboard mode
          if (hasAccess && email) owned.push(pObj);
          else available.push(pObj);
      }
    }
  }

  return withPublicCacheVersion_({ status: "success", owned, available, total_komisi: totalKomisi, partners: partners.reverse() }, "catalog");
}

function getDashboardData(d) {
  try {
    const dashboardCacheVersion = publicCacheVersionToken_("dashboard");
    const cfg = getSettingsMap_();
    
    // 1. Get User ID & Admin ID from Users Sheet
    const email = String(d.email || "").trim().toLowerCase();
    const users = mustSheet_("Users").getDataRange().getValues();
    let userId = "";
    let userNama = "";
    let adminId = "";
    
    for(let i=1; i<users.length; i++) {
        // Check for Admin (fallback upline)
        if(String(users[i][4]).toLowerCase() === "admin" && !adminId) {
            adminId = String(users[i][0]);
        }
        // Check for Current User
        if(String(users[i][1]).toLowerCase() === email) {
            userId = String(users[i][0]);
            userNama = String(users[i][3]);
        }
    }
    
    // 1b. Find Upline (Sponsor) from Orders History
    let uplineId = "";
    const orders = mustSheet_("Orders").getDataRange().getValues();
    
    if(userId) {
        // Search from oldest order (top) to find the first referrer
        for(let k=1; k<orders.length; k++) {
             if(String(orders[k][1]).toLowerCase() === email) {
                 const aff = String(orders[k][9] || "").trim();
                 if(aff && aff !== "-" && aff !== "" && aff !== "GUEST") {
                     uplineId = aff;
                     break; // Found the first sponsor
                 }
             }
        }
    }
    // Default to Admin if no upline found
    if(!uplineId) uplineId = adminId;

    // 1c. Get Upline Name
    let uplineName = "Admin";
    if(uplineId) {
         for(let m=1; m<users.length; m++) {
             if(String(users[m][0]) === uplineId) {
                 uplineName = String(users[m][3]);
                 break;
             }
         }
    }
    
    // 2. Get Products (reuse existing logic + pass cached orders)
    const productsData = getProducts(d, cfg, orders);
    const dashboardProducts = productsData && typeof productsData === "object" ? Object.assign({}, productsData) : {};
    delete dashboardProducts.cache_version;
    
    // 3. Get Global Pages (Affiliate Tools - ADMIN owned)
    const globalPages = getAllPages({ ...d, owner_id: "" });
    
    // 4. Get My Pages (User owned)
    let myPages = { data: [] };
    if(userId) {
        myPages = getAllPages({ ...d, owner_id: userId, only_mine: true });
    }
    
    // 5. Get Affiliate Pixels (User specific)
    let myPixels = [];
    if(userId) {
        const s = ss.getSheetByName("Affiliate_Pixels");
        if (s) {
            const data = s.getDataRange().getValues();
            for (let i = 1; i < data.length; i++) {
                if (String(data[i][0]) === userId) {
                    myPixels.push({
                        product_id: data[i][1],
                        pixel_id: data[i][2],
                        pixel_token: data[i][3],
                        pixel_test_code: data[i][4]
                    });
                }
            }
        }
    }
    
    return {
      status: "success",
      cache_version: dashboardCacheVersion,
      data: {
        user: { id: userId, nama: userNama, upline_id: uplineId, upline_name: uplineName },
        settings: { 
            site_name: getCfgFrom_(cfg, "site_name"),
            site_logo: sanitizeAssetUrl_(getCfgFrom_(cfg, "site_logo")),
            site_favicon: sanitizeAssetUrl_(getCfgFrom_(cfg, "site_favicon")),
            wa_admin: getCfgFrom_(cfg, "wa_admin")
        },
        products: dashboardProducts,
        pages: globalPages.data || [],
        my_pages: myPages.data || [],
        affiliate_pixels: myPixels
      }
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/* =========================
   LOGIN + PAGE + ADMIN
========================= */
function loginUser(d) {
  const u = mustSheet_("Users").getDataRange().getValues();
  const e = String(d.email || "").trim().toLowerCase();
  const inputPass = String(d.password || "").trim();

  if (!e || !inputPass) {
    return { status: "error", message: "Email dan password wajib diisi." };
  }

  for (let i = 1; i < u.length; i++) {
    if (String(u[i][1]).trim().toLowerCase() === e) {
      const storedPass = String(u[i][2]).trim();
      if (verifyPassword_(inputPass, storedPass)) {
        return { status: "success", data: { id: u[i][0], nama: u[i][3], email: u[i][1], role: String(u[i][4] || "member") } };
      }
      return { status: "error", message: "Password salah. Silakan cek kembali." };
    }
  }
  return { status: "error", message: "Gagal Login: Email tidak ditemukan." };
}

function loginAndDashboard(d) {
  const loginResult = loginUser(d);
  if (loginResult.status !== "success") return loginResult;

  const email = String((loginResult.data && loginResult.data.email) || d.email || "").trim().toLowerCase();
  const dashboardResult = getDashboardData({ email: email });

  if (dashboardResult.status !== "success") {
    return {
      status: "success",
      data: loginResult.data,
      dashboard: null,
      warning: dashboardResult.message || "Dashboard bootstrap gagal dimuat."
    };
  }

  return {
    status: "success",
    data: loginResult.data,
    dashboard: dashboardResult.data
  };
}

function getPageContent(d) {
  try {
    const r = mustSheet_("Pages").getDataRange().getValues();
    for (let i = 1; i < r.length; i++) {
      if (String(r[i][1]) === String(d.slug)) {
          return withPublicCacheVersion_({ 
              status: "success", 
              title: r[i][2], 
              content: r[i][3],
              pixel_id: r[i][7] || "",
              pixel_token: r[i][8] || "",
              pixel_test_code: r[i][9] || "",
              theme_mode: r[i][10] || "light"
          }, "pages");
      }
    }
    return { status: "error" };
  } catch (e) {
    return { status: "error" };
  }
}

function getAllPages(d) {
  try {
    const r = mustSheet_("Pages").getDataRange().getValues();
    const data = [];
    const filterOwner = String(d.owner_id || "").trim();
    const onlyMine = d.only_mine === true;

    for (let i = 1; i < r.length; i++) {
      if (String(r[i][4]) === "Active") {
        // Kolom 7 (index 6) adalah Owner ID. Jika kosong, anggap milik ADMIN (Global)
        const pageOwner = String(r[i][6] || "ADMIN").trim(); 

        if (onlyMine) {
            // Mode "Halaman Saya": Hanya tampilkan milik user ini
            if (pageOwner === filterOwner) data.push(r[i]);
        } else {
            // Mode Default (Global): Tampilkan halaman ADMIN (untuk affiliate link)
            if (pageOwner === "ADMIN") data.push(r[i]);
        }
      }
    }
    return withPublicCacheVersion_({ status: "success", data: data }, "pages");
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function adminLogin(d) {
  const e = String(d.email || "").trim().toLowerCase();
  const inputPass = String(d.password || "").trim();

  if (!e || !inputPass) {
    return { status: "error", message: "Email dan password wajib diisi." };
  }

  const u = mustSheet_("Users").getDataRange().getValues();

  for (let i = 1; i < u.length; i++) {
    if (String(u[i][1]).trim().toLowerCase() === e) {
      const storedPass = String(u[i][2]).trim();
      const role = String(u[i][4]).trim().toLowerCase();

      if (verifyPassword_(inputPass, storedPass) && role === "admin") {
        const session = createAdminSession_({
          id: String(u[i][0] || ""),
          email: e,
          name: String(u[i][3] || "Admin"),
          role: "admin"
        });
        return {
          status: "success",
          data: {
            id: String(u[i][0] || ""),
            nama: String(u[i][3] || "Admin"),
            email: e,
            role: "admin",
            session_token: session.token,
            expires_at: session.expires_at
          }
        };
      }

      if (verifyPassword_(inputPass, storedPass) && role !== "admin") {
        return { status: "error", message: "Akun ditemukan tapi bukan admin. Role: " + u[i][4] };
      }

      return { status: "error", message: "Password salah. Silakan cek kembali." };
    }
  }

  return { status: "error", message: "Email " + e + " tidak ditemukan di database." };
}

/* =========================
   DIAGNOSTIC: Debug Login Data
========================= */
function debugLogin(d) {
  try {
    const u = mustSheet_("Users").getDataRange().getValues();
    const targetEmail = String(d.email || "").trim().toLowerCase();
    const inputPass = String(d.password || "");
    const results = [];

    for (let i = 1; i < u.length; i++) {
      const rawEmail = u[i][1];
      const rawPass = u[i][2];
      const rawRole = u[i][4];
      const emailStr = String(rawEmail);
      const passStr = String(rawPass);
      const roleStr = String(rawRole);

      if (emailStr.trim().toLowerCase() === targetEmail || !targetEmail) {
        // Get charCodes of password to detect hidden characters
        const passChars = [];
        for (let c = 0; c < passStr.length; c++) {
          passChars.push({ char: passStr[c], code: passStr.charCodeAt(c) });
        }

        const inputChars = [];
        for (let c = 0; c < inputPass.length; c++) {
          inputChars.push({ char: inputPass[c], code: inputPass.charCodeAt(c) });
        }

        results.push({
          row: i + 1,
          email: { raw: emailStr, trimmed: emailStr.trim(), type: typeof rawEmail, length: emailStr.length, trimmed_length: emailStr.trim().length },
          password: { raw_length: passStr.length, trimmed: passStr.trim(), trimmed_length: passStr.trim().length, type: typeof rawPass, charCodes: passChars },
          input_password: { raw: inputPass, trimmed: inputPass.trim(), length: inputPass.length, charCodes: inputChars },
          password_match: { raw: passStr === inputPass, trimmed: passStr.trim() === inputPass.trim() },
          role: { raw: roleStr, trimmed: roleStr.trim(), lowercase: roleStr.trim().toLowerCase(), type: typeof rawRole, is_admin: roleStr.trim().toLowerCase() === "admin" }
        });
      }
    }

    return { status: "success", data: results, total_users: u.length - 1 };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/* =========================
   UNIT TESTS: Authentication
========================= */
function runAuthTests() {
  const results = [];
  const u = mustSheet_("Users").getDataRange().getValues();

  // Test 1: Users sheet has data
  results.push({ test: "Users sheet exists and has data", pass: u.length > 1, detail: "Rows: " + u.length });

  // Test 2: Header structure
  const expectedHeaders = ["user_id", "email", "password", "nama_lengkap", "role"];
  const headers = u[0].map(h => String(h).trim().toLowerCase());
  const headerMatch = expectedHeaders.every(h => headers.includes(h));
  results.push({ test: "Headers match expected structure", pass: headerMatch, detail: "Found: " + headers.slice(0, 5).join(", ") });

  // Test 3: Find admin user
  let adminRow = null;
  for (let i = 1; i < u.length; i++) {
    if (String(u[i][4]).trim().toLowerCase() === "admin") {
      adminRow = { index: i, email: String(u[i][1]), pass: String(u[i][2]), name: String(u[i][3]), role: String(u[i][4]) };
      break;
    }
  }
  results.push({ test: "Admin user exists in Users sheet", pass: !!adminRow, detail: adminRow ? "Email: " + adminRow.email : "No admin found" });

  if (adminRow) {
    // Test 4: Admin password has no hidden characters
    const passStr = adminRow.pass;
    const hasHidden = passStr.length !== passStr.trim().length;
    results.push({ test: "Admin password has no trailing/leading spaces", pass: !hasHidden, 
      detail: "Raw length: " + passStr.length + ", Trimmed: " + passStr.trim().length });

    // Test 5: Admin email has no hidden characters
    const emailStr = adminRow.email;
    const emailHasHidden = emailStr.length !== emailStr.trim().length;
    results.push({ test: "Admin email has no trailing/leading spaces", pass: !emailHasHidden,
      detail: "Raw length: " + emailStr.length + ", Trimmed: " + emailStr.trim().length });

    // Test 6: loginUser works for admin (should succeed — tests email+pass)
    const loginResult = loginUser({ email: adminRow.email.trim(), password: adminRow.pass.trim() });
    results.push({ test: "loginUser() succeeds for admin credentials", pass: loginResult.status === "success",
      detail: JSON.stringify(loginResult) });

    // Test 7: adminLogin works for admin (should succeed — tests email+pass+role)
    const adminResult = adminLogin({ email: adminRow.email.trim(), password: adminRow.pass.trim() });
    results.push({ test: "adminLogin() succeeds for admin credentials", pass: adminResult.status === "success",
      detail: JSON.stringify(adminResult) });

    const adminSessionToken = adminResult && adminResult.data ? String(adminResult.data.session_token || "") : "";
    const adminSession = adminSessionToken ? getAdminSession_(adminSessionToken) : null;
    results.push({
      test: "Admin session disimpan persisten tanpa expiry otomatis",
      pass: !!adminSessionToken && !!adminSession && Number(adminResult && adminResult.data ? adminResult.data.expires_at : 0) === 0,
      detail: JSON.stringify({
        token_exists: !!adminSessionToken,
        session_exists: !!adminSession,
        expires_at: adminResult && adminResult.data ? Number(adminResult.data.expires_at || 0) : null
      })
    });

    if (adminSessionToken) {
      const logoutResult = adminLogout({ auth_session_token: adminSessionToken });
      const revokedSession = getAdminSession_(adminSessionToken);
      results.push({
        test: "adminLogout() mencabut session admin secara manual",
        pass: logoutResult.status === "success" && !revokedSession,
        detail: JSON.stringify({
          logout: logoutResult,
          session_exists_after_logout: !!revokedSession
        })
      });
    }
  }

  // Test 8: Find member user
  let memberRow = null;
  for (let i = 1; i < u.length; i++) {
    if (String(u[i][4]).trim().toLowerCase() === "member") {
      memberRow = { index: i, email: String(u[i][1]), pass: String(u[i][2]), name: String(u[i][3]), role: String(u[i][4]) };
      break;
    }
  }

  if (memberRow) {
    // Test 9: loginUser works for member
    const memberResult = loginUser({ email: memberRow.email.trim(), password: memberRow.pass.trim() });
    results.push({ test: "loginUser() succeeds for member credentials", pass: memberResult.status === "success",
      detail: JSON.stringify(memberResult) });

    // Test 10: adminLogin rejects member (should fail — not admin role)
    const memberAdminResult = adminLogin({ email: memberRow.email.trim(), password: memberRow.pass.trim() });
    results.push({ test: "adminLogin() correctly rejects member user", pass: memberAdminResult.status === "error",
      detail: JSON.stringify(memberAdminResult) });
  }

  // Test 11: Empty credentials rejected
  const emptyResult = adminLogin({ email: "", password: "" });
  results.push({ test: "adminLogin() rejects empty credentials", pass: emptyResult.status === "error",
    detail: emptyResult.message });

  // Test 12: Wrong password rejected
  if (adminRow) {
    const wrongPassResult = adminLogin({ email: adminRow.email, password: "wrongpass123" });
    results.push({ test: "adminLogin() rejects wrong password", pass: wrongPassResult.status === "error",
      detail: wrongPassResult.message });
  }

  const passed = results.filter(r => r.pass).length;
  const failed = results.filter(r => !r.pass).length;

  return { status: "success", summary: passed + " passed, " + failed + " failed, " + results.length + " total", tests: results };
}

function runMootaValidationTests() {
  const cases = [
    {
      test: "Menerima URL HTTPS dan Secret Token alphanumeric valid",
      input: { gasUrl: "https://example.com/webhook/moota", token: "Secret123" },
      expectedErrors: []
    },
    {
      test: "Menolak link webhook kosong",
      input: { gasUrl: "", token: "Secret123" },
      expectedErrors: ["Link webhook Moota wajib diisi."]
    },
    {
      test: "Menolak link webhook non-HTTPS",
      input: { gasUrl: "http://example.com/webhook/moota", token: "Secret123" },
      expectedErrors: ["Format link webhook Moota tidak valid. Gunakan URL HTTPS tanpa query string."]
    },
    {
      test: "Menolak link webhook Google Apps Script langsung",
      input: { gasUrl: "https://script.google.com/macros/s/abc/exec", token: "Secret123" },
      expectedErrors: ["Link webhook Moota tidak boleh langsung ke Google Apps Script. Gunakan endpoint Cloudflare Worker atau proxy publik agar header Signature bisa diteruskan."]
    },
    {
      test: "Menolak Secret Token kosong",
      input: { gasUrl: "https://example.com/webhook/moota", token: "" },
      expectedErrors: ["Secret Token Moota wajib diisi."]
    },
    {
      test: "Menolak Secret Token kurang dari 8 karakter",
      input: { gasUrl: "https://example.com/webhook/moota", token: "Abc1234" },
      expectedErrors: ["Format Secret Token Moota tidak valid. Gunakan minimal 8 karakter alphanumeric tanpa spasi."]
    },
    {
      test: "Menolak Secret Token dengan karakter non-alphanumeric",
      input: { gasUrl: "https://example.com/webhook/moota", token: "Secret-123" },
      expectedErrors: ["Format Secret Token Moota tidak valid. Gunakan minimal 8 karakter alphanumeric tanpa spasi."]
    }
  ];

  const results = cases.map(function(item) {
    const errors = validateMootaConfigFormat_(item.input);
    const pass = JSON.stringify(errors) === JSON.stringify(item.expectedErrors);
    return {
      test: item.test,
      pass: pass,
      input: item.input,
      expected: item.expectedErrors,
      actual: errors
    };
  });

  const passed = results.filter(function(result) { return result.pass; }).length;
  const failed = results.length - passed;

  return {
    status: "success",
    summary: passed + " passed, " + failed + " failed, " + results.length + " total",
    tests: results
  };
}

function runMootaSignatureTests() {
  const payload = JSON.stringify([{
    amount: 50000,
    type: "CR",
    description: "Testing webhook moota",
    created_at: "2019-11-10 14:33:01"
  }]);
  const secret = "Secret123";
  const expectedSignature = computeMootaSignatureHex_(payload, secret);
  const prefixedSignature = "sha256=" + expectedSignature.toUpperCase();
  const cases = [
    {
      test: "Normalisasi signature menerima prefix sha256 dan huruf besar",
      actual: normalizeMootaSignature_(prefixedSignature),
      expected: expectedSignature
    },
    {
      test: "Verifikasi signature valid",
      actual: verifyMootaSignature_(payload, secret, expectedSignature).code,
      expected: "ok"
    },
    {
      test: "Verifikasi signature valid dengan prefix sha256",
      actual: verifyMootaSignature_(payload, secret, prefixedSignature).code,
      expected: "ok"
    },
    {
      test: "Verifikasi signature gagal saat signature kosong",
      actual: verifyMootaSignature_(payload, secret, "").code,
      expected: "missing_signature"
    },
    {
      test: "Verifikasi signature gagal saat secret kosong",
      actual: verifyMootaSignature_(payload, "", expectedSignature).code,
      expected: "missing_secret"
    },
    {
      test: "Verifikasi signature gagal saat signature tidak cocok",
      actual: verifyMootaSignature_(payload, secret, "deadbeef").code,
      expected: "invalid_signature"
    },
    {
      test: "Meta membaca flag verifikasi Worker",
      actual: extractMootaSignatureMeta_({
        parameter: {
          moota_signature: expectedSignature,
          moota_sig_verified: "1",
          moota_sig_verified_by: "worker"
        }
      }).workerVerifiedSignature,
      expected: true
    },
    {
      test: "Meta membaca sumber verifikasi Worker",
      actual: extractMootaSignatureMeta_({
        parameter: {
          moota_signature: expectedSignature,
          moota_sig_verified: "1",
          moota_sig_verified_by: "worker"
        }
      }).workerVerificationSource,
      expected: "worker"
    },
    {
      test: "Helper mendeteksi URL Google Apps Script langsung",
      actual: isDirectAppsScriptUrl_("https://script.google.com/macros/s/abc/exec"),
      expected: true
    },
    {
      test: "Klasifikasi missing signature untuk URL Google Apps Script langsung",
      actual: classifyMootaSignatureMissing_(
        { gasUrl: "https://script.google.com/macros/s/abc/exec" },
        { forwardedByWorker: false, workerSawSignature: false }
      ).code,
      expected: "direct_apps_script_url"
    },
    {
      test: "Klasifikasi missing signature saat Worker tidak terdeteksi",
      actual: classifyMootaSignatureMissing_(
        { gasUrl: "https://example.com/webhook/moota" },
        { forwardedByWorker: false, workerSawSignature: false }
      ).code,
      expected: "worker_not_detected"
    },
    {
      test: "Klasifikasi missing signature saat Worker hidup tapi header tidak ada",
      actual: classifyMootaSignatureMissing_(
        { gasUrl: "https://example.com/webhook/moota" },
        { forwardedByWorker: true, workerSawSignature: false }
      ).code,
      expected: "worker_missing_signature_header"
    }
  ];

  const results = cases.map(function(item) {
    const pass = JSON.stringify(item.actual) === JSON.stringify(item.expected);
    return {
      test: item.test,
      pass: pass,
      expected: item.expected,
      actual: item.actual
    };
  });

  const passed = results.filter(function(result) { return result.pass; }).length;
  const failed = results.length - passed;

  return {
    status: "success",
    summary: passed + " passed, " + failed + " failed, " + results.length + " total",
    tests: results,
    sample_signature: maskMootaSignatureForLog_(expectedSignature)
  };
}

function getAdminData(d, cfg) {
  try {
    const session = requireAdminSession_(d, { actionName: "get_admin_data" });
    cfg = cfg || getSettingsMap_();
    const o = mustSheet_("Orders").getDataRange().getValues();
    const u = mustSheet_("Users").getDataRange().getValues();
    const s = mustSheet_("Settings").getDataRange().getValues();
    const p = mustSheet_("Access_Rules").getDataRange().getValues();
    const pg = mustSheet_("Pages").getDataRange().getValues();

    let rev = 0;
    for (let i = 1; i < o.length; i++) {
      if (String(o[i][7]) === "Lunas") rev += Number(o[i][6] || 0);
    }

    let t = {};
    for (let i = 1; i < s.length; i++) {
      if (s[i][0]) t[s[i][0]] = s[i][1];
    }
    const resolvedMootaCfg = resolveMootaConfig_({}, cfg);
    t.moota_gas_url = normalizeMootaUrl_(resolvedMootaCfg.gasUrl || t.moota_gas_url || getCurrentWebAppUrl_());
    t.moota_token = "";
    t.moota_token_configured = !!resolvedMootaCfg.token;
    t.ik_private_key = "";
    t.ik_private_key_configured = !!getSecret_("ik_private_key", cfg);

    const result = {
      status: "success",
      role: session.role,
      session_expires_at: session.expires_at,
      stats: { users: u.length - 1, orders: o.length - 1, rev: rev },
      orders: o.slice(1).reverse().slice(0, 20),
      products: p.slice(1).map(normalizeProductRow_),
      pages: pg.slice(1),
      settings: t,
      users: u.slice(1).reverse().slice(0, 20),
      has_more_orders: (o.length - 1) > 20,
      has_more_users: (u.length - 1) > 20
    };
    return result;
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/* =========================
   SAVE PRODUCT / PAGE / SETTINGS
========================= */
function saveProduct(d) {
  try {
    requireAdminSession_(d, { actionName: "save_product" });
    const s = mustSheet_("Access_Rules");
    const descValidation = validateProductDescription_(d.desc);
    if (descValidation.errors.length) {
      return { status: "error", message: descValidation.errors[0], errors: descValidation.errors };
    }
    const productId = normalizePlainText_(d.id);
    const productTitle = normalizePlainText_(d.title);
    const productDesc = descValidation.value;
    const productUrl = String(d.url || "").trim();
    const productStatus = normalizePlainText_(d.status || "Active") || "Active";
    const landingPageUrl = String(d.lp_url || "").trim();
    const imageUrl = String(d.image_url || "").trim();
    const pixelId = normalizePlainText_(d.pixel_id);
    const pixelToken = String(d.pixel_token || "").trim();
    const pixelTestCode = normalizePlainText_(d.pixel_test_code);
    const commission = String(d.commission || "").trim();
    
    // Ensure we have enough columns (12 columns needed)
    if (s.getMaxColumns() < 12) s.insertColumnsAfter(s.getMaxColumns(), 12 - s.getMaxColumns());
    
    const dataRow = [productId, productTitle, productDesc, productUrl, d.harga, productStatus, landingPageUrl, imageUrl, pixelId, pixelToken, pixelTestCode, commission];
    const isEdit = String(d.is_edit) === "true";

    if (isEdit) {
      const r = s.getDataRange().getValues();
      for (let i = 1; i < r.length; i++) {
        if (String(r[i][0]).trim() === productId) {
          s.getRange(i + 1, 1, 1, 12).setValues([dataRow]);
          invalidateCaches_(["access_rules"]);
          return withPublicCacheState_({ status: "success" }, bumpPublicCacheState_(["catalog", "dashboard"]));
        }
      }
      return { status: "error", message: "ID Produk tidak ditemukan untuk diedit" };
    } else {
      // Check for duplicate ID before appending
      const r = s.getDataRange().getValues();
      for (let i = 1; i < r.length; i++) {
        if (String(r[i][0]).trim() === productId) {
           return { status: "error", message: "ID Produk sudah digunakan. Mohon refresh halaman." };
        }
      }
      s.appendRow(dataRow);
      invalidateCaches_(["access_rules"]);
      return withPublicCacheState_({ status: "success" }, bumpPublicCacheState_(["catalog", "dashboard"]));
    }
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function deleteProduct(d) {
  try {
    requireAdminSession_(d, { actionName: "delete_product" });
    const s = mustSheet_("Access_Rules");
    const r = s.getDataRange().getValues();
    const id = String(d.id).trim();

    for (let i = 1; i < r.length; i++) {
      if (String(r[i][0]).trim() === id) {
        s.deleteRow(i + 1);
        invalidateCaches_(["access_rules"]);
        return withPublicCacheState_({ status: "success", message: "Produk berhasil dihapus" }, bumpPublicCacheState_(["catalog", "dashboard"]));
      }
    }
    return { status: "error", message: "ID Produk tidak ditemukan" };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function savePage(d) {
  try {
    requireAdminSession_(d, { actionName: "save_page" });
    const s = mustSheet_("Pages");
    const isEdit = String(d.is_edit) === "true";
    const ownerId = String(d.owner_id || "ADMIN").trim(); // Default ke ADMIN
    const slug = String(d.slug).trim();
    const id = String(d.id).trim();

    const r = s.getDataRange().getValues();

    // 1. Cek Unik Slug (Global Check)
    for (let i = 1; i < r.length; i++) {
        const rowSlug = String(r[i][1]).trim();
        const rowId = String(r[i][0]).trim();
        
        if (rowSlug === slug) {
            // Jika slug sama, pastikan ini adalah halaman yang sama (sedang diedit)
            // Jika ID beda, berarti slug sudah dipakai orang lain
            if (isEdit && rowId === id) {
                // Ini halaman kita sendiri, lanjut
            } else {
                return { status: "error", message: "Slug URL sudah digunakan. Pilih slug lain." };
            }
        }
    }

    // Check if columns exist
    const maxCols = s.getMaxColumns();
    if (maxCols < 11) s.insertColumnsAfter(maxCols, 11 - maxCols);

    if (isEdit) {
      for (let i = 1; i < r.length; i++) {
        if (String(r[i][0]).trim() === id) {
          // Hanya izinkan edit jika owner cocok (atau admin bisa edit semua)
          const existingOwner = String(r[i][6] || "ADMIN").trim();
          
           if (existingOwner !== ownerId && ownerId !== "ADMIN") { 
              return { status: "error", message: "Anda tidak memiliki izin mengedit halaman ini." };
          }

          s.getRange(i + 1, 1, 1, 4).setValues([[d.id, slug, d.title, d.content]]);
          // Update Meta Pixel Columns (Col 8, 9, 10) + Theme Mode (Col 11)
          s.getRange(i + 1, 8, 1, 4).setValues([[d.meta_pixel_id || "", d.meta_pixel_token || "", d.meta_pixel_test_event || "", d.theme_mode || "light"]]);
          return withPublicCacheState_({ status: "success" }, bumpPublicCacheState_(["pages", "dashboard"]));
        }
      }
      return { status: "error", message: "ID Halaman tidak ditemukan" };
    } else {
      const newId = "PG-" + Date.now();
      // Tambahkan Owner ID di kolom ke-7 (index 6) + Meta Pixel (7,8,9) + Theme Mode (10)
      s.appendRow([newId, slug, d.title, d.content, "Active", toISODate_(), ownerId, d.meta_pixel_id || "", d.meta_pixel_token || "", d.meta_pixel_test_event || "", d.theme_mode || "light"]);
      return withPublicCacheState_({ status: "success" }, bumpPublicCacheState_(["pages", "dashboard"]));
    }
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function deletePage(d) {
  try {
    requireAdminSession_(d, { actionName: "delete_page" });
    const s = mustSheet_("Pages");
    const id = String(d.id).trim();
    const ownerId = String(d.owner_id || "ADMIN").trim();

    const r = s.getDataRange().getValues();
    for (let i = 1; i < r.length; i++) {
      if (String(r[i][0]).trim() === id) {
        // Security Check: Only Owner or Admin can delete
        const pageOwner = String(r[i][6] || "ADMIN").trim();
        if (pageOwner !== ownerId && ownerId !== "ADMIN") {
            return { status: "error", message: "Anda tidak memiliki izin menghapus halaman ini." };
        }
        
        s.deleteRow(i + 1);
        return withPublicCacheState_({ status: "success", message: "Halaman berhasil dihapus" }, bumpPublicCacheState_(["pages", "dashboard"]));
      }
    }
    return { status: "error", message: "ID Halaman tidak ditemukan" };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function checkSlug(d) {
  try {
    const s = mustSheet_("Pages");
    const slug = String(d.slug).trim();
    const excludeId = String(d.exclude_id || "").trim(); // For edit mode
    
    const r = s.getDataRange().getValues();
    for (let i = 1; i < r.length; i++) {
      const rowSlug = String(r[i][1]).trim();
      const rowId = String(r[i][0]).trim();
      
      if (rowSlug === slug) {
          if (excludeId && rowId === excludeId) {
              // Same page, it's fine
          } else {
              return { status: "success", available: false, message: "Slug URL sudah digunakan" };
          }
      }
    }
    return { status: "success", available: true, message: "Slug URL tersedia" };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function updateSettings(d) {
  requireAdminSession_(d, { actionName: "update_settings" });
  const cfg = getSettingsMap_();
  const payload = Object.assign({}, (d && d.payload && typeof d.payload === "object") ? d.payload : {});
  if (Object.prototype.hasOwnProperty.call(payload, "moota_secret") && !Object.prototype.hasOwnProperty.call(payload, "moota_token")) {
    payload.moota_token = payload.moota_secret;
  }
  if (Object.prototype.hasOwnProperty.call(payload, "moota_secret")) {
    delete payload.moota_secret;
  }
  const hasMootaPayload = Object.prototype.hasOwnProperty.call(payload, "moota_gas_url")
    || Object.prototype.hasOwnProperty.call(payload, "moota_token");
  if (hasMootaPayload) {
    const mootaCfg = resolveMootaConfig_(payload, cfg);
    const shouldValidateMoota = !!(mootaCfg.gasUrl || mootaCfg.token);
    if (shouldValidateMoota) {
      const mootaErrors = validateMootaConfigFormat_(mootaCfg);
      if (mootaErrors.length) {
        return { status: "error", message: mootaErrors[0], errors: mootaErrors };
      }
    }
  }

  const s = mustSheet_("Settings");
  const r = s.getDataRange().getValues();
  const propertyOnlyKeys = {
    ik_private_key: true,
    moota_token: true
  };
  for (let k in payload) {
    let nextValue = payload[k];
    if (k === "site_logo" || k === "site_favicon") {
      nextValue = sanitizeAssetUrl_(nextValue);
    }
    if (k === "moota_gas_url") {
      nextValue = normalizeMootaUrl_(nextValue);
    }
    const storeInPropertiesOnly = !!propertyOnlyKeys[k];
    if (storeInPropertiesOnly) {
      const props = PropertiesService.getScriptProperties();
      nextValue = String(nextValue || "").trim();
      if (nextValue) {
        props.setProperty(k, nextValue);
        if (k === "moota_token") props.deleteProperty("moota_secret");
      } else {
        props.deleteProperty(k);
      }
    }
    let f = false;
    for (let i = 1; i < r.length; i++) {
      if (r[i][0] === k) {
        s.getRange(i + 1, 2).setValue(storeInPropertiesOnly ? "" : nextValue);
        f = true;
        break;
      }
    }
    if (!f && !storeInPropertiesOnly) s.appendRow([k, nextValue]);
  }
  invalidateCaches_(["settings_map"]);
  return withPublicCacheState_({ status: "success" }, bumpPublicCacheState_(["settings", "dashboard"]));
}

function updateMootaGatewaySettings(d) {
  requireAdminSession_(d, { allowDemo: false, actionName: "update_moota_gateway" });
  const payload = (d && d.payload && typeof d.payload === "object") ? d.payload : d || {};
  return updateSettings({
    auth_session_token: getAdminSessionToken_(d),
    payload: {
      moota_gas_url: payload.moota_gas_url,
      moota_token: payload.moota_token !== undefined ? payload.moota_token : payload.moota_secret
    }
  });
}

function updateImageKitMediaSettings(d) {
  requireAdminSession_(d, { actionName: "update_imagekit_media" });
  const payload = (d && d.payload && typeof d.payload === "object") ? d.payload : d || {};
  return updateSettings({
    auth_session_token: getAdminSessionToken_(d),
    payload: {
      ik_public_key: payload.ik_public_key,
      ik_endpoint: payload.ik_endpoint,
      ik_private_key: payload.ik_private_key
    }
  });
}

function importMootaConfig(d) {
  requireAdminSession_(d, { allowDemo: false, actionName: "import_moota_config" });
  const payload = (d && d.payload && typeof d.payload === "object") ? d.payload : d || {};
  return updateSettings({
    auth_session_token: getAdminSessionToken_(d),
    payload: {
      moota_gas_url: payload.moota_gas_url,
      moota_token: payload.moota_token !== undefined ? payload.moota_token : payload.moota_secret
    }
  });
}

/* =========================
   IMAGEKIT AUTH
========================= */
function testImageKitConfig(d, cfg) {
  requireAdminSession_(d, { actionName: "test_ik_config" });
  cfg = cfg || getSettingsMap_();
  const ikCfg = resolveImageKitConfig_(d, cfg);
  const errors = validateImageKitConfigFormat_(ikCfg, { requireEndpoint: false });
  if (errors.length) {
    return { status: "error", message: errors[0], errors: errors };
  }

  const result = fetchImageKitFiles_(ikCfg.privateKey, 1);
  if (!result.ok) return { status: "error", message: result.message };

  const sampleFile = result.files.length ? result.files[0] : null;
  const sampleUrl = sampleFile ? String(sampleFile.url || "") : "";
  const inferredEndpoint = inferImageKitEndpointFromUrl_(sampleUrl);
  const warnings = [];

  if (!ikCfg.endpoint && inferredEndpoint) {
    warnings.push("URL endpoint berhasil dideteksi otomatis dari file yang ada di akun.");
  } else if (ikCfg.endpoint && inferredEndpoint && sampleUrl && sampleUrl.indexOf(ikCfg.endpoint) !== 0) {
    warnings.push("URL endpoint yang diisi tidak cocok dengan contoh URL file di akun. Periksa kembali URL endpoint ImageKit Anda.");
  }

  return {
    status: "success",
    message: "Koneksi ImageKit berhasil.",
    endpoint: ikCfg.endpoint || inferredEndpoint,
    inferred_endpoint: inferredEndpoint,
    sample_file_url: sampleUrl,
    warnings: warnings
  };
}

function getImageKitAuth(d, cfg) {
  requireAdminSession_(d, { actionName: "get_ik_auth" });
  cfg = cfg || getSettingsMap_();
  const ikCfg = resolveImageKitConfig_(d, cfg);
  const errors = validateImageKitConfigFormat_(ikCfg, { requirePublic: false, requireEndpoint: false, requirePrivate: true });
  if (errors.length) return { status: "error", message: errors[0] };

  const t = Utilities.getUuid();
  const exp = Math.floor(Date.now() / 1000) + 2400;
  const toSign = t + exp;

  const sig = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, toSign, ikCfg.privateKey)
    .map(b => ("0" + (b & 255).toString(16)).slice(-2))
    .join("");

  return { status: "success", token: t, expire: exp, signature: sig };
}

/* =========================
   CHANGE PASSWORD
========================= */
function changeUserPassword(d) {
  try {
    const s = mustSheet_("Users");
    const r = s.getDataRange().getValues();
    const email = String(d.email).trim().toLowerCase();
    const oldPass = String(d.old_password);
    const newPass = String(d.new_password);

    for (let i = 1; i < r.length; i++) {
      if (String(r[i][1]).trim().toLowerCase() === email) {
        if (verifyPassword_(oldPass, String(r[i][2] || ""))) {
          s.getRange(i + 1, 3).setValue(hashPassword_(newPass));
          return { status: "success", message: "Password berhasil diubah" };
        } else {
          return { status: "error", message: "Password lama salah!" };
        }
      }
    }
    return { status: "error", message: "Email pengguna tidak ditemukan." };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/* =========================
   UPDATE PROFILE (NAMA & EMAIL)
========================= */
function updateUserProfile(d) {
  try {
    const s = mustSheet_("Users");
    const r = s.getDataRange().getValues();
    const currentEmail = String(d.email).trim().toLowerCase();
    const newName = String(d.new_name).trim();
    const newEmail = String(d.new_email).trim().toLowerCase();
    const password = String(d.password); // Verify password before updating sensitive info

    if (!newName || !newEmail) return { status: "error", message: "Nama dan Email baru wajib diisi." };

    let userRowIndex = -1;
    let currentData = null;

    // 1. Verify User & Check duplicate email if changed
    for (let i = 1; i < r.length; i++) {
      const rowEmail = String(r[i][1]).trim().toLowerCase();
      
      // Find current user
      if (rowEmail === currentEmail) {
        if (!verifyPassword_(password, String(r[i][2] || ""))) return { status: "error", message: "Password salah!" };
        userRowIndex = i + 1;
        currentData = r[i];
      } 
      
      // Check if new email is already taken by SOMEONE ELSE
      if (rowEmail === newEmail && rowEmail !== currentEmail) {
        return { status: "error", message: "Email baru sudah digunakan oleh pengguna lain." };
      }
    }

    if (userRowIndex === -1) return { status: "error", message: "Pengguna tidak ditemukan." };

    // 2. Update Users Sheet
    // Col 2: Email (index 1), Col 4: Nama (index 3)
    // Note: getRange(row, col) is 1-based.
    s.getRange(userRowIndex, 2).setValue(newEmail);
    s.getRange(userRowIndex, 4).setValue(newName);

    // 3. Update Orders Sheet if email changed (Consistency)
    if (newEmail !== currentEmail) {
      const oS = mustSheet_("Orders");
      const oR = oS.getDataRange().getValues();
      for (let j = 1; j < oR.length; j++) {
        if (String(oR[j][1]).toLowerCase() === currentEmail) {
          oS.getRange(j + 1, 2).setValue(newEmail);
          oS.getRange(j + 1, 3).setValue(newName); // Update name as well
        }
      }
    } else {
       // Just update name in Orders if email same
      const oS = mustSheet_("Orders");
      const oR = oS.getDataRange().getValues();
      for (let j = 1; j < oR.length; j++) {
        if (String(oR[j][1]).toLowerCase() === currentEmail) {
          oS.getRange(j + 1, 3).setValue(newName);
        }
      }
    }

    return { status: "success", message: "Profil berhasil diperbarui", new_email: newEmail, new_name: newName };

  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/* =========================
   AFFILIATE PIXEL SETTINGS
========================= */
function saveAffiliatePixel(d) {
  try {
    const sName = "Affiliate_Pixels";
    let s = ss.getSheetByName(sName);
    if (!s) {
      s = ss.insertSheet(sName);
      s.appendRow(["user_id", "product_id", "pixel_id", "pixel_token", "pixel_test_code"]);
    }
    
    // 1. Get User ID from Email (Secure way: use login token if available, but here we trust email for now as it's backend call from trusted client logic)
    // Ideally we should use session token, but current system uses email.
    const email = String(d.email || "").trim().toLowerCase();
    if (!email) return { status: "error", message: "Email wajib diisi" };

    const uS = mustSheet_("Users");
    const uR = uS.getDataRange().getValues();
    let userId = "";
    
    for (let i = 1; i < uR.length; i++) {
      if (String(uR[i][1]).toLowerCase() === email) { 
        userId = String(uR[i][0]); 
        break; 
      }
    }
    
    if (!userId) return { status: "error", message: "User tidak ditemukan" };
    
    const productId = String(d.product_id).trim();
    const pixelId = String(d.pixel_id || "").trim();
    const pixelToken = String(d.pixel_token || "").trim();
    const pixelTest = String(d.pixel_test_code || "").trim();

    const r = s.getDataRange().getValues();
    let found = false;

    for (let i = 1; i < r.length; i++) {
      if (String(r[i][0]) === userId && String(r[i][1]) === productId) {
        // Update existing row (Col 3, 4, 5 -> index 2, 3, 4)
        s.getRange(i + 1, 3, 1, 3).setValues([[pixelId, pixelToken, pixelTest]]);
        found = true;
        break;
      }
    }

    if (!found) {
      s.appendRow([userId, productId, pixelId, pixelToken, pixelTest]);
    }
    
    return { status: "success", message: "Pixel berhasil disimpan" };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/* =========================
   PERMISSION WARMUP
========================= */
function pancinganIzin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) ss.getName();
  MailApp.getRemainingDailyQuota();
  try {
    UrlFetchApp.fetch("https://google.com");
  } catch (e) {
    // Ignore fetch errors
  }
  Logger.log("Pancingan sukses! Izin berhasil di-refresh.");
}

/* =========================
   AUTO-PAYMENT SYSTEM (MOOTA WEBHOOK)
========================= */
function handleMootaWebhook(mutations, cfg) {
  try {
    cfg = cfg || getSettingsMap_();

    // LOG: Raw incoming webhook for debugging
    logMoota_("WEBHOOK_IN", "Mutations count: " + mutations.length + " | Data masked");

    const s = mustSheet_("Orders");
    const orders = s.getDataRange().getValues();
    const siteName = getCfgFrom_(cfg, "site_name") || "Sistem Premium";
    const adminWA = getCfgFrom_(cfg, "wa_admin");

    const MAX_AGE_HOURS = 72; // Extended from 48 to 72 hours for better matching
    const matched = [];
    const debugLog = [];

    debugLog.push("MUTATIONS: " + mutations.length);

    for (let m = 0; m < mutations.length; m++) {
      const mutasi = mutations[m];
      const type = String(mutasi.type || "").toUpperCase();

      // Filter Credit only (Uang Masuk)
      if (type !== "CR" && type !== "CREDIT") {
        debugLog.push(`SKIP [${m}] Type=${type} (Not CR)`);
        logMoota_("SKIP_TYPE", "Mutation " + m + " type=" + type + " (not CR/CREDIT)");
        continue;
      }

      // Robust Amount Parsing (Handle number or string)
      let nominalTransfer = 0;
      if (typeof mutasi.amount === 'number') {
        nominalTransfer = mutasi.amount;
      } else {
        nominalTransfer = parseFloat(String(mutasi.amount || 0).replace(/[^0-9.-]/g, "")) || 0;
      }
      // Round to integer to avoid floating point issues
      nominalTransfer = Math.round(nominalTransfer);

      if (nominalTransfer <= 0) {
        debugLog.push(`SKIP [${m}] Amount=0`);
        logMoota_("SKIP_ZERO", "Mutation " + m + " amount=0 or negative");
        continue;
      }

        debugLog.push(`CHECKING Amount=${nominalTransfer}`);

      let foundMatch = false;
      // Collect pending orders info for debugging if no match
      let pendingOrders = [];

      // Iterate Orders to find match
      for (let i = 1; i < orders.length; i++) {
        const statusOrder = String(orders[i][7] || "").trim();
        
        // Hanya proses yang statusnya Pending
        if (statusOrder !== "Pending") continue;

        // Cek umur order
        if (MAX_AGE_HOURS > 0) {
          const dtStr = String(orders[i][8] || "").trim();
          const dt = new Date(dtStr);
          if (!isNaN(dt.getTime())) {
            const ageHours = (Date.now() - dt.getTime()) / 36e5;
            if (ageHours > MAX_AGE_HOURS) continue;
          }
        }

        const tagihanOrder = Math.round(toNumberSafe_(orders[i][6])); // Round to integer
        pendingOrders.push({ inv: orders[i][0], tagihan: tagihanOrder });
        
        // MATCHING LOGIC: Exact Amount (Rounded integers)
        if (tagihanOrder === nominalTransfer) {
          debugLog.push(`  MATCH FOUND Row ${i+1}: Inv=${orders[i][0]}`);
          logMoota_("MATCH", "Inv=" + orders[i][0] + " Amount=" + nominalTransfer + " Row=" + (i+1));
          
          // 1. UPDATE SHEET STATUS
          s.getRange(i + 1, 8).setValue("Lunas");
          orders[i][7] = "Lunas"; // Prevent double matching

          const inv = orders[i][0];
          const uEmail = orders[i][1];
          const uName = orders[i][2];
          const uWA = orders[i][3];
          const pId = orders[i][4];
          const pName = orders[i][5];

          // 2. GET ACCESS URL
          let accessUrl = "";
          const pS = ss.getSheetByName("Access_Rules");
          if (pS) {
            const pData = pS.getDataRange().getValues();
            for (let k = 1; k < pData.length; k++) {
              if (String(pData[k][0]) === String(pId)) { accessUrl = pData[k][3]; break; }
            }
          }

          // 3. SEND NOTIFICATIONS
          
          // LOG: Debug WA target before sending (diagnose Lunas WA failures)
          logWA_("DEBUG_MOOTA_LUNAS", String(uWA), "raw=" + JSON.stringify(uWA) + " type=" + typeof uWA + " normalized=" + normalizePhone_(uWA) + " | Inv=" + inv);

          // A) WA Customer
          sendWA(
            uWA,
            `🎉 *PEMBAYARAN DITERIMA!* 🎉\n\nHalo *${uName}*, pembayaran Anda sebesar Rp ${Number(nominalTransfer).toLocaleString('id-ID')} telah berhasil diverifikasi otomatis.\n\nPesanan *${pName}* (Invoice: #${inv}) kini *AKTIF*.\n\n🚀 *AKSES MATERI:* \n${accessUrl}\n\nTerima kasih!\n*Tim ${siteName}*`,
            cfg
          );

          // B) Email Customer
          const emailHtml = `
            <div style="font-family: sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px;">
                <h2 style="color: #10b981;">Pembayaran Berhasil! ✅</h2>
                <p>Halo <b>${uName}</b>,</p>
                <p>Pembayaran invoice <b>#${inv}</b> sebesar <b>Rp ${Number(nominalTransfer).toLocaleString('id-ID')}</b> telah diterima.</p>
                <p>Silakan akses produk <b>${pName}</b> melalui tombol di bawah ini:</p>
                <div style="text-align: center; margin: 30px 0;">
                    <a href="${accessUrl}" style="background-color: #4f46e5; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold;">Akses Materi</a>
                </div>
                <p>Terima kasih,<br><b>Tim ${siteName}</b></p>
            </div>`;
          sendEmail(uEmail, `Pembayaran Sukses: #${inv} - ${siteName}`, emailHtml, cfg);

          // C) WA Admin
          sendWA(
            adminWA,
            `💰 *MOOTA PAYMENT RECEIVED* 💰\n\nInv: #${inv}\nAmt: Rp ${Number(nominalTransfer).toLocaleString('id-ID')}\nUser: ${uName}\nProduk: ${pName}\n\nStatus: Auto-Lunas by System.`,
            cfg
          );

          foundMatch = true;
          matched.push(inv);
          break; // Stop searching orders for this mutation
        }
      }

      if (!foundMatch) {
        const pendingInfo = pendingOrders.map(o => o.inv + "=" + o.tagihan).join(", ");
        debugLog.push(`NO MATCH for Amount=${nominalTransfer} | Pending orders: ${pendingInfo}`);
        logMoota_("NO_MATCH", "Amount=" + nominalTransfer + " | Pending orders: " + pendingInfo);
        
        // Alert admin about unmatched payment (only for significant amounts)
        if (adminWA && nominalTransfer >= 10000) {
          sendWA(
            adminWA,
            `⚠️ *UNMATCHED PAYMENT* ⚠️\n\nTransfer masuk Rp ${Number(nominalTransfer).toLocaleString('id-ID')} dari Moota TIDAK COCOK dengan order manapun.\n\nDeskripsi: ${String(mutasi.description || "-").substring(0, 100)}\n\nPending Orders:\n${pendingOrders.length > 0 ? pendingOrders.slice(0, 5).map(o => "• " + o.inv + " = Rp " + Number(o.tagihan).toLocaleString('id-ID')).join("\n") : "(tidak ada order pending)"}\n\nMohon cek manual di dashboard.`,
            cfg
          );
        }
      }
    }

    const resultSummary = matched.length > 0
      ? "PROCESSED: " + matched.join(", ")
      : "NO_MATCHING_ORDER";
    logMoota_("RESULT", resultSummary + " | Logs: " + debugLog.join(" | "));
      
    return ContentService.createTextOutput(JSON.stringify({
       status: "success", 
       processed: matched, 
       logs: debugLog 
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    logMoota_("ERROR", e.toString());
    return ContentService.createTextOutput(JSON.stringify({
       status: "error", 
       message: e.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/* =========================
   FORGOT PASSWORD
========================= */
function forgotPassword(d) {
  try {
    const s = mustSheet_("Users");
    const r = s.getDataRange().getValues();
    const email = String(d.email).trim().toLowerCase();
    const cfg = getSettingsMap_();
    const siteName = getCfgFrom_(cfg, "site_name") || "Sistem Premium";
    
    let found = false;
    let nama = "";
    let rowIndex = -1;
    let tempPass = "";
    
    for (let i = 1; i < r.length; i++) {
      if (String(r[i][1]).trim().toLowerCase() === email) {
        rowIndex = i + 1;
        nama = r[i][3];
        found = true;
        break;
      }
    }
    
    if (found) {
        // Send Email
        const subject = `Lupa Password - ${siteName}`;
        tempPass = Math.random().toString(36).slice(-10);
        s.getRange(rowIndex, 3).setValue(hashPassword_(tempPass));

        const body = `
          <div style="font-family: sans-serif; padding: 20px;">
            <h3>Halo ${nama},</h3>
            <p>Anda meminta reset password akun.</p>
            <p>Password sementara Anda adalah:</p>
            <p><strong>Email:</strong> ${email}<br>
            <strong>Password Sementara:</strong> ${tempPass}</p>
            <p>Silakan login kembali lalu segera ganti password Anda.</p>
            <br>
            <p>Salam,<br>Tim ${siteName}</p>
          </div>
        `;
        
        sendEmail(email, subject, body, cfg);
        return { status: "success", message: "Password telah dikirim ke email anda." };
    }
    
    return { status: "error", message: "Email tidak ditemukan." };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/* =========================
   PAGINATION ACTIONS
========================= */
function getAdminOrders(d) {
  try {
    requireAdminSession_(d, { actionName: "get_admin_orders" });
    const page = Number(d.page) || 1;
    const limit = Number(d.limit) || 20;
    const o = mustSheet_("Orders").getDataRange().getValues();
    const data = o.slice(1).reverse();
    const start = (page - 1) * limit;
    const end = start + limit;
    
    return {
      status: "success",
      data: data.slice(start, end),
      has_more: data.length > end
    };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

function getAdminUsers(d) {
  try {
    requireAdminSession_(d, { actionName: "get_admin_users" });
    const page = Number(d.page) || 1;
    const limit = Number(d.limit) || 20;
    const u = mustSheet_("Users").getDataRange().getValues();
    const data = u.slice(1).reverse();
    const start = (page - 1) * limit;
    const end = start + limit;
    
    return {
      status: "success",
      data: data.slice(start, end),
      has_more: data.length > end
    };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

/* =========================
   DIAGNOSTIC & TEST FUNCTIONS
========================= */
function getEmailLogs_() {
  try {
    const s = ss.getSheetByName("Email_Logs");
    if (!s || s.getLastRow() <= 1) return { status: "success", data: [], message: "No email logs yet" };
    const data = s.getDataRange().getValues();
    return { status: "success", data: data.slice(1).reverse().slice(0, 50) };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function getMootaLogs_() {
  try {
    const s = ss.getSheetByName("Moota_Logs");
    if (!s || s.getLastRow() <= 1) return { status: "success", data: [], message: "No moota logs yet" };
    const data = s.getDataRange().getValues();
    return { status: "success", data: data.slice(1).reverse().slice(0, 50) };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function testEmailDelivery(d) {
  try {
    const email = String(d.email || "").trim();
    if (!email) return { status: "error", message: "Email target wajib diisi" };
    
    const cfg = getSettingsMap_();
    const siteName = getCfgFrom_(cfg, "site_name") || "Sistem Premium";
    
    const testHtml = '<div style="font-family: sans-serif; padding: 20px; max-width: 500px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 8px;">' +
      '<h2 style="color: #4f46e5;">✅ Test Email Berhasil!</h2>' +
      '<p>Ini adalah email test dari sistem <b>' + siteName + '</b>.</p>' +
      '<p><b>Waktu:</b> ' + new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" }) + '</p>' +
      '<p><b>Quota Tersisa:</b> ' + MailApp.getRemainingDailyQuota() + ' email</p>' +
      '<p>Jika Anda menerima email ini, berarti sistem email berfungsi normal.</p>' +
      '</div>';
    
    const result = sendEmail(email, "[TEST] Email Test - " + siteName, testHtml, cfg);
    return { status: "success", message: "Test email sent", result: result };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function testMootaWebhook() {
  try {
    const cfg = getSettingsMap_();
    const orders = mustSheet_("Orders").getDataRange().getValues();
    
    // Find a Pending order to simulate
    var testAmount = 0;
    var testInv = "";
    for (var i = orders.length - 1; i >= 1; i--) {
      if (String(orders[i][7]).trim() === "Pending") {
        testAmount = toNumberSafe_(orders[i][6]);
        testInv = orders[i][0];
        break;
      }
    }
    
    if (!testAmount) {
      return { status: "warning", message: "Tidak ada order Pending untuk di-test. Buat order test terlebih dahulu." };
    }
    
    // DRY RUN: simulate matching only, DO NOT actually update status
    return {
      status: "success",
      message: "Dry run - order ditemukan untuk matching",
      test_data: {
        invoice: testInv,
        amount: testAmount,
        would_match: true,
        note: "Ini hanya simulasi. Order TIDAK diubah statusnya. Untuk test penuh, kirim webhook asli dari Moota."
      }
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function getSystemHealth() {
  try {
    const cfg = getSettingsMap_();
    const emailQuota = MailApp.getRemainingDailyQuota();
    
    // Count pending orders
    const orders = mustSheet_("Orders").getDataRange().getValues();
    var pendingCount = 0;
    var oldPendingCount = 0;
    for (var i = 1; i < orders.length; i++) {
      if (String(orders[i][7]).trim() === "Pending") {
        pendingCount++;
        var dt = new Date(String(orders[i][8]));
        if (!isNaN(dt.getTime()) && (Date.now() - dt.getTime()) / 36e5 > 72) {
          oldPendingCount++;
        }
      }
    }
    
    // Check config
    const mootaCfg = resolveMootaConfig_({}, cfg);
    const mootaToken = mootaCfg.token;
    const mootaGasUrl = normalizeMootaUrl_(mootaCfg.gasUrl || getCurrentWebAppUrl_());
    const fonnteToken = getSecret_("fonnte_token", cfg);
    
    // Email log stats
    var emailLogCount = 0, emailFailCount = 0;
    var emailSheet = ss.getSheetByName("Email_Logs");
    if (emailSheet && emailSheet.getLastRow() > 1) {
      var eLogs = emailSheet.getDataRange().getValues();
      emailLogCount = eLogs.length - 1;
      for (var j = 1; j < eLogs.length; j++) {
        if (String(eLogs[j][1]) === "FAILED" || String(eLogs[j][1]) === "QUOTA_EXCEEDED") emailFailCount++;
      }
    }
    
    // Moota log stats
    var mootaLogCount = 0, mootaNoMatch = 0;
    var mootaSheet = ss.getSheetByName("Moota_Logs");
    if (mootaSheet && mootaSheet.getLastRow() > 1) {
      var mLogs = mootaSheet.getDataRange().getValues();
      mootaLogCount = mLogs.length - 1;
      for (var k = 1; k < mLogs.length; k++) {
        if (String(mLogs[k][1]) === "NO_MATCH") mootaNoMatch++;
      }
    }
    
    // WA log stats
    var waSentCount = 0, waFailCount = 0, waRejectedCount = 0, waLogCount = 0;
    var waSheet = ss.getSheetByName("WA_Logs");
    if (waSheet && waSheet.getLastRow() > 1) {
      var wLogs = waSheet.getDataRange().getValues();
      waLogCount = wLogs.length - 1;
      for (var w = 1; w < wLogs.length; w++) {
        var wStatus = String(wLogs[w][1]);
        if (wStatus === "SENT" || wStatus === "SENT_UNVERIFIED") waSentCount++;
        else if (wStatus === "REJECTED") waRejectedCount++;
        else if (wStatus === "HTTP_ERROR" || wStatus === "EXCEPTION" || wStatus === "NO_TOKEN") waFailCount++;
      }
    }
    
    return {
      status: "success",
      health: {
        email: {
          quota_remaining: emailQuota,
          quota_warning: emailQuota < 10,
          total_logs: emailLogCount,
          failed_count: emailFailCount
        },
        whatsapp: {
          total_logs: waLogCount,
          sent_count: waSentCount,
          rejected_count: waRejectedCount,
          failed_count: waFailCount,
          sent_rate: waLogCount > 0 ? Math.round((waSentCount / waLogCount) * 100) + "%" : "N/A"
        },
        moota: {
          gas_url_configured: !!mootaGasUrl,
          secret_token_configured: !!mootaToken,
          total_webhooks: mootaLogCount,
          unmatched_count: mootaNoMatch
        },
        orders: {
          pending_count: pendingCount,
          stale_pending: oldPendingCount
        },
        integrations: {
          fonnte_configured: !!fonnteToken,
          moota_configured: !!mootaToken && !!mootaGasUrl
        }
      }
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function getWALogs_() {
  try {
    var s = ss.getSheetByName("WA_Logs");
    if (!s || s.getLastRow() <= 1) return { status: "success", data: [], message: "No WA logs yet" };
    var data = s.getDataRange().getValues();
    return { status: "success", data: data.slice(1).reverse().slice(0, 50) };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function testWADelivery(d) {
  try {
    var target = String(d.target || d.whatsapp || "").trim();
    if (!target) return { status: "error", message: "Nomor WhatsApp target wajib diisi (parameter: target)" };
    
    var cfg = getSettingsMap_();
    var siteName = getCfgFrom_(cfg, "site_name") || "Sistem Premium";
    var testMessage = "✅ *TEST WA BERHASIL!*\n\nIni adalah pesan test dari sistem *" + siteName + "*.\n\nWaktu: " + new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" }) + "\n\nJika Anda menerima pesan ini, berarti koneksi WhatsApp via Fonnte berfungsi normal.";
    
    var result = sendWA(target, testMessage, cfg);
    return { status: "success", message: "Test WA sent to " + target, result: result };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/**
 * testLunasNotification — Simulates the EXACT Lunas notification flow.
 * Finds a pending/existing order and sends WA + Email using the same code path 
 * as updateOrderStatus. Does NOT change the order status.
 * 
 * Call: {"action":"test_lunas_notification","invoice":"INV-XXXXX"}
 * Or:   {"action":"test_lunas_notification"} (auto-finds the latest pending order)
 */
function testLunasNotification(d) {
  try {
    var cfg = getSettingsMap_();
    var s = mustSheet_("Orders");
    var pS = mustSheet_("Access_Rules");
    var r = s.getDataRange().getValues();
    var siteName = getCfgFrom_(cfg, "site_name") || "Sistem Premium";
    var targetInv = String(d.invoice || d.id || "").trim();
    
    // Find order (specific or latest pending)
    var orderRow = null;
    var orderRowIdx = -1;
    for (var i = r.length - 1; i >= 1; i--) {
      if (targetInv) {
        if (String(r[i][0]) === targetInv) { orderRow = r[i]; orderRowIdx = i; break; }
      } else {
        if (String(r[i][7]).trim() === "Pending") { orderRow = r[i]; orderRowIdx = i; break; }
      }
    }
    
    if (!orderRow) {
      return { status: "error", message: targetInv ? "Invoice " + targetInv + " tidak ditemukan" : "Tidak ada order Pending. Buat order test dulu." };
    }
    
    var inv = orderRow[0];
    var uEmail = orderRow[1];
    var uName = orderRow[2];
    var uWA = orderRow[3];
    var pId = orderRow[4];
    var pName = orderRow[5];
    
    // Debug: capture raw data from sheet
    var debugInfo = {
      invoice: inv,
      row_index: orderRowIdx + 1,
      wa_raw_value: uWA,
      wa_raw_type: typeof uWA,
      wa_json: JSON.stringify(uWA),
      wa_normalized: normalizePhone_(uWA),
      email: uEmail,
      name: uName,
      product: pName,
      current_status: orderRow[7]
    };
    
    // Get access URL
    var accessUrl = "";
    var pData = pS.getDataRange().getValues();
    for (var k = 1; k < pData.length; k++) {
      if (String(pData[k][0]) === String(pId)) { accessUrl = pData[k][3]; break; }
    }
    debugInfo.access_url = accessUrl;
    
    // SEND WA (same message as real Lunas flow)
    logWA_("TEST_LUNAS", String(uWA), "Testing Lunas notification for " + inv + " | WA raw=" + JSON.stringify(uWA) + " type=" + typeof uWA);
    var waResult = sendWA(
      uWA,
      "🎉 *[TEST] PEMBAYARAN TERVERIFIKASI!* 🎉\n\nHalo *" + uName + "*, ini adalah TEST notifikasi Lunas.\n\nProduk *" + pName + "* (Invoice: #" + inv + ")\n\n🚀 *AKSES MATERI:*\n" + accessUrl + "\n\nIni pesan test. Jika terkirim berarti notifikasi Lunas berfungsi normal.\n*Tim " + siteName + "*",
      cfg
    );
    
    // SEND EMAIL (same template as real Lunas flow)
    var emailHtml = '<div style="font-family:sans-serif;max-width:600px;margin:0 auto;padding:20px;border:1px solid #e2e8f0;border-radius:8px;">' +
      '<h2 style="color:#10b981;">[TEST] Akses Terbuka! 🎉</h2>' +
      '<p>Halo <b>' + uName + '</b>,</p>' +
      '<p>Ini adalah TEST notifikasi Lunas untuk produk <b>' + pName + '</b>.</p>' +
      '<div style="text-align:center;margin:30px 0;">' +
      '<a href="' + accessUrl + '" style="background-color:#4f46e5;color:white;padding:12px 24px;text-decoration:none;border-radius:6px;font-weight:bold;">Akses Materi</a>' +
      '</div>' +
      '<p>Jika Anda menerima email ini, notifikasi Lunas berfungsi normal.</p>' +
      '<p>Tim <b>' + siteName + '</b></p></div>';
    var emailResult = sendEmail(uEmail, "[TEST] Akses Terbuka - " + siteName, emailHtml, cfg);
    
    return {
      status: "success",
      message: "Test Lunas notification sent for " + inv,
      debug: debugInfo,
      results: {
        wa: waResult,
        email: emailResult
      }
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}
