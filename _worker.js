const buildCorsHeaders = (origin, allowedOrigins) => {
  const allowOrigin = allowedOrigins.includes(origin)
    ? origin
    : allowedOrigins[0] || "*";

  return {
    "Access-Control-Allow-Origin": allowOrigin,
    "Access-Control-Allow-Methods": "GET,POST,PUT,PATCH,DELETE,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type,Authorization",
    "Access-Control-Max-Age": "86400",
    Vary: "Origin",
  };
};

export default {
  async fetch(request, env) {
    const origin = request.headers.get("Origin") || "";
    const allowedOrigins = (env.ALLOWED_ORIGINS || "")
      .split(",")
      .map((item) => item.trim())
      .filter(Boolean);

    const corsHeaders = buildCorsHeaders(origin, allowedOrigins);

    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: corsHeaders,
      });
    }

    const url = new URL(request.url);
    if (url.pathname === "/api/health") {
      return new Response(
        JSON.stringify({
          ok: true,
          worker: "zhostecourse",
          domain: "ecourse.zhost.digital",
        }),
        {
          status: 200,
          headers: {
            "content-type": "application/json; charset=utf-8",
            ...corsHeaders,
          },
        }
      );
    }

    if (env.ASSETS && typeof env.ASSETS.fetch === "function") {
      const assetResponse = await env.ASSETS.fetch(request);
      const headers = new Headers(assetResponse.headers);
      Object.entries(corsHeaders).forEach(([key, value]) => {
        headers.set(key, value);
      });
      return new Response(assetResponse.body, {
        status: assetResponse.status,
        statusText: assetResponse.statusText,
        headers,
      });
    }

    return new Response("Asset binding is not configured.", {
      status: 500,
      headers: {
        "content-type": "text/plain; charset=utf-8",
        ...corsHeaders,
      },
    });
  },
};
