/**
 * Assistant ingest — Cloudflare Workers örneği.
 * Deploy: wrangler.toml ile bu dosyayı entry yapın; ortamda BRANCH_KEY veya ALLOWED_BRANCH_KEYS tanımlayın.
 *
 * Cloudflare "error code: 1101" = yakalanmamış istisna (genelde env.R2 vb. binding eksik veya prod kodda hata).
 *
 * İstek (Electron): POST JSON, header X-Branch-Key
 * Gövde: { version, exportKind, fileName, fileBase64, grid, sentAt, reportDate }
 * reportDate: YYYY-MM-DD (gün sonu=tarih seçici; stok=seçilen ayın son günü; diğer=bugün)
 */

export default {
  async fetch(request, env, _ctx) {
    try {
      return await handleRequest(request, env);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      console.error("[ingest] unhandled:", e);
      return json(
        {
          ok: false,
          error: "internal",
          message: msg,
          hint:
            "Wrangler’da R2/D1/KV binding’i tanımlı mı? Dashboard → Workers → Logs’ta aynı zamanı kontrol edin.",
        },
        500,
      );
    }
  },
};

async function handleRequest(request, env) {
  if (request.method === "OPTIONS") {
    return new Response(null, {
      headers: corsHeaders(),
    });
  }

  const url = new URL(request.url);
  if (
    request.method !== "POST" ||
    (url.pathname !== "/" && url.pathname !== "/ingest")
  ) {
    if (request.method === "GET") {
      return new Response("Assistant ingest — POST / veya /ingest", {
        status: 200,
        headers: corsHeaders(),
      });
    }
    return new Response("Not found", { status: 404, headers: corsHeaders() });
  }

  const key = (request.headers.get("X-Branch-Key") || "").trim();
  const e = env || {};
  const single = String(e.BRANCH_KEY || "").trim();
  const list = String(e.ALLOWED_BRANCH_KEYS || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
  const authorized =
    (single && key === single) || (list.length > 0 && list.includes(key));

  if (!authorized) {
    return json({ ok: false, error: "unauthorized" }, 401);
  }

  let body;
  try {
    body = await request.json();
  } catch {
    return json({ ok: false, error: "invalid_json" }, 400);
  }

  if (!body?.fileName || !body?.fileBase64) {
    return json({ ok: false, error: "missing_file" }, 400);
  }

  // --- Burada R2 / D1 / Telegram entegrasyonu ---
  // wrangler.toml’da [[r2_buckets]] yoksa env.R2 kullanmayın; yoksa 1101 benzeri çöker.
  // const bytes = Uint8Array.from(atob(body.fileBase64), (c) => c.charCodeAt(0));
  // await env.R2.put(`${key}/${body.fileName}`, bytes);

  return json(
    {
      ok: true,
      fileName: body.fileName,
      exportKind: body.exportKind || "generic",
      receivedAt: new Date().toISOString(),
      message: "Worker örneği: dosya alındı (depolama henüz bağlanmadı).",
    },
    200,
  );
}

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, X-Branch-Key",
  };
}

function json(data, status) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json",
      ...corsHeaders(),
    },
  });
}

/*
 * İsteğe bağlı: Telegram’da indirme linki
 * - R2’ye public custom domain bağlamak, veya GET ?file=... ile Worker’da R2.get + stream.
 * - Presigned URL: süreli link; S3 uyumlu imza veya R2 API ile ayrıca kurulur.
 */
