/**
 * Cloudflare Worker — OMS Onflow → Apps Script proxy
 *
 * Mục đích: OMS UI gửi OPTIONS preflight CORS khi click "Xác thực" callback URL.
 * Apps Script Web App KHÔNG hỗ trợ OPTIONS → trả 405 HTML → OMS báo URL invalid.
 *
 * Worker này:
 *   - Trả 204 + CORS headers cho OPTIONS preflight (OMS xác thực pass)
 *   - Forward POST/GET sang Apps Script (URL + token đọc từ env vars — không hardcode)
 *
 * Secrets (set qua `wrangler secret put NAME`, không commit):
 *   - APPS_SCRIPT_URL  — URL Web App Wellhome (gồm deploymentId)
 *   - OMS_TOKEN        — token verify webhook giữa Apps Script và Worker
 *
 * Deploy: `wrangler deploy` từ thư mục này.
 */

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-OMS-Signature, X-Webhook-Signature',
  'Access-Control-Max-Age': '86400',
};

export default {
  async fetch(request, env) {
    // OPTIONS preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS_HEADERS });
    }

    if (!env.APPS_SCRIPT_URL || !env.OMS_TOKEN) {
      return new Response(
        JSON.stringify({ ok: false, error: 'Worker secrets not configured' }),
        { status: 500, headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' } }
      );
    }

    const upstreamUrl = `${env.APPS_SCRIPT_URL}?action=oms_webhook&oms_token=${env.OMS_TOKEN}`;

    let body;
    if (request.method !== 'GET' && request.method !== 'HEAD') {
      try { body = await request.text(); } catch (err) { body = ''; }
    }

    try {
      const upstream = await fetch(upstreamUrl, {
        method: request.method,
        headers: {
          'Content-Type': request.headers.get('Content-Type') || 'application/json',
        },
        body,
        redirect: 'follow',
      });
      const text = await upstream.text();
      return new Response(text, {
        status: 200,
        headers: { ...CORS_HEADERS, 'Content-Type': 'application/json; charset=utf-8' },
      });
    } catch (err) {
      return new Response(
        JSON.stringify({ ok: false, error: 'Proxy fetch fail: ' + err.message }),
        { status: 502, headers: { ...CORS_HEADERS, 'Content-Type': 'application/json; charset=utf-8' } }
      );
    }
  },
};
