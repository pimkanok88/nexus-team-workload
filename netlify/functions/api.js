/**
 * Netlify Function proxy
 * Browser -> /api -> Netlify Function -> Apps Script Web App -> Google Sheet
 *
 * Required Netlify environment variables:
 *   APPS_SCRIPT_API_URL
 *   APPS_SCRIPT_API_SECRET
 */

const JSON_HEADERS = {
  'content-type': 'application/json; charset=utf-8',
  'cache-control': 'no-store'
};

function reply(statusCode, payload) {
  return {
    statusCode,
    headers: JSON_HEADERS,
    body: JSON.stringify(payload)
  };
}

export const handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers: JSON_HEADERS, body: '' };
  }

  if (event.httpMethod !== 'POST') {
    return reply(405, { ok: false, error: 'Method not allowed' });
  }

  const appsScriptUrl = String(process.env.APPS_SCRIPT_API_URL || '').trim();
  const secret = String(process.env.APPS_SCRIPT_API_SECRET || '').trim();

  if (!appsScriptUrl || !secret) {
    return reply(500, {
      ok: false,
      error: 'ยังไม่ได้ตั้ง APPS_SCRIPT_API_URL หรือ APPS_SCRIPT_API_SECRET ใน Netlify'
    });
  }

  let incoming;
  try {
    incoming = JSON.parse(event.body || '{}');
  } catch (err) {
    return reply(400, { ok: false, error: 'Request JSON ไม่ถูกต้อง' });
  }

  if (!incoming.fn) {
    return reply(400, { ok: false, error: 'ไม่พบชื่อ API function' });
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 25000);

  try {
    const upstream = await fetch(appsScriptUrl, {
      method: 'POST',
      headers: {
        // text/plain avoids any special JSON handling and Apps Script still gets postData.contents.
        'content-type': 'text/plain;charset=UTF-8'
      },
      body: JSON.stringify({
        fn: incoming.fn,
        args: Array.isArray(incoming.args) ? incoming.args : [],
        actor: String(incoming.actor || ''),
        secret
      }),
      redirect: 'follow',
      signal: controller.signal
    });

    const text = await upstream.text();

    let payload;
    try {
      payload = JSON.parse(text);
    } catch (err) {
      const looksLikeHtml = /<html|<!doctype/i.test(text);
      return reply(502, {
        ok: false,
        error: looksLikeHtml
          ? 'Apps Script ส่งหน้า HTML กลับมาแทน API กรุณาตรวจว่า Web App deploy เป็น Execute as: Me และ Who has access: Anyone'
          : 'Apps Script ตอบกลับไม่ใช่ JSON',
        detail: text.slice(0, 240)
      });
    }

    if (!upstream.ok) {
      return reply(502, {
        ok: false,
        error: payload.error || `Apps Script HTTP ${upstream.status}`
      });
    }

    // Pass Apps Script's {ok,data,error} envelope back to the browser.
    return reply(payload.ok ? 200 : 400, payload);
  } catch (err) {
    const message = err && err.name === 'AbortError'
      ? 'Apps Script ใช้เวลาตอบเกิน 25 วินาที'
      : (err && err.message ? err.message : String(err));

    return reply(502, {
      ok: false,
      error: 'เชื่อมต่อ Apps Script API ไม่สำเร็จ: ' + message
    });
  } finally {
    clearTimeout(timer);
  }
};
