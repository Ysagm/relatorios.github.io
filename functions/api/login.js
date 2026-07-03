const enc = new TextEncoder();

function b64url(buf) {
  return btoa(String.fromCharCode(...new Uint8Array(buf)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}

async function createJWT(payload, secret) {
  const header  = b64url(enc.encode(JSON.stringify({ alg: 'HS256', typ: 'JWT' })));
  const body    = b64url(enc.encode(JSON.stringify({
    ...payload,
    iat: Math.floor(Date.now() / 1000),
    exp: Math.floor(Date.now() / 1000) + 86400 * 7,   // 7 days
  })));
  const key = await crypto.subtle.importKey(
    'raw', enc.encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']
  );
  const sig = await crypto.subtle.sign('HMAC', key, enc.encode(`${header}.${body}`));
  return `${header}.${body}.${b64url(sig)}`;
}

export async function onRequestPost(context) {
  const { request, env } = context;

  let hash;
  try {
    ({ hash } = await request.json());
  } catch {
    return json({ error: 'Bad request' }, 400);
  }

  if (!hash || typeof hash !== 'string') return json({ error: 'Bad request' }, 400);

  let config;
  try {
    config = JSON.parse(env.PASSWORDS_CONFIG);
  } catch {
    return json({ error: 'Server configuration error' }, 500);
  }

  const session = config[hash.toLowerCase()];
  if (!session) return json({ error: 'Senha incorreta. Tente novamente.' }, 401);

  const token = await createJWT(session, env.JWT_SECRET);
  return json({ token, session });
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}
