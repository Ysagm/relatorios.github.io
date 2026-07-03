const enc = new TextEncoder();

function b64urlDecode(str) {
  str = str.replace(/-/g, '+').replace(/_/g, '/');
  while (str.length % 4) str += '=';
  return Uint8Array.from(atob(str), c => c.charCodeAt(0));
}

async function verifyJWT(token, secret) {
  const parts = token.split('.');
  if (parts.length !== 3) return null;
  const [header, payload, sig] = parts;

  const key = await crypto.subtle.importKey(
    'raw', enc.encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['verify']
  );
  const valid = await crypto.subtle.verify(
    'HMAC', key, b64urlDecode(sig), enc.encode(`${header}.${payload}`)
  );
  if (!valid) return null;

  const session = JSON.parse(new TextDecoder().decode(b64urlDecode(payload)));
  if (session.exp && session.exp < Math.floor(Date.now() / 1000)) return null;
  return session;
}

export async function onRequest(context) {
  const { request, next, env } = context;
  const url = new URL(request.url);

  // Only intercept requests to /data/* — everything else is served normally
  if (!url.pathname.startsWith('/data/')) return next();

  // Extract JWT from Authorization header
  const authHeader = request.headers.get('Authorization') || '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;

  if (!token) return new Response('Unauthorized', { status: 401 });

  let session;
  try {
    session = await verifyJWT(token, env.JWT_SECRET);
  } catch {
    return new Response('Unauthorized', { status: 401 });
  }
  if (!session) return new Response('Unauthorized', { status: 401 });

  // Determine which aircraft file is being requested
  // Expected paths: /data/owner.json  or  /data/PP-AGN.json
  const match = url.pathname.match(/^\/data\/(.+)\.json$/);
  if (!match) return new Response('Not Found', { status: 404 });
  const requested = match[1];   // "owner" or a registration like "PP-AGN"

  if (session.role === 'owner') {
    // Owner may only access the owner bundle
    if (requested !== 'owner') return new Response('Forbidden', { status: 403 });
    return next();
  }

  if (session.role === 'aircraft' && requested === session.allowedAircraft) {
    return next();
  }

  return new Response('Forbidden', { status: 403 });
}
