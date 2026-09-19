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

// Token pode vir no header Authorization (fetch de /data/*) ou no cookie
// HttpOnly ab_session (navegacao direta a um PDF em /docs/*, que nao manda header).
function getToken(request) {
  const auth = request.headers.get('Authorization') || '';
  if (auth.startsWith('Bearer ')) return auth.slice(7);
  const cookie = request.headers.get('Cookie') || '';
  const m = cookie.match(/(?:^|;\s*)ab_session=([^;]+)/);
  return m ? decodeURIComponent(m[1]) : null;
}

export async function onRequest(context) {
  const { request, next, env } = context;
  const url = new URL(request.url);

  const isData = url.pathname.startsWith('/data/');   // JSON de manutencao (fetch)
  const isDocs = url.pathname.startsWith('/docs/');    // PDFs sigilosos (navegacao)
  if (!isData && !isDocs) return next();               // resto do site segue normal

  // Sem sessao valida: /data -> 401 (e fetch); /docs -> manda pra tela de login.
  const unauthorized = () => isDocs
    ? Response.redirect(new URL('/', url).toString(), 302)
    : new Response('Unauthorized', { status: 401 });

  const token = getToken(request);
  if (!token) return unauthorized();

  let session;
  try {
    session = await verifyJWT(token, env.JWT_SECRET);
  } catch {
    session = null;
  }
  if (!session) return unauthorized();

  if (isData) {
    // Esperado: /data/owner.json  ou  /data/<REGISTRO>.json
    const match = url.pathname.match(/^\/data\/(.+)\.json$/);
    if (!match) return new Response('Not Found', { status: 404 });
    const requested = match[1];
    if (session.role === 'owner') {
      return requested === 'owner' ? next() : new Response('Forbidden', { status: 403 });
    }
    if (session.role === 'aircraft' && requested === session.allowedAircraft) return next();
    return new Response('Forbidden', { status: 403 });
  }

  // isDocs: /docs/<REGISTRO>/arquivo.pdf — cada cliente so acessa a propria aeronave.
  const m = url.pathname.match(/^\/docs\/([^/]+)\//);
  if (!m) return new Response('Not Found', { status: 404 });
  const reg = m[1];
  if (session.role === 'owner') return next();
  if (session.role === 'aircraft' && reg === session.allowedAircraft) return next();
  return new Response('Forbidden', { status: 403 });
}
