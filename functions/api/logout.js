// Limpa o cookie de sessao HttpOnly (o cliente nao consegue apaga-lo via JS).
export async function onRequestPost() {
  const cookie = 'ab_session=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0';
  return new Response(JSON.stringify({ ok: true }), {
    status: 200,
    headers: { 'Content-Type': 'application/json', 'Set-Cookie': cookie },
  });
}
