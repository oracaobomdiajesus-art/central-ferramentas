export async function onRequest(context) {
  const { request, next, env } = context;
  const auth = request.headers.get("Authorization");

  if (auth) {
    const [scheme, encoded] = auth.split(" ");
    if (scheme === "Basic" && encoded) {
      const decoded = atob(encoded);
      const sep = decoded.indexOf(":");
      const pass = decoded.slice(sep + 1);
      if (env.SITE_PASSWORD && pass === env.SITE_PASSWORD) {
        return next();
      }
    }
  }

  return new Response("Autenticação necessária", {
    status: 401,
    headers: { "WWW-Authenticate": 'Basic realm="Area restrita"' }
  });
}
