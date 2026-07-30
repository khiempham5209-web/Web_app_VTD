module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    res.status(204).end();
    return;
  }
  if (req.method !== "POST") {
    res.status(405).json({ ok: false, message: "Method not allowed." });
    return;
  }

  try {
    const body = typeof req.body === "object" && req.body !== null ? req.body : JSON.parse(String(req.body || "{}"));
    const targetUrl = String(body.targetUrl || "").trim();
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[-\w]+\/exec/i.test(targetUrl)) {
      res.status(400).json({ ok: false, message: "Target Apps Script URL khong hop le." });
      return;
    }
    const payload = body.payload && typeof body.payload === "object"
      ? body.payload
      : { action: body.action || "", params: body.params || {} };

    const upstream = await fetch(targetUrl, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
      redirect: "follow"
    });
    const text = await upstream.text();
    const contentType = String(upstream.headers.get("content-type") || "");
    try {
      JSON.parse(text);
      res.setHeader("Content-Type", "application/json; charset=utf-8");
      res.status(upstream.status).send(text);
    } catch (parseErr) {
      const compact = text.replace(/\s+/g, " ").slice(0, 300);
      const targetTail = targetUrl.replace(/^https:\/\/script\.google\.com\/macros\/s\//i, "").slice(0, 18);
      res.status(502).json({
        ok: false,
        message: "Apps Script tra ve HTML/khong phai JSON. Target " + targetTail + "... status " + upstream.status + ". Preview: " + compact,
        upstreamStatus: upstream.status,
        upstreamContentType: contentType,
        targetTail,
        preview: compact
      });
    }
  } catch (err) {
    res.status(502).json({ ok: false, message: err && err.message ? err.message : String(err) });
  }
};
