const MAIL_SHEET = "https://docs.google.com/spreadsheets/d/1e5WsZIUj2d-nqOgFHSSgaWosXtowq3Ccwkpc43CtTWo/edit";
const CACHE_MS = 2 * 60 * 1000;
const EMAIL = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;
let cache = null;

function exportUrl(link) {
  const id = (link.match(/\/spreadsheets\/d\/([A-Za-z0-9_-]+)/) || [])[1];
  return id ? `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx` : link;
}

const unescapeXml = (s) => s
  .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
  .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(+d))
  .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const attr = (tag, name) => unescapeXml((tag.match(new RegExp(`\\s${name}="([^"]*)"`)) || [])[1] || "");
const texts = (xml) => [...xml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => unescapeXml(m[1])).join("");

async function unzip(buf) {
  const view = new DataView(buf), bytes = new Uint8Array(buf), dec = new TextDecoder();
  let eocd = -1;
  for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("not-xlsx");
  const files = new Map();
  let p = view.getUint32(eocd + 16, true);
  for (let k = view.getUint16(eocd + 10, true); k > 0; k--) {
    const method = view.getUint16(p + 10, true), size = view.getUint32(p + 20, true);
    const nameLen = view.getUint16(p + 28, true), extra = view.getUint16(p + 30, true), note = view.getUint16(p + 32, true);
    const local = view.getUint32(p + 42, true);
    const name = dec.decode(bytes.subarray(p + 46, p + 46 + nameLen));
    const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    files.set(name, {method, data: bytes.subarray(start, start + size)});
    p += 46 + nameLen + extra + note;
  }
  return async (name) => {
    const f = files.get(name);
    if (!f) return "";
    if (f.method === 0) return dec.decode(f.data);
    const stream = new Blob([f.data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return dec.decode(await new Response(stream).arrayBuffer());
  };
}

async function readMailBook(buf) {
  const read = await unzip(buf);
  const rels = new Map([...(await read("xl/_rels/workbook.xml.rels")).matchAll(/<Relationship\b[^>]*>/g)]
    .map(([tag]) => [attr(tag, "Id"), attr(tag, "Target")]));
  const shared = [...(await read("xl/sharedStrings.xml")).matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => texts(m[1]));
  const tabs = [];
  for (const [tag] of (await read("xl/workbook.xml")).matchAll(/<sheet\b[^>]*>/g)) {
    const target = rels.get(attr(tag, "r:id")) || "";
    const path = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
    const xml = await read(path);
    const values = [...xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)].map(([, a, body = ""]) => {
      const v = (body.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
      return attr(a, "t") === "s" ? shared[+v] || "" : texts(body) || unescapeXml(v || "");
    });
    const seen = new Set(), emails = [];
    for (const m of values.join("\n").matchAll(EMAIL)) {
      const key = m[0].toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      emails.push(m[0]);
    }
    tabs.push({name: attr(tag, "name"), emails});
  }
  return tabs;
}

async function loadMailLists() {
  if (cache && Date.now() - cache.time < CACHE_MS) return cache.result;
  let result, reason;
  try {
    const res = await fetch(exportUrl(MAIL_SHEET), {credentials: "include", cache: "no-store"});
    const host = new URL(res.url).hostname, type = res.headers.get("content-type") || "";
    if (host.startsWith("accounts.")) reason = "Google bắt đăng nhập";
    else if (res.status === 401 || res.status === 403) reason = `không có quyền tải sheet, HTTP ${res.status}`;
    else if (!res.ok) reason = `HTTP ${res.status}`;
    else if (type.includes("text/html")) reason = "Google trả về trang web thay vì file";
    else result = {ok: true, tabs: await readMailBook(await res.arrayBuffer())};
  } catch (e) {
    reason = e?.name === "TypeError" ? "bị chặn khi gọi Google, có thể Chrome không gửi kèm đăng nhập" : `lỗi đọc file: ${e.message}`;
  }
  if (!result) result = {ok: false, link: MAIL_SHEET, error: `Không đọc được sheet mail (${reason}).`};
  else cache = {time: Date.now(), result};
  return result;
}

if (globalThis.chrome?.runtime?.onMessage) {
  chrome.runtime.onMessage.addListener((msg, _sender, reply) => {
    if (msg?.type !== "mail-lists") return false;
    loadMailLists().then(reply);
    return true;
  });
}
