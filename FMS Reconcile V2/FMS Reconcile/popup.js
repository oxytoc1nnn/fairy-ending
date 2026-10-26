const FMS_URL = "https://spx.shopee.vn/#/orderTracking";
const input = document.getElementById("lt");
const button = document.getElementById("go");
const status = document.getElementById("status");
let tab = null;

const say = (text, bad = false) => {
  status.textContent = text;
  status.className = bad ? "bad" : "";
};

async function findFmsTab() {
  const [active] = await chrome.tabs.query({active: true, currentWindow: true});
  if (active?.url?.startsWith("https://spx.shopee.vn/")) return active;
  const tabs = await chrome.tabs.query({url: "https://spx.shopee.vn/*"});
  return tabs.sort((a, b) => (b.lastAccessed || 0) - (a.lastAccessed || 0))[0] || null;
}

async function init() {
  tab = await findFmsTab();
  if (tab) {
    say("Sẽ chạy trên tab FMS đang mở.");
    input.focus();
  } else {
    button.textContent = "Mở FMS";
    say("Chưa mở FMS. Mở và đăng nhập trước, rồi bấm lại icon.");
  }
}

async function start() {
  if (!tab) {
    await chrome.tabs.create({url: FMS_URL});
    window.close();
    return;
  }
  const raw = input.value.toUpperCase().split(/[\s,;]+/).filter(Boolean);
  const lts = [...new Set(raw)];
  const bad = lts.filter((x) => !/^LT[A-Z0-9]+$/.test(x));
  if (!lts.length) return say("Nhập ít nhất 1 mã LT.", true);
  if (bad.length) return say(`Mã không hợp lệ: ${bad.join(", ")}`, true);

  button.disabled = true;
  try {
    const target = {tabId: tab.id};
    await chrome.scripting.executeScript({target, files: ["xlsx.js", "job.js"]});
    const [{result}] = await chrome.scripting.executeScript({
      target,
      args: [lts],
      func: (codes) => {
        if (globalThis.FmsReconcile.isRunning()) return "busy";
        globalThis.FmsReconcile.run(codes);
        return "started";
      },
    });
    if (result === "busy") {
      button.disabled = false;
      return say("Tab FMS đang xuất dở. Chờ xong rồi chạy tiếp.", true);
    }
    await chrome.tabs.update(tab.id, {active: true});
    await chrome.windows.update(tab.windowId, {focused: true});
    window.close();
  } catch (e) {
    button.disabled = false;
    say(`Không chạy được trên tab FMS (${e.message}). Tải lại trang FMS rồi thử lại.`, true);
  }
}

button.addEventListener("click", start);
input.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) start();
});
init();
