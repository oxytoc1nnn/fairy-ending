(() => {
  if (globalThis.FmsReconcile) return;

  const SOURCES = [
    {list: "/api/admin/transportation/trip/history/list", loading: "/api/admin/transportation/trip/history/loading/list",
     extra: {}, label: "Completed"},
    {list: "/api/admin/transportation/trip/list", loading: "/api/admin/transportation/trip/loading/list",
     extra: {query_type: 2}, label: "Pending Handover"},
  ];
  const TO_DETAIL = "/api/in-station/general_to/detail/search";
  const PAGE_SIZE = 50, MAX_PAGES = 200, DELAY = 400, SEQ = 1;

  class FmsError extends Error {}
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  const api = {
    calls: 0,
    async get(path, params) {
      const url = `${location.origin}${path}?${new URLSearchParams(params)}`;
      let last = "";
      for (let attempt = 0; attempt < 4; attempt++) {
        await sleep(attempt === 0 ? DELAY : 1000 * 2 ** attempt);
        let res;
        try {
          res = await fetch(url, {credentials: "include", headers: {Accept: "application/json"}});
        } catch (e) {
          last = e.name;
          continue;
        }
        this.calls++;
        if (res.status === 401 || res.status === 403) throw new FmsError("Hết phiên FMS. Đăng nhập lại rồi chạy lại.");
        if (res.status === 429 || res.status >= 500) {
          last = `HTTP ${res.status}`;
          continue;
        }
        let body;
        try {
          body = await res.json();
        } catch {
          throw new FmsError(`${path}: HTTP ${res.status}, không phải JSON`);
        }
        if (body.retcode !== 0) throw new FmsError(`${path}: retcode ${body.retcode} ${body.message || ""}`);
        return body.data || {};
      }
      throw new FmsError(`${path}: lỗi sau 4 lần thử (${last})`);
    },
  };

  async function fetchAll(api, path, params) {
    let first = null;
    const rows = [];
    for (let page = 1; page <= MAX_PAGES; page++) {
      const data = await api.get(path, {...params, pageno: page, count: PAGE_SIZE});
      first = first || data;
      const batch = data.list || [];
      rows.push(...batch);
      if (!batch.length || rows.length >= (data.total || 0)) break;
    }
    return {first: first || {}, rows, total: (first && first.total) || 0};
  }

  const isBulky = (row) => !!row.scan_number && row.scan_number !== row.to_number;

  async function collectTrip(api, lt, onProgress = () => {}) {
    let trip = null, source = null;
    for (const src of SOURCES) {
      const data = await api.get(src.list, {trip_number: lt, pageno: 1, count: 24, ...src.extra});
      trip = (data.list || []).find((t) => String(t.trip_number || "").toUpperCase() === lt);
      if (trip) {
        source = src;
        break;
      }
    }
    if (!trip) throw new FmsError(`Không tìm thấy trip ${lt}`);
    const codes = trip.next_station_code_list || [], names = trip.next_station_name_list || [];
    const stations = new Map((trip.next_station_list || []).map((id, i) => [id, [codes[i] || "", names[i] || ""]]));

    onProgress("Đang lấy danh sách TO");
    const {first: head, rows: toRows, total: totalTo} = await fetchAll(
      api, source.loading, {trip_id: trip.id, loaded_sequence_number: SEQ, type: "outbound"});
    const bags = toRows.filter((r) => !isBulky(r));
    const rows = [], flags = [], bagCount = new Map();
    let n = 0;
    for (const row of toRows) {
      const toNo = row.to_number || "", qty = row.to_parcel_quantity;
      const planned = row.unloaded_station, actual = row.actual_unloaded_station;
      const [code, name] = stations.get(planned) || ["", row.unloaded_station_name || ""];
      if (actual && actual !== planned) {
        flags.push([toNo, `Dỡ tại ${row.actual_unloaded_station_name}, kế hoạch ${name}`]);
      }
      if (isBulky(row)) {
        rows.push({tracking: row.scan_number, to: toNo, code, name, bulky: true});
        if (qty !== 1 || row.loading_type !== 2) flags.push([toNo, "Bulky nhưng parcel khác 1 hoặc loading_type khác 2"]);
        continue;
      }
      onProgress(`Bao ${++n}/${bags.length}: ${toNo}`);
      const {rows: orders} = await fetchAll(api, TO_DETAIL, {to_number: toNo});
      for (const o of orders) rows.push({tracking: o.fleet_order_id || "", to: toNo, code, name, bulky: false});
      bagCount.set(toNo, {qty, detail: orders.length});
    }

    const groups = new Map();
    rows.forEach((r, i) => r.tracking && groups.set(r.tracking, [...(groups.get(r.tracking) || []), i]));
    const over = (i) => !rows[i].bulky && bagCount.get(rows[i].to).detail > bagCount.get(rows[i].to).qty;
    const drop = new Set();
    for (const [tracking, idx] of groups) {
      if (idx.length < 2) continue;
      const keep = idx.find((i) => rows[i].bulky) ?? idx.find((i) => !over(i)) ?? idx[0];
      const sure = rows[keep].bulky || !over(keep);
      const dropped = idx.filter((i) => i !== keep);
      dropped.forEach((i) => drop.add(i));
      flags.push([tracking, `Trùng ${idx.length} lần: giữ ở ${rows[keep].to}${rows[keep].bulky ? " (bulky)" : ""}, ` +
        `bỏ khỏi ${dropped.map((i) => rows[i].to).join(", ")}${sure ? "" : ". Không chắc, kiểm tra tay"}`]);
    }
    for (const i of drop) if (bagCount.has(rows[i].to)) bagCount.get(rows[i].to).detail--;
    for (const [toNo, {qty, detail}] of bagCount) {
      if (detail !== qty) flags.push([toNo, `TO detail ra ${detail} đơn, loading list ghi ${qty}`]);
    }
    const out = rows.filter((_, i) => !drop.has(i)).map((r) => [r.tracking, r.to, r.code, r.name]);
    const blank = out.filter((r) => !r[0]).length;
    if (blank) flags.push(["(trống)", `${blank} dòng không có tracking`]);
    if (toRows.length !== totalTo) flags.push(["(trip)", `Lấy được ${toRows.length} TO, FMS ghi ${totalTo}`]);
    return {out, flags, totalParcel: head.total_parcel, bulky: toRows.length - bags.length, bags: bags.length,
      status: source.label, removed: drop.size};
  }

  function stamp() {
    const d = new Date(), p = (v) => String(v).padStart(2, "0");
    return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
  }

  function download(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 60000);
  }

  function panel() {
    document.getElementById("fms-reconcile-panel")?.remove();
    const host = document.createElement("div");
    host.id = "fms-reconcile-panel";
    host.style.cssText = "position:fixed;right:20px;bottom:20px;z-index:2147483647";
    const root = host.attachShadow({mode: "open"});
    root.innerHTML = `<style>
      :host { all: initial; }
      .box { width: 340px; max-height: 60vh; display: flex; flex-direction: column; background: #fff; color: #1b2733;
        border: 1px solid #c9d6e3; border-left: 4px solid #1f4e79; border-radius: 6px;
        box-shadow: 0 6px 24px rgba(16, 38, 61, .18); font: 13px/1.45 "Segoe UI", system-ui, sans-serif; }
      .top { display: flex; align-items: center; justify-content: space-between; padding: 10px 12px 6px; }
      h1 { margin: 0; font-size: 14px; font-weight: 600; color: #1f4e79; }
      button { font: inherit; border: 0; background: none; color: #4a5a6a; cursor: pointer; padding: 2px 6px; border-radius: 4px; }
      button:hover { background: #eef3f8; } button:focus-visible { outline: 2px solid #1f4e79; }
      .body { overflow: auto; padding: 0 12px 12px; }
      .trip { padding: 8px 0; border-top: 1px solid #e3eaf1; }
      .lt { font-weight: 600; font-variant-numeric: tabular-nums; }
      .msg { color: #4a5a6a; } .ok { color: #1e7a3c; font-weight: 600; } .bad { color: #b42318; font-weight: 600; }
      ul { margin: 4px 0 0; padding-left: 18px; color: #b42318; } li { margin: 2px 0; }
    </style>
    <div class="box" role="status" aria-live="polite"><div class="top"><h1>FMS Reconcile</h1>
    <button type="button" aria-label="Đóng">Đóng</button></div><div class="body"></div></div>`;
    root.querySelector("button").onclick = () => host.remove();
    document.body.appendChild(host);
    const body = root.querySelector(".body");
    return {
      trip(lt) {
        const el = document.createElement("div");
        el.className = "trip";
        el.innerHTML = '<div class="lt"></div><div class="msg"></div>';
        el.querySelector(".lt").textContent = lt;
        body.appendChild(el);
        const msg = el.querySelector(".msg");
        return {
          set(text, cls = "msg") { msg.className = cls; msg.textContent = text; },
          flags(list) {
            if (!list.length) return;
            const ul = document.createElement("ul");
            for (const [code, issue] of list) {
              const li = document.createElement("li");
              li.textContent = `${code}: ${issue}`;
              ul.appendChild(li);
            }
            el.appendChild(ul);
          },
        };
      },
    };
  }

  let running = false;

  async function run(lts) {
    if (running) return "busy";
    running = true;
    const ui = panel();
    try {
      for (const lt of lts) {
        const line = ui.trip(lt);
        api.calls = 0;
        try {
          line.set("Đang tìm trip");
          const r = await collectTrip(api, lt, (t) => line.set(t));
          const filename = `${lt}_${stamp()}.xlsx`;
          const blob = await globalThis.FmsXlsx.build(
            "Chi tiết", ["Tracking Number", "TO Number", "Dest Code", "Dest Name"], r.out, [24, 22, 14, 28]);
          download(blob, filename);
          const ok = r.out.length === r.totalParcel;
          line.set(`Trip ${r.status}. ${r.out.length}/${r.totalParcel} đơn, ${ok ? "khớp" : "lệch"}` +
            `${r.removed ? `, đã bỏ ${r.removed} dòng trùng` : ""}. ${r.bulky} bulky, ${r.bags} bao. Đã tải ${filename}`,
            ok ? "ok" : "bad");
          line.flags(r.flags);
        } catch (e) {
          line.set(e instanceof FmsError ? e.message : `Lỗi: ${e.name}`, "bad");
        }
      }
    } finally {
      running = false;
    }
    return "done";
  }

  globalThis.FmsReconcile = {run, collectTrip, FmsError, isRunning: () => running};
})();