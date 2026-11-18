// Swiggy / Instamart Order Exporter - runs inside the site tab (injected by the popup).
// Pages through the site's own order-history API (newest first), keeps orders inside
// the chosen date range, and downloads an .xlsx with a Summary tab + one tab per month.

(function () {
  'use strict';

  const PAGE_DELAY_MS = 350; // be gentle with the APIs
  const DETAIL_DELAY_MS = 250;
  const MAX_PAGES = 2000; // hard safety stop
  const IST_OFFSET_MS = 5.5 * 3600 * 1000;

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const num = (v) => {
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  };
  const round2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
  const moneyText = (t) => {
    const n = parseFloat(String(t || '').replace(/[^0-9.\-]/g, ''));
    return Number.isFinite(n) ? n : 0;
  };

  // epoch ms -> { ymd, hms } in India time
  function istParts(ms) {
    const d = new Date(ms + IST_OFFSET_MS);
    const p = (n) => String(n).padStart(2, '0');
    return {
      ymd: `${d.getUTCFullYear()}-${p(d.getUTCMonth() + 1)}-${p(d.getUTCDate())}`,
      hms: `${p(d.getUTCHours())}:${p(d.getUTCMinutes())}:${p(d.getUTCSeconds())}`,
    };
  }
  function excelDate(ymd) {
    const [y, m, d] = ymd.split('-').map(Number);
    return Math.round(Date.UTC(y, m - 1, d) / 86400000) + 25569;
  }
  function excelTime(hms) {
    const [h, mi, s] = hms.split(':').map(Number);
    return (h * 3600 + mi * 60 + (s || 0)) / 86400;
  }

  async function getJson(url) {
    let lastErr;
    for (let attempt = 1; attempt <= 4; attempt++) {
      try {
        const res = await fetch(url, { credentials: 'include' });
        if (res.status === 401 || res.status === 403) throw new Error('Not logged in (or session expired). Log in and retry.');
        if (!res.ok) throw new Error('HTTP ' + res.status);
        const json = await res.json();
        if (json.statusCode !== 0) throw new Error(json.statusMessage || 'Server returned status ' + json.statusCode);
        return json.data || {};
      } catch (e) {
        lastErr = e;
        if (/Not logged in/.test(e.message)) throw e;
        await sleep(800 * attempt);
      }
    }
    throw lastErr;
  }

  // =========================================================================
  // Source: Swiggy Food  (www.swiggy.com)
  // =========================================================================
  const swiggy = {
    label: 'Swiggy',
    filePrefix: 'swiggy-orders',
    firstCursor: () => '',
    async fetchPage(cursor) {
      const data = await getJson(`/dapi/order/all?order_id=${cursor}`);
      const orders = data.orders || [];
      return { orders, next: orders.length ? orders[orders.length - 1].order_id : cursor };
    },
    dayOf: (o) => String(o.order_time || '').slice(0, 10),
    idOf: (o) => o.order_id,
    isDelivered: (o) => String(o.order_status).toLowerCase() === 'delivered',
    toRow(o) {
      const [ymd, hms = '00:00:00'] = String(o.order_time || '').split(' ');
      const items = (o.order_items || []).map((i) => `${i.name} x${i.quantity}`);
      const itemCount = (o.order_items || []).reduce((s, i) => s + num(i.quantity), 0);
      const tip = num(o.order_total_with_tip) - num(o.order_total);
      return {
        month: ymd.slice(0, 7), ymd, hms,
        orderId: String(o.order_id),
        restaurant: o.restaurant_name || '',
        area: o.restaurant_area_name || '',
        city: o.restaurant_city_name || '',
        items: items.join('; '),
        itemCount,
        itemTotal: num(o.item_total),
        discount: num(o.order_discount),
        coupon: o.coupon_code || '',
        delivery: num(o.order_delivery_charge),
        packing: num(o.restaurant_packing_charges),
        platformFee: num(o.convenience_fees),
        taxes: num(o.total_tax_including_platform_fee || o.total_tax),
        tip: round2(Math.max(tip, 0)),
        paid: num(o.order_total_with_tip || o.order_total),
        payment: o.payment_method || o.order_payment_method || '',
        status: o.order_status || '',
      };
    },
    columns: [
      ['Order ID', 'orderId', 18, 's'], ['Date', 'ymd', 12, 'date'], ['Time', 'hms', 9, 'time'],
      ['Restaurant', 'restaurant', 28, 's'], ['Area', 'area', 16, 's'], ['City', 'city', 12, 's'],
      ['Items', 'items', 60, 's'], ['Qty', 'itemCount', 6, 'n'],
      ['Item total', 'itemTotal', 11, 'money'], ['Discount', 'discount', 10, 'money'], ['Coupon', 'coupon', 14, 's'],
      ['Delivery fee', 'delivery', 12, 'money'], ['Packing', 'packing', 9, 'money'], ['Platform fee', 'platformFee', 12, 'money'],
      ['Taxes', 'taxes', 9, 'money'], ['Tip', 'tip', 8, 'money'], ['Amount paid', 'paid', 12, 'money'],
      ['Payment', 'payment', 12, 's'], ['Status', 'status', 12, 's'],
    ],
    sum: [['Items', 'itemCount'], ['Item total', 'itemTotal'], ['Discount', 'discount'], ['Delivery fee', 'delivery'],
      ['Packing', 'packing'], ['Platform fee', 'platformFee'], ['Taxes', 'taxes'], ['Tip', 'tip'], ['Amount paid', 'paid']],
  };

  // =========================================================================
  // Source: Instamart  (instamart.in)
  // =========================================================================
  function instamartBill(detail) {
    const out = { itemBill: 0, delivery: 0, handling: 0, otherFees: 0, discount: 0, items: null };
    const lines = (detail && detail.billDetail && detail.billDetail.billLineItems) || [];
    for (const l of lines) {
      const title = String((l.title && l.title.title) || '').trim().toLowerCase();
      const amt = moneyText(l.amount && l.amount.text);
      if (/item bill|item total/.test(title)) out.itemBill += amt;
      else if (/delivery/.test(title)) out.delivery += amt;
      else if (/handling/.test(title)) out.handling += amt;
      else if (/discount|saving|coupon/.test(title) || amt < 0) out.discount += Math.abs(amt);
      else out.otherFees += amt;
    }
    const ship = (detail && detail.shipmentDetails) || [];
    const names = [];
    ship.forEach((s) => (s.items || []).forEach((i) => { if (!i.removed) names.push(`${i.description} x${i.quantity}`); }));
    if (names.length) out.items = names;
    return out;
  }

  const instamart = {
    label: 'Instamart',
    filePrefix: 'instamart-orders',
    hasDetail: true,
    firstCursor: () => Date.now(),
    async fetchPage(cursor) {
      const data = await getJson(`/api/instamart/orders?order_type=DASH&count=20&from_time=${cursor}`);
      const orders = data.orders || [];
      return { orders, next: orders.length ? orders[orders.length - 1].created_at : cursor };
    },
    dayOf: (o) => istParts(num(o.created_at)).ymd,
    idOf: (o) => o.order_id,
    isDelivered: (o) => String(o.history_status).toUpperCase() === 'DELIVERED',
    fetchDetail: (o) => getJson(`/api/order/v3/getOrderDetails/${o.order_id}`),
    toRow(o, detail) {
      const { ymd, hms } = istParts(num(o.created_at));
      const v2 = o.order_data_v2 || {};
      const listItems = [];
      (v2.shipments || []).forEach((s) => (s.items || []).forEach((i) => listItems.push(i)));
      const bill = instamartBill(detail);
      const itemList = bill.items || listItems.map((i) => `${i.name} x${i.quantity}`);
      const itemCount = listItems.reduce((s, i) => s + num(i.quantity), 0);
      const paid = v2.total ? num(v2.total.units) + num(v2.total.nanos) / 1e9 : num(o.order_data && o.order_data.payment_info_v2 && o.order_data.payment_info_v2.amount_collected);
      const ptx = o.order_data && o.order_data.payment_info_v2 && o.order_data.payment_info_v2.payment_transaction_reference;
      return {
        month: ymd.slice(0, 7), ymd, hms,
        orderId: String(o.order_id),
        area: v2.description || '',
        items: itemList.join('; '),
        itemCount,
        itemBill: bill.itemBill,
        delivery: bill.delivery,
        handling: bill.handling,
        otherFees: round2(bill.otherFees),
        discount: bill.discount,
        paid: round2(paid),
        payment: (ptx && ptx[0] && ptx[0].primary_payment_method) || '',
        status: o.history_status || '',
        _detail: !!detail,
      };
    },
    columns: [
      ['Order ID', 'orderId', 18, 's'], ['Date', 'ymd', 12, 'date'], ['Time', 'hms', 9, 'time'],
      ['Area', 'area', 18, 's'], ['Items', 'items', 70, 's'], ['Qty', 'itemCount', 6, 'n'],
      ['Item bill', 'itemBill', 11, 'money'], ['Delivery fee', 'delivery', 12, 'money'], ['Handling fee', 'handling', 12, 'money'],
      ['Other fees', 'otherFees', 11, 'money'], ['Discount', 'discount', 10, 'money'], ['Amount paid', 'paid', 12, 'money'],
      ['Payment', 'payment', 12, 's'], ['Status', 'status', 12, 's'],
    ],
    sum: [['Items', 'itemCount'], ['Item bill', 'itemBill'], ['Delivery fee', 'delivery'], ['Handling fee', 'handling'],
      ['Other fees', 'otherFees'], ['Discount', 'discount'], ['Amount paid', 'paid']],
  };

  const SOURCES = { swiggy, instamart };

  // =========================================================================
  // Workbook
  // =========================================================================
  function cell(kind, row, key) {
    const v = row[key];
    if (kind === 'date') return { t: 'n', v: excelDate(v), z: 'yyyy-mm-dd' };
    if (kind === 'time') return { t: 'n', v: excelTime(v), z: 'hh:mm' };
    if (kind === 'money') return { t: 'n', v: num(v), z: '#,##0.00' };
    if (kind === 'n') return { t: 'n', v: num(v) };
    return { t: 's', v: String(v == null ? '' : v) };
  }

  function monthsBetween(startYmd, endYmd) {
    const out = [];
    let [y, m] = startYmd.slice(0, 7).split('-').map(Number);
    const [ey, em] = endYmd.slice(0, 7).split('-').map(Number);
    while (y < ey || (y === ey && m <= em)) {
      out.push(`${y}-${String(m).padStart(2, '0')}`);
      m++;
      if (m > 12) { m = 1; y++; }
    }
    return out;
  }

  function buildWorkbook(XLSX, rows, startYmd, endYmd, schema) {
    const wb = XLSX.utils.book_new();
    const byMonth = {};
    rows.forEach((r) => (byMonth[r.month] = byMonth[r.month] || []).push(r));

    const money = (v) => ({ t: 'n', v: round2(v), z: '#,##0.00' });
    const header = ['Month', 'Orders', ...schema.sum.map(([h]) => h), 'Avg order'];
    const totals = { orders: 0 };
    schema.sum.forEach(([, k]) => (totals[k] = 0));
    const line = (label, s) => [
      { t: 's', v: label },
      { t: 'n', v: s.orders },
      ...schema.sum.map(([, k]) => (k === 'itemCount' ? { t: 'n', v: s[k] } : money(s[k]))),
      money(s.orders ? s.paid / s.orders : 0),
    ];
    const aoa = [header.map((h) => ({ t: 's', v: h }))];
    for (const mk of monthsBetween(startYmd, endYmd)) {
      const list = byMonth[mk] || [];
      const s = { orders: list.length };
      schema.sum.forEach(([, k]) => (s[k] = list.reduce((a, r) => a + num(r[k]), 0)));
      Object.keys(totals).forEach((k) => (totals[k] += s[k]));
      aoa.push(line(mk, s));
    }
    aoa.push(line('TOTAL', totals));
    const wsSum = XLSX.utils.aoa_to_sheet(aoa);
    wsSum['!cols'] = header.map((h, i) => ({ wch: i === 0 ? 12 : 13 }));
    XLSX.utils.book_append_sheet(wb, wsSum, 'Summary');

    Object.keys(byMonth).sort().forEach((mk) => {
      const list = byMonth[mk].slice().sort((a, b) => (b.ymd + b.hms).localeCompare(a.ymd + a.hms));
      const sheet = [schema.columns.map(([h]) => ({ t: 's', v: h }))];
      list.forEach((r) => sheet.push(schema.columns.map(([, key, , kind]) => cell(kind, r, key))));
      const ws = XLSX.utils.aoa_to_sheet(sheet);
      ws['!cols'] = schema.columns.map(([, , w]) => ({ wch: w }));
      ws['!autofilter'] = { ref: ws['!ref'] };
      XLSX.utils.book_append_sheet(wb, ws, mk);
    });
    return wb;
  }

  // =========================================================================
  // Progress banner
  // =========================================================================
  function banner() {
    let el = document.getElementById('__order_export_banner');
    if (!el) {
      el = document.createElement('div');
      el.id = '__order_export_banner';
      el.style.cssText =
        'position:fixed;z-index:2147483647;right:16px;bottom:16px;max-width:340px;padding:12px 14px;' +
        'background:#1b1b1b;color:#fff;font:13px/1.4 -apple-system,Segoe UI,Roboto,sans-serif;' +
        'border-radius:10px;box-shadow:0 6px 24px rgba(0,0,0,.3)';
      el.innerHTML = '<div id="__oxb_msg"></div><button id="__oxb_btn" style="margin-top:8px;padding:4px 10px;border:0;border-radius:6px;background:#fc8019;color:#fff;cursor:pointer">Cancel</button>';
      document.body.appendChild(el);
    }
    const msg = el.querySelector('#__oxb_msg');
    const btn = el.querySelector('#__oxb_btn');
    btn.onclick = null;
    btn.textContent = 'Cancel';
    return {
      set: (t) => (msg.textContent = t),
      onCancel: (fn) => (btn.onclick = fn),
      finish: (t) => {
        msg.textContent = t;
        btn.textContent = 'Close';
        btn.onclick = () => el.remove();
      },
    };
  }

  // =========================================================================
  // Main
  // =========================================================================
  async function runExport(opts) {
    const { start, end, deliveredOnly } = opts;
    const src = SOURCES[opts.source];
    const ui = banner();
    let cancelled = false;
    ui.onCancel(() => (cancelled = true));
    try {
      if (!src) throw new Error('Unknown source: ' + opts.source);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) throw new Error('Pick a valid start and end date.');
      if (start > end) throw new Error('Start date is after end date.');

      const matched = [];
      const seen = new Set();
      let cursor = src.firstCursor();
      let pages = 0, scanned = 0, skippedStatus = 0;

      while (!cancelled && pages < MAX_PAGES) {
        const { orders, next } = await src.fetchPage(cursor);
        pages++;
        if (!orders.length) break;

        for (const o of orders) {
          scanned++;
          const day = src.dayOf(o);
          if (!day || day > end || day < start) continue;
          const id = src.idOf(o);
          if (seen.has(id)) continue;
          if (deliveredOnly && !src.isDelivered(o)) { skippedStatus++; continue; }
          seen.add(id);
          matched.push(o);
        }

        const lastDay = src.dayOf(orders[orders.length - 1]);
        ui.set(`${src.label} export: scanned ${scanned} orders (back to ${lastDay}), ${matched.length} match so far...`);
        if (lastDay && lastDay < start) break; // passed the start date; older orders can't match
        if (String(next) === String(cursor)) break;
        cursor = next;
        await sleep(PAGE_DELAY_MS);
      }

      if (cancelled) { ui.finish('Export cancelled.'); return { cancelled: true }; }
      if (!matched.length) { ui.finish(`No ${src.label} orders found between ${start} and ${end}` + (deliveredOnly ? ' (delivered only).' : '.')); return { rows: 0 }; }

      // Instamart: per-order bill details (fees, discounts, item prices)
      const details = new Map();
      let detailFailures = 0;
      if (src.hasDetail && opts.withDetails !== false) {
        for (let i = 0; i < matched.length && !cancelled; i++) {
          ui.set(`${src.label} export: reading bill details ${i + 1}/${matched.length}...`);
          try { details.set(src.idOf(matched[i]), await src.fetchDetail(matched[i])); } catch (e) { detailFailures++; }
          await sleep(DETAIL_DELAY_MS);
        }
        if (cancelled) { ui.finish('Export cancelled.'); return { cancelled: true }; }
      }

      const rows = matched.map((o) => src.toRow(o, details.get(src.idOf(o))));
      const wb = buildWorkbook(XLSX, rows, start, end, src);
      XLSX.writeFile(wb, `${src.filePrefix}_${start}_to_${end}.xlsx`);
      const total = rows.reduce((s, r) => s + r.paid, 0);
      ui.finish(
        `Done: ${rows.length} orders, Rs ${round2(total).toLocaleString('en-IN')} paid. File downloaded` +
        (skippedStatus ? ` (${skippedStatus} non-delivered skipped)` : '') +
        (detailFailures ? ` (${detailFailures} orders missing bill details)` : '') + '.'
      );
      return { rows: rows.length };
    } catch (e) {
      ui.finish('Export failed: ' + e.message);
      return { error: e.message };
    }
  }

  const api = { SOURCES, istParts, instamartBill, monthsBetween, buildWorkbook, runExport };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else {
    window.__orderExport = runExport;
    window.__orderExporter = api;
  }
})();
