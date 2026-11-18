const $ = (id) => document.getElementById(id);
const ymd = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

function setRange(kind) {
  const now = new Date();
  let s, e;
  if (kind === 'last') {
    s = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    e = new Date(now.getFullYear(), now.getMonth(), 0);
  } else if (kind === 'this') {
    s = new Date(now.getFullYear(), now.getMonth(), 1);
    e = now;
  } else {
    s = new Date(now.getFullYear(), now.getMonth() - 3, 1);
    e = new Date(now.getFullYear(), now.getMonth(), 0);
  }
  $('start').value = ymd(s);
  $('end').value = ymd(e);
}

function status(msg, isErr) {
  $('status').textContent = msg;
  $('status').className = isErr ? 'err' : '';
}

function sourceFor(url) {
  const host = (() => { try { return new URL(url).hostname; } catch (e) { return ''; } })();
  if (host === 'www.swiggy.com' || host === 'swiggy.com') return 'swiggy';
  if (host === 'instamart.in' || host.endsWith('.instamart.in')) return 'instamart';
  return null;
}

document.querySelectorAll('.quick button').forEach((b) => b.addEventListener('click', () => setRange(b.dataset.q)));
setRange('last'); // default: previous full month (month-end expense review)

(async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  const source = tab && sourceFor(tab.url || '');
  $('site').textContent = source === 'swiggy' ? 'Detected: Swiggy Food' : source === 'instamart' ? 'Detected: Instamart' : 'Open swiggy.com or instamart.in first';
  $('detailRow').style.display = source === 'instamart' ? 'flex' : 'none';
})();

$('go').addEventListener('click', async () => {
  const opts = { start: $('start').value, end: $('end').value, deliveredOnly: $('delivered').checked, withDetails: $('details').checked };
  if (!opts.start || !opts.end) return status('Pick both dates.', true);
  if (opts.start > opts.end) return status('Start date is after end date.', true);

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  opts.source = tab && sourceFor(tab.url || '');
  if (!opts.source) return status('Open swiggy.com or instamart.in (logged in) in this tab first.', true);

  try {
    await chrome.scripting.executeScript({ target: { tabId: tab.id }, files: ['lib/xlsx.full.min.js', 'exporter.js'] });
    // Not awaited on purpose: progress shows on the page, so the popup can close.
    chrome.scripting.executeScript({ target: { tabId: tab.id }, func: (o) => window.__orderExport(o), args: [opts] });
    status('Started. Watch the progress box on the page.');
    setTimeout(() => window.close(), 1200);
  } catch (e) {
    status('Could not start: ' + e.message, true);
  }
});
