(() => {
  const $ = (sel) => document.querySelector(sel);
  const form = $('#form');
  const preview = $('#preview');
  const urlInput = $('#imgurl');
  const list = $('#events');
  const status = $('#status');
  const timeFormat = list.dataset.timeFormat === '24' ? 24 : 12;
  let events = [];
  let editingId = null;
  let view = 'split';

  // ---------- preview + image address
  function imageUrl() {
    return `${location.origin}/calendar.png?view=${view}`;
  }
  function refreshPreview() {
    preview.src = `/calendar.png?view=${view}&t=${Date.now()}`;
    urlInput.value = imageUrl();
  }
  document.querySelectorAll('input[name=view]').forEach((r) =>
    r.addEventListener('change', () => { view = r.value; refreshPreview(); }));
  $('#copy').addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(urlInput.value);
      say('Address copied.');
    } catch {
      urlInput.select();
      say('Press Ctrl+C (or Cmd+C) to copy the selected address.');
    }
  });

  // ---------- helpers
  function say(msg, bad = false) {
    status.textContent = msg;
    status.className = bad ? 'bad' : '';
  }
  async function api(path, options = {}) {
    const res = await fetch(path, {
      headers: { 'Content-Type': 'application/json' },
      ...options,
    });
    const data = res.status === 204 ? null : await res.json().catch(() => null);
    return { ok: res.ok, status: res.status, data };
  }
  const parseDay = (iso) => {
    const [y, m, d] = iso.split('-').map(Number);
    return new Date(y, m - 1, d);
  };
  const fmtDay = (iso) => parseDay(iso).toLocaleDateString(undefined, { weekday: 'short', day: 'numeric', month: 'short' });
  const fmtMonth = (iso) => parseDay(iso).toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
  function fmtTime(hhmm) {
    const [h, m] = hhmm.split(':').map(Number);
    if (timeFormat === 24) return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    return `${h % 12 || 12}:${String(m).padStart(2, '0')} ${h < 12 ? 'AM' : 'PM'}`;
  }
  function todayIso() {
    const n = new Date();
    return `${n.getFullYear()}-${String(n.getMonth() + 1).padStart(2, '0')}-${String(n.getDate()).padStart(2, '0')}`;
  }
  function isPast(ev) {
    const last = ev.repeat !== 'none' ? ev.repeat_until : (ev.end_date || ev.date);
    return last ? last < todayIso() : false;
  }
  const repeatText = { daily: 'every day', weekly: 'every week', monthly: 'every month', yearly: 'every year' };

  // ---------- list (built with textContent so titles can never inject HTML)
  function el(tag, cls, text) {
    const node = document.createElement(tag);
    if (cls) node.className = cls;
    if (text !== undefined) node.textContent = text;
    return node;
  }

  function row(ev) {
    const r = el('div', 'ev' + (isPast(ev) ? ' past' : ''));
    r.append(el('div', 'when', fmtDay(ev.date)));

    const what = el('div', 'what');
    what.append(document.createTextNode(ev.title));
    const bits = [];
    if (ev.end_date) bits.push(`to ${fmtDay(ev.end_date)}`);
    if (ev.start_time) bits.push(fmtTime(ev.start_time) + (ev.end_time ? `–${fmtTime(ev.end_time)}` : ''));
    else bits.push('all day');
    if (ev.location) bits.push(ev.location);
    if (ev.repeat !== 'none') {
      bits.push(`repeats ${repeatText[ev.repeat]}` + (ev.repeat_until ? ` until ${fmtDay(ev.repeat_until)}` : ''));
    }
    what.append(el('span', 'meta', bits.join(', ')));
    r.append(what);

    const ops = el('div', 'ops');
    const edit = el('button', 'link', 'Edit');
    edit.type = 'button';
    edit.setAttribute('aria-label', `Edit ${ev.title}`);
    edit.addEventListener('click', () => startEdit(ev));
    const del = el('button', 'link danger', 'Delete');
    del.type = 'button';
    del.setAttribute('aria-label', `Delete ${ev.title}`);
    del.addEventListener('click', () => remove(ev));
    ops.append(edit, del);
    r.append(ops);
    return r;
  }

  function group(title, items) {
    const g = el('div', 'group');
    g.append(el('h3', '', title));
    items.forEach((ev) => g.append(row(ev)));
    return g;
  }

  function renderList() {
    const showPast = $('#showPast').checked;
    const shown = events.filter((ev) => showPast || !isPast(ev));
    list.replaceChildren();
    if (!shown.length) {
      list.append(el('p', 'empty', events.length
        ? 'No upcoming events. Tick “Show past events” to see older ones.'
        : 'No events yet. Add your first one using the form.'));
      return;
    }
    const repeating = shown.filter((ev) => ev.repeat !== 'none');
    if (repeating.length) list.append(group('Repeating', repeating));
    const months = new Map();
    shown.filter((ev) => ev.repeat === 'none').forEach((ev) => {
      const key = fmtMonth(ev.date);
      if (!months.has(key)) months.set(key, []);
      months.get(key).push(ev);
    });
    months.forEach((items, key) => list.append(group(key, items)));
  }
  $('#showPast').addEventListener('change', renderList);

  async function load() {
    const res = await api('/api/events');
    if (res.ok) { events = res.data; renderList(); }
    else say('Could not load events.', true);
  }

  // ---------- form
  const fields = ['title', 'date', 'end_date', 'start_time', 'end_time', 'location', 'repeat', 'repeat_until'];

  function syncForm() {
    const allDay = $('#all_day').checked;
    $('#times').classList.toggle('off', allDay);
    $('#start_time').disabled = allDay;
    $('#end_time').disabled = allDay;
    $('#untilField').hidden = $('#repeat').value === 'none';
  }
  $('#all_day').addEventListener('change', syncForm);
  $('#repeat').addEventListener('change', syncForm);

  function clearErrors() {
    form.querySelectorAll('.err').forEach((p) => { p.textContent = ''; });
    form.querySelectorAll('[aria-invalid]').forEach((i) => i.removeAttribute('aria-invalid'));
  }
  function showErrors(errors) {
    let first = null;
    Object.entries(errors).forEach(([name, msg]) => {
      const p = form.querySelector(`.err[data-for="${name}"]`);
      const input = form.elements[name];
      if (p) p.textContent = msg;
      if (input) { input.setAttribute('aria-invalid', 'true'); first = first || input; }
    });
    if (first) first.focus();
  }

  function resetForm() {
    editingId = null;
    form.reset();
    clearErrors();
    $('#form-title').textContent = 'Add an event';
    $('#save').textContent = 'Save event';
    $('#cancel').hidden = true;
    syncForm();
  }

  function startEdit(ev) {
    editingId = ev.id;
    clearErrors();
    fields.forEach((f) => { form.elements[f].value = ev[f] || (f === 'repeat' ? 'none' : ''); });
    $('#all_day').checked = !ev.start_time;
    $('#form-title').textContent = 'Edit event';
    $('#save').textContent = 'Update event';
    $('#cancel').hidden = false;
    syncForm();
    say('');
    form.scrollIntoView({ behavior: 'smooth', block: 'center' });
    form.elements.title.focus();
  }
  $('#cancel').addEventListener('click', () => { resetForm(); say(''); });

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    clearErrors();
    const body = {};
    fields.forEach((f) => { body[f] = form.elements[f].value; });
    body.all_day = $('#all_day').checked;
    const wasEditing = editingId !== null;
    const res = await api(wasEditing ? `/api/events/${editingId}` : '/api/events', {
      method: wasEditing ? 'PUT' : 'POST',
      body: JSON.stringify(body),
    });
    if (res.status === 400 && res.data && res.data.errors) {
      showErrors(res.data.errors);
      say('Fix the highlighted fields and save again.', true);
      return;
    }
    if (!res.ok) {
      say((res.data && res.data.error) || 'Something went wrong. Try again.', true);
      return;
    }
    resetForm();
    say(wasEditing ? 'Event updated.' : 'Event saved.');
    await load();
    refreshPreview();
  });

  async function remove(ev) {
    if (!window.confirm(`Delete “${ev.title}”? This removes every repeat of it.`)) return;
    const res = await api(`/api/events/${ev.id}`, { method: 'DELETE' });
    if (!res.ok) { say('Could not delete that event.', true); return; }
    if (editingId === ev.id) resetForm();
    say('Event deleted.');
    await load();
    refreshPreview();
  }

  // ---------- start
  urlInput.value = imageUrl();
  syncForm();
  load();
})();
