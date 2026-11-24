(() => {
  const $ = (id) => document.getElementById(id);
  const form = $('form'), input = $('code'), go = $('go'), err = $('err');
  const scan = $('scan'), result = $('result');
  const RESET_SECONDS = 45;
  let resetTimer = null, countdown = null;

  const fmtMoney = (n) => 'RD$ ' + Number(n || 0).toLocaleString('es-DO', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmtDays = (n) => {
    const v = Number(n || 0);
    const s = Number.isInteger(v) ? String(v) : v.toLocaleString('es-DO', { maximumFractionDigits: 2 });
    return s + (Math.abs(v) === 1 ? ' día' : ' días');
  };
  const fmtDate = (iso) => {
    if (!iso) return '—';
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return iso;
    const onlyDate = /^\d{4}-\d{2}-\d{2}$/.test(iso);
    return onlyDate
      ? d.toLocaleDateString('es-DO', { day: '2-digit', month: 'short', year: 'numeric', timeZone: 'UTC' })
      : d.toLocaleString('es-DO', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
  };

  // ---- estado de los datos (pie y pastilla superior) ----
  function paintDatos(d) {
    if (!d) return;
    const st = $('status');
    if (d.online === true) { st.textContent = 'SharePoint en vivo'; st.className = 'status online'; }
    else if (d.online === false) { st.textContent = 'Sin conexión · copia guardada'; st.className = 'status offline'; }
    else if (/power automate/i.test(d.via || '')) { st.textContent = 'Sincronizado desde SharePoint'; st.className = 'status online'; }
    else { st.textContent = d.via ? `Datos: ${d.via}` : 'Sin datos'; st.className = 'status'; }
    const etiqueta = /flujo/i.test(d.via || '') ? 'Último cambio del Excel detectado' : 'Excel actualizado por RRHH';
    $('datos').textContent = d.fileModified
      ? `${etiqueta}: ${fmtDate(d.fileModified)}${d.online === false ? ' (última copia disponible)' : ''}`
      : 'Sin datos cargados';
  }
  async function loadStatus() {
    try { const r = await fetch('/api/status'); paintDatos(await r.json()); } catch { /* sin red al servidor */ }
  }

  // ---- foco permanente para el lector de codigos ----
  function focusInput() { if (!result.hidden) return; input.focus(); input.select(); }
  document.addEventListener('click', (e) => { if (!e.target.closest('a,button,table,.result')) focusInput(); });
  window.addEventListener('focus', focusInput);
  // si el lector "escribe" mientras se muestra un resultado, empezamos una consulta nueva con lo tecleado
  document.addEventListener('keydown', (e) => {
    if (result.hidden || e.ctrlKey || e.metaKey || e.altKey) return;
    if (e.key.length === 1 && /[a-z0-9]/i.test(e.key)) { reset(); input.value = ''; }
  });

  function showError(msg) { err.textContent = msg; err.hidden = false; }
  function clearError() { err.hidden = true; err.textContent = ''; }

  function badgeClass(estado) {
    const s = String(estado || '').toLowerCase();
    if (s.includes('saldad')) return 'ok';
    if (s.includes('pendiente')) return 'warn';
    return 'bad';
  }

  function render(e) {
    $('r-nombre').textContent = e.nombre || e.codigo;
    $('r-codigo').textContent = e.codigo;
    $('r-depto').textContent = [e.departamento, e.tipo].filter(Boolean).join(' · ');
    const b = $('r-estado'); b.textContent = e.estado; b.className = 'badge ' + badgeClass(e.estado);

    $('r-dias-ini').textContent = fmtDays(e.diasAdeudados);
    $('r-deuda-ini').textContent = fmtMoney(e.deudaInicial);
    $('r-dias-pend').textContent = fmtDays(e.diasPendientes);
    $('r-deuda-rest').textContent = fmtMoney(e.deudaRestante);
    $('r-dias-pag').textContent = fmtDays(e.diasPagados);
    $('r-deuda-pag').textContent = fmtMoney(e.deudaPagada);
    const pct = Number(e.porcentajePagado || 0);
    $('r-bar').style.width = pct + '%';
    $('r-pct').textContent = e.deudaInicial > 0 ? `${pct.toLocaleString('es-DO')}% de la deuda pagada` : 'Sin deuda registrada';

    const movs = e.movimientos || [];
    const tb = $('r-movs'); tb.innerHTML = '';
    if (movs.length) {
      let total = 0;
      for (const m of movs) {
        total += Number(m.monto || 0);
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${fmtDate(m.fecha)}</td><td>${m.horas ?? '—'}</td><td>${m.tipoHora ?? '—'}</td><td class="num">${fmtMoney(m.monto)}</td>`;
        tb.appendChild(tr);
      }
      const tr = document.createElement('tr');
      tr.innerHTML = `<td colspan="3"><b>Total aplicado</b></td><td class="num"><b>${fmtMoney(total)}</b></td>`;
      tb.appendChild(tr);
    }
    $('movs').hidden = !movs.length;

    scan.hidden = true; result.hidden = false;
    startCountdown();
  }

  function startCountdown() {
    let left = RESET_SECONDS;
    const auto = $('auto');
    auto.textContent = `La pantalla se limpia en ${left} s`;
    clearInterval(countdown); clearTimeout(resetTimer);
    countdown = setInterval(() => { left--; auto.textContent = `La pantalla se limpia en ${left} s`; }, 1000);
    resetTimer = setTimeout(reset, RESET_SECONDS * 1000);
  }
  function reset() {
    clearInterval(countdown); clearTimeout(resetTimer);
    result.hidden = true; scan.hidden = false;
    input.value = ''; clearError(); focusInput();
  }
  $('again').addEventListener('click', reset);

  form.addEventListener('submit', async (ev) => {
    ev.preventDefault();
    const code = input.value.trim();
    if (!code) return;
    clearError(); go.disabled = true;
    try {
      const r = await fetch('/api/empleado/' + encodeURIComponent(code));
      const j = await r.json();
      paintDatos(j.datos);
      if (!r.ok) {
        showError(j.error || 'No se pudo consultar');
        input.select();
        return;
      }
      render(j);
    } catch {
      showError('No hay conexión con el servidor. Inténtalo de nuevo.');
    } finally { go.disabled = false; }
  });

  loadStatus();
  setInterval(loadStatus, 60_000);
  focusInput();
})();
