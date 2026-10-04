/* Lectores locales de PDF y XLSX. No envían archivos a servicios externos. */
(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const state = { doc: 'informe', pdf: null, page: 1, task: null, version: 0, sheet: 0, book: null };
  const bytes = base64 => Uint8Array.from(atob(base64), c => c.charCodeAt(0));
  const message = text => { $('viewer-status').textContent = text; };

  async function openDocument(key) {
    state.doc = key;
    const version = ++state.version;
    if (state.task) { state.task.cancel(); try { await state.task.promise; } catch (_) {} state.task = null; }
    document.querySelectorAll('[data-doc]').forEach(button => {
      const active = button.dataset.doc === key;
      button.setAttribute('aria-selected', active);
      button.tabIndex = active ? 0 : -1;
    });
    $('document-panel').setAttribute('aria-labelledby', 'tab-' + key);
    $('document-title').textContent = DOCUMENTOS[key].title;
    $('pdf-viewer').hidden = key === 'excel';
    $('xlsx-viewer').hidden = key !== 'excel';
    message('Cargando el archivo…');
    try {
      if (key === 'excel') {
        state.book ||= await readWorkbook(bytes(DOCUMENTOS.excel.base64));
        if (version !== state.version) return;
        renderSheet();
        message('Excel integrado · ' + state.book.length + ' hojas · Solo lectura');
      } else {
        if (state.pdf) { await state.pdf.destroy(); state.pdf = null; }
        const pdf = await pdfjsLib.getDocument({ data: bytes(DOCUMENTOS[key].base64), useWasm: false, isEvalSupported: false, useSystemFonts: true }).promise;
        if (version !== state.version) { await pdf.destroy(); return; }
        state.pdf = pdf;
        state.page = 1;
        $('pdf-page').max = pdf.numPages;
        $('pdf-total').textContent = 'de ' + pdf.numPages;
        $('pdf-zoom').value = 'fit';
        await renderPage();
      }
    } catch (error) {
      if (version !== state.version) return;
      message('No se pudo mostrar el documento. Usa un navegador actualizado o descarga el archivo con el botón del visor.');
      console.error(error);
    }
  }

  let renderSequence = 0;
  async function renderPage() {
    if (!state.pdf || state.doc === 'excel') return;
    const sequence = ++renderSequence;
    const documentVersion = state.version;
    if (state.task) { state.task.cancel(); try { await state.task.promise; } catch (_) {} }
    if (sequence !== renderSequence) return;
    const pageNumber = state.page;
    try {
      const page = await state.pdf.getPage(pageNumber);
      if (sequence !== renderSequence || documentVersion !== state.version) return;
      const base = page.getViewport({ scale: 1 });
      const zoom = $('pdf-zoom').value;
      const scale = zoom === 'fit' ? Math.min(2, Math.max(.2, ($('pdf-stage').clientWidth - 32) / base.width)) : Number(zoom);
      const viewport = page.getViewport({ scale });
      const canvas = $('pdf-canvas');
      const density = Math.min(window.devicePixelRatio || 1, 2);
      canvas.width = Math.round(viewport.width * density);
      canvas.height = Math.round(viewport.height * density);
      canvas.style.width = viewport.width + 'px';
      canvas.style.height = viewport.height + 'px';
      canvas.setAttribute('aria-label', DOCUMENTOS[state.doc].title + ', página ' + pageNumber);
      $('pdf-page').value = pageNumber;
      $('pdf-prev').disabled = pageNumber <= 1;
      $('pdf-next').disabled = pageNumber >= state.pdf.numPages;
      message('Mostrando página ' + pageNumber + ' de ' + state.pdf.numPages);
      state.task = page.render({ canvasContext: canvas.getContext('2d'), viewport, transform: density === 1 ? null : [density, 0, 0, density, 0, 0] });
      await state.task.promise;
      if (sequence !== renderSequence || documentVersion !== state.version) return;
      state.task = null;
      const text = await page.getTextContent();
      if (sequence !== renderSequence || documentVersion !== state.version) return;
      $('pdf-text-content').textContent = text.items.map(item => item.str + (item.hasEOL ? '\n' : ' ')).join('') || 'Este PDF contiene una imagen. Usa el zoom para ampliar su contenido; la misma información está disponible en las fichas de la línea de tiempo.';
    } catch (error) {
      if (error.name !== 'RenderingCancelledException' && documentVersion === state.version) {
        message('No se pudo dibujar esta página. Prueba otra página o descarga el documento.');
        console.error(error);
      }
    }
  }

  /* Lector OOXML mínimo: descomprime el XLSX, lee las hojas y sus valores
     guardados. No ejecuta macros, fórmulas ni relaciones externas. */
  async function unzip(data) {
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    let end = data.length - 22;
    while (end >= Math.max(0, data.length - 65557) && view.getUint32(end, true) !== 0x06054b50) end--;
    if (end < 0) throw new Error('El XLSX no es un ZIP válido.');
    const count = view.getUint16(end + 10, true);
    let offset = view.getUint32(end + 16, true);
    const files = {};
    const decoder = new TextDecoder();
    for (let i = 0; i < count; i++) {
      if (view.getUint32(offset, true) !== 0x02014b50) throw new Error('Entrada ZIP inválida.');
      const method = view.getUint16(offset + 10, true);
      const size = view.getUint32(offset + 20, true);
      const nameLength = view.getUint16(offset + 28, true);
      const extraLength = view.getUint16(offset + 30, true);
      const commentLength = view.getUint16(offset + 32, true);
      const local = view.getUint32(offset + 42, true);
      const name = decoder.decode(data.subarray(offset + 46, offset + 46 + nameLength));
      const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
      if (/\.(xml|rels)$/.test(name)) {
        let content = data.slice(start, start + size);
        if (method === 8) content = new Uint8Array(await new Response(new Blob([content]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());
        else if (method !== 0) throw new Error('Compresión no compatible.');
        files[name] = decoder.decode(content);
      }
      offset += 46 + nameLength + extraLength + commentLength;
    }
    return files;
  }
  const xml = value => new DOMParser().parseFromString(value, 'application/xml');
  const nodes = (element, name) => [...element.getElementsByTagNameNS('*', name)];
  const textNodes = element => nodes(element, 't').map(n => n.textContent).join('');
  async function readWorkbook(data) {
    const files = await unzip(data);
    const shared = files['xl/sharedStrings.xml'] ? nodes(xml(files['xl/sharedStrings.xml']), 'si').map(textNodes) : [];
    const relations = Object.fromEntries(nodes(xml(files['xl/_rels/workbook.xml.rels']), 'Relationship').map(r => [r.getAttribute('Id'), r.getAttribute('Target')]));
    return nodes(xml(files['xl/workbook.xml']), 'sheet').map(sheet => {
      const rid = sheet.getAttribute('r:id');
      const target = relations[rid];
      const path = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      const rows = nodes(xml(files[path]), 'row').map(row => {
        const values = [];
        nodes(row, 'c').forEach(cell => {
          const letters = cell.getAttribute('r').match(/^[A-Z]+/)[0];
          const column = [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
          const raw = nodes(cell, 'v')[0]?.textContent || '';
          const type = cell.getAttribute('t');
          values[column] = type === 's' ? shared[Number(raw)] : type === 'inlineStr' ? textNodes(cell) : type === 'b' ? (raw === '1' ? 'Verdadero' : 'Falso') : raw;
        });
        return { number: Number(row.getAttribute('r')), values };
      });
      return { name: sheet.getAttribute('name'), rows };
    });
  }
  const normalize = value => String(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLocaleLowerCase('es');
  function renderSheet() {
    const sheet = state.book[state.sheet];
    const tabs = $('sheet-tabs'); tabs.replaceChildren();
    state.book.forEach((s, i) => {
      const button = document.createElement('button');
      button.textContent = s.name; button.className = 'action';
      button.setAttribute('role', 'tab'); button.setAttribute('aria-selected', i === state.sheet);
      button.tabIndex = i === state.sheet ? 0 : -1;
      button.onclick = () => { state.sheet = i; $('sheet-search').value = ''; renderSheet(); tabs.children[i].focus(); };
      tabs.append(button);
    });
    const header = sheet.rows.find(row => row.number === 6);
    const records = sheet.rows.filter(row => row.number > 6 && row.values.some(v => v !== ''));
    const query = normalize($('sheet-search').value.trim());
    const selected = records.filter(row => normalize(row.values.join(' ')).includes(query));
    const table = $('sheet-table'); table.replaceChildren();
    const caption = table.createCaption();
    caption.textContent = [2,3,4].map(n => sheet.rows.find(r => r.number === n)?.values[0]).filter(Boolean).join(' · ');
    const widths = [56, ...header.values.map(label => /^(ID|Año|Fuentes)$/.test(label) ? 88 : /Análisis y conectividad/.test(label) ? 540 : /Descripción|Impacto|Conexión|Referencia|Enlaces/.test(label) ? 350 : /Título|Hito$/.test(label) ? 270 : 180)];
    table.style.width = widths.reduce((a,b) => a+b, 0) + 'px';
    const columns = document.createElement('colgroup');
    widths.forEach(width => { const col = document.createElement('col'); col.style.width = width + 'px'; columns.append(col); });
    table.append(columns);
    const thead = table.createTHead().insertRow();
    ['Fila', ...header.values].forEach(label => { const th = document.createElement('th'); th.scope = 'col'; th.textContent = label; thead.append(th); });
    const tbody = table.createTBody();
    selected.forEach(row => {
      const tr = tbody.insertRow();
      const number = document.createElement('th'); number.scope = 'row'; number.textContent = row.number; tr.append(number);
      header.values.forEach((_, i) => { const td = tr.insertCell(); td.textContent = row.values[i] || ''; });
    });
    $('sheet-summary').textContent = sheet.name + ' · ' + selected.length + ' de ' + records.length + ' registros · ' + header.values.length + ' columnas';
    if (!selected.length) { const td = tbody.insertRow().insertCell(); td.colSpan = header.values.length + 1; td.textContent = 'No hay resultados para esta búsqueda.'; }
  }

  function enableArrowTabs(container, selector) {
    container.addEventListener('keydown', event => {
      if (!['ArrowLeft','ArrowRight','Home','End'].includes(event.key)) return;
      const buttons = [...container.querySelectorAll(selector)];
      const current = buttons.indexOf(document.activeElement);
      if (current < 0) return;
      event.preventDefault();
      const next = event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length-1 : (current + (event.key === 'ArrowRight' ? 1 : -1) + buttons.length) % buttons.length;
      buttons[next].focus(); buttons[next].click();
    });
  }
  document.querySelectorAll('[data-doc]').forEach(button => button.onclick = () => openDocument(button.dataset.doc));
  enableArrowTabs(document.querySelector('.doc-tabs'), 'button');
  enableArrowTabs($('sheet-tabs'), 'button');
  $('pdf-prev').onclick = () => { if (state.page > 1) { state.page--; renderPage(); $('pdf-stage').scrollTop = 0; } };
  $('pdf-next').onclick = () => { if (state.pdf && state.page < state.pdf.numPages) { state.page++; renderPage(); $('pdf-stage').scrollTop = 0; } };
  $('pdf-page').onchange = () => { if (!state.pdf) return; state.page = Math.max(1, Math.min(state.pdf.numPages, Math.round(Number($('pdf-page').value)) || 1)); renderPage(); $('pdf-stage').scrollTop = 0; };
  $('pdf-zoom').onchange = renderPage;
  $('sheet-search').oninput = renderSheet;
  $('document-download').onclick = () => {
    const item = DOCUMENTOS[state.doc];
    const url = URL.createObjectURL(new Blob([bytes(item.base64)], {type: state.doc === 'excel' ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/pdf'}));
    const a = document.createElement('a'); a.href = url; a.download = item.file; a.click(); setTimeout(() => URL.revokeObjectURL(url), 10000);
  };
  const closeExpanded = () => { $('document-panel').classList.remove('expanded'); document.body.classList.remove('viewer-open'); $('viewer-expand').setAttribute('aria-expanded','false'); $('viewer-expand').textContent = 'Ampliar visor'; renderPage(); };
  $('viewer-expand').onclick = () => {
    if ($('document-panel').classList.contains('expanded')) return closeExpanded();
    $('document-panel').classList.add('expanded'); document.body.classList.add('viewer-open');
    $('viewer-expand').setAttribute('aria-expanded','true'); $('viewer-expand').textContent = 'Reducir visor'; renderPage();
  };
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('document-panel').classList.contains('expanded')) closeExpanded(); });
  let timer;
  window.addEventListener('resize', () => { clearTimeout(timer); timer = setTimeout(() => { if ($('pdf-zoom').value === 'fit') renderPage(); }, 180); });
  // Carga inicial cuando el lector se aproxima a la pantalla.
  if ('IntersectionObserver' in window) {
    const observer = new IntersectionObserver(entries => { if (entries.some(e => e.isIntersecting)) { observer.disconnect(); openDocument(state.doc); } }, {rootMargin:'400px'});
    observer.observe($('entrega'));
  } else openDocument('informe');
})();
