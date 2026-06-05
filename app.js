/* ---- переключение тем ---- */
function setTheme(theme) {
  if (!theme || theme === 'light') document.documentElement.removeAttribute('data-theme');
  else document.documentElement.setAttribute('data-theme', theme);
  try { localStorage.setItem('docfmt-theme', theme || 'light'); } catch (e) {}
}

// при загрузке: выставить в выпадающем списке сохранённую тему
document.addEventListener('DOMContentLoaded', function () {
  let saved = 'light';
  try { saved = localStorage.getItem('docfmt-theme') || 'light'; } catch (e) {}
  const sel = document.getElementById('themeSelect');
  if (sel) sel.value = saved;
  setTheme(saved);

  // Защита поля "Красная строка" от ввода значений ниже минимума
  const firstIndentInput = document.getElementById('firstIndent');

  firstIndentInput.addEventListener('blur', function () {
    const value = parseFloat(this.value);
    if (isNaN(value) || value < 1) {
      this.value = 1;
    }
  });

  firstIndentInput.addEventListener('input', function () {
    const value = parseFloat(this.value);
    if (!isNaN(value) && value < 1) {
      this.value = 1;
    }
  });
});

/* ---- вкладки ---- */
function showTab(tabName) {
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById(tabName).classList.add('active');
  if (typeof event !== 'undefined' && event.target) event.target.classList.add('active');
}

/* ---- имя выбранного файла ---- */
document.getElementById('fileInput').addEventListener('change', function () {
  const fileName = this.files[0] ? this.files[0].name : 'Файл не выбран';
  document.getElementById('fileName').textContent = fileName;
});

/* =========================================================================
   Namespaces и канонический порядок дочерних элементов (по схеме OOXML).
   ========================================================================= */
const NS = {
  w:   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r:   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  ct:  'http://schemas.openxmlformats.org/package/2006/content-types',
  xml: 'http://www.w3.org/XML/1998/namespace',
};
const PPR_ORDER = ['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'];
const RPR_ORDER = ['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath'];
const SECT_ORDER = ['headerReference','footerReference','footnotePr','endnotePr','type','pgSz','pgMar','paperSrc','pgBorders','lnNumType','pgNumType','cols','formProt','vAlign','noEndnote','titlePg','textDirection','bidi','rtlGutter','docGrid','printerSettings','sectPrChange'];
const LVL_ORDER = ['start','numFmt','lvlRestart','pStyle','isLgl','suff','lvlText','lvlPicBulletId','legacy','lvlJc','pPr','rPr'];

/* ---- DOM-хелперы ---- */
function parseXml(s){ return new DOMParser().parseFromString(s, 'application/xml'); }
function serializeXml(doc){ return new XMLSerializer().serializeToString(doc); }
function mk(doc, qname, attrs){
  const pfx = qname.split(':')[0];
  const el = doc.createElementNS(NS[pfx], qname);
  if (attrs) for (const k of Object.keys(attrs)){
    const apfx = k.split(':')[0];
    if (k.includes(':') && NS[apfx]) el.setAttributeNS(NS[apfx], k, String(attrs[k]));
    else el.setAttribute(k, String(attrs[k]));
  }
  return el;
}
function mkText(doc, qname, attrs, text){ const e = mk(doc, qname, attrs); e.appendChild(doc.createTextNode(text)); return e; }
function directChildren(parent){ return Array.from(parent.childNodes).filter(n => n.nodeType === 1); }
function firstChildLocal(parent, local){ return directChildren(parent).find(c => c.localName === local) || null; }
function removeByLocal(parent, names){ for (const c of directChildren(parent)) if (names.includes(c.localName)) parent.removeChild(c); }
function insertOrdered(parent, el, order){
  const idx = order.indexOf(el.localName); let ref = null;
  for (const c of directChildren(parent)){ const ci = order.indexOf(c.localName); if (ci > idx){ ref = c; break; } }
  parent.insertBefore(el, ref);
}
function ensurePr(doc, parent, qname){
  let pr = firstChildLocal(parent, qname.split(':')[1]);
  if (!pr){ pr = mk(doc, qname); parent.insertBefore(pr, parent.firstChild); }
  return pr;
}
function xmlEscape(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function ensureRootNs(xml){
  const m = xml.match(/<w:document\b[^>]*>/); if (!m) return xml;
  let tag = m[0];
  if (!/xmlns:w=/.test(tag)) tag = tag.replace('<w:document', `<w:document xmlns:w="${NS.w}"`);
  if (!/xmlns:r=/.test(tag)) tag = tag.replace('<w:document', `<w:document xmlns:r="${NS.r}"`);
  return xml.replace(m[0], tag);
}

/* ---- содержимое колонтитулов ---- */
function rPrStr(font, sz){
  return `<w:rPr><w:rFonts w:ascii="${xmlEscape(font)}" w:hAnsi="${xmlEscape(font)}" w:cs="${xmlEscape(font)}"/>` +
    `<w:b w:val="0"/><w:i w:val="0"/><w:color w:val="000000"/>` +
    `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/><w:u w:val="none"/></w:rPr>`;
}
function pageNumPara(rp, alignVal){
  return `<w:p><w:pPr><w:jc w:val="${alignVal}"/></w:pPr>` +
    `<w:r>${rp}<w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r>${rp}<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>` +
    `<w:r>${rp}<w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r>${rp}<w:t>1</w:t></w:r>` +
    `<w:r>${rp}<w:fldChar w:fldCharType="end"/></w:r></w:p>`;
}
function textPara(rp, text){
  return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${rp}<w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r></w:p>`;
}
function wrapHdr(parasXml){
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="${NS.w}" xmlns:r="${NS.r}">${parasXml}</w:hdr>`;
}
function wrapFtr(parasXml){
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="${NS.w}" xmlns:r="${NS.r}">${parasXml}</w:ftr>`;
}
function readRelMap(relsXml){
  const map = {}; const re = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/?>/g; let m;
  while ((m = re.exec(relsXml))){
    let t = m[2], path;
    if (t.startsWith('/')) path = t.slice(1);
    else if (t.startsWith('../')) path = t.replace(/^\.\.\//, '');
    else path = 'word/' + t;
    map[m[1]] = path;
  }
  return map;
}
function nextRid(relsXml){ let max = 0, m, re = /Id="rId(\d+)"/g; while ((m = re.exec(relsXml))) max = Math.max(max, parseInt(m[1], 10)); return 'rId' + (max + 1); }
function freeName(existing, base, ext){ let n = base + ext, i = 1; while (existing.has(n)){ n = base + (++i) + ext; } existing.add(n); return n; }

/* =========================================================================
   processDocx — применяет форматирование к PizZip-архиву (синхронно)
   ========================================================================= */
function processDocx(zip, params){
  const ALIGN = { 'По левому краю':'left', 'По центру':'center', 'По правому краю':'right', 'По ширине':'both' };
  const align = ALIGN[params.alignment] || params.alignment || 'both';
  const cmToTwips = cm => Math.round(cm * 567);
  const firstIndent = cmToTwips(params.firstIndent);
  const lineVal = Math.round(params.lineSpacing * 240);
  const sz = Math.round(params.fontSize * 2);
  const font = params.fontName;
  const wantToc = params.toc !== false;
  // колонтитулы: тексты для 1-й и остальных страниц + позиция нумерации
  const headerFirst = (params.headerFirst || '').trim();
  const footerFirst = (params.footerFirst || '').trim();
  const headerOther = (params.headerOther || '').trim();
  const footerOther = (params.footerOther || '').trim();
  const numPos   = params.numberingPos || 'none';      // 'none' | 'top' | 'bottom'
  const numAlign = params.numberingAlign || 'center';   // 'center' | 'right'
  // нормализация начертания: курсив/жирный/подчёркивание всегда выключаются, цвет — чёрный.
  // заголовки по умолчанию тоже без жирного (правило «жирный отключать»); можно вернуть флагом.
  const boldHeadings = params.boldHeadings === true;

  // --- document.xml ---
  const doc = parseXml(ensureRootNs(zip.file('word/document.xml').asText()));
  const body = doc.getElementsByTagNameNS(NS.w, 'body')[0];

  // стили, которые сами задают нумерацию (List Number, List Bullet, рус. «a3» и т.п.) —
  // абзацы с такими стилями тоже считаются элементами списка
  const listStyleIds = new Set();
  const stylesFile0 = zip.file('word/styles.xml');
  if (stylesFile0) {
    const sdoc = parseXml(stylesFile0.asText());
    let stylesChanged = false;
    for (const st of Array.from(sdoc.getElementsByTagNameNS(NS.w, 'style'))) {
      const sppr = firstChildLocal(st, 'pPr');
      if (sppr && firstChildLocal(sppr, 'numPr')) {
        const id = st.getAttributeNS(NS.w, 'styleId');
        if (id) listStyleIds.add(id);
      }
      // стили оглавления (TOC1..9, TOC Heading) — приводим начертание к чистому
      const sid = (st.getAttributeNS(NS.w, 'styleId') || '');
      if (/^toc/i.test(sid)) {
        let srpr = firstChildLocal(st, 'rPr');
        if (!srpr) { srpr = mk(sdoc, 'w:rPr'); st.appendChild(srpr); }
        removeByLocal(srpr, ['b', 'bCs', 'i', 'iCs', 'u', 'color']);
        insertOrdered(srpr, mk(sdoc, 'w:b',   { 'w:val': boldHeadings ? '1' : '0' }), RPR_ORDER);
        insertOrdered(srpr, mk(sdoc, 'w:i',   { 'w:val': '0' }), RPR_ORDER);
        insertOrdered(srpr, mk(sdoc, 'w:color', { 'w:val': '000000' }), RPR_ORDER);
        insertOrdered(srpr, mk(sdoc, 'w:u',   { 'w:val': 'none' }), RPR_ORDER);
        stylesChanged = true;
      }
    }
    if (stylesChanged) zip.file('word/styles.xml', serializeXml(sdoc));
  }

  // 1) абзацы и раны
  const paras = Array.from(doc.getElementsByTagNameNS(NS.w, 'p'));
  for (const p of paras){
    const text = Array.from(p.getElementsByTagNameNS(NS.w, 't')).map(t => t.textContent).join('').trim();

    const pPr = ensurePr(doc, p, 'w:pPr');
    const styleEl = firstChildLocal(pPr, 'pStyle');
    const styleId = styleEl ? styleEl.getAttributeNS(NS.w, 'val') : null;
    const isList = !!firstChildLocal(pPr, 'numPr') || (styleId && listStyleIds.has(styleId));
    const isHeading = !isList && text.length > 0 && text.length <= 120 && text === text.toUpperCase() && /[A-ZА-ЯЁ]/.test(text);

    // pStyle НЕ удаляем (иначе слетают стили списков и др.). У списков сохраняем их собственный отступ.
    removeByLocal(pPr, ['jc', 'spacing', 'outlineLvl']);
    insertOrdered(pPr, mk(doc, 'w:spacing', { 'w:line': lineVal, 'w:lineRule': 'auto' }), PPR_ORDER);
    if (!isList) {
      removeByLocal(pPr, ['ind']);
      insertOrdered(pPr, mk(doc, 'w:ind', { 'w:firstLine': isHeading ? 0 : firstIndent }), PPR_ORDER);
    }
    insertOrdered(pPr, mk(doc, 'w:jc', { 'w:val': isHeading ? 'center' : align }), PPR_ORDER);
    if (isHeading) insertOrdered(pPr, mk(doc, 'w:outlineLvl', { 'w:val': '0' }), PPR_ORDER);

    for (const run of Array.from(p.getElementsByTagNameNS(NS.w, 'r'))){
      const rPr = ensurePr(doc, run, 'w:rPr');
      // единый шрифт/кегль, всегда чёрный, без курсива/жирного/подчёркивания (даже если их задаёт стиль)
      removeByLocal(rPr, ['rFonts', 'sz', 'szCs', 'color', 'b', 'bCs', 'i', 'iCs', 'u']);
      insertOrdered(rPr, mk(doc, 'w:rFonts', { 'w:ascii': font, 'w:hAnsi': font, 'w:cs': font }), RPR_ORDER);
      insertOrdered(rPr, mk(doc, 'w:b',   { 'w:val': '0' }), RPR_ORDER);
      insertOrdered(rPr, mk(doc, 'w:bCs', { 'w:val': '0' }), RPR_ORDER);
      insertOrdered(rPr, mk(doc, 'w:i',   { 'w:val': '0' }), RPR_ORDER);
      insertOrdered(rPr, mk(doc, 'w:iCs', { 'w:val': '0' }), RPR_ORDER);
      insertOrdered(rPr, mk(doc, 'w:color', { 'w:val': '000000' }), RPR_ORDER);
      insertOrdered(rPr, mk(doc, 'w:sz',   { 'w:val': sz }), RPR_ORDER);
      insertOrdered(rPr, mk(doc, 'w:szCs', { 'w:val': sz }), RPR_ORDER);
      insertOrdered(rPr, mk(doc, 'w:u',   { 'w:val': 'none' }), RPR_ORDER);
      if (isHeading && boldHeadings){
        removeByLocal(rPr, ['b', 'bCs']);
        insertOrdered(rPr, mk(doc, 'w:b',   { 'w:val': '1' }), RPR_ORDER);
        insertOrdered(rPr, mk(doc, 'w:bCs', { 'w:val': '1' }), RPR_ORDER);
      }
    }
  }

  // маркеры списков (numbering.xml): размер — на всех уровнях; шрифт — только на нумерованных.
  const numFile = zip.file('word/numbering.xml');
  if (numFile) {
    const numDoc = parseXml(numFile.asText());
    for (const lvl of Array.from(numDoc.getElementsByTagNameNS(NS.w, 'lvl'))) {
      const fmt = firstChildLocal(lvl, 'numFmt');
      const isBullet = fmt && fmt.getAttributeNS(NS.w, 'val') === 'bullet';
      let rPr = firstChildLocal(lvl, 'rPr');
      if (!rPr) { rPr = mk(numDoc, 'w:rPr'); insertOrdered(lvl, rPr, LVL_ORDER); }
      removeByLocal(rPr, ['sz', 'szCs', 'color', 'b', 'bCs', 'i', 'iCs', 'u']);
      if (!isBullet) {
        removeByLocal(rPr, ['rFonts']);
        insertOrdered(rPr, mk(numDoc, 'w:rFonts', { 'w:ascii': font, 'w:hAnsi': font, 'w:cs': font }), RPR_ORDER);
      }
      insertOrdered(rPr, mk(numDoc, 'w:b',   { 'w:val': '0' }), RPR_ORDER);
      insertOrdered(rPr, mk(numDoc, 'w:i',   { 'w:val': '0' }), RPR_ORDER);
      insertOrdered(rPr, mk(numDoc, 'w:color', { 'w:val': '000000' }), RPR_ORDER);
      insertOrdered(rPr, mk(numDoc, 'w:sz', { 'w:val': sz }), RPR_ORDER);
      insertOrdered(rPr, mk(numDoc, 'w:szCs', { 'w:val': sz }), RPR_ORDER);
      insertOrdered(rPr, mk(numDoc, 'w:u',   { 'w:val': 'none' }), RPR_ORDER);
    }
    zip.file('word/numbering.xml', serializeXml(numDoc));
  }

  // поля страницы во всех секциях
  const sectPrs = Array.from(doc.getElementsByTagNameNS(NS.w, 'sectPr'));
  for (const sect of sectPrs){
    let pg = firstChildLocal(sect, 'pgMar');
    if (!pg){ pg = mk(doc, 'w:pgMar'); insertOrdered(sect, pg, SECT_ORDER); }
    pg.setAttributeNS(NS.w, 'w:top', String(cmToTwips(params.marginTop)));
    pg.setAttributeNS(NS.w, 'w:right', String(cmToTwips(params.marginRight)));
    pg.setAttributeNS(NS.w, 'w:bottom', String(cmToTwips(params.marginBottom)));
    pg.setAttributeNS(NS.w, 'w:left', String(cmToTwips(params.marginLeft)));
    if (!pg.getAttributeNS(NS.w, 'header')) pg.setAttributeNS(NS.w, 'w:header', '720');
    if (!pg.getAttributeNS(NS.w, 'footer')) pg.setAttributeNS(NS.w, 'w:footer', '720');
    if (!pg.getAttributeNS(NS.w, 'gutter')) pg.setAttributeNS(NS.w, 'w:gutter', '0');
  }

  // колонтитулы (1-я страница vs остальные + позиция нумерации)
  let relsXml = zip.file('word/_rels/document.xml.rels').asText();
  let ctXml   = zip.file('[Content_Types].xml').asText();
  let relMap  = readRelMap(relsXml);
  const existingNames = new Set(Object.keys(zip.files));
  const filesToWrite = {}; const relAdd = []; const ctAdd = [];
  const REL_TYPE = { footer:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
                     header:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header' };
  const CT_TYPE  = { footer:'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
                     header:'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml' };
  const rp = rPrStr(font, sz);
  const partCache = {};   // content -> {rid, part} чтобы не плодить одинаковые части

  // ставит/переписывает ссылку нужного типа во всех секциях
  function setRef(kind, type, content){      // kind: header|footer; type: default|first
    if (!content) return;
    const refLocal = kind + 'Reference';
    for (const sect of sectPrs){
      const ref = directChildren(sect).find(c => c.localName === refLocal && c.getAttributeNS(NS.w, 'type') === type);
      if (ref){
        const rid = ref.getAttributeNS(NS.r, 'id'); const part = relMap[rid];
        if (part){ filesToWrite[part] = content; continue; }
      }
      let cached = partCache[content];
      if (!cached){
        const part = freeName(existingNames, 'word/' + kind + 'Fmt', '.xml');
        const rid  = nextRid(relsXml + relAdd.map(r => `Id="${r.id}"`).join(''));
        filesToWrite[part] = content;
        relAdd.push({ id: rid, type: REL_TYPE[kind], target: part.replace('word/', '') });
        ctAdd.push({ part: '/' + part, ct: CT_TYPE[kind] });
        relMap[rid] = part;
        cached = partCache[content] = { rid, part };
      }
      insertOrdered(sect, mk(doc, 'w:' + refLocal, { 'w:type': type, 'r:id': cached.rid }), SECT_ORDER);
    }
  }

  const alignVal = (numAlign === 'right') ? 'right' : 'center';
  const numTop    = numPos === 'top'    ? pageNumPara(rp, alignVal) : '';
  const numBottom = numPos === 'bottom' ? pageNumPara(rp, alignVal) : '';

  // пустое поле «1-я страница» = берём текст как у остальных страниц
  const effHeaderFirst = headerFirst || headerOther;
  const effFooterFirst = footerFirst || footerOther;

  // номер страницы — только на «остальных» (со 2-й). Титульная (1-я) — без номера.
  const defHeader   = (headerOther ? textPara(rp, headerOther) : '') + numTop;
  const defFooter   = (footerOther ? textPara(rp, footerOther) : '') + numBottom;
  const firstHeader = (effHeaderFirst ? textPara(rp, effHeaderFirst) : '');
  const firstFooter = (effFooterFirst ? textPara(rp, effFooterFirst) : '');

  // отдельная 1-я страница нужна, если титул отличается ИЛИ есть нумерация (она идёт со 2-й)
  const needTitlePg = (effHeaderFirst !== headerOther) || (effFooterFirst !== footerOther) || (numPos !== 'none');
  // трогаем колонтитулы только если пользователь что-то задал во вкладке
  const anyColontitul = !!(headerFirst || footerFirst || headerOther || footerOther) || numPos !== 'none';

  if (anyColontitul){
    // полностью переопределяем колонтитулы: убираем прежние ссылки и titlePg во всех секциях
    for (const sect of sectPrs) removeByLocal(sect, ['headerReference', 'footerReference', 'titlePg']);

    if (defHeader) setRef('header', 'default', wrapHdr(defHeader));
    if (defFooter) setRef('footer', 'default', wrapFtr(defFooter));

    if (needTitlePg){
      // 1-я страница: только заданный текст (без номера); пустые части глушат «остальные» на титуле
      setRef('header', 'first', wrapHdr(firstHeader || '<w:p/>'));
      setRef('footer', 'first', wrapFtr(firstFooter || '<w:p/>'));
      for (const sect of sectPrs) insertOrdered(sect, mk(doc, 'w:titlePg'), SECT_ORDER);
    }
  }

  // оглавление в начало
  if (wantToc){
    const tocTitle = mk(doc, 'w:p');
    const tPr = mk(doc, 'w:pPr');
    tPr.appendChild(mk(doc, 'w:spacing', { 'w:line': lineVal, 'w:lineRule':'auto', 'w:after':'120' }));
    tPr.appendChild(mk(doc, 'w:jc', { 'w:val':'center' }));
    tocTitle.appendChild(tPr);
    const tr = mk(doc, 'w:r'), trp = mk(doc, 'w:rPr');
    trp.appendChild(mk(doc, 'w:rFonts', { 'w:ascii':font, 'w:hAnsi':font, 'w:cs':font }));
    trp.appendChild(mk(doc, 'w:b', { 'w:val': boldHeadings ? '1' : '0' }));
    trp.appendChild(mk(doc, 'w:i', { 'w:val': '0' }));
    trp.appendChild(mk(doc, 'w:color', { 'w:val': '000000' }));
    trp.appendChild(mk(doc, 'w:sz', { 'w:val': sz + 4 }));
    trp.appendChild(mk(doc, 'w:szCs', { 'w:val': sz + 4 }));
    trp.appendChild(mk(doc, 'w:u', { 'w:val': 'none' }));
    tr.appendChild(trp); tr.appendChild(mkText(doc, 'w:t', null, 'Оглавление')); tocTitle.appendChild(tr);

    const fld = mk(doc, 'w:p');
    const fldRun = t => { const r = mk(doc, 'w:r'); r.appendChild(mk(doc, 'w:fldChar', { 'w:fldCharType': t })); return r; };
    fld.appendChild(fldRun('begin'));
    const ir = mk(doc, 'w:r');
    ir.appendChild(mkText(doc, 'w:instrText', { 'xml:space':'preserve' }, ' TOC \\o "1-3" \\h \\z \\u '));
    fld.appendChild(ir);
    fld.appendChild(fldRun('separate'));
    const ph = mk(doc, 'w:r');
    const phPr = mk(doc, 'w:rPr');
    phPr.appendChild(mk(doc, 'w:rFonts', { 'w:ascii':font, 'w:hAnsi':font, 'w:cs':font }));
    phPr.appendChild(mk(doc, 'w:b', { 'w:val':'0' }));
    phPr.appendChild(mk(doc, 'w:i', { 'w:val':'0' }));
    phPr.appendChild(mk(doc, 'w:color', { 'w:val':'000000' }));
    phPr.appendChild(mk(doc, 'w:sz', { 'w:val': sz }));
    phPr.appendChild(mk(doc, 'w:szCs', { 'w:val': sz }));
    phPr.appendChild(mk(doc, 'w:u', { 'w:val':'none' }));
    ph.appendChild(phPr);
    ph.appendChild(mkText(doc, 'w:t', { 'xml:space':'preserve' }, 'Нет элементов. ПКМ по полю → «Обновить поле» (или F9).'));
    fld.appendChild(ph);
    fld.appendChild(fldRun('end'));

    const pb = mk(doc, 'w:p'); const pbr = mk(doc, 'w:r'); pbr.appendChild(mk(doc, 'w:br', { 'w:type':'page' })); pb.appendChild(pbr);

    const anchor = body.firstChild;
    for (const node of [tocTitle, fld, pb]) body.insertBefore(node, anchor);
  }

  // --- сохраняем document.xml ---
  zip.file('word/document.xml', serializeXml(doc));

  
  if (relAdd.length){
    const rels = relAdd.map(r => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"/>`).join('');
    relsXml = relsXml.replace('</Relationships>', rels + '</Relationships>');
    zip.file('word/_rels/document.xml.rels', relsXml);
  }

  if (ctAdd.length){
    const ovs = ctAdd.filter(o => !ctXml.includes(`PartName="${o.part}"`))
      .map(o => `<Override PartName="${o.part}" ContentType="${o.ct}"/>`).join('');
    if (ovs){ ctXml = ctXml.replace('</Types>', ovs + '</Types>'); zip.file('[Content_Types].xml', ctXml); }
  }
  // --- части колонтитулов ---
  for (const path of Object.keys(filesToWrite)) zip.file(path, filesToWrite[path]);

  // --- settings.xml: обновлять поля при открытии ---
  const setFile = zip.file('word/settings.xml');
  if (setFile && wantToc){
    let s = setFile.asText();
    if (!/<w:updateFields\b/.test(s)){
      s = s.replace(/(<w:settings\b[^>]*>)/, `$1<w:updateFields w:val="true"/>`);
      zip.file('word/settings.xml', s);
    }
  }
  return zip;
}

// чтение полей вкладки «Колонтитулы»
function byId(ids){ for (const id of ids){ const el = document.getElementById(id); if (el) return el; } return null; }

function inputNear(label){
  if (label.htmlFor){ const t = document.getElementById(label.htmlFor); if (t) return t; }
  const inside = label.querySelector('input, textarea, select'); if (inside) return inside;
  let sib = label.nextElementSibling;
  while (sib){
    if (sib.matches && sib.matches('input, textarea, select')) return sib;
    const q = sib.querySelector && sib.querySelector('input, textarea, select'); if (q) return q;
    sib = sib.nextElementSibling;
  }
  const cont = label.parentElement;
  if (cont){ const q = cont.querySelector('input, textarea, select'); if (q && q !== label) return q; }
  return null;
}
function findByLabel(rx){
  for (const lab of document.querySelectorAll('label, .label, .field-label')){
    if (rx.test(lab.textContent || '')){ const el = inputNear(lab); if (el) return el; }
  }
  return null;
}
function fieldValue(ids, rx){
  const el = byId(ids) || findByLabel(rx);
  return el ? (el.value || '') : '';
}
function checkValue(ids, rx, fallback){
  const el = byId(ids) || findByLabel(rx);
  return el ? !!el.checked : fallback;
}
// список «Нумерация страниц»: ищем select с пунктом «Не добавлять», читаем выбранное значение
function findNumberingSelect(){
  const byid = byId(['pageNumbering', 'numbering', 'pageNumber', 'pageNum']);
  if (byid && byid.tagName === 'SELECT') return byid;
  const lab = findByLabel(/нумерац/i);
  if (lab && lab.tagName === 'SELECT') return lab;
  return Array.from(document.querySelectorAll('select')).find(s => /не\s*добав/i.test(s.textContent || '')) || null;
}
function parseNumbering(){
  const sel = findNumberingSelect();
  if (!sel) return { numberingPos: 'none', numberingAlign: 'center' };
  // читаем ВИДИМЫЙ текст выбранного пункта (а не value — у него может быть "0"/"none" и т.п.)
  let opt = null;
  if (sel.options && typeof sel.selectedIndex === 'number' && sel.selectedIndex >= 0) opt = sel.options[sel.selectedIndex];
  if (!opt && sel.selectedOptions) opt = sel.selectedOptions[0];
  const text = ((opt && opt.textContent) || '').toLowerCase().trim();
  const val  = (sel.value || '').toLowerCase().trim();
  const raw  = text || val;
  if (!raw || /не\s*добав|none|без|нет|off|^0$/.test(raw)) return { numberingPos: 'none', numberingAlign: 'center' };
  const pos   = /(верх|вверх|сверху|top)/.test(raw) ? 'top' : 'bottom';
  const align = /(справ|право|right)/.test(raw) ? 'right' : 'center';
  return { numberingPos: pos, numberingAlign: align };
}


async function applyFormatting(){
  const fileInput = document.getElementById('fileInput');
  const status = document.getElementById('status');

  if (!fileInput.files[0]){
    status.className = 'status error';
    status.textContent = 'Пожалуйста, выберите файл!';
    return;
  }
  if (typeof PizZip === 'undefined'){
    status.className = 'status error';
    status.textContent = 'Не подключена библиотека PizZip (проверь <script> в index.html).';
    return;
  }

  status.className = 'status loading';
  status.textContent = 'Обрабатываем документ...';

  try {
    const file = fileInput.files[0];
    const arrayBuffer = await file.arrayBuffer();

    const num = parseNumbering();
    const params = {
      fontName:    document.getElementById('fontName').value,
      fontSize:    parseInt(document.getElementById('fontSize').value) || 14,
      alignment:   document.getElementById('alignment').value,
      lineSpacing: parseFloat(document.getElementById('lineSpacing').value) || 1.5,
      firstIndent: Math.max(1, parseFloat(document.getElementById('firstIndent').value) || 1),
      marginTop:    parseFloat(document.getElementById('marginTop').value)    || 2,
      marginBottom: parseFloat(document.getElementById('marginBottom').value) || 2,
      marginLeft:   parseFloat(document.getElementById('marginLeft').value)   || 3,
      marginRight:  parseFloat(document.getElementById('marginRight').value)  || 1.5,
      // колонтитулы — берутся из вкладки «Колонтитулы» (по id или по подписям полей)
      headerFirst: fieldValue(['headerFirst', 'headerFirstPage'], /верхн.*(1|перв).*страниц/i),
      footerFirst: fieldValue(['footerFirst', 'footerFirstPage'], /нижн.*(1|перв).*страниц/i),
      headerOther: fieldValue(['headerOther', 'headerRest'],      /верхн.*(остальн|друг).*страниц/i),
      footerOther: fieldValue(['footerOther', 'footerRest'],      /нижн.*(остальн|друг).*страниц/i),
      numberingPos:   num.numberingPos,
      numberingAlign: num.numberingAlign,
      // оглавление: подхватится, если есть переключатель; иначе по умолчанию включено
      toc: checkValue(['toc', 'optToc', 'tocEnabled', 'addToc'], /оглавл|содержан/i, true),
    };

    const zip = new PizZip(arrayBuffer);
    if (!zip.file('word/document.xml')) throw new Error('Это не похоже на документ Word (.docx).');

    processDocx(zip, params);

    const output = zip.generate({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });

    const url = URL.createObjectURL(output);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'formatted_' + file.name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 4000);

    status.className = 'status success';
    status.textContent = '✔ Готово! Файл скачан. В Word согласись обновить поля (или Ctrl+A → F9) для оглавления и номеров.';

  } catch (err) {
    status.className = 'status error';
    status.textContent = 'Ошибка: ' + (err && err.message ? err.message : err);
    console.error(err);
  }
}

/* =========================================================================
   Быстрая настройка: пресеты ГОСТ + разбор форматирования из текста.
   Оба механизма только ЗАПОЛНЯЮТ форму — применяет всё кнопка
   «Применить форматирование» (как обычно).
   ========================================================================= */

/* Пресеты. Значения подставляются в поля формы.
   pageNumbers — это видимый текст пункта в списке «Нумерация страниц». */
const PRESETS = {
  // Отчёт о НИР / реферат. Поля: лево 30мм, право 15мм, верх/низ 20мм; номер внизу по центру.
  gost7_32:    { label: 'ГОСТ 7.32-2017', fontName: 'Times New Roman', fontSize: 14, lineSpacing: 1.5,
                 alignment: 'По ширине', firstIndent: 1.25, marginTop: 2, marginBottom: 2, marginLeft: 3, marginRight: 1.5,
                 pageNumbers: 'Внизу по центру' },
  // Диссертации/авторефераты — те же базовые метрики.
  gost_r_7011: { label: 'ГОСТ Р 7.0.11', fontName: 'Times New Roman', fontSize: 14, lineSpacing: 1.5,
                 alignment: 'По ширине', firstIndent: 1.25, marginTop: 2, marginBottom: 2, marginLeft: 3, marginRight: 1.5,
                 pageNumbers: 'Внизу по центру' },
  // Типовая курсовая/реферат вуза — часто номер внизу справа.
  coursework:  { label: 'Курсовая / реферат', fontName: 'Times New Roman', fontSize: 14, lineSpacing: 1.5,
                 alignment: 'По ширине', firstIndent: 1.25, marginTop: 2, marginBottom: 2, marginLeft: 3, marginRight: 1.5,
                 pageNumbers: 'Внизу справа' },
  // Короткие работы (доклад/эссе) без нумерации, поля по 2 см.
  plain:       { label: 'Доклад / эссе', fontName: 'Times New Roman', fontSize: 14, lineSpacing: 1.5,
                 alignment: 'По ширине', firstIndent: 1.25, marginTop: 2, marginBottom: 2, marginLeft: 2, marginRight: 2,
                 pageNumbers: 'Не добавлять' },
};

/* ---- установка значений в поля (поддержка input и select) ---- */
function setVal(id, v){ const el = document.getElementById(id); if (el) el.value = v; }
function setSelectNumber(id, num){
  const sel = document.getElementById(id); if (!sel) return;
  let opt = Array.from(sel.options).find(o => parseFloat((o.value || o.textContent).replace(',', '.')) === num);
  if (!opt){ opt = new Option(String(num), String(num)); sel.add(opt); }   // нет такого варианта — добавляем
  sel.value = opt.value;
}
function setSelectText(id, text){
  const sel = document.getElementById(id); if (!sel) return;
  let opt = Array.from(sel.options).find(o => o.textContent.trim() === text);
  if (!opt){ opt = new Option(text, text); sel.add(opt); }
  sel.value = opt.value;
}

/* применяет нормализованный набор параметров к форме */
function applyParams(p){
  if (p.fontName)            setVal('fontName', p.fontName);
  if (p.fontSize)            setSelectNumber('fontSize', p.fontSize);
  if (p.alignment)           setSelectText('alignment', p.alignment);
  if (p.lineSpacing)         setSelectNumber('lineSpacing', p.lineSpacing);
  if (p.firstIndent != null) setVal('firstIndent', Math.max(1, p.firstIndent));
  if (p.marginTop != null)   setSelectNumber('marginTop', p.marginTop);
  if (p.marginBottom != null)setSelectNumber('marginBottom', p.marginBottom);
  if (p.marginLeft != null)  setSelectNumber('marginLeft', p.marginLeft);
  if (p.marginRight != null) setSelectNumber('marginRight', p.marginRight);
  if (p.pageNumbers)         setSelectText('pageNumbers', p.pageNumbers);
}

function quickMsg(ok, text){
  const s = document.getElementById('quickStatus');
  if (s){ s.className = 'quick-status ' + (ok ? 'ok' : 'err'); s.textContent = text; }
}

/* ---- кнопка-пресет ---- */
function applyPreset(name){
  const p = PRESETS[name];
  if (!p) return;
  applyParams(p);
  quickMsg(true, 'Подставлены параметры: ' + p.label + '. Проверьте вкладки и нажмите «Применить форматирование».');
}

/* ---- разбор форматирования из текста (правило-ориентированный, офлайн) ---- */
function parseFormatting(raw){
  const text = ' ' + String(raw).toLowerCase().replace(/ё/g, 'е') + ' ';
  const params = {}; const rec = [];

  // шрифт
  const FONTS = [
    ['Times New Roman', /times\s*new\s*roman|таймс|times/],
    ['Arial', /arial|ариал/],
    ['Calibri', /calibri|калибри/],
    ['Georgia', /georgia|джорджия/],
    ['Verdana', /verdana|вердана/],
    ['Tahoma', /tahoma|тахома/],
    ['Courier New', /courier\s*new|courier|курьер/],
    ['Garamond', /garamond|гарамонд/],
  ];
  for (const [name, re] of FONTS){ if (re.test(text)){ params.fontName = name; rec.push('шрифт ' + name); break; } }

  // размер шрифта
  let m = text.match(/(?:кегл[а-яё]*|размер[а-яё]*)\s*[:=]?\s*(\d{1,2})/)
       || text.match(/(\d{1,2})\s*(?:пт|pt)\b/)
       || text.match(/(?:roman|arial|calibri|georgia|verdana|tahoma|courier|garamond|таймс|ариал|калибри|джорджия|вердана|тахома|курьер|гарамонд)\D{0,6}(\d{1,2})\b/);
  if (m){ const v = parseInt(m[1], 10); if (v >= 8 && v <= 72){ params.fontSize = v; rec.push('размер ' + v); } }

  // межстрочный интервал
  if (/полуторн/.test(text)){ params.lineSpacing = 1.5; rec.push('интервал 1.5'); }
  else if (/двойн/.test(text)){ params.lineSpacing = 2.0; rec.push('интервал 2.0'); }
  else if (/одинарн|одинар/.test(text)){ params.lineSpacing = 1.0; rec.push('интервал 1.0'); }
  else {
    m = text.match(/(?:интервал[а-яё]*|межстрочн[а-яё]*)\s*[:=]?\s*(\d(?:[.,]\d)?)/) || text.match(/(\d[.,]\d)\s*(?:интервал|межстрочн)/);
    if (m){ const v = parseFloat(m[1].replace(',', '.')); params.lineSpacing = v; rec.push('интервал ' + v); }
  }

  // нумерация — вырезаем её фразу, чтобы «по центру» из неё не сбило выравнивание
  let work = text;
  const numM = work.match(/(?:без\s+|не\s+)?(?:нумер[а-яё]*|номер[а-яё]*(?:\s*страниц[а-яё]*)?)[^,.;]*/);
  if (numM){
    const clause = numM[0];
    work = work.replace(clause, ' ');
    if (/(?:^|\s)(?:без|не)(?:\s|$)/.test(clause)){ params.pageNumbers = 'Не добавлять'; rec.push('без нумерации'); }
    else {
      const top = /(верх|вверх|сверху)/.test(clause);
      const right = /(справ|право)/.test(clause);
      params.pageNumbers = (top ? 'Вверху' : 'Внизу') + (right ? ' справа' : ' по центру');
      rec.push('нумерация — ' + params.pageNumbers.toLowerCase());
    }
  }

  // выравнивание (из остатка text без фразы нумерации)
  if (/по\s*ширин|ширине|ширину/.test(work)){ params.alignment = 'По ширине'; rec.push('по ширине'); }
  else if (/лев\w*\s*кра|по\s*лев|влево/.test(work)){ params.alignment = 'По левому краю'; rec.push('по левому краю'); }
  else if (/прав\w*\s*кра|по\s*прав|вправо/.test(work)){ params.alignment = 'По правому краю'; rec.push('по правому краю'); }
  else if (/по\s*центр|центрир/.test(work)){ params.alignment = 'По центру'; rec.push('по центру'); }

  // красная строка / абзацный отступ
  m = text.match(/(?:красн[а-яё]*\s*строк[а-яё]*|абзацн[а-яё]*\s*отступ[а-яё]*|отступ[а-яё]*\s*перв[а-яё]*\s*строк[а-яё]*)\D{0,8}(\d(?:[.,]\d+)?)/);
  if (m){ const v = Math.max(1, parseFloat(m[1].replace(',', '.'))); params.firstIndent = v; rec.push('красная строка ' + v + ' см'); }

  // поля
  if (/пол[а-яё]*\s+по\s+гост/.test(text)){
    params.marginTop = 2; params.marginBottom = 2; params.marginLeft = 3; params.marginRight = 1.5;
    rec.push('поля по ГОСТ (2/2/3/1.5)');
  }
  const side = re => { const mm = text.match(re); return mm ? parseFloat(mm[1].replace(',', '.')) : undefined; };
  const t = side(/верхн[а-яё]*\s*(?:поле[а-яё]*)?\s*[:=]?\s*(\d(?:[.,]\d)?)\s*см/);
  const b = side(/нижн[а-яё]*\s*(?:поле[а-яё]*)?\s*[:=]?\s*(\d(?:[.,]\d)?)\s*см/);
  const l = side(/лев[а-яё]*\s*(?:поле[а-яё]*)?\s*[:=]?\s*(\d(?:[.,]\d)?)\s*см/);
  const r = side(/прав[а-яё]*\s*(?:поле[а-яё]*)?\s*[:=]?\s*(\d(?:[.,]\d)?)\s*см/);
  const sidesGiven = [t, b, l, r].some(x => x != null);
  const mAll = text.match(/пол[а-яё]*\s*(?:со\s+всех\s+сторон\s*)?[:=]?\s*(\d(?:[.,]\d)?)\s*см\s*(?:со\s+всех\s+сторон)?/);
  if (!sidesGiven && !/пол[а-яё]*\s+по\s+гост/.test(text) && mAll){
    const v = parseFloat(mAll[1].replace(',', '.'));
    params.marginTop = v; params.marginBottom = v; params.marginLeft = v; params.marginRight = v;
    rec.push('поля ' + v + ' см со всех сторон');
  }
  if (t != null) params.marginTop = t;
  if (b != null) params.marginBottom = b;
  if (l != null) params.marginLeft = l;
  if (r != null) params.marginRight = r;
  if (sidesGiven){
    rec.push('поля: ' + ['верх ' + (params.marginTop ?? '—'), 'низ ' + (params.marginBottom ?? '—'),
                         'лево ' + (params.marginLeft ?? '—'), 'право ' + (params.marginRight ?? '—')].join(', '));
  }

  return { params, recognized: rec };
}

/* ---- кнопка «Применить описание» ---- */
function applyFromText(){
  const el = document.getElementById('nlInput');
  const text = el ? el.value : '';
  if (!text.trim()){ quickMsg(false, 'Введите описание форматирования словами.'); return; }
  const { params, recognized } = parseFormatting(text);
  applyParams(params);
  if (recognized.length){
    quickMsg(true, 'Распознано: ' + recognized.join(', ') + '. Проверьте и нажмите «Применить форматирование».');
  } else {
    quickMsg(false, 'Не удалось распознать параметры. Укажите шрифт, размер, интервал, выравнивание, поля или нумерацию.');
  }
}