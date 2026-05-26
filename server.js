require('dotenv').config();
const express = require('express');
const multer = require('multer');
const axios = require('axios');
const mammoth = require('mammoth');
const pdfParse = require('pdf-parse');
const path = require('path');
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, PageBreak } = require('docx');

const app = express();
const upload = multer({ dest: 'uploads/' });
app.use(express.json());
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

// ?¼íŠ¸ë¹??œì? ?‰ìƒ
const BLUE = "5B9BD5";   // ???¤ë” ë°°ê²½
const LIGHT = "DEEAF6";  // ???°ì´??ë°°ê²½
const bd = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
const borders = { top: bd, bottom: bd, left: bd, right: bd };

// ???¤ë” ?€ (?Œë? ë°°ê²½, ??ê¸€?? Pretendard)
function hCell(text, w) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: { fill: BLUE, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: String(text||''), bold: true, color: "FFFFFF", size: 20, font: "Pretendard" })]
    })]
  });
}

// ???°ì´???€ (?°íŒŒ??ë°°ê²½, ?¨ì´ˆë¡¬ë°”??
function dCell(text, w, align) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: { fill: LIGHT, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: String(text||''), size: 20, font: "Pretendard" })]
    })]
  });
}

// ë¹??€
function eCell(w) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: "", size: 20 })] })]
  });
}

// ë³¸ë¬¸ ???°ì´???€ (Pretendard, ?¼ìª½ ?•ë ¬)
function bCell(text, w) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: { fill: LIGHT, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: AlignmentType.BOTH,
      children: [new TextRun({ text: String(text||''), size: 20, font: "Pretendard" })]
    })]
  });
}

// ë¬¸ë‹¨ ?ì„±
function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : (opts.left ? AlignmentType.LEFT : AlignmentType.BOTH),
    indent: opts.indent ? { left: opts.indent, hanging: opts.hanging || 0 } : undefined,
    spacing: { before: opts.before || 120, after: opts.after || 120, line: 276 },
    children: [new TextRun({
      text: String(text||''),
      bold: opts.bold || false,
      size: opts.size || 22,
      font: "Pretendard"
    })]
  });
}

// ë³¸ë¬¸ ???ì„±
function buildBodyTable(rows) {
  if (!rows.length) return null;
  const headers = rows[0].split('|').map(s => s.trim());
  const colCount = headers.length;
  const totalW = 9026;
  const colW = Math.floor(totalW / colCount);
  const colWidths = headers.map((_, i) => i === colCount-1 ? totalW - colW*(colCount-1) : colW);
  const tableRows = [new TableRow({ children: headers.map((h, i) => hCell(h, colWidths[i])) })];
  for (let r = 1; r < rows.length; r++) {
    const cells = rows[r].split('|').map(s => s.trim());
    while (cells.length < colCount) cells.push('');
    tableRows.push(new TableRow({ children: cells.map((c, i) => bCell(c, colWidths[i])) }));
  }
  return new Table({ width: { size: totalW, type: WidthType.DXA }, columnWidths: colWidths, rows: tableRows });
}

function buildDoc(aiText, meta) {
  const children = [];

  // ===== ?œì? =====
  children.push(p(""), p(""), p(""), p(""));
  children.push(p(meta.title, { bold: true, size: 36, center: true }));
  children.push(p(""));
  children.push(p("ì£¼ì‹?Œì‚¬ ?¼íŠ¸ë¹?, { bold: true, size: 26, center: true }));
  children.push(p(""));
  children.push(p("Version " + (meta.version||"1.0"), { size: 22, center: true }));
  children.push(p(""), p(""), p(""));
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ===== ?¬ê·œ?´ì—­??=====
  children.push(p("?¬ê·œ?´ì—­", { bold: true, size: 24, left: true }));
  children.push(p(""));
  children.push(new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [1800, 2713, 1800, 2713],
    rows: [
      new TableRow({ children: [hCell("ê´€ë¦¬ë²ˆ??,1800), dCell(meta.manageNum,2713), hCell("?¬ê·œëª?,1800), dCell(meta.title,2713)] }),
      new TableRow({ children: [hCell("ê¸°ì•ˆë¶€??,1800), dCell(meta.dept,2713), hCell("ê³µí¬?¼ìž",1800), dCell(meta.pubDate,2713)] }),
      new TableRow({ children: [hCell("?¹ì¸??,1800), dCell(meta.approver,2713), hCell("?œí–‰?¼ìž",1800), dCell(meta.effDate,2713)] }),
    ]
  }));
  children.push(new Paragraph({ spacing: { before: 0, after: 0 } }));

  // ===== ê°œì •?´ë ¥??=====
  children.push(p("ê°œì •?´ë ¥", { bold: true, size: 24, left: true }));
  children.push(p(""));
  const histRows = [new TableRow({ children: [hCell("ë²„ì „",900),hCell("ë³€ê²½ë‚´??,2726),hCell("?œÂ·ê°œ??ê³µí¬?¼ìž",1500),hCell("?œí–‰?¼ìž",1500),hCell("?‘ì„±??,1200),hCell("?¹ì¸??,1200)] })];
  meta.history.forEach(h => {
    histRows.push(new TableRow({ children: [dCell(h.ver,900), dCell(h.content,2726), dCell(h.pubDate,1500), dCell(h.effDate,1500), dCell(h.author,1200), dCell(h.approver,1200)] }));
  });
  histRows.push(new TableRow({ children: [eCell(900),eCell(2726),eCell(1500),eCell(1500),eCell(1200),eCell(1200)] }));
  histRows.push(new TableRow({ children: [eCell(900),eCell(2726),eCell(1500),eCell(1500),eCell(1200),eCell(1200)] }));
  children.push(new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [900,2726,1500,1500,1200,1200], rows: histRows }));
  children.push(p(""));
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ===== ëª©ì°¨ =====
  children.push(p("ëª? ì°?, { bold: true, size: 24, center: true }));
  children.push(p(""));
  meta.toc.forEach(t => children.push(p(t, { size: 22, left: true })));
  children.push(p(""));
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ===== ë³¸ë¬¸ =====
  const lines = aiText.split('\n');
  let inBody = false, inTable = false, tableRows2 = [];

  for (const line of lines) {
    const t = line.trim();
    if (t.includes('[ë³¸ë¬¸]') || t === '## ë³¸ë¬¸') { inBody = true; continue; }
    if (!inBody) continue;

    if (t === '[TABLE_START]') { inTable = true; tableRows2 = []; continue; }
    if (t === '[TABLE_END]') {
      inTable = false;
      const tbl = buildBodyTable(tableRows2);
      if (tbl) { children.push(tbl); children.push(p("")); }
      continue;
    }
    if (inTable) { if (t) tableRows2.push(t); continue; }
    if (!t) { children.push(p("")); continue; }

    const clean = t.replace(/^#{1,4}\s*/, '').trim();
    if (!clean) continue;

    if (/^??s*\d+\s*??.test(clean)) {
      // ???œëª©: êµµê²Œ, ê°€?´ë°, 14pt
      children.push(p(""));
      children.push(p(clean, { bold: true, size: 28, center: true }));
      children.push(p(""));
    } else if (/^??s*\d+\s*ì¡?.test(clean)) {
      // ì¡??œëª©: êµµê²Œ, ?¼ìª½
      children.push(p(clean, { bold: true, size: 22, left: true }));
    } else if (/^[? â‘¡?¢â‘£?¤â‘¥?¦â‘§?¨â‘©]/.test(clean)) {
      // ?? ?¤ì—¬?°ê¸° 1?¨ê³„
      children.push(p(clean, { size: 22, indent: 600 }));
    } else if (/^[1-9]\.\s/.test(clean)) {
      // ?? ?¤ì—¬?°ê¸° 2?¨ê³„
      children.push(p(clean, { size: 22, indent: 1200 }));
    } else if (/^[ê°€?˜ë‹¤?¼ë§ˆë°”ì‚¬?„ìžì°¨ì¹´?€?Œí•˜]\.\s/.test(clean)) {
      // ?¸ë?ëª? ?¤ì—¬?°ê¸° 3?¨ê³„
      children.push(p(clean, { size: 22, indent: 1800 }));
    } else if (/^(ë¶€\s*ì¹?ë¶€\s{2}ì¹?/.test(clean)) {
      children.push(p(""));
      children.push(p("ë¶€  ì¹?, { bold: true, size: 28, center: true }));
      children.push(p(""));
    } else {
      // ?¼ë°˜ ë³¸ë¬¸
      children.push(p(clean, { size: 22 }));
    }
  }

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1701, bottom: 1440, left: 1440, right: 1440 }
        }
      },
      children
    }]
  });
}

app.post('/convert', upload.single('file'), async (req, res) => {
  try {
    const apiKey = process.env.API_KEY;
    if (!apiKey) return res.status(500).json({ error: 'API ?¤ê? ?†ìŠµ?ˆë‹¤.' });
    if (!req.file) return res.status(400).json({ error: '?Œì¼???†ìŠµ?ˆë‹¤.' });

    const ext = (req.file.originalname || '').split('.').pop().toLowerCase();
    let docText = '';
    if (ext === 'pdf') {
      const pdfBuf = fs.readFileSync(req.file.path);
      const pdfData = await pdfParse(pdfBuf);
      docText = pdfData.text;
    } else if (ext === 'docx') {
      const r = await mammoth.extractRawText({ path: req.file.path });
      docText = r.value;
    } else {
      docText = fs.readFileSync(req.file.path, 'utf8');
    }
    fs.unlinkSync(req.file.path);

    const prompt = `?¹ì‹ ?€ ?¼íŠ¸ë¹?ì£¼ì‹?Œì‚¬ ?¼íŠ¸ë¹? ?¬ê·œ ë¬¸ì„œ ë³€???„ë¬¸ê°€?…ë‹ˆ??
?„ëž˜ ?ë³¸ ë¬¸ì„œë¥??¼íŠ¸ë¹??¬ê·œ ?œì? ?‘ì‹?¼ë¡œ ë³€?˜í•˜?¸ìš”.
?ˆë? ?”ì•½?˜ê±°??ë¶„ì„?˜ì? ë§ˆì„¸?? ?ë¬¸ ?´ìš©??ê·¸ë?ë¡?? ì??˜ë©´???„ëž˜ ?•ì‹?¼ë¡œë§?ì¶œë ¥?˜ì„¸??

(?¬ê·œëª?
ì£¼ì‹?Œì‚¬ ?¼íŠ¸ë¹?Version 1.0

[?¬ê·œ?´ì—­]
ê´€ë¦¬ë²ˆ?? (?ë³¸?ì„œ ì°¾ê¸°, ?†ìœ¼ë©?ë¹ˆì¹¸)
?¬ê·œëª? (?ë³¸ ?œëª©)
ê¸°ì•ˆë¶€?? (?ë³¸?ì„œ ì°¾ê¸°)
ê³µí¬?¼ìž: (?ë³¸?ì„œ ì°¾ê¸°)
?¹ì¸?? (?ë³¸?ì„œ ì°¾ê¸°)
?œí–‰?¼ìž: (?ë³¸?ì„œ ì°¾ê¸°)

[ê°œì •?´ë ¥]
ë²„ì „ | ë³€ê²½ë‚´??| ê³µí¬?¼ìž | ?œí–‰?¼ìž | ?‘ì„±??| ?¹ì¸??(?ë³¸ ?´ìš©)

[ëª©ì°¨]
??N ??(???œëª©) Â·Â·Â·Â· N

[ë³¸ë¬¸]
??N ?? (???œëª©)

??N ì¡°ã€?ì¡??œëª©)????(???´ìš©)
??(???´ìš©)
1. (???´ìš©)
2. (???´ìš©)
ê°€. (?¸ë?ëª??´ìš©)

?œê? ?ˆìœ¼ë©?ë°˜ë“œ???„ëž˜ ?•ì‹?¼ë¡œ:
[TABLE_START]
?¤ë”1 | ?¤ë”2 | ?¤ë”3
?°ì´?? | ?°ì´?? | ?°ì´??
[TABLE_END]

ë¶€  ì¹?
??1 ì¡°ã€ì‹œ?‰ì¼??(?´ìš©)

ì£¼ì˜?¬í•­:
- ?ë¬¸??ëª¨ë“  ì¡°í•­ê³??œë? ë¹ ì§?†ì´ ?¬í•¨?˜ì„¸??- "?Œì‚¬"??"?¼íŠ¸ë¹?ë¡?ë³€ê²½í•˜?¸ìš”
- ?„â€?§?ì? ?„Â·ì§?ìœ¼ë¡?ë³€ê²½í•˜?¸ìš”
- ?¤ëª…?´ë‚˜ ë¶„ì„?€ ?ˆë? ?°ì? ë§ˆì„¸??
===?ë³¸ ë¬¸ì„œ===
${docText}`;

    const r = await axios.post(
      'https://api.anthropic.com/v1/messages',
      { model: 'claude-sonnet-4-5', max_tokens: 8000, messages: [{ role: 'user', content: prompt }] },
      { headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' } }
    );

    let aiText = r.data.content.map(b => b.text||'').join('\n').trim();
    // ??????ì¤‘ê°„??    aiText = aiText.replace(/??g, 'Â·');

    const lines = aiText.split('\n');
    const meta = { title: '', version: '1.0', manageNum: '', dept: '', pubDate: '', approver: '', effDate: '', history: [], toc: [] };
    let section = '';

    for (const line of lines) {
      const t = line.trim();
      if (!t) continue;
      if (t.startsWith('[?¬ê·œ?´ì—­]')) { section = 'meta'; continue; }
      if (t.startsWith('[ê°œì •?´ë ¥]')) { section = 'hist'; continue; }
      if (t.startsWith('[ëª©ì°¨]')) { section = 'toc'; continue; }
      if (t.startsWith('[ë³¸ë¬¸]') || t === '## ë³¸ë¬¸') break;

      if (section === '') {
        const clean = t.replace(/^#+\s*/, '');
        if (!meta.title && !clean.startsWith('ì£¼ì‹?Œì‚¬') && !clean.startsWith('Version')) meta.title = clean;
        if (clean.startsWith('Version')) meta.version = clean.replace('Version','').trim();
      } else if (section === 'meta') {
        if (t.startsWith('ê´€ë¦¬ë²ˆ??')) meta.manageNum = t.replace('ê´€ë¦¬ë²ˆ??','').trim();
        if (t.startsWith('?¬ê·œëª?')) meta.title = meta.title || t.replace('?¬ê·œëª?','').trim();
        if (t.startsWith('ê¸°ì•ˆë¶€??')) meta.dept = t.replace('ê¸°ì•ˆë¶€??','').trim();
        if (t.startsWith('ê³µí¬?¼ìž:')) meta.pubDate = t.replace('ê³µí¬?¼ìž:','').trim();
        if (t.startsWith('?¹ì¸??')) meta.approver = t.replace('?¹ì¸??','').trim();
        if (t.startsWith('?œí–‰?¼ìž:')) meta.effDate = t.replace('?œí–‰?¼ìž:','').trim();
      } else if (section === 'hist') {
        const parts = t.split('|').map(s => s.trim());
        if (parts.length >= 4 && !parts[0].startsWith('ë²„ì „') && !parts[0].startsWith('[')) {
          meta.history.push({ ver: parts[0], content: parts[1], pubDate: parts[2]||'', effDate: parts[3]||'', author: parts[4]||'', approver: parts[5]||'' });
        }
      } else if (section === 'toc') {
        meta.toc.push(t);
      }
    }
    if (!meta.history.length) {
      meta.history.push({ ver: 'v.1', content: '?œì •ë³?, pubDate: meta.pubDate, effDate: meta.effDate, author: '', approver: meta.approver });
    }

    const doc = buildDoc(aiText, meta);
    const buf = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="sentbe_saryu.docx"');
    res.send(buf);

  } catch (e) {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    console.error(e.response?.data || e.message);
    res.status(500).json({ error: e.response?.data?.error?.message || e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => console.log('?œë²„ ?¤í–‰ ì¤? ' + PORT));
