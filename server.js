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

// 센트비 표준 색상
const BLUE = "5B9BD5";   // 표 헤더 배경
const LIGHT = "DEEAF6";  // 표 데이터 배경
const bd = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
const borders = { top: bd, bottom: bd, left: bd, right: bd };

// 표 헤더 셀 (파란 배경, 흰 글씨, Pretendard)
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

// 표 데이터 셀 (연파랑 배경, 함초롬바탕)
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

// 빈 셀
function eCell(w) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: "", size: 20 })] })]
  });
}

// 본문 표 데이터 셀 (Pretendard, 왼쪽 정렬)
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

// 문단 생성
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

// 본문 표 생성
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

  // ===== 표지 =====
  children.push(p(""), p(""), p(""), p(""));
  children.push(p(meta.title, { bold: true, size: 36, center: true }));
  children.push(p(""));
  children.push(p("주식회사 센트비", { bold: true, size: 26, center: true }));
  children.push(p(""));
  children.push(p("Version " + (meta.version||"1.0"), { size: 22, center: true }));
  children.push(p(""), p(""), p(""));
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ===== 사규내역표 =====
  children.push(p("사규내역", { bold: true, size: 24, left: true }));
  children.push(p(""));
  children.push(new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [1800, 2713, 1800, 2713],
    rows: [
      new TableRow({ children: [hCell("관리번호",1800), dCell(meta.manageNum,2713), hCell("사규명",1800), dCell(meta.title,2713)] }),
      new TableRow({ children: [hCell("기안부서",1800), dCell(meta.dept,2713), hCell("공포일자",1800), dCell(meta.pubDate,2713)] }),
      new TableRow({ children: [hCell("승인자",1800), dCell(meta.approver,2713), hCell("시행일자",1800), dCell(meta.effDate,2713)] }),
    ]
  }));
  children.push(new Paragraph({ spacing: { before: 0, after: 0 } }));

  // ===== 개정이력표 =====
  children.push(p("개정이력", { bold: true, size: 24, left: true }));
  children.push(p(""));
  const histRows = [new TableRow({ children: [hCell("버전",900),hCell("변경내용",2726),hCell("제·개정/공포일자",1500),hCell("시행일자",1500),hCell("작성자",1200),hCell("승인자",1200)] })];
  meta.history.forEach(h => {
    histRows.push(new TableRow({ children: [dCell(h.ver,900), dCell(h.content,2726), dCell(h.pubDate,1500), dCell(h.effDate,1500), dCell(h.author,1200), dCell(h.approver,1200)] }));
  });
  histRows.push(new TableRow({ children: [eCell(900),eCell(2726),eCell(1500),eCell(1500),eCell(1200),eCell(1200)] }));
  histRows.push(new TableRow({ children: [eCell(900),eCell(2726),eCell(1500),eCell(1500),eCell(1200),eCell(1200)] }));
  children.push(new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [900,2726,1500,1500,1200,1200], rows: histRows }));
  children.push(p(""));
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ===== 목차 =====
  children.push(p("목  차", { bold: true, size: 24, center: true }));
  children.push(p(""));
  meta.toc.forEach(t => children.push(p(t, { size: 22, left: true })));
  children.push(p(""));
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ===== 본문 =====
  const lines = aiText.split('\n');
  let inBody = false, inTable = false, tableRows2 = [];

  for (const line of lines) {
    const t = line.trim();
    if (t.includes('[본문]') || t === '## 본문') { inBody = true; continue; }
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

    if (/^제\s*\d+\s*장/.test(clean)) {
      // 장 제목: 굵게, 가운데, 14pt
      children.push(p(""));
      children.push(p(clean, { bold: true, size: 28, center: true }));
      children.push(p(""));
    } else if (/^제\s*\d+\s*조/.test(clean)) {
      // 조 제목: 굵게, 왼쪽
      children.push(p(clean, { bold: true, size: 22, left: true }));
    } else if (/^[①②③④⑤⑥⑦⑧⑨⑩]/.test(clean)) {
      // 항: 들여쓰기 1단계
      children.push(p(clean, { size: 22, indent: 600 }));
    } else if (/^[1-9]\.\s/.test(clean)) {
      // 호: 들여쓰기 2단계
      children.push(p(clean, { size: 22, indent: 1200 }));
    } else if (/^[가나다라마바사아자차카타파하]\.\s/.test(clean)) {
      // 세부목: 들여쓰기 3단계
      children.push(p(clean, { size: 22, indent: 1800 }));
    } else if (/^(부\s*칙|부\s{2}칙)/.test(clean)) {
      children.push(p(""));
      children.push(p("부  칙", { bold: true, size: 28, center: true }));
      children.push(p(""));
    } else {
      // 일반 본문
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
    if (!apiKey) return res.status(500).json({ error: 'API 키가 없습니다.' });
    if (!req.file) return res.status(400).json({ error: '파일이 없습니다.' });

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

    const prompt = `당신은 센트비(주식회사 센트비) 사규 문서 변환 전문가입니다.
아래 원본 문서를 센트비 사규 표준 양식으로 변환하세요.
절대 요약하거나 분석하지 마세요. 원문 내용을 그대로 유지하면서 아래 형식으로만 출력하세요.

(사규명)
주식회사 센트비
Version 1.0

[사규내역]
관리번호: (원본에서 찾기, 없으면 빈칸)
사규명: (원본 제목)
기안부서: (원본에서 찾기)
공포일자: (원본에서 찾기)
승인자: (원본에서 찾기)
시행일자: (원본에서 찾기)

[개정이력]
버전 | 변경내용 | 공포일자 | 시행일자 | 작성자 | 승인자
(원본 내용)

[목차]
제 N 장 (장 제목) ···· N

[본문]
제 N 장  (장 제목)

제 N 조【(조 제목)】
① (항 내용)
② (항 내용)
1. (호 내용)
2. (호 내용)
가. (세부목 내용)

표가 있으면 반드시 아래 형식으로:
[TABLE_START]
헤더1 | 헤더2 | 헤더3
데이터1 | 데이터2 | 데이터3
[TABLE_END]

부  칙

제 1 조【시행일】
(내용)

주의사항:
- 원문의 모든 조항과 표를 빠짐없이 포함하세요
- "회사"는 "센트비"로 변경하세요
- 임•직원은 임·직원으로 변경하세요
- 설명이나 분석은 절대 쓰지 마세요

===원본 문서===
${docText}`;

    const r = await axios.post(
      'https://api.anthropic.com/v1/messages',
      { model: 'claude-sonnet-4-5', max_tokens: 8000, messages: [{ role: 'user', content: prompt }] },
      { headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' } }
    );

    let aiText = r.data.content.map(b => b.text||'').join('\n').trim();
    // 큰 점 → 중간점
    aiText = aiText.replace(/•/g, '·');

    const lines = aiText.split('\n');
    const meta = { title: '', version: '1.0', manageNum: '', dept: '', pubDate: '', approver: '', effDate: '', history: [], toc: [] };
    let section = '';

    for (const line of lines) {
      const t = line.trim();
      if (!t) continue;
      if (t.startsWith('[사규내역]')) { section = 'meta'; continue; }
      if (t.startsWith('[개정이력]')) { section = 'hist'; continue; }
      if (t.startsWith('[목차]')) { section = 'toc'; continue; }
      if (t.startsWith('[본문]') || t === '## 본문') break;

      if (section === '') {
        const clean = t.replace(/^#+\s*/, '');
        if (!meta.title && !clean.startsWith('주식회사') && !clean.startsWith('Version')) meta.title = clean;
        if (clean.startsWith('Version')) meta.version = clean.replace('Version','').trim();
      } else if (section === 'meta') {
        if (t.startsWith('관리번호:')) meta.manageNum = t.replace('관리번호:','').trim();
        if (t.startsWith('사규명:')) meta.title = meta.title || t.replace('사규명:','').trim();
        if (t.startsWith('기안부서:')) meta.dept = t.replace('기안부서:','').trim();
        if (t.startsWith('공포일자:')) meta.pubDate = t.replace('공포일자:','').trim();
        if (t.startsWith('승인자:')) meta.approver = t.replace('승인자:','').trim();
        if (t.startsWith('시행일자:')) meta.effDate = t.replace('시행일자:','').trim();
      } else if (section === 'hist') {
        const parts = t.split('|').map(s => s.trim());
        if (parts.length >= 4 && !parts[0].startsWith('버전') && !parts[0].startsWith('[')) {
          meta.history.push({ ver: parts[0], content: parts[1], pubDate: parts[2]||'', effDate: parts[3]||'', author: parts[4]||'', approver: parts[5]||'' });
        }
      } else if (section === 'toc') {
        meta.toc.push(t);
      }
    }
    if (!meta.history.length) {
      meta.history.push({ ver: 'v.1', content: '제정본', pubDate: meta.pubDate, effDate: meta.effDate, author: '', approver: meta.approver });
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
app.listen(PORT, '0.0.0.0', () => console.log('서버 실행 중: ' + PORT));
