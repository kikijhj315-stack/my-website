// app.js — 주간영업보고 자동화 (템플릿 없이 엑셀 처음부터 생성)

// ──────────────────────────────────────────────
// 헬퍼 함수들
// ──────────────────────────────────────────────

function getCellValue(cell) {
  if (!cell) return '';
  const val = cell.value;
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') {
    if (val.formula)   return val.result !== undefined ? val.result : '';
    if (val.richText)  return val.richText.map(t => t.text).join('');
  }
  return val;
}

function parseExcelDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') {
    return new Date((val - 25569) * 86400 * 1000);
  }
  if (typeof val === 'string') {
    const cleaned = val.replace(/\./g, '-').replace(/\//g, '-').trim();
    const parsed = new Date(cleaned);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = e => resolve(e.target.result);
    reader.onerror = e => reject(e);
    reader.readAsArrayBuffer(file);
  });
}

function downloadBlob(buffer, filename) {
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// 열 문자 → 숫자
function colLetterToNum(col) {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num;
}

// 첫 행에서 열 탐색 (정확 매치 우선 → 포함 매치)
function findColumnIndex(sheet, possibleNames, defaultIdx) {
  if (!sheet) return defaultIdx;
  const headerRow = sheet.getRow(1);
  if (!headerRow) return defaultIdx;

  const map = [];
  headerRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
    const v = String(getCellValue(cell)).trim();
    if (v) map.push({ col: colNum, norm: v.toLowerCase().replace(/\s+/g, '') });
  });

  for (const name of possibleNames) {
    const n = name.toLowerCase().replace(/\s+/g, '');
    for (const h of map) { if (h.norm === n) return h.col; }
  }
  for (const name of possibleNames) {
    const n = name.toLowerCase().replace(/\s+/g, '');
    for (const h of map) { if (h.norm.includes(n)) return h.col; }
  }
  return defaultIdx;
}

// 병합 해제 (model._merges 직접 조작)
function cleanSheetData(worksheet) {
  if (!worksheet) return;
  try {
    const mergeRefs = [];
    if (worksheet.model && Array.isArray(worksheet.model.merges)) {
      mergeRefs.push(...worksheet.model.merges);
    }
    if (worksheet._merges) {
      const existing = new Set(mergeRefs);
      for (const k of Object.keys(worksheet._merges)) {
        if (!existing.has(k)) mergeRefs.push(k);
      }
    }

    const savedValues = {};
    mergeRefs.forEach(ref => {
      try {
        const [startRef, endRef = startRef] = ref.split(':');
        const sm = startRef.match(/^([A-Z]+)(\d+)$/);
        const em = endRef.match(/^([A-Z]+)(\d+)$/);
        if (!sm || !em) return;

        const sc = colLetterToNum(sm[1]), sr = parseInt(sm[2], 10);
        const ec = colLetterToNum(em[1]), er = parseInt(em[2], 10);

        const masterCell = worksheet.getRow(sr).getCell(sc);
        const val = masterCell ? getCellValue(masterCell) : null;

        for (let r = sr; r <= er; r++) {
          for (let c = sc; c <= ec; c++) {
            if (val !== null && val !== undefined && String(val).trim() !== '') {
              savedValues[`${r}:${c}`] = val;
            }
          }
        }
      } catch(e) { /* 무시 */ }
    });

    // 병합 완전 초기화
    if (worksheet.model) worksheet.model.merges = [];
    if (worksheet._merges) {
      for (const k of Object.keys(worksheet._merges)) delete worksheet._merges[k];
    }

    // 값 복원
    for (const [key, val] of Object.entries(savedValues)) {
      const [r, c] = key.split(':').map(Number);
      try { worksheet.getRow(r).getCell(c).value = val; } catch(e) { /* 무시 */ }
    }

    // 모든 열에 대해 Fill-down (병합이 해제된 빈 셀을 위 값으로 채움)
    const maxRow = worksheet.actualRowCount || worksheet.rowCount || 3000;
    const maxCol = worksheet.actualColumnCount || worksheet.columnCount || 20;
    for (let c = 1; c <= maxCol; c++) {
      let lastVal = '';
      for (let r = 2; r <= maxRow; r++) {
        const row  = worksheet.getRow(r);
        if (!row) continue;
        const cell = row.getCell(c);
        const v    = String(getCellValue(cell)).trim();
        if (v && v !== 'null') {
          lastVal = v;
        } else if (lastVal) {
          cell.value = lastVal;
        }
      }
    }
  } catch(err) {
    console.error('[cleanSheetData]', err);
  }
}

// ──────────────────────────────────────────────
// 업로드된 파일 버퍼 저장소
// ──────────────────────────────────────────────
let uploadedData = {
  rawData:       null,
  topSales:      null,
  productMaster: null,
  coupang:       null
};

// ──────────────────────────────────────────────
// 스타일 헬퍼
// ──────────────────────────────────────────────
const COLORS = {
  titleBg:     'FF1E1B4B',  // 짙은 남색
  titleFont:   'FFFFFFFF',
  headerBg:    'FF2D2B8E',  // 보라 계열
  headerFont:  'FFFFFFFF',
  subHeaderBg: 'FF4A4890',
  subHeaderFont: 'FFFFFFFF',
  rowEven:     'FFF0EFFF',
  rowOdd:      'FFFFFFFF',
  catBg:       'FFEDE9FE',  // 연보라
  border:      'FFCCCCCC',
  numFont:     'FF1A237E',
};

function borderStyle(color = COLORS.border) {
  const s = { style: 'thin', color: { argb: color } };
  return { top: s, left: s, bottom: s, right: s };
}

function applyCell(cell, {
  value, bold = false, italic = false, fontSize = 10,
  fontColor = '00000000', bgColor = null,
  hAlign = 'left', vAlign = 'middle',
  numFmt = null, border = true, wrapText = false
} = {}) {
  if (value !== undefined) cell.value = value;
  cell.font      = { bold, italic, size: fontSize, color: { argb: fontColor }, name: '맑은 고딕' };
  cell.alignment = { horizontal: hAlign, vertical: vAlign, wrapText };
  if (bgColor) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
  if (numFmt)  cell.numFmt = numFmt;
  if (border)  cell.border = borderStyle();
}

// ──────────────────────────────────────────────
// 메인 처리 함수
// ──────────────────────────────────────────────
async function processFilesAndGenerate({ startDate, endDate, managerName, selectedClients, clientRankMap, selectedCategories }) {
  if (!uploadedData.rawData || !uploadedData.topSales || !uploadedData.productMaster) {
    throw new Error('원본데이터, 매출TOP, Product_Master 파일을 모두 업로드해 주세요.');
  }

  // ── 1. 워크북 로드 ──
  const rawWorkbook    = new ExcelJS.Workbook();
  await rawWorkbook.xlsx.load(uploadedData.rawData.slice(0));
  const topWorkbook    = new ExcelJS.Workbook();
  await topWorkbook.xlsx.load(uploadedData.topSales.slice(0));
  const masterWorkbook = new ExcelJS.Workbook();
  await masterWorkbook.xlsx.load(uploadedData.productMaster.slice(0));

  const rawSheet    = rawWorkbook.worksheets[0];
  const topSheet    = topWorkbook.worksheets[0];
  const masterSheet = masterWorkbook.worksheets[0];

  if (!rawSheet || !topSheet || !masterSheet) {
    throw new Error('업로드된 파일의 시트를 읽을 수 없습니다.');
  }

  // ── 2. 전처리 ──
  cleanSheetData(rawSheet);
  cleanSheetData(topSheet);

  // ── 헤더 탐색 — 실제 파일 열 위치 확정값 사용 —
  // B열(2): 몰주문일, D열(4): 몰주문번호, F열(6): 주문상태
  // G열(7): 매출처, K열(11): 수량, M열(13): 판매가
  const rawDateCol   = findColumnIndex(rawSheet, ['몰주문일', '주문일', '주문일자', '일자'], 2);
  const rawOrderCol  = findColumnIndex(rawSheet, ['몰주문번호', '주문번호'], 4);
  const rawStatusCol = findColumnIndex(rawSheet, ['주문상태', '상태'], 6);
  const rawClientCol = findColumnIndex(rawSheet, ['매출처', '고객사', '판매처'], 7);
  const rawQtyCol    = findColumnIndex(rawSheet, ['매출수량', '판매수량', '수량'], 11);
  const rawAmountCol = findColumnIndex(rawSheet, ['판매가', '판매금액', '판매액', '매출액'], 13);

  const topClientCol     = 1; // A열: 매출처
  const topModelGroupCol = 2; // B열: 모델그룹
  const topQtyCol        = findColumnIndex(topSheet, ['판매수량', '매출수량', '수량'], 4);
  const topAmountCol     = findColumnIndex(topSheet, ['매출금액', '매출액', '판매금액', '금액'], 5);

  const pmCatCol         = 6; // F열: 국내분류
  const pmModelGroupCol  = 7; // G열: 모델그룹

  console.log('[DEBUG] 열 매핑 →', { rawClientCol, rawDateCol, rawStatusCol, rawQtyCol, rawAmountCol, rawOrderCol });
  console.log('[DEBUG] TOP 열 →', { topClientCol, topModelGroupCol, topQtyCol, topAmountCol });
  console.log('[DEBUG] Master 열 →', { pmCatCol, pmModelGroupCol });

  // ── 3. Product_Master 파싱 ──
  const productCategoryMap = {}; // 모델그룹명(하이픈 이후 텍스트) → 국내분류
  masterSheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    let modelGroup = String(getCellValue(row.getCell(pmModelGroupCol))).trim();
    const cat      = String(getCellValue(row.getCell(pmCatCol))).trim();
    if (modelGroup && cat) {
      // 하이픈(-) 앞의 텍스트와 하이픈 제거 (예: TM-MUA05 -> MUA05)
      const dashIdx = modelGroup.indexOf('-');
      if (dashIdx !== -1) {
        modelGroup = modelGroup.substring(dashIdx + 1).trim();
      }
      productCategoryMap[modelGroup] = cat;
    }
  });
  console.log('[DEBUG] Product_Master 모델수:', Object.keys(productCategoryMap).length);

  // ── 4. 원본데이터 파싱 (시트1) ──
  const start = new Date(startDate); start.setHours(0, 0, 0, 0);
  const end   = new Date(endDate);   end.setHours(23, 59, 59, 999);

  // { client: { orderNos: Set, qty, amount } }
  const clientMap = {};
  let skippedStatus = 0, skippedDate = 0, passedRows = 0;

  rawSheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return;

    // 주문상태 필터
    const status = String(getCellValue(row.getCell(rawStatusCol))).trim();
    if (status === '품절취소' || status === '주문취소') { skippedStatus++; return; }

    // 날짜 필터
    const dateVal   = getCellValue(row.getCell(rawDateCol));
    const orderDate = parseExcelDate(dateVal);
    if (orderDate) {
      const t = new Date(orderDate); t.setHours(12, 0, 0, 0);
      if (t < start || t > end) { skippedDate++; return; }
    }

    const client  = String(getCellValue(row.getCell(rawClientCol))).trim();
    if (!client || client === 'null') return;

    const orderNo   = String(getCellValue(row.getCell(rawOrderCol))).trim();
    const qty       = parseFloat(getCellValue(row.getCell(rawQtyCol)))    || 0;
    const unitPrice = parseFloat(getCellValue(row.getCell(rawAmountCol))) || 0;
    const amount    = qty * unitPrice; // 매출금액 = 수량 × 판매가 (행별 곱 후 합산)

    if (!clientMap[client]) clientMap[client] = { orderNos: new Set(), qty: 0, amount: 0 };
    if (orderNo) clientMap[client].orderNos.add(orderNo);
    clientMap[client].qty    += qty;
    clientMap[client].amount += amount;
    passedRows++;

  });

  console.log(`[DEBUG] 원본데이터 처리: 상태제외=${skippedStatus}, 날짜제외=${skippedDate}, 통과=${passedRows}`);
  console.log('[DEBUG] 매출처 목록:', Object.keys(clientMap));

  // 선택된 매출처만 필터링
  const sheet1Rows = selectedClients
    .filter(c => clientMap[c])
    .map(c => ({
      client:     c,
      orderCount: clientMap[c].orderNos.size,
      qty:        clientMap[c].qty,
      amount:     clientMap[c].amount
    }));

  // 데이터가 없어도 선택된 매출처는 빈 행으로 포함
  selectedClients.forEach(c => {
    if (!clientMap[c]) {
      sheet1Rows.push({ client: c, orderCount: 0, qty: 0, amount: 0 });
    }
  });

  // ── 4-2. 쿠팡로켓 파싱 (선택) ──
  if (uploadedData.coupang) {
    const cpWorkbook = new ExcelJS.Workbook();
    await cpWorkbook.xlsx.load(uploadedData.coupang.slice(0));
    const cpSheet = cpWorkbook.worksheets[0];
    if (cpSheet) {
      let cpOrderCount = 0;
      let cpQty = 0;
      let cpAmount = 0;
      
      cpSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const dateVal = getCellValue(row.getCell(1)); // A열 몰주문일
        const orderDate = parseExcelDate(dateVal);
        
        if (orderDate) {
          const t = new Date(orderDate); t.setHours(12, 0, 0, 0);
          if (t >= start && t <= end) {
            // G열(7) 수량, H열(8) 단가
            const q = parseFloat(getCellValue(row.getCell(7))) || 0;
            const p = parseFloat(getCellValue(row.getCell(8))) || 0;
            cpOrderCount++; // 모델명 건수가 주문건수 (행 기준 집계)
            cpQty += q;
            cpAmount += (q * p); // G열 * H열
          }
        }
      });

      if (cpOrderCount > 0) {
        sheet1Rows.push({
          client: '쿠팡로켓',
          orderCount: cpOrderCount,
          qty: cpQty,
          amount: cpAmount
        });
      }
    }
  }

  // ── 4-3. 최종 정렬 (매핑표 순위 기준) ──
  if (clientRankMap) {
    sheet1Rows.sort((a, b) => {
      const rankA = clientRankMap[a.client] !== undefined ? clientRankMap[a.client] : 9999;
      const rankB = clientRankMap[b.client] !== undefined ? clientRankMap[b.client] : 9999;
      return rankA - rankB;
    });
  }

  // ── 5. 매출TOP 파싱 (기본 데이터 수집) ──
  const parsedTopSales = [];
  let lastTopClient = ''; 

  topSheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return;

    let client = String(getCellValue(row.getCell(topClientCol))).trim();
    if (client && client !== 'null') {
      lastTopClient = client;
    } else {
      client = lastTopClient;
    }

    let modelGroup = String(getCellValue(row.getCell(topModelGroupCol))).trim();
    const dashIdx = modelGroup.indexOf('-');
    if (dashIdx !== -1) {
      modelGroup = modelGroup.substring(dashIdx + 1).trim();
    }

    const qty    = parseFloat(getCellValue(row.getCell(topQtyCol)))    || 0;
    const amount = parseFloat(getCellValue(row.getCell(topAmountCol))) || 0;

    if (!client || !modelGroup || client === '테슬라재팬') return;

    const cat = productCategoryMap[modelGroup] || '미분류';
    parsedTopSales.push({ client, modelGroup, qty, amount, cat });
  });

  // ── 6. 엑셀 생성 ──
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Weekly Sales Automation';
  wb.created  = new Date();

  buildSheet1(wb, { startDate, endDate, managerName, sheet1Rows });

  // 카테고리별 시트 생성 로직
  const categoriesToBuild = (selectedCategories && selectedCategories.length > 0)
    ? selectedCategories
    : ['전체'];

  categoriesToBuild.forEach(targetCat => {
    const sheetClientMap = {}; 
    const pivotModelGroupsSet = new Set();
    const modelGroupTotalQty = {}; // 모델그룹별 총 수량 (열 내림차순 정렬용)

    parsedTopSales.forEach(row => {
      if (targetCat !== '전체' && row.cat !== targetCat) return;

      if (!sheetClientMap[row.client]) {
        sheetClientMap[row.client] = { qty: 0, amount: 0, models: {} };
      }
      sheetClientMap[row.client].qty    += row.qty;
      sheetClientMap[row.client].amount += row.amount;

      if (!sheetClientMap[row.client].models[row.modelGroup]) {
        sheetClientMap[row.client].models[row.modelGroup] = 0;
      }
      sheetClientMap[row.client].models[row.modelGroup] += row.qty;
      
      pivotModelGroupsSet.add(row.modelGroup);
      
      if (!modelGroupTotalQty[row.modelGroup]) modelGroupTotalQty[row.modelGroup] = 0;
      modelGroupTotalQty[row.modelGroup] += row.qty;
    });

    const sheet2Data = Object.entries(sheetClientMap).map(([client, data]) => ({
      client,
      qty: data.qty,
      amount: data.amount,
      models: data.models
    }));
    
    // 행 정렬 (매출처별 수량 내림차순)
    sheet2Data.sort((a, b) => b.qty - a.qty);

    // 열 정렬 (모델그룹 전체 합계 수량 내림차순)
    const pivotModelGroups = Array.from(pivotModelGroupsSet);
    pivotModelGroups.sort((a, b) => modelGroupTotalQty[b] - modelGroupTotalQty[a]);

    buildSheet2(wb, { sheet2Data, targetCat, pivotModelGroups });
  });

  const buffer = await wb.xlsx.writeBuffer();
  // 반환값에 들어갈 임의의 첫 번째 sheet2Data 설정 (메타데이터 반환용)
  return { buffer, meta: { startDate, endDate, managerName }, sheet1Rows };
}

// ──────────────────────────────────────────────
// 시트1: 주간 영업보고
// ──────────────────────────────────────────────
function buildSheet1(wb, { startDate, endDate, managerName, sheet1Rows }) {
  const ws = wb.addWorksheet('주간 영업보고');
  ws.views = [{ zoomScale: 90 }]; // 화면 보기 확대/축소 90%

  // 열 너비
  ws.columns = [
    { width: 22 }, // A: 매출처
    { width: 15 }, // B: 주문건수
    { width: 15 }, // C: 주문수량
    { width: 22 }, // D: 매출금액
    { width: 15 }, // E: 평균판가
    { width: 40 }, // F: 영업내용
  ];

  const numFmt = '#,##0';

  // ── 행 1: 제목 ──
  ws.mergeCells('A1:F1');
  applyCell(ws.getCell('A1'), {
    value: '주간 영업보고',
    bold: true, fontSize: 20,
    fontColor: COLORS.titleFont, bgColor: COLORS.titleBg,
    hAlign: 'center', border: true
  });
  ws.getRow(1).height = 45;

  // ── 행 2: 담당자 / 기간 ──
  ws.mergeCells('A2:C2');

  applyCell(ws.getCell('A2'), {
    value: `담당자: ${managerName}`,
    bold: true, fontSize: 14,
    fontColor: COLORS.titleFont, bgColor: COLORS.titleBg,
    hAlign: 'left', border: true
  });
  ws.mergeCells('D2:F2');
  applyCell(ws.getCell('D2'), {
    value: `기간: ${startDate} ~ ${endDate}`,
    bold: false, fontSize: 14,
    fontColor: COLORS.titleFont, bgColor: COLORS.titleBg,
    hAlign: 'right', border: true
  });
  ws.getRow(2).height = 20;

  // ── 행 3: 헤더 ──

  const headers = ['매출처', '주문건수', '주문수량', '매출금액', '평균판가', '영업내용'];
  headers.forEach((h, i) => {
    applyCell(ws.getRow(3).getCell(i + 1), {
      value: h,
      bold: true, fontSize: 18,
      fontColor: COLORS.headerFont, bgColor: COLORS.headerBg,
      hAlign: 'center', border: true
    });
  });
  ws.getRow(3).height = 40;

  // ── 행 4~: 데이터 ──
  sheet1Rows.forEach((row, idx) => {
    const rowNum  = idx + 4;
    const r       = ws.getRow(rowNum);
    const bgColor = idx % 2 === 0 ? COLORS.rowEven : COLORS.rowOdd;
    const avgVal  = row.qty > 0 ? Math.round(row.amount / row.qty) : 0;

    applyCell(r.getCell(1), { value: row.client,     bold: true,  fontSize: 18, bgColor, hAlign: 'center', border: true });
    applyCell(r.getCell(2), { value: row.orderCount, bold: false, fontSize: 18, bgColor, hAlign: 'center', border: true, numFmt });
    applyCell(r.getCell(3), { value: row.qty,        bold: false, fontSize: 18, bgColor, hAlign: 'right',  border: true, numFmt });
    applyCell(r.getCell(4), { value: row.amount,     bold: false, fontSize: 18, bgColor, hAlign: 'right',  border: true, numFmt });
    applyCell(r.getCell(5), { value: avgVal,         bold: false, fontSize: 18, bgColor, hAlign: 'right',  border: true, numFmt });
    applyCell(r.getCell(6), { value: '',             bold: false, fontSize: 18, bgColor, hAlign: 'left',   border: true, wrapText: true });

    r.height = 80; // 행 높이 80
  });

  // ── 합계 행 ──
  const dataRowCount  = sheet1Rows.length;
  const totalRowNum   = dataRowCount + 4; // 4행부터 데이터 시작
  const totalOrderCnt = sheet1Rows.reduce((s, r) => s + r.orderCount, 0);
  const totalQty      = sheet1Rows.reduce((s, r) => s + r.qty, 0);
  const totalAmount   = sheet1Rows.reduce((s, r) => s + r.amount, 0);
  const totalAvg      = totalQty > 0 ? Math.round(totalAmount / totalQty) : 0;

  const totalBg = 'FF2D2B8E'; // 헤더와 같은 보라 배경
  const tr = ws.getRow(totalRowNum);
  applyCell(tr.getCell(1), { value: '합  계', bold: true, fontSize: 18, bgColor: totalBg, fontColor: COLORS.headerFont, hAlign: 'center', border: true });
  applyCell(tr.getCell(2), { value: totalOrderCnt, bold: true, fontSize: 18, bgColor: totalBg, fontColor: COLORS.headerFont, hAlign: 'center', border: true, numFmt });
  applyCell(tr.getCell(3), { value: totalQty,      bold: true, fontSize: 18, bgColor: totalBg, fontColor: COLORS.headerFont, hAlign: 'right',  border: true, numFmt });
  applyCell(tr.getCell(4), { value: totalAmount,   bold: true, fontSize: 18, bgColor: totalBg, fontColor: COLORS.headerFont, hAlign: 'right',  border: true, numFmt });
  applyCell(tr.getCell(5), { value: totalAvg,      bold: true, fontSize: 18, bgColor: totalBg, fontColor: COLORS.headerFont, hAlign: 'right',  border: true, numFmt });
  applyCell(tr.getCell(6), { value: '',            bold: false, fontSize: 18, bgColor: totalBg, border: true });
  tr.height = 40;

  // 빈 행이면 최소 1행 표시
  if (sheet1Rows.length === 0) {
    const r = ws.getRow(4);
    ws.mergeCells('A4:F4');
    applyCell(r.getCell(1), { value: '해당 조건에 일치하는 데이터가 없습니다.', hAlign: 'center', border: true });
    r.height = 20;
  }
}

// ──────────────────────────────────────────────
// 시트2: 카테고리별 판매현황
// ──────────────────────────────────────────────
function buildSheet2(wb, { sheet2Data, targetCat, pivotModelGroups }) {
  const sheetName = targetCat === '전체' ? '판매현황' : `${targetCat} 판매현황`;
  // 시트 이름이 이미 존재할 경우 처리 (예: 동일 이름 중복)
  let ws;
  try {
    ws = wb.addWorksheet(sheetName);
  } catch (e) {
    ws = wb.addWorksheet(`${sheetName}_${Math.floor(Math.random() * 1000)}`);
  }
  
  ws.views = [{ zoomScale: 90 }]; // 화면 보기 확대/축소 90%

  ws.columns = [
    { width: 24 }, // A: 매출처
    { width: 14 }, // B: 판매수량
    { width: 18 }, // C: 판매금액
  ];

  const numFmt = '#,##0';

  // ── 행 1: 제목 ──
  ws.mergeCells('A1:C1');
  applyCell(ws.getCell('A1'), {
    value: `${targetCat} 판매현황`,
    bold: true, fontSize: 16,
    fontColor: COLORS.titleFont, bgColor: COLORS.titleBg,
    hAlign: 'center', border: true
  });
  ws.getRow(1).height = 36;

  // ── 행 2: 헤더 ──
  ['매출처', '판매수량', '판매금액'].forEach((h, i) => {
    applyCell(ws.getRow(2).getCell(i + 1), {
      value: h,
      bold: true, fontSize: 10,
      fontColor: COLORS.headerFont, bgColor: COLORS.headerBg,
      hAlign: 'center', border: true
    });
  });
  ws.getRow(2).height = 22;

  // 필터 지정
  ws.autoFilter = 'A2:C2';

  // ── 행 3~: 데이터 ──
  let currentRow = 3;

  if (sheet2Data.length === 0) {
    ws.mergeCells(`A3:C3`);
    applyCell(ws.getRow(3).getCell(1), { value: '매출TOP 데이터가 없습니다.', hAlign: 'center', border: true });
    return;
  }

  sheet2Data.forEach((row, idx) => {
    const r = ws.getRow(currentRow);
    const bgColor = currentRow % 2 === 0 ? COLORS.rowEven : COLORS.rowOdd;
    applyCell(r.getCell(1), { value: row.client, bold: false, bgColor, hAlign: 'left',  border: true });
    applyCell(r.getCell(2), { value: row.qty,    bold: false, bgColor, hAlign: 'right', border: true, numFmt });
    applyCell(r.getCell(3), { value: row.amount, bold: false, bgColor, hAlign: 'right', border: true, numFmt });
    r.height = 20;
    currentRow++;
  });

  // ── 합계 행 ──
  const totalQty = sheet2Data.reduce((s, r) => s + r.qty, 0);
  const totalAmount = sheet2Data.reduce((s, r) => s + r.amount, 0);

  const totalBg = 'FF2D2B8E';
  const tr = ws.getRow(currentRow);
  applyCell(tr.getCell(1), { value: '합  계', bold: true, fontSize: 10, bgColor: totalBg, fontColor: COLORS.headerFont, hAlign: 'center', border: true });
  applyCell(tr.getCell(2), { value: totalQty, bold: true, fontSize: 10, bgColor: totalBg, fontColor: COLORS.headerFont, hAlign: 'right', border: true, numFmt });
  applyCell(tr.getCell(3), { value: totalAmount, bold: true, fontSize: 10, bgColor: totalBg, fontColor: COLORS.headerFont, hAlign: 'right', border: true, numFmt });
  tr.height = 28;

  // ── 피벗 매트릭스 렌더링 (G열부터) ──
  const pivotStartCol = 7; // G열

  if (pivotModelGroups && pivotModelGroups.length > 0) {
    // 1행 제목 (병합 해제, 단독 셀 G1에 작성)
    applyCell(ws.getRow(1).getCell(pivotStartCol), {
      value: '모델그룹별 판매수량', bold: true, fontSize: 16,
      fontColor: COLORS.titleFont, bgColor: COLORS.titleBg,
      hAlign: 'center', border: true
    });

    // 2행 헤더
    applyCell(ws.getRow(2).getCell(pivotStartCol), {
      value: '매출처별', bold: true, fontSize: 10,
      fontColor: COLORS.headerFont, bgColor: COLORS.headerBg,
      hAlign: 'center', border: true
    });
    ws.getColumn(pivotStartCol).width = 24;

    pivotModelGroups.forEach((modelGroup, idx) => {
      const colIdx = pivotStartCol + 1 + idx;
      applyCell(ws.getRow(2).getCell(colIdx), {
        value: modelGroup, bold: true, fontSize: 10,
        fontColor: COLORS.headerFont, bgColor: COLORS.headerBg,
        hAlign: 'center', border: true
      });
      ws.getColumn(colIdx).width = 12;
    });

    // 3행~ 데이터 (내림차순 정렬 요청됨 - sheet2Data는 이미 내림차순 정렬 상태)
    let pRow = 3;
    sheet2Data.forEach(pRowData => {
      const r = ws.getRow(pRow);
      const bgColor = pRow % 2 === 0 ? COLORS.rowEven : COLORS.rowOdd;

      applyCell(r.getCell(pivotStartCol), {
        value: pRowData.client, bold: false, bgColor, hAlign: 'left', border: true
      });

      pivotModelGroups.forEach((modelGroup, idx) => {
        const colIdx = pivotStartCol + 1 + idx;
        const qty = pRowData.models[modelGroup] || 0;
        applyCell(r.getCell(colIdx), {
          value: qty > 0 ? qty : '', // 0 대신 빈칸
          bold: false, bgColor, hAlign: 'right', border: true, numFmt: '#,##0'
        });
      });
      
      r.height = 20;
      pRow++;
    });

    // 맨 하단 총합계
    const pTotalRow = ws.getRow(pRow);
    const totalBg2 = 'FF2D2B8E';
    applyCell(pTotalRow.getCell(pivotStartCol), {
      value: '총 합계', bold: true, fontSize: 10, bgColor: totalBg2, fontColor: COLORS.headerFont, hAlign: 'center', border: true
    });

    pivotModelGroups.forEach((modelGroup, idx) => {
      const colIdx = pivotStartCol + 1 + idx;
      let sumQty = 0;
      sheet2Data.forEach(p => { sumQty += (p.models[modelGroup] || 0); });
      
      applyCell(pTotalRow.getCell(colIdx), {
        value: sumQty > 0 ? sumQty : '', 
        bold: true, fontSize: 10, bgColor: totalBg2, fontColor: COLORS.headerFont, hAlign: 'right', border: true, numFmt: '#,##0'
      });
    });
    pTotalRow.height = 28;
  }
}
