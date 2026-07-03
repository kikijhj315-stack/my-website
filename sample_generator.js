// sample_generator.js — 테스트용 샘플 파일 4종 생성

async function generateSampleManagerMap() {
  const workbook = new ExcelJS.Workbook();
  const sheet    = workbook.addWorksheet('Sheet1');

  sheet.columns = [
    { header: '담당자', key: 'manager', width: 15 },
    { header: '매출처', key: 'client',  width: 20 }
  ];

  // 샘플 매핑 데이터
  const mapData = [
    { manager: '홍길동', client: '무신사' },
    { manager: '홍길동', client: '29CM' },
    { manager: '김철수', client: 'W컨셉' },
    { manager: '이영희', client: '지그재그' }
  ];

  sheet.getRow(1).font = { bold: true };
  mapData.forEach(d => sheet.addRow(d));

  return await workbook.xlsx.writeBuffer();
}

async function generateSampleCoupang() {
  const workbook = new ExcelJS.Workbook();
  const sheet    = workbook.addWorksheet('Sheet1');

  // 쿠팡 로켓 샘플 구조: A:몰주문일, G:수량, H:단가, 중간에는 임의 열
  sheet.columns = [
    { header: '몰주문일', key: 'date', width: 15 },
    { header: '임의2', key: 'c2', width: 10 },
    { header: '임의3', key: 'c3', width: 10 },
    { header: '모델명', key: 'model', width: 20 },
    { header: '임의5', key: 'c5', width: 10 },
    { header: '임의6', key: 'c6', width: 10 },
    { header: '수량', key: 'qty', width: 10 },
    { header: '단가', key: 'price', width: 15 }
  ];

  const today = new Date();
  const yyyymmdd = today.toISOString().substring(0, 10);

  const coupangData = [
    { date: yyyymmdd, c2: '', c3: '', model: 'CP-MODEL-A', c5: '', c6: '', qty: 10, price: 5000 },
    { date: yyyymmdd, c2: '', c3: '', model: 'CP-MODEL-B', c5: '', c6: '', qty: 5,  price: 20000 },
    { date: yyyymmdd, c2: '', c3: '', model: 'CP-MODEL-A', c5: '', c6: '', qty: 2,  price: 5000 }
  ];

  sheet.getRow(1).font = { bold: true };
  coupangData.forEach(d => sheet.addRow(d));

  return await workbook.xlsx.writeBuffer();
}


async function generateSampleRawData() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');

  sheet.columns = [
    { header: '매출처',     key: 'client',      width: 15 },
    { header: '매출처_상세', key: 'client_detail', width: 15 },
    { header: '몰주문일',   key: 'order_date',  width: 15 },
    { header: '담당자',     key: 'manager',     width: 12 },
    { header: '몰주문번호', key: 'order_no',    width: 15 },
    { header: '주문상태',   key: 'status',      width: 12 },
    { header: '매출수량',   key: 'qty',         width: 12 },
    { header: '판매금액',   key: 'amount',      width: 15 }
  ];

  const rawData = [
    { client: '무신사',  order_date: '2026-07-02', manager: '홍길동', order_no: 'M0001', status: '배송완료', qty: 10, amount: 250000 },
    { client: '무신사',  order_date: '2026-07-03', manager: '홍길동', order_no: 'M0002', status: '배송완료', qty: 5,  amount: 125000 },
    { client: '무신사',  order_date: '2026-07-02', manager: '홍길동', order_no: 'M0003', status: '주문취소', qty: 2,  amount: 50000  }, // 주문취소 테스트
    { client: '29CM',   order_date: '2026-07-04', manager: '홍길동', order_no: 'C0001', status: '배송완료', qty: 8,  amount: 320000 },
    { client: '29CM',   order_date: '2026-07-05', manager: '홍길동', order_no: 'C0002', status: '품절취소', qty: 12, amount: 480000 }, // 품절취소 테스트
    { client: 'W컨셉',  order_date: '2026-07-03', manager: '김철수', order_no: 'W0001', status: '배송완료', qty: 15, amount: 450000 },
    { client: 'W컨셉',  order_date: '2026-07-04', manager: '김철수', order_no: 'W0002', status: '배송완료', qty: 20, amount: 600000 },
    { client: '지그재그', order_date: '2026-07-02', manager: '이영희', order_no: 'Z0001', status: '배송완료', qty: 25, amount: 500000 },
    { client: '지그재그', order_date: '2026-07-03', manager: '이영희', order_no: 'Z0002', status: '배송완료', qty: 30, amount: 600000 }
  ];

  sheet.getRow(1).font = { bold: true };

  let currentRowIdx = 2;
  const clients = ['무신사', '29CM', 'W컨셉', '지그재그'];

  clients.forEach(cName => {
    const clientRows = rawData.filter(d => d.client === cName);
    const startRow   = currentRowIdx;

    clientRows.forEach((r, idx) => {
      sheet.addRow({
        client:       idx === 0 ? cName : '',
        client_detail: '',
        order_date:   r.order_date,
        manager:      r.manager,
        order_no:     r.order_no,
        status:       r.status,
        qty:          r.qty,
        amount:       r.amount
      });
      currentRowIdx++;
    });

    const endRow = currentRowIdx - 1;
    sheet.mergeCells(`A${startRow}:B${endRow}`);
    sheet.getCell(`A${startRow}`).value = cName;
  });

  return await workbook.xlsx.writeBuffer();
}

async function generateSampleTopSales() {
  const workbook = new ExcelJS.Workbook();
  const sheet    = workbook.addWorksheet('Sheet1');

  sheet.columns = [
    { header: '매출처',     key: 'client',        width: 15 },
    { header: '모델그룹',   key: 'modelGroup',    width: 15 },
    { header: '모델명',     key: 'model',         width: 20 },
    { header: '판매수량',   key: 'qty',           width: 12 },
    { header: '매출금액',   key: 'amount',        width: 15 }
  ];

  const topData = [
    { client: '무신사',  modelGroup: 'TM-MUA01', model: '반팔_화이트', qty: 50, amount: 1000000 },
    { client: '29CM',    modelGroup: 'TM-MUA05', model: '청바지_블루', qty: 25, amount: 1200000 },
    { client: 'W컨셉',   modelGroup: 'TM-MUA10', model: '슬랙스_블랙', qty: 40, amount: 800000  },
    { client: '테슬라재팬', modelGroup: 'TM-MUA05', model: 'TSL-X', qty: 200, amount: 20000000 },
  ];

  topData.forEach((d, idx) => {
    const r = sheet.getRow(idx + 2);
    r.getCell(1).value = d.client;
    r.getCell(2).value = d.modelGroup;
    r.getCell(3).value = d.model;
    r.getCell(4).value = d.qty;
    r.getCell(5).value = d.amount;
  });

  // G열부터 매트릭스 데이터 (2행 헤더, 3행부터 값)
  sheet.getCell('G1').value = '';
  sheet.getCell('G2').value = '매출처별';
  sheet.getCell('H2').value = 'MUA01';
  sheet.getCell('I2').value = 'MUA05';
  sheet.getCell('J2').value = 'MUA10';
  sheet.getCell('K2').value = 'MUA20'; // 카테고리에 없는 그룹(테스트)

  const pivotData = [
    { client: '무신사', m1: 50, m2: 0, m3: 15, m4: 5 },
    { client: '29CM',   m1: 0, m2: 25, m3: 100, m4: 0 },
    { client: 'W컨셉',  m1: 5, m2: 0, m3: 40, m4: 10 }
  ];

  pivotData.forEach((p, i) => {
    const r = sheet.getRow(i + 3);
    r.getCell(7).value = p.client; // G
    r.getCell(8).value = p.m1;     // H
    r.getCell(9).value = p.m2;     // I
    r.getCell(10).value = p.m3;    // J
    r.getCell(11).value = p.m4;    // K
  });

  sheet.getRow(1).font = { bold: true };
  sheet.getRow(2).font = { bold: true };

  return await workbook.xlsx.writeBuffer();
}

async function generateSampleProductMaster() {
  const workbook = new ExcelJS.Workbook();
  const sheet    = workbook.addWorksheet('Sheet1');

  sheet.columns = [
    { header: '모델명', key: 'model',    width: 20 },
    { header: '',       key: 'dummy1',   width: 10 },
    { header: '',       key: 'dummy2',   width: 10 },
    { header: '',       key: 'dummy3',   width: 10 },
    { header: '',       key: 'dummy4',   width: 10 },
    { header: '국내분류', key: 'category', width: 15 },
    { header: '모델그룹', key: 'modelGroup', width: 15 }
  ];

  const masterData = [
    { model: '반팔_화이트', dummy1: '', dummy2: '', dummy3: '', dummy4: '', category: '반팔', modelGroup: 'TM-MUA01' },
    { model: '반팔_블랙',   dummy1: '', dummy2: '', dummy3: '', dummy4: '', category: '반팔', modelGroup: 'TM-MUA02' },
    { model: '청바지_블루', dummy1: '', dummy2: '', dummy3: '', dummy4: '', category: '바지', modelGroup: 'TM-MUA05' },
    { model: '슬랙스_블랙', dummy1: '', dummy2: '', dummy3: '', dummy4: '', category: '바지', modelGroup: 'TM-MUA10' }
  ];

  sheet.getRow(1).font = { bold: true };
  masterData.forEach(d => sheet.addRow(d));

  return await workbook.xlsx.writeBuffer();
}
