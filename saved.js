/* =========================================
   7단계: 저장된 계산 목록 + 엑셀 내려받기
   (index.html의 계산 기능과 따로 관리하려고 분리한 파일이에요)
   ========================================= */

// 목록 전용 사물함 번호표 (입력값 사물함과 따로 써요)
const LIST_KEY = 'margin-calc-list';

// 저장 목록에 쓰는 입력칸 이름표들
const INPUT_IDS = ['cost', 'shipping', 'etc', 'fee', 'cs', 'ad', 'margin', 'adTotal'];

// 페이지를 열 때 사물함에서 저장 목록 꺼내기
let savedList = loadList();

function loadList() {
  try { return JSON.parse(localStorage.getItem(LIST_KEY)) || []; }
  catch (e) { return []; }
}

function storeList() {
  try { localStorage.setItem(LIST_KEY, JSON.stringify(savedList)); } catch (e) {}
}

// 목록 카드에 안내 문구 띄우기 (초록 = 완료, 빨강 = 문제)
function savedSay(text, type) {
  const box = document.getElementById('savedMsg');
  box.textContent = text;
  box.className = type;
  // 완료 문구(초록)만 3초 뒤 사라지고, 문제 문구(빨강)는 남겨둬요
  if (text && type === 'ok') {
    setTimeout(function () { if (box.textContent === text) box.textContent = ''; }, 3000);
  }
}

// ---------- [이 결과 저장] 버튼 ----------
function saveResult() {
  const nameBox = document.getElementById('itemName');
  const name = nameBox.value.trim();

  if (!lastResult) {                 // 계산이 안 된 상태
    savedSay('먼저 계산이 되도록 값을 입력해 주세요.', 'error');
    return;
  }
  if (name === '') {                 // 상품명이 비어 있음
    savedSay('상품명을 입력해 주세요.', 'error');
    nameBox.focus();
    return;
  }

  // 그때의 입력값 모으기 (나중에 다시 불러오려고)
  const inputValues = {};
  INPUT_IDS.forEach(function (id) { inputValues[id] = document.getElementById(id).value; });

  // 목록 맨 아래에 새 줄로 추가
  savedList.push({
    name: name,
    savedAt: nowText(),
    inputs: inputValues,
    result: lastResult
  });
  storeList();
  renderList();
  savedSay('"' + name + '" 저장했어요.', 'ok');
}

// ---------- 목록 화면 그리기 ----------
function renderList() {
  const body = document.getElementById('savedBody');
  body.innerHTML = '';  // 기존 줄 모두 지우고 다시 그리기

  savedList.forEach(function (item, i) {
    const tr = document.createElement('tr');

    // 상품명 (누르면 다시 불러오기)
    const nameCell = document.createElement('td');
    const nameBtn = document.createElement('button');
    nameBtn.className = 'name-link';
    nameBtn.textContent = item.name;
    nameBtn.onclick = function () { loadItem(i); };
    nameCell.appendChild(nameBtn);
    tr.appendChild(nameCell);

    // 숫자 칸 4개: 원가, 판매가, 순마진, 건당 광고비
    [Number(item.inputs.cost), item.result.price, item.result.profit, item.result.adCost]
      .forEach(function (n) {
        const td = document.createElement('td');
        td.className = 'num';
        td.textContent = Math.round(n).toLocaleString('ko-KR');
        tr.appendChild(td);
      });

    // 삭제 버튼 (두 번 눌러야 삭제)
    const delCell = document.createElement('td');
    const delBtn = document.createElement('button');
    delBtn.className = 'del-btn';
    delBtn.textContent = '🗑';
    delBtn.title = '삭제';
    delBtn.onclick = function () {
      twoTap(delBtn, '삭제?', function () {
        savedList.splice(i, 1);  // i번째 줄 빼기
        storeList();
        renderList();
      });
    };
    delCell.appendChild(delBtn);
    tr.appendChild(delCell);

    body.appendChild(tr);
  });

  // 개수 표시, 빈 목록 안내, 표 보이기/숨기기
  const has = savedList.length > 0;
  document.getElementById('savedCount').textContent = has ? '(' + savedList.length + '개)' : '';
  document.getElementById('savedTable').style.display = has ? 'table' : 'none';
  document.getElementById('savedHint').style.display  = has ? 'block' : 'none';
  document.getElementById('savedEmpty').style.display = has ? 'none' : 'block';
}

// ---------- 상품명 누르면 다시 불러오기 ----------
function loadItem(i) {
  const item = savedList[i];
  INPUT_IDS.forEach(function (id) { document.getElementById(id).value = item.inputs[id]; });
  document.getElementById('itemName').value = item.name;
  calculate();    // 다시 계산
  saveInputs();   // 입력값 사물함에도 반영
  window.scrollTo({ top: 0, behavior: 'smooth' });
  savedSay('"' + item.name + '" 입력값을 불러왔어요.', 'ok');
}

// ---------- [엑셀로 내려받기] 버튼 ----------
// 엑셀 제작 도구(xlsx-js-style)는 버튼을 누를 때만 인터넷에서 불러와요
function downloadExcel() {
  if (savedList.length === 0) {
    savedSay('저장된 계산이 없어요. 먼저 결과를 저장해 주세요.', 'error');
    return;
  }
  if (window.XLSX) { makeExcel(); return; }   // 이미 불러왔으면 바로 만들기

  savedSay('엑셀 도구를 불러오는 중이에요...', 'ok');
  const s = document.createElement('script');
  s.src = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';
  s.onload = makeExcel;
  s.onerror = function () {
    savedSay('엑셀 도구를 불러오지 못했어요. 인터넷 연결을 확인해 주세요.', 'error');
  };
  document.head.appendChild(s);
}

function makeExcel() {
  // 1) 제목 줄
  const header = [
    '저장일시', '상품명',
    '원가', '배송비', '기타비용',
    '수수료율(%)', 'CS비율(%)', '광고비율(%)', '목표마진율(%)',
    '추천 판매가', '건당 광고비', '순마진', 'VAT', '실제 남는 돈',
    '월 광고비', '적자 (개 이하)', '목표 달성 (개 이상)', '손익분기 ROAS(%)'
  ];

  // 2) 저장된 계산을 한 줄씩 표로 만들기
  const rows = savedList.map(function (item) {
    const v = item.inputs, r = item.result;
    const num = function (x) { return x === '' || x === undefined ? 0 : Number(x); };
    const opt = function (x) { return x === undefined || x === null ? '' : x; }; // 없으면 빈칸
    return [
      item.savedAt, item.name,
      num(v.cost), num(v.shipping), num(v.etc),
      num(v.fee), num(v.cs), num(v.ad), num(v.margin),
      Math.round(r.price), Math.round(r.adCost), Math.round(r.profit),
      Math.round(r.vat), Math.round(r.afterVat),
      r.adTotal ? r.adTotal : '', opt(r.lossMax), opt(r.goalMin), opt(r.roas)
    ];
  });

  const sheet = XLSX.utils.aoa_to_sheet([header].concat(rows));

  // 3) 꾸미기: 제목 줄 남색 바탕 + 흰 글씨, 금액은 천 단위 쉼표, 구간은 색 글씨
  const moneyCols = [2, 3, 4, 9, 10, 11, 12, 13, 14];  // 금액 칸 번호 (0부터 셈)
  const range = XLSX.utils.decode_range(sheet['!ref']);
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: R, c: C })];
      if (!cell) continue;
      if (R === 0) {
        cell.s = {
          font: { bold: true, color: { rgb: 'FFFFFFFF' } },
          fill: { fgColor: { rgb: 'FF1F3A5F' } },
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
        };
        continue;
      }
      if (moneyCols.indexOf(C) !== -1) cell.z = '#,##0';
      if (C === 9)  cell.s = { font: { bold: true, color: { rgb: 'FF1F3A5F' } } }; // 추천 판매가 강조
      if (C === 15) cell.s = { font: { color: { rgb: 'FFC62828' } } };             // 적자: 빨강
      if (C === 16) cell.s = { font: { color: { rgb: 'FF2E7D32' } } };             // 목표 달성: 초록
    }
  }

  // 4) 칸 너비
  sheet['!cols'] = [
    { wch: 17 }, { wch: 14 },
    { wch: 10 }, { wch: 9 }, { wch: 9 },
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 9 },
    { wch: 11 }, { wch: 10 }, { wch: 10 }, { wch: 9 }, { wch: 11 },
    { wch: 12 }, { wch: 10 }, { wch: 11 }, { wch: 10 }
  ];

  // 5) 파일로 내려받기 (파일 이름에 오늘 날짜)
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, '마진계산');
  XLSX.writeFile(book, '마진계산_' + nowText().slice(0, 10) + '.xlsx');
  savedSay('엑셀 파일을 내려받았어요.', 'ok');
}

// "2026-11-07 15:10" 같은 현재 날짜·시간 글자 만들기
function nowText() {
  const d = new Date();
  const two = function (n) { return String(n).padStart(2, '0'); };
  return d.getFullYear() + '-' + two(d.getMonth() + 1) + '-' + two(d.getDate()) +
         ' ' + two(d.getHours()) + ':' + two(d.getMinutes());
}

// 페이지를 열 때 저장 목록 그리기
renderList();
