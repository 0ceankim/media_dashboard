// shared_process.js - 두 대시보드가 공유하는 전처리 로직
// 이 파일을 두 HTML에서 모두 참조

const EXCLUDE_PATTERNS = ['(알파)', '(원스, 데브)', '(데브)', '(원스, 알파)'];

// === 공통 유틸 ===
function convertDate(dateStr) {
  const match = String(dateStr).match(/(\d{4})-(\d{2})-(\d{2})\s*\((.)\)/);
  if (match) return parseInt(match[2]) + '/' + parseInt(match[3]) + '(' + match[4] + ')';
  return String(dateStr);
}

function identifyWeeks(dates) {
  const weeks = [];
  let weekStart = 0;
  for (let i = 0; i < dates.length; i++) {
    const match = dates[i].match(/\((.)\)/);
    if (match && match[1] === '목' && i > 0) {
      weeks.push({ start: weekStart, end: i - 1 });
      weekStart = i;
    }
  }
  weeks.push({ start: weekStart, end: dates.length - 1 });
  return weeks;
}

function getWeekLabel(dates, week) {
  const start = dates[week.start].replace(/\(.\)/, '').trim();
  const end = dates[week.end].replace(/\(.\)/, '').trim();
  return start + ' ~ ' + end;
}

function isTargetMedia(name) {
  return name.includes('한게임') || name.includes('윈조이') || name.includes('WPL');
}

function cleanMediaName(name) {
  name = name.replace(/\s*\(\d+\)\s*$/, '');
  if (name.includes('윈조이 포커')) name = '윈조이 포커';
  if (name.startsWith('WPL')) name = 'WPL';
  const match = name.match(/^(.+?)\s*[:：]\s*.+?(\(원스토어\))?\s*$/);
  if (match) name = match[1].trim() + (match[2] ? ' ' + match[2] : '');
  return name.trim();
}

function getCompany(name) {
  return name.includes('한게임') ? 'NHN 한게임' : '잼팟 주식회사';
}

// === 엑셀 파싱 (원본 데이터 추출) ===
function parseExcelRaw(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const header = [...raw[0]];
  const dataRows = raw.slice(1).map(r => [...r]);
  
  // 연동 형태 제거 (인덱스 2)
  header.splice(2, 1);
  dataRows.forEach(r => r.splice(2, 1));
  
  // 날짜 변환
  const dates = [];
  for (let i = 2; i < header.length - 1; i++) {
    header[i] = convertDate(header[i]);
    dates.push(header[i]);
  }
  
  // 매체 데이터 추출 (전체)
  const allMedia = [];
  dataRows.forEach(row => {
    const name = String(row[0] || '');
    const os = String(row[1] || '');
    if (!name || name === '평균' || name === '합계') return;
    const daily = [];
    for (let i = 2; i < 2 + dates.length; i++) {
      daily.push(Number(row[i]) || 0);
    }
    if (daily.reduce((a, b) => a + b, 0) > 0) {
      allMedia.push({ name, os, daily });
    }
  });
  
  return { dates, allMedia };
}

// === 매체 대시보드용 전처리 (전체 기간) ===
function processForAllMedia(rawData) {
  const { dates, allMedia } = rawData;
  const weeks = identifyWeeks(dates);
  
  // 마지막 2주 비교 (WoW용)
  const lastWeek = weeks[weeks.length - 1];
  const prevWeek = weeks.length >= 2 ? weeks[weeks.length - 2] : null;
  
  const mediaResults = allMedia.map(m => {
    const w1 = prevWeek ? m.daily.slice(prevWeek.start, prevWeek.end + 1).reduce((a, b) => a + b, 0) : 0;
    const w2 = m.daily.slice(lastWeek.start, lastWeek.end + 1).reduce((a, b) => a + b, 0);
    const total = w1 + w2;
    const pct = w1 > 0 ? Math.round((w2 - w1) / w1 * 1000) / 10 : (w2 > 0 ? 999.9 : 0);
    return { name: m.name, os: m.os, w1, w2, total, diff: w2 - w1, pct };
  }).filter(m => m.total > 0);
  
  const totalW1 = mediaResults.reduce((a, m) => a + m.w1, 0);
  const totalW2 = mediaResults.reduce((a, m) => a + m.w2, 0);
  
  const sortedByDiff = [...mediaResults].sort((a, b) => a.diff - b.diff);
  const pctUpPool = mediaResults.filter(m => m.w1 >= 50000 && m.diff >= 50000);
  const pctDownPool = mediaResults.filter(m => m.w1 >= 100000 && m.diff <= -50000);
  
  return {
    meta: {
      updatedAt: new Date().toISOString().split('T')[0],
      period: dates[0] + ' ~ ' + dates[dates.length - 1],
      mediaCount: mediaResults.length,
      w1Label: prevWeek ? getWeekLabel(dates, prevWeek) : '-',
      w2Label: getWeekLabel(dates, lastWeek)
    },
    total: { w1: totalW1, w2: totalW2 },
    topDiffUp: sortedByDiff.slice(-10).reverse(),
    topDiffDown: sortedByDiff.slice(0, 10),
    topPctUp: [...pctUpPool].sort((a, b) => b.pct - a.pct).slice(0, 10),
    topPctDown: [...pctDownPool].sort((a, b) => a.pct - b.pct).slice(0, 10),
    allMedia: [...mediaResults].sort((a, b) => b.total - a.total).slice(0, 30)
  };
}

// === 고포류 매체 대시보드용 전처리 (최근 2주만) ===
function processForGophoryu(rawData) {
  const { dates, allMedia } = rawData;
  const weeks = identifyWeeks(dates);
  
  // 최근 2주만 추출
  let last2WeeksStart, last2WeeksEnd;
  if (weeks.length >= 2) {
    last2WeeksStart = weeks[weeks.length - 2].start;
    last2WeeksEnd = weeks[weeks.length - 1].end;
  } else {
    last2WeeksStart = 0;
    last2WeeksEnd = dates.length - 1;
  }
  
  const slicedDates = dates.slice(last2WeeksStart, last2WeeksEnd + 1);
  const slicedWeeks = identifyWeeks(slicedDates);
  
  // 고포류 매체 필터링 + 회사별 합산
  const companyDaily = {};
  allMedia.forEach(m => {
    if (!isTargetMedia(m.name)) return;
    if (EXCLUDE_PATTERNS.some(p => m.name.includes(p))) return;
    const cleaned = cleanMediaName(m.name);
    const company = getCompany(cleaned);
    if (!companyDaily[company]) companyDaily[company] = new Array(slicedDates.length).fill(0);
    for (let i = 0; i < slicedDates.length; i++) {
      companyDaily[company][i] += m.daily[last2WeeksStart + i];
    }
  });
  
  // 개별 매체 리스트 (회사별 합산이 아닌 매체별)
  const lastWeek = slicedWeeks[slicedWeeks.length - 1];
  const prevWeek = slicedWeeks.length >= 2 ? slicedWeeks[slicedWeeks.length - 2] : null;
  const mediaDetail = [];
  allMedia.forEach(m => {
    if (!isTargetMedia(m.name)) return;
    if (EXCLUDE_PATTERNS.some(p => m.name.includes(p))) return;
    const daily = [];
    for (let i = 0; i < slicedDates.length; i++) {
      daily.push(m.daily[last2WeeksStart + i]);
    }
    const w1 = prevWeek ? daily.slice(prevWeek.start, prevWeek.end + 1).reduce((a, b) => a + b, 0) : 0;
    const w2 = daily.slice(lastWeek.start, lastWeek.end + 1).reduce((a, b) => a + b, 0);
    const total = w1 + w2;
    if (total > 0) {
      const pct = w1 > 0 ? Math.round((w2 - w1) / w1 * 1000) / 10 : (w2 > 0 ? 999.9 : 0);
      mediaDetail.push({ company: getCompany(m.name), name: cleanMediaName(m.name), os: m.os, w1, w2, total, diff: w2 - w1, pct });
    }
  });
  mediaDetail.sort((a, b) => b.total - a.total);

  return {
    meta: {
      updatedAt: new Date().toISOString().split('T')[0],
      period: slicedDates[0] + ' ~ ' + slicedDates[slicedDates.length - 1],
      weekCount: slicedWeeks.length
    },
    dates: slicedDates,
    companies: companyDaily,
    weeks: slicedWeeks.map(w => ({ label: getWeekLabel(slicedDates, w), start: w.start, end: w.end })),
    mediaDetail: mediaDetail
  };
}

// === localStorage 공유 ===
function saveRawData(rawData, fileName) {
  const toSave = { rawData, fileName, savedAt: new Date().toISOString() };
  localStorage.setItem('media_dashboard_raw', JSON.stringify(toSave));
}

function loadRawData() {
  const saved = localStorage.getItem('media_dashboard_raw');
  if (!saved) return null;
  return JSON.parse(saved);
}
