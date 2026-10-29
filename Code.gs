/**
 * [나중에 다시 보고싶은 영상들] Google Apps Script (단일 파일 올인원 버전)
 * 스프레드시트 ID: YOUR_SPREADSHEET_ID
 */

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID';

function getTargetSheet() {
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return ss.getSheets()[0];
}

function doGet(e) {
  const htmlContent = getHtmlContent();
  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle('나중에 다시 보고싶은 영상들')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function getCategories() {
  const defaultCategories = ['개발', '운동', '경제', '엔터테인먼트', '덕질', '음악'];
  try {
    const sheet = getTargetSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 4) {
      const bValues = sheet.getRange(4, 2, lastRow - 3, 1).getValues();
      const sheetCategories = bValues
        .map(row => (row[0] ? String(row[0]).trim() : ''))
        .filter(val => val.length > 0);
      return [...new Set([...defaultCategories, ...sheetCategories])];
    }
  } catch (err) {
    console.error('카테고리 조회 중 오류:', err);
  }
  return defaultCategories;
}

function saveLink(category, link) {
  if (!category || !String(category).trim()) {
    throw new Error('카테고리를 선택하거나 입력해주세요.');
  }
  if (!link || !String(link).trim()) {
    throw new Error('링크(URL)를 입력해주세요.');
  }

  const cleanCategory = String(category).trim();
  const cleanLink = String(link).trim();
  const sheet = getTargetSheet();
  const maxRows = sheet.getMaxRows();
  
  const scanLimit = Math.min(maxRows, 500);
  const numRowsToScan = scanLimit - 3;
  
  let targetRow = null;
  let firstEmptyRow = null;

  if (numRowsToScan > 0) {
    const values = sheet.getRange(4, 1, numRowsToScan, 3).getValues();

    for (let i = 0; i < values.length; i++) {
      const currentRow = 4 + i;
      const rowCategory = values[i][1] ? String(values[i][1]).trim() : '';
      const rowLink = values[i][2] ? String(values[i][2]).trim() : '';

      // 1순위: 카테고리가 같고 링크가 비어있는 행
      if (rowCategory === cleanCategory && rowLink === '') {
        targetRow = currentRow;
        break;
      }

      // 2순위: 카테고리와 링크 모두 비어있는 첫 행
      if (firstEmptyRow === null && rowCategory === '' && rowLink === '') {
        firstEmptyRow = currentRow;
      }
    }
  }

  if (!targetRow && firstEmptyRow) {
    targetRow = firstEmptyRow;
  }
  if (!targetRow) {
    targetRow = sheet.getLastRow() + 1;
  }

  // A열 체크박스
  const checkboxCell = sheet.getRange(targetRow, 1);
  checkboxCell.insertCheckboxes();
  checkboxCell.setValue(false);

  // B열 카테고리, C열 링크
  sheet.getRange(targetRow, 2).setValue(cleanCategory);
  sheet.getRange(targetRow, 3).setValue(cleanLink);

  return {
    success: true,
    message: `[${cleanCategory}] 링크가 성공적으로 저장되었습니다!`,
    row: targetRow
  };
}

function getHtmlContent() {
  return `<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="default">
  <title>나중에 볼 영상 저장</title>
  <script src="https://cdn.tailwindcss.com"><\/script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;500;600;700&display=swap');
    body {
      font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
      -webkit-tap-highlight-color: transparent;
      user-select: none;
    }
    input, select, textarea {
      user-select: text;
    }
    .custom-scrollbar::-webkit-scrollbar {
      display: none;
    }
    .custom-scrollbar {
      -ms-overflow-style: none;
      scrollbar-width: none;
    }
    * {
      -webkit-overflow-scrolling: touch;
    }
  </style>
</head>
<body class="bg-slate-100 min-h-screen text-slate-800 flex flex-col justify-between p-4 sm:p-6">

  <div class="max-w-md w-full mx-auto space-y-4">
    
    <!-- 상단 헤더 -->
    <header class="bg-gradient-to-r from-emerald-600 to-teal-700 text-white p-5 rounded-2xl shadow-lg shadow-emerald-700/20 relative overflow-hidden">
      <div class="relative z-10">
        <div class="flex items-center justify-between">
          <div class="flex items-center space-x-2.5">
            <div class="w-10 h-10 bg-white/20 backdrop-blur-md rounded-xl flex items-center justify-center text-xl shadow-inner">
              <i class="fa-solid fa-play text-emerald-100 ml-0.5"></i>
            </div>
            <div>
              <h1 class="text-lg font-bold tracking-tight">나중에 볼 영상</h1>
              <p class="text-xs text-emerald-100/90 font-medium">구글 시트 자동 저장</p>
            </div>
          </div>
          <a href="https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit" 
             target="_blank" 
             class="px-2.5 py-1.5 bg-white/15 hover:bg-white/25 active:scale-95 transition-all text-xs rounded-lg flex items-center gap-1.5 border border-white/20">
            <i class="fa-solid fa-table-cells"></i>
            <span>시트 열기</span>
          </a>
        </div>
      </div>
      <div class="absolute -right-6 -bottom-8 w-28 h-28 bg-white/10 rounded-full blur-xl pointer-events-none"></div>
    </header>

    <!-- 입력 폼 카드 -->
    <main class="bg-white rounded-2xl shadow-sm border border-slate-200/80 p-5 space-y-5">
      
      <!-- 1. 카테고리 영역 -->
      <div class="space-y-2.5">
        <div class="flex items-center justify-between">
          <label class="text-sm font-semibold text-slate-700 flex items-center gap-1.5">
            <i class="fa-solid fa-folder-open text-emerald-600"></i>
            <span>카테고리 선택</span>
          </label>
          <button type="button" id="toggleDirectInputBtn" class="text-xs font-semibold text-emerald-600 hover:text-emerald-700 active:scale-95 transition flex items-center gap-1">
            <i class="fa-solid fa-plus-circle"></i>
            <span id="toggleDirectText">직접 추가하기</span>
          </button>
        </div>

        <!-- 빠른 선택 칩 -->
        <div id="categoryChips" class="flex gap-2 overflow-x-auto pb-1.5 custom-scrollbar">
        </div>

        <!-- 직접 입력 인풋 -->
        <div id="customCategoryContainer" class="hidden pt-1">
          <div class="relative">
            <input type="text" id="customCategoryInput" 
                   placeholder="새 카테고리 이름 입력..." 
                   class="w-full px-3.5 py-2.5 bg-slate-50 border border-emerald-400 focus:border-emerald-600 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 transition">
            <button type="button" id="confirmNewCategoryBtn" class="absolute right-2 top-2 px-2.5 py-1 bg-emerald-600 text-white text-xs font-medium rounded-lg hover:bg-emerald-700 active:scale-95 transition">
              추가
            </button>
          </div>
        </div>
      </div>

      <!-- 2. 링크 붙여넣기 영역 -->
      <div class="space-y-2">
        <div class="flex items-center justify-between">
          <label for="linkInput" class="text-sm font-semibold text-slate-700 flex items-center gap-1.5">
            <i class="fa-solid fa-link text-emerald-600"></i>
            <span>링크 (URL)</span>
          </label>
          <button type="button" id="pasteBtn" class="text-xs font-semibold px-2.5 py-1 bg-emerald-50 text-emerald-700 hover:bg-emerald-100 active:scale-95 border border-emerald-200 rounded-lg flex items-center gap-1 transition">
            <i class="fa-regular fa-clipboard"></i>
            <span>클립보드 붙여넣기</span>
          </button>
        </div>
        
        <div class="relative">
          <input type="url" id="linkInput" 
                 inputmode="url"
                 placeholder="https://... 영상 링크 붙여넣기" 
                 class="w-full pl-3.5 pr-9 py-3 bg-slate-50 border border-slate-300 focus:border-emerald-600 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 transition">
          <button type="button" id="clearLinkBtn" class="hidden absolute right-3 top-3.5 text-slate-400 hover:text-slate-600 text-sm">
            <i class="fa-solid fa-circle-xmark"></i>
          </button>
        </div>
      </div>

      <!-- 3. 저장 버튼 -->
      <div class="pt-2">
        <button type="button" id="saveBtn" class="w-full py-3.5 px-4 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white font-bold rounded-xl shadow-md shadow-emerald-600/30 flex items-center justify-center gap-2 transition active:scale-[0.99] disabled:opacity-50 disabled:pointer-events-none">
          <i class="fa-solid fa-check text-lg"></i>
          <span id="saveBtnText" class="text-base tracking-wide">시트에 저장하기</span>
        </button>
      </div>

    </main>

    <!-- 최근 저장 카드 -->
    <div id="recentSaveCard" class="hidden bg-emerald-50 border border-emerald-200 rounded-xl p-3.5 transition-all duration-300">
      <div class="flex items-start gap-2.5">
        <div class="w-6 h-6 rounded-full bg-emerald-500 text-white flex items-center justify-center flex-shrink-0 mt-0.5 text-xs">
          <i class="fa-solid fa-check"></i>
        </div>
        <div class="text-xs space-y-0.5 overflow-hidden">
          <p class="font-bold text-emerald-900" id="recentCategory">저장 완료</p>
          <p class="text-emerald-700 truncate" id="recentLink">링크</p>
        </div>
      </div>
    </div>

    <footer class="text-center text-xs text-slate-400 py-1">
      <p>💡 홈 화면에 바로가기를 추가하면 앱처럼 쓸 수 있습니다.</p>
    </footer>

  </div>

  <!-- 토스트 알림창 -->
  <div id="toast" class="fixed top-5 left-1/2 -translate-x-1/2 z-50 transform -translate-y-20 opacity-0 transition-all duration-300 pointer-events-none max-w-xs w-full px-4">
    <div class="bg-slate-900/90 backdrop-blur text-white text-xs font-medium py-3 px-4 rounded-xl shadow-xl flex items-center gap-2.5 border border-slate-700">
      <i id="toastIcon" class="fa-solid fa-circle-check text-emerald-400 text-sm"></i>
      <span id="toastMsg">저장되었습니다!</span>
    </div>
  </div>

  <script>
    let categories = ['개발', '운동', '경제', '엔터테인먼트', '덕질', '음악'];
    let selectedCategory = '개발';
    let isDirectInputMode = false;

    const categoryChips = document.getElementById('categoryChips');
    const toggleDirectInputBtn = document.getElementById('toggleDirectInputBtn');
    const toggleDirectText = document.getElementById('toggleDirectText');
    const customCategoryContainer = document.getElementById('customCategoryContainer');
    const customCategoryInput = document.getElementById('customCategoryInput');
    const confirmNewCategoryBtn = document.getElementById('confirmNewCategoryBtn');
    const linkInput = document.getElementById('linkInput');
    const clearLinkBtn = document.getElementById('clearLinkBtn');
    const pasteBtn = document.getElementById('pasteBtn');
    const saveBtn = document.getElementById('saveBtn');
    const saveBtnText = document.getElementById('saveBtnText');
    const toast = document.getElementById('toast');
    const toastMsg = document.getElementById('toastMsg');
    const toastIcon = document.getElementById('toastIcon');
    const recentSaveCard = document.getElementById('recentSaveCard');
    const recentCategory = document.getElementById('recentCategory');
    const recentLink = document.getElementById('recentLink');

    window.addEventListener('DOMContentLoaded', () => {
      renderCategoryChips();
      fetchCategories();
      setupEventListeners();
    });

    function renderCategoryChips() {
      categoryChips.innerHTML = '';
      categories.forEach(cat => {
        const isSelected = (cat === selectedCategory && !isDirectInputMode);
        const chip = document.createElement('button');
        chip.type = 'button';
        chip.className = 'px-3.5 py-1.5 rounded-full text-xs font-semibold whitespace-nowrap transition-all active:scale-95 ' +
          (isSelected 
            ? 'bg-emerald-600 text-white shadow-sm shadow-emerald-600/30' 
            : 'bg-slate-100 text-slate-600 hover:bg-slate-200');
        chip.textContent = cat;
        chip.onclick = () => {
          selectedCategory = cat;
          isDirectInputMode = false;
          customCategoryContainer.classList.add('hidden');
          toggleDirectText.textContent = '직접 추가하기';
          renderCategoryChips();
        };
        categoryChips.appendChild(chip);
      });
    }

    function fetchCategories() {
      if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler((serverCategories) => {
            if (Array.isArray(serverCategories) && serverCategories.length > 0) {
              categories = serverCategories;
              if (!categories.includes(selectedCategory)) {
                selectedCategory = categories[0];
              }
              renderCategoryChips();
            }
          })
          .getCategories();
      }
    }

    function setupEventListeners() {
      toggleDirectInputBtn.addEventListener('click', () => {
        isDirectInputMode = !isDirectInputMode;
        if (isDirectInputMode) {
          customCategoryContainer.classList.remove('hidden');
          customCategoryInput.focus();
          toggleDirectText.textContent = '목록에서 선택';
        } else {
          customCategoryContainer.classList.add('hidden');
          toggleDirectText.textContent = '직접 추가하기';
        }
        renderCategoryChips();
      });

      confirmNewCategoryBtn.addEventListener('click', addCustomCategory);
      customCategoryInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          addCustomCategory();
        }
      });

      pasteBtn.addEventListener('click', async () => {
        try {
          if (navigator.clipboard && navigator.clipboard.readText) {
            const text = await navigator.clipboard.readText();
            if (text && text.trim()) {
              linkInput.value = text.trim();
              clearLinkBtn.classList.remove('hidden');
              showToast('링크가 붙여넣어졌습니다!', 'success');
              if (navigator.vibrate) navigator.vibrate(30);
              return;
            }
          }
          showToast('클립보드에 복사된 내용이 없습니다.', 'info');
        } catch (err) {
          showToast('입력창을 길게 눌러 붙여넣기 해주세요.', 'info');
          linkInput.focus();
        }
      });

      linkInput.addEventListener('input', () => {
        if (linkInput.value.length > 0) {
          clearLinkBtn.classList.remove('hidden');
        } else {
          clearLinkBtn.classList.add('hidden');
        }
      });

      clearLinkBtn.addEventListener('click', () => {
        linkInput.value = '';
        clearLinkBtn.classList.add('hidden');
        linkInput.focus();
      });

      saveBtn.addEventListener('click', handleSave);
      linkInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          handleSave();
        }
      });
    }

    function addCustomCategory() {
      const val = customCategoryInput.value.trim();
      if (!val) {
        showToast('카테고리 이름을 입력해주세요.', 'info');
        customCategoryInput.focus();
        return;
      }
      if (!categories.includes(val)) {
        categories.push(val);
      }
      selectedCategory = val;
      isDirectInputMode = false;
      customCategoryInput.value = '';
      customCategoryContainer.classList.add('hidden');
      toggleDirectText.textContent = '직접 추가하기';
      renderCategoryChips();
      showToast("'" + val + "' 카테고리 선택됨", 'success');
    }

    function handleSave() {
      let finalCategory = selectedCategory;

      if (isDirectInputMode) {
        const customVal = customCategoryInput.value.trim();
        if (customVal) {
          finalCategory = customVal;
          if (!categories.includes(finalCategory)) {
            categories.push(finalCategory);
          }
        } else {
          showToast('카테고리 이름을 입력해주세요.', 'info');
          customCategoryInput.focus();
          return;
        }
      }

      const link = linkInput.value.trim();
      if (!link) {
        showToast('저장할 링크(URL)를 입력해주세요.', 'info');
        linkInput.focus();
        return;
      }

      setLoading(true);

      if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler((response) => {
            setLoading(false);
            onSaveSuccess(finalCategory, link);
          })
          .withFailureHandler((error) => {
            setLoading(false);
            showToast('저장 실패: ' + (error.message || '네트워크 오류'), 'error');
          })
          .saveLink(finalCategory, link);
      } else {
        setTimeout(() => {
          setLoading(false);
          onSaveSuccess(finalCategory, link);
        }, 500);
      }
    }

    function onSaveSuccess(cat, link) {
      if (navigator.vibrate) {
        navigator.vibrate([40, 60, 40]);
      }
      showToast('시트에 성공적으로 저장되었습니다! 🎉', 'success');

      recentCategory.textContent = '[' + cat + '] 방금 저장됨';
      recentLink.textContent = link;
      recentSaveCard.classList.remove('hidden');

      linkInput.value = '';
      clearLinkBtn.classList.add('hidden');
      if (isDirectInputMode) {
        isDirectInputMode = false;
        customCategoryContainer.classList.add('hidden');
        toggleDirectText.textContent = '직접 추가하기';
      }
      selectedCategory = cat;
      renderCategoryChips();
    }

    function setLoading(isLoading) {
      saveBtn.disabled = isLoading;
      if (isLoading) {
        saveBtnText.textContent = '저장 중...';
        saveBtn.querySelector('i').className = 'fa-solid fa-spinner fa-spin text-lg';
      } else {
        saveBtnText.textContent = '시트에 저장하기';
        saveBtn.querySelector('i').className = 'fa-solid fa-check text-lg';
      }
    }

    let toastTimeout;
    function showToast(message, type = 'success') {
      clearTimeout(toastTimeout);
      toastMsg.textContent = message;
      
      if (type === 'success') {
        toastIcon.className = 'fa-solid fa-circle-check text-emerald-400 text-sm';
      } else if (type === 'error') {
        toastIcon.className = 'fa-solid fa-triangle-exclamation text-rose-400 text-sm';
      } else {
        toastIcon.className = 'fa-solid fa-circle-info text-sky-400 text-sm';
      }

      toast.classList.remove('-translate-y-20', 'opacity-0');
      toast.classList.add('translate-y-0', 'opacity-100');

      toastTimeout = setTimeout(() => {
        toast.classList.remove('translate-y-0', 'opacity-100');
        toast.classList.add('-translate-y-20', 'opacity-0');
      }, 2500);
    }
  </script>
</body>
</html>`;
}
