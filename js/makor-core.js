
    const defaultConfig = {
      page_title: t('projectTitle'),
      default_word: t('defaultWord'),
      primary_color: '#3b82f6',
      secondary_color: '#8b5cf6',
      text_color: '#1f2937',
      background_color: '#f0f4ff',
      button_color: '#3b82f6',
      font_family: 'Arial',
      font_size: 16
    };

    // Page orientation: 'portrait' or 'landscape'
    let pageOrientation = 'portrait';

    let stickers = [];
    let selectedSticker = null;
    let selectedWord = null;
    let selectedImage = null;
    let draggedElement = null;
    let resizingImage = null;
    let resizingSticker = null;
    let resizingWord = null;
    let undoStack = [];
    let redoStack = [];
    let isRestoringHistory = false;
    let offsetX = 0;
    let offsetY = 0;
    let wordIdCounter = 0;
    let wordSeriesCounter = 0;
    let imageIdCounter = 0;
    let imageSeriesCounter = 0;
    let syncMoveEnabled = false;
    let syncDeleteEnabled = false;
    let autoArrangeEnabled = true;
    let initialDragPosition = null;
    let currentProjectFileName = null;
    const EXPORT_QUALITY = { pdfScale: 4, imageScale: 5, zipMaxBytes: 15 * 1024 * 1024, jpegQuality: 0.98, pdfCompression: 'NONE', pdfDpi: 360 };

    const MM_TO_PX = 3.7795275591;

    let stickerLayoutConfig = {
      uploadLimit: 0,
      stickersPerRow: 2,
      edgeMargin: 1,
      gap: 1,
      sizeMode: 'width'
    };
    
    // Numbers mode variables
    let currentMode = 'words'; // 'words' or 'numbers' or 'lottery'
    let numberedStickers = [];
    let singleStickerTemplate = null;
    let numberDragStart = null;
    
    // Lottery mode variables
    let lotteryNumbers = [];
    
    // Elements library
    let elementsLibrary = [];
    
    // GitHub repository configuration
    const GITHUB_REPO = {
      stickers: 'https://raw.githubusercontent.com/yoelyoel111/automatic-fishstick/main/%D7%94%D7%92%D7%A8%D7%95%D7%9C%D7%9E%D7%98%20%D7%AA%D7%9E%D7%95%D7%A0%D7%95%D7%AA/%D7%9E%D7%93%D7%91%D7%A7%D7%95%D7%AA/',
      elements: 'https://raw.githubusercontent.com/yoelyoel111/automatic-fishstick/main/%D7%94%D7%92%D7%A8%D7%95%D7%9C%D7%9E%D7%98%20%D7%AA%D7%9E%D7%95%D7%A0%D7%95%D7%AA/%D7%90%D7%9C%D7%9E%D7%A0%D7%98%D7%99%D7%9D/'
    };

    function focusWordInput() {
      if (typeof setActiveToolsSection === 'function') {
        setActiveToolsSection('Text');
      }
      const wordInput = document.getElementById('wordInput');
      if (!wordInput) return;
      setTimeout(() => {
        try {
          wordInput.focus({ preventScroll: true });
          const end = (wordInput.value || '').length;
          if (typeof wordInput.setSelectionRange === 'function') {
            wordInput.setSelectionRange(end, end);
          }
        } catch (_) {}
      }, 0);
    }
    
    // Available files in GitHub repo - organized by categories
    const GITHUB_FILES = {
      stickers: {
        'מלבנים': ['35.png', '36.png', '37.png', '40.png', 'שבלונה להזנת מספרים.png'],
        'מקומות קדושים': ['41.png', '42.png', '43.png', '44.png', 'ארון_הקודש_בישיבת_פוניבז.jpg', 'עטרתשלמה.png'],
        'נופים': ['50.png', '51.png', '52.png', '53.png', '54.png', '55.png', '56.png'],
        'עיגולים': ['16.png', '26.png', '27.png', '28.png', '29.png', '30.png', '31.png', '32.png'],
        'ריבועים': ['57.png', '58.png', '59.png', '60.png', '61.png']
      },
      elements: {
        'מאכלים': ['17.png', '18.png', '19.png', '20.png', '21.png', '22.png', '23.png', '24.png', '25.png'],
        'ספרים': ['מקראות גדולות.png', 'משנהברורה.png', 'שספנינים.png'],
        'פרסים': ['35.png', '36.png', '37.png', '38.png'],
        'רבנים': ['45.png', '46.png', '47.png']
      }
    };

    const categoryTranslations = {
      'מלבנים': 'catRectangles',
      'מקומות קדושים': 'catHolyPlaces',
      'נופים': 'catLandscapes',
      'עיגולים': 'catCircles',
      'ריבועים': 'catSquares',
      'מאכלים': 'catFood',
      'ספרים': 'catBooks',
      'פרסים': 'catPrizes',
      'רבנים': 'catRabbis'
    };

    function cloneHistoryState() {
      return {
        stickers: JSON.parse(JSON.stringify(stickers)),
        stickerLayoutConfig: JSON.parse(JSON.stringify(stickerLayoutConfig)),
        pageOrientation,
        selectedSticker,
        selectedWord,
        selectedImage,
        syncMoveEnabled,
        syncDeleteEnabled,
        autoArrangeEnabled,
        wordIdCounter,
        wordSeriesCounter,
        imageIdCounter,
        imageSeriesCounter
      };
    }

    function updateUndoRedoButtons() {
      const undoBtn = document.getElementById('undoBtn');
      const redoBtn = document.getElementById('redoBtn');

      if (undoBtn) {
        const disabled = undoStack.length === 0;
        undoBtn.disabled = disabled;
        undoBtn.classList.toggle('opacity-50', disabled);
        undoBtn.classList.toggle('cursor-not-allowed', disabled);
      }

      if (redoBtn) {
        const disabled = redoStack.length === 0;
        redoBtn.disabled = disabled;
        redoBtn.classList.toggle('opacity-50', disabled);
        redoBtn.classList.toggle('cursor-not-allowed', disabled);
      }
    }

    function pushHistory() {
      if (isRestoringHistory) return;
      redoStack = [];
      undoStack.push(cloneHistoryState());
      if (undoStack.length > 20) undoStack.shift();
      updateUndoRedoButtons();
      
      // Trigger auto-save
      if (typeof GoogleDriveManager !== 'undefined' && GoogleDriveManager.markUnsavedChanges) {
        GoogleDriveManager.markUnsavedChanges();
      }
    }

    function restoreFromHistoryState(state) {
      if (!state) return;

      stickers = Array.isArray(state.stickers) ? state.stickers : [];
      stickerLayoutConfig = state.stickerLayoutConfig || stickerLayoutConfig;
      
      // Restore page orientation
      if (state.pageOrientation === 'landscape' || state.pageOrientation === 'portrait') {
        pageOrientation = state.pageOrientation;
        // Update dropdown display
        const orientationDropdownIcon = document.getElementById('orientationDropdownIcon');
        const orientationDropdownText = document.getElementById('orientationDropdownText');
        if (orientationDropdownIcon && orientationDropdownText) {
          if (pageOrientation === 'portrait') {
            orientationDropdownIcon.textContent = '📄';
            orientationDropdownText.textContent = 'כיוון הדף: לאורך';
          } else {
            orientationDropdownIcon.textContent = '📃';
            orientationDropdownText.textContent = 'כיוון הדף: לרוחב';
          }
        }
      }
      
      selectedSticker = (state.selectedSticker === null || Number.isFinite(state.selectedSticker)) ? state.selectedSticker : null;
      selectedWord = state.selectedWord ?? null;
      selectedImage = state.selectedImage ?? null;
      syncMoveEnabled = !!state.syncMoveEnabled;
      syncDeleteEnabled = !!state.syncDeleteEnabled;
      autoArrangeEnabled = !!state.autoArrangeEnabled;
      wordIdCounter = Number.isFinite(state.wordIdCounter) ? state.wordIdCounter : wordIdCounter;
      wordSeriesCounter = Number.isFinite(state.wordSeriesCounter) ? state.wordSeriesCounter : wordSeriesCounter;
      imageIdCounter = Number.isFinite(state.imageIdCounter) ? state.imageIdCounter : imageIdCounter;
      imageSeriesCounter = Number.isFinite(state.imageSeriesCounter) ? state.imageSeriesCounter : imageSeriesCounter;

      applyStickerLayoutConfigToUI();
      renderStickers();
      updateFileCount();
      updateUndoRedoButtons();
    }

    function undoLastAction() {
      if (undoStack.length === 0) {
        showStatus(t('noActionsToUndo'), true);
        return;
      }

      const current = cloneHistoryState();
      const prev = undoStack.pop();
      redoStack.push(current);
      if (redoStack.length > 20) redoStack.shift();

      isRestoringHistory = true;
      restoreFromHistoryState(prev);
      isRestoringHistory = false;
    }

    function redoLastAction() {
      if (redoStack.length === 0) {
        showStatus(t('noActionsToRedo'), true);
        return;
      }

      const current = cloneHistoryState();
      const next = redoStack.pop();
      undoStack.push(current);
      if (undoStack.length > 20) undoStack.shift();

      isRestoringHistory = true;
      restoreFromHistoryState(next);
      isRestoringHistory = false;
    }

    async function fetchImageAsDataUrl(url) {
      const res = await fetch(url, { cache: 'no-store' });
      if (!res.ok) {
        throw new Error(`Failed to fetch: ${url} (HTTP ${res.status})`);
      }

      const blob = await res.blob();
      return await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onerror = () => reject(new Error('Failed to read image blob'));
        reader.onload = () => resolve(reader.result);
        reader.readAsDataURL(blob);
      });
    }

    async function loadStickersFromGithub() {
      try {
        showStatus(t('loadingFromRepo'));
        
        // Flatten all sticker files from all categories
        const allFiles = [];
        if (GITHUB_FILES && GITHUB_FILES.stickers) {
          Object.keys(GITHUB_FILES.stickers).forEach(category => {
            const categoryFiles = GITHUB_FILES.stickers[category];
            if (Array.isArray(categoryFiles)) {
              categoryFiles.forEach(fileName => {
                allFiles.push({ fileName, category });
              });
            }
          });
        }
        
        if (allFiles.length === 0) {
          showStatus(t('noFilesToLoad'), true);
          return;
        }

        const cfg = getStickerLayoutConfigFromUI();
        const desiredCount = Number.isFinite(cfg.uploadLimit) && cfg.uploadLimit > 0 ? cfg.uploadLimit : 0;
        const actualCount = calculateActualStickerCount(desiredCount);
        const filesToLoad = actualCount > 0 ? actualCount : allFiles.length;

        pushHistory();

        for (let i = 0; i < filesToLoad; i++) {
          const fileInfo = allFiles[i % allFiles.length];
          const url = `${GITHUB_REPO.stickers}${encodeURIComponent(fileInfo.category)}/${encodeURIComponent(fileInfo.fileName)}`;
          const dataUrl = await fetchImageAsDataUrl(url);

          const img = await new Promise((resolve, reject) => {
            const image = new Image();
            image.onload = () => resolve(image);
            image.onerror = () => reject(new Error(`Image failed to load: ${url}`));
            image.src = dataUrl;
          });

          const originalWidth = img.width;
          const originalHeight = img.height;

          stickers.push({
            id: `sticker-github-${Date.now()}-${i}`,
            dataUrl,
            fileName: fileInfo.fileName,
            category: fileInfo.category,
            page: 0,
            x: 0,
            y: 0,
            width: originalWidth,
            height: originalHeight,
            originalWidth,
            originalHeight,
            words: [],
            images: []
          });
        }

        // תמיד קוראים ל-reflowStickers כדי לחשב גדלים נכון
        reflowStickers();
        renderStickers();
        updateFileCount();
        showStatus(t('filesLoadedSuccess', { count: filesToLoad }));
      } catch (error) {
        console.error('GitHub Stickers Error:', error);
        showStatus(t('errorLoadingFromRepo'), true);
      }
    }

    async function onConfigChange(config) {
      const titleElement = document.getElementById('pageTitle');
      
      if (titleElement) {
        titleElement.textContent = config.page_title || defaultConfig.page_title;
        titleElement.style.color = config.text_color || defaultConfig.text_color;
        titleElement.style.fontFamily = `${config.font_family || defaultConfig.font_family}, Arial, sans-serif`;
        titleElement.style.fontSize = `${(config.font_size || defaultConfig.font_size) * 2}px`;
      }
      document.body.style.fontFamily = `${config.font_family || defaultConfig.font_family}, Arial, sans-serif`;
      document.body.style.fontSize = `${config.font_size || defaultConfig.font_size}px`;
    }

    function renderLotteryResults() {
      const emptyState = document.getElementById('lotteryEmptyState');
      const resultsSection = document.getElementById('lotteryResultsSection');
      const results = document.getElementById('lotteryResults');

      if (!resultsSection || !results || !emptyState) return;

      if (!lotteryNumbers || lotteryNumbers.length === 0) {
        results.innerHTML = '';
        emptyState.classList.remove('hidden');
        resultsSection.classList.add('hidden');
        return;
      }

      emptyState.classList.add('hidden');
      resultsSection.classList.remove('hidden');
      results.innerHTML = '';

      const perRow = Math.max(1, Math.min(40, Number(document.getElementById('lotteryPerRow')?.value) || 5));
      const fontSize = Math.max(12, Number(document.getElementById('lotteryFontSize')?.value) || 48);
      const color = document.getElementById('lotteryColor')?.value || '#000000';

      const table = document.createElement('table');
      table.style.width = '100%';
      table.style.borderCollapse = 'separate';
      table.style.borderSpacing = '0';
      table.style.tableLayout = 'fixed';
      table.style.overflow = 'hidden';
      table.style.borderRadius = '14px';
      table.style.border = '2px solid rgba(0,0,0,0.08)';

      const tbody = document.createElement('tbody');
      const rowCount = Math.ceil(lotteryNumbers.length / perRow);

      for (let r = 0; r < rowCount; r++) {
        const tr = document.createElement('tr');
        tr.style.background = (r % 2 === 0) ? '#ffffff' : '#f3f4f6';

        for (let c = 0; c < perRow; c++) {
          const idx = r * perRow + c;
          const td = document.createElement('td');
          td.style.padding = '12px 8px';
          td.style.textAlign = 'center';
          td.style.borderBottom = '1px solid rgba(0,0,0,0.06)';
          td.style.borderLeft = c === 0 ? 'none' : '1px solid rgba(0,0,0,0.06)';

          if (idx < lotteryNumbers.length) {
            const el = document.createElement('div');
            el.textContent = String(lotteryNumbers[idx]);
            el.style.fontSize = `${fontSize}px`;
            el.style.fontWeight = '800';
            el.style.color = color;
            el.style.lineHeight = '1';
            td.appendChild(el);
          } else {
            td.innerHTML = '&nbsp;';
          }

          tr.appendChild(td);
        }

        tbody.appendChild(tr);
      }

      table.appendChild(tbody);
      results.appendChild(table);
    }

    function generateLottery() {
      const min = Number(document.getElementById('lotteryMin')?.value);
      const max = Number(document.getElementById('lotteryMax')?.value);
      const count = Number(document.getElementById('lotteryCount')?.value);

      if (!Number.isFinite(min) || !Number.isFinite(max) || !Number.isFinite(count) || min < 1 || max < 1 || count < 1 || max < min) {
        showStatus('טווח/כמות לא תקינים', true);
        return;
      }

      const rangeSize = max - min + 1;
      if (count > rangeSize) {
        showStatus('הכמות גדולה מהטווח (אי אפשר מספרים כפולים)', true);
        return;
      }

      const set = new Set();
      while (set.size < count) {
        const n = Math.floor(Math.random() * rangeSize) + min;
        set.add(n);
      }

      lotteryNumbers = Array.from(set);
      for (let i = lotteryNumbers.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        const tmp = lotteryNumbers[i];
        lotteryNumbers[i] = lotteryNumbers[j];
        lotteryNumbers[j] = tmp;
      }

      renderLotteryResults();
      showStatus(`הוגרלו ${lotteryNumbers.length} מספרים!`);
    }


    async function downloadLotteryAsPDF() {
      if (!lotteryNumbers || lotteryNumbers.length === 0) {
        showStatus('אין תוצאות להורדה', true);
        return;
      }

      const resultsSection = document.getElementById('lotteryResultsSection');
      if (!resultsSection) {
        showStatus('שגיאה בתצוגת ההגרלה', true);
        return;
      }

      const results = document.getElementById('lotteryResults');
      if (!results) {
        showStatus('שגיאה בתצוגת ההגרלה', true);
        return;
      }

      const { jsPDF } = window.jspdf;
      showStatus('מכין PDF...');

      try {
        const jpegQuality = (EXPORT_QUALITY && Number.isFinite(EXPORT_QUALITY.jpegQuality)) ? EXPORT_QUALITY.jpegQuality : 0.98;
        const pdfCompression = (EXPORT_QUALITY && EXPORT_QUALITY.pdfCompression) ? EXPORT_QUALITY.pdfCompression : 'NONE';
        const pdfDpr = (typeof window !== 'undefined' && window.devicePixelRatio)
          ? Math.min(2, window.devicePixelRatio)
          : 1;

        const wrapper = document.createElement('div');
        wrapper.style.position = 'fixed';
        wrapper.style.left = '-99999px';
        wrapper.style.top = '0';
        wrapper.style.width = '210mm';
        wrapper.style.height = '297mm';
        wrapper.style.background = '#ffffff';
        wrapper.style.padding = '12mm';
        wrapper.style.boxSizing = 'border-box';
        wrapper.style.overflow = 'hidden';

        const title = document.createElement('div');
        title.textContent = 'תוצאות ההגרלה';
        title.style.fontSize = '18px';
        title.style.fontWeight = '800';
        title.style.textAlign = 'center';
        title.style.marginBottom = '12px';
        wrapper.appendChild(title);

        const clone = results.cloneNode(true);
        clone.style.width = '100%';
        wrapper.appendChild(clone);

        document.body.appendChild(wrapper);

        const canvas = await captureElementToCanvas(wrapper, { scale: EXPORT_QUALITY.pdfScale, dpr: pdfDpr });

        document.body.removeChild(wrapper);

        const pdf = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4', compress: true });
        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfHeight = pdf.internal.pageSize.getHeight();

        const normalized = normalizeCanvasForPdf(canvas, pdfWidth, pdfHeight);
        const imgData = normalized.toDataURL('image/jpeg', jpegQuality);

        const margin = 8;
        const maxW = pdfWidth - margin * 2;
        const maxH = pdfHeight - margin * 2;

        const imgRatio = canvas.width / canvas.height;
        let drawW = maxW;
        let drawH = drawW / imgRatio;
        if (drawH > maxH) {
          drawH = maxH;
          drawW = drawH * imgRatio;
        }

        const x = (pdfWidth - drawW) / 2;
        const y = (pdfHeight - drawH) / 2;

        pdf.addImage(imgData, 'JPEG', x, y, drawW, drawH, undefined, pdfCompression);
        pdf.save('הגרלת-מספרים.pdf');
        showStatus('PDF הורד בהצלחה! ✓');
      } catch (error) {
        console.error('Lottery PDF Error:', error);
        try {
          const leftover = document.querySelector('body > div[style*="210mm"][style*="-99999px"]');
          if (leftover) document.body.removeChild(leftover);
        } catch (_) {}
        showStatus('שגיאה בהורדת PDF', true);
      }
    }

    function renderStickers() {
      const preview = document.getElementById('printPreview');
      const previewInner = document.getElementById('printPreviewInner') || preview;
      const printPreviewSection = document.getElementById('printPreviewSection');
      
      if (!preview || !printPreviewSection) {
        console.error('renderStickers: Required DOM elements not found');
        return;
      }

      // Always show the preview section
      printPreviewSection.classList.remove('hidden');

      if (stickers.length === 0) {
        // Show empty page with correct orientation
        const pageEl = document.createElement('div');
        pageEl.className = pageOrientation === 'landscape' ? 'print-page landscape' : 'print-page portrait';
        pageEl.dataset.pageIndex = 0;
        previewInner.replaceChildren(pageEl);
        applyPrintPreviewScale();
        updateButtonStates(); // Update button states when no stickers
        return;
      }

      try {
        const fragment = document.createDocumentFragment();

        const maxPageIndex = stickers.reduce((max, s) => Math.max(max, Number.isFinite(s.page) ? s.page : 0), 0);
        const pageCount = Math.max(1, maxPageIndex + 1);

        const pages = [];
        for (let p = 0; p < pageCount; p++) {
          const pageEl = document.createElement('div');
          pageEl.className = pageOrientation === 'landscape' ? 'print-page landscape' : 'print-page portrait';
          pageEl.dataset.pageIndex = p;
          fragment.appendChild(pageEl);
          pages.push(pageEl);
        }

        stickers.forEach((sticker, index) => {
          const pageIndex = Number.isFinite(sticker.page) ? sticker.page : 0;
          const pageEl = pages[Math.max(0, Math.min(pageIndex, pages.length - 1))];

          const stickerDiv = document.createElement('div');
          stickerDiv.className = 'sticker-container';
          stickerDiv.style.left = `${sticker.x}px`;
          stickerDiv.style.top = `${sticker.y}px`;
          stickerDiv.style.width = `${sticker.width}px`;
          stickerDiv.style.height = `${sticker.height}px`;
          stickerDiv.dataset.stickerIndex = index;

          stickerDiv.addEventListener('mousedown', (e) => {
            if (e.button !== 0) return;
            if (e.target && e.target.closest) {
              if (e.target.closest('.sticker-controls')) return;
              if (e.target.closest('.sticker-resize-handle')) return;
              if (e.target.closest('.text-word')) return;
              if (e.target.closest('.sticker-image')) return;
            }
            e.preventDefault();
            selectSticker(index);
            startStickerDrag(e, index);
          });
          
          const img = document.createElement('img');
          img.src = sticker.dataUrl;
          img.className = 'sticker-main-image';
          img.style.width = '100%';
          img.style.height = '100%';
          img.style.objectFit = 'fill';
          img.style.pointerEvents = 'none';
          
          // Apply opacity only to the main sticker image, not the container
          // This allows child elements (text, images) to have their own independent opacity
          if (sticker.opacity !== undefined) {
            img.style.opacity = sticker.opacity;
          }
          
          stickerDiv.appendChild(img);

          const controlsDiv = document.createElement('div');
          controlsDiv.className = 'sticker-controls no-print';
          
          const deleteBtn = document.createElement('button');
          deleteBtn.className = 'sticker-control-btn delete-sticker-btn';
          deleteBtn.textContent = '×';
          deleteBtn.onclick = (e) => {
            e.stopPropagation();
            deleteStickerByIndex(index);
          };
          
          const duplicateBtn = document.createElement('button');
          duplicateBtn.className = 'sticker-control-btn duplicate-sticker-btn';
          duplicateBtn.textContent = '+';
          duplicateBtn.onclick = (e) => {
            e.stopPropagation();
            duplicateSticker(index);
          };
          
          controlsDiv.appendChild(duplicateBtn);
          controlsDiv.appendChild(deleteBtn);
          
          // הסרנו את כפתור הסרת הרקע - עכשיו הוא בסרגל העליון

          stickerDiv.appendChild(controlsDiv);
          
          const resizeHandle = document.createElement('div');
          resizeHandle.className = 'sticker-resize-handle no-print';
          resizeHandle.addEventListener('mousedown', (e) => {
            e.stopPropagation();
            startStickerResize(e, index);
          });
          stickerDiv.appendChild(resizeHandle);
          
          // תצוגת מידות המדבקה (רוחב למעלה, גובה בצד ימין)
          const widthInCm = (sticker.width / MM_TO_PX / 10).toFixed(1);
          const heightInCm = (sticker.height / MM_TO_PX / 10).toFixed(1);
          
          const widthLabel = document.createElement('div');
          widthLabel.className = 'sticker-dimension sticker-dimension-width no-print';
          widthLabel.textContent = `${widthInCm} ס"מ`;
          stickerDiv.appendChild(widthLabel);
          
          const heightLabel = document.createElement('div');
          heightLabel.className = 'sticker-dimension sticker-dimension-height no-print';
          heightLabel.textContent = `${heightInCm} ס"מ`;
          stickerDiv.appendChild(heightLabel);
          
          sticker.images = sticker.images || [];
          sticker.images.forEach(image => {
            const imageEl = createImageElement(image, index);
            stickerDiv.appendChild(imageEl);
          });
          
          sticker.words = sticker.words || [];
          sticker.words.forEach(word => {
            const wordEl = createWordElement(word, index);
            stickerDiv.appendChild(wordEl);
          });
          
          stickerDiv.addEventListener('click', (e) => {
            if (e.target === stickerDiv || e.target === img) {
              selectSticker(index);
              
              // פתח טקסט רק אם טאב הטקסט כבר פתוח
              const textSection = document.getElementById('toolsSectionContentText');
              if (textSection && !textSection.classList.contains('hidden')) {
                focusWordInput();
              }
            }
          });
          
          pageEl.appendChild(stickerDiv);
        });

        previewInner.replaceChildren(fragment);
        applyPrintPreviewScale();
        updateButtonStates(); // Update button states after rendering
      } catch (error) {
        console.error('Error in renderStickers:', error);
        showStatus('שגיאה בהצגת המדבקות (ראה Console)', true);
      }
    }

    function applyPrintPreviewScale() {
      const preview = document.getElementById('printPreview');
      const inner = document.getElementById('printPreviewInner');
      const section = document.getElementById('printPreviewSection');

      if (!preview || !inner || !section) return;

      const pages = inner.querySelectorAll('.print-page');
      if (pages.length === 0) {
        preview.style.width = '';
        preview.style.height = '';
        inner.style.transform = '';
        updatePrintRuler();
        return;
      }

      inner.style.transform = '';
      preview.style.width = '';
      preview.style.height = '';

      // Always use A4 dimensions for scaling, not content dimensions
      const A4_WIDTH_PX = pageOrientation === 'landscape' ? 297 * MM_TO_PX : 210 * MM_TO_PX;
      const A4_HEIGHT_PX = pageOrientation === 'landscape' ? 210 * MM_TO_PX : 297 * MM_TO_PX;
      
      // Force pages to maintain A4 dimensions with aggressive CSS enforcement
      pages.forEach(page => {
        // Remove any existing inline styles that might interfere
        page.style.cssText = '';
        
        // Apply A4 dimensions with maximum specificity and force override
        page.style.setProperty('width', `${A4_WIDTH_PX}px`, 'important');
        page.style.setProperty('height', `${A4_HEIGHT_PX}px`, 'important');
        page.style.setProperty('min-width', `${A4_WIDTH_PX}px`, 'important');
        page.style.setProperty('min-height', `${A4_HEIGHT_PX}px`, 'important');
        page.style.setProperty('max-width', `${A4_WIDTH_PX}px`, 'important');
        page.style.setProperty('max-height', `${A4_HEIGHT_PX}px`, 'important');
        page.style.setProperty('box-sizing', 'border-box', 'important');
        page.style.setProperty('flex-shrink', '0', 'important');
        page.style.setProperty('overflow', 'visible', 'important');
        page.style.setProperty('position', 'relative', 'important');
        page.style.setProperty('background', 'white', 'important');
        
        // Force dimensions using attributes as well (backup method)
        page.setAttribute('data-forced-width', A4_WIDTH_PX);
        page.setAttribute('data-forced-height', A4_HEIGHT_PX);
        
        // Ensure correct class is applied
        if (pageOrientation === 'landscape') {
          page.classList.add('landscape');
          page.classList.remove('portrait');
        } else {
          page.classList.remove('landscape');
          page.classList.add('portrait');
        }
      });

      // Force a reflow to ensure styles are applied
      pages.forEach(page => {
        page.offsetHeight; // Force reflow
      });

      const maxScale = 1;
      const availableWidth = Math.max(0, preview.clientWidth - 2);
      const scale = A4_WIDTH_PX > 0 ? Math.min(maxScale, availableWidth / A4_WIDTH_PX) : 1;

      inner.style.transformOrigin = 'top right';
      inner.style.transform = `scale(${scale})`;
      preview.style.width = '';
      preview.style.height = '';

      updatePrintRuler();
      
      // Double-check after a short delay to ensure dimensions stick
      setTimeout(() => {
        pages.forEach(page => {
          if (page.offsetWidth !== A4_WIDTH_PX || page.offsetHeight !== A4_HEIGHT_PX) {
            console.log('Re-forcing A4 dimensions');
            page.style.setProperty('width', `${A4_WIDTH_PX}px`, 'important');
            page.style.setProperty('height', `${A4_HEIGHT_PX}px`, 'important');
          }
        });
      }, 100);
    }

    function updatePrintRuler() {
      const ruler = document.getElementById('printRuler');
      const track = document.getElementById('printRulerTrack');
      const ticks = document.getElementById('printRulerTicks');
      const label = document.getElementById('printRulerLabel');
      const inner = document.getElementById('printPreviewInner');
      const preview = document.getElementById('printPreview');
      const canvas = document.getElementById('printRulerCanvas');
      if (!ruler || !track || !ticks || !label || !inner) return;

      const page = inner.querySelector('.print-page');
      if (!page) {
        track.style.width = '0px';
        track.style.transform = 'translateX(0px)';
        if (canvas) {
          canvas.width = 0;
          canvas.height = 0;
        }
        ticks.replaceChildren();
        label.textContent = '';
        return;
      }

      const rect = page.getBoundingClientRect();
      const w = Math.max(0, rect.width);
      const h = Math.max(0, track.clientHeight || 44);

      track.style.width = `${Math.floor(w)}px`;

      if (preview) {
        const previewRect = preview.getBoundingClientRect();
        const rulerRect = ruler.getBoundingClientRect();
        // הסרגל צריך להיות מיושר עם הדף
        // הדף מיושר לימין בתוך ה-preview (בגלל align-items: flex-end ו-direction: rtl)
        // צריך לחשב את המרחק מצד ימין של הסרגל לצד ימין של הדף
        const pageRight = rect.right;
        const rulerRight = rulerRect.right;
        // ה-offset הוא כמה צריך להזיז את ה-track כדי שהצד הימני שלו יהיה באותו מקום כמו הצד הימני של הדף
        // ב-flex עם justify-content: flex-end, ה-track מתחיל מימין
        // אם הדף יותר שמאלה מהסרגל, צריך להזיז את ה-track שמאלה (offset שלילי)
        const offsetFromRight = rulerRight - pageRight;
        // ב-justify-content: flex-end, ה-track מיושר לימין
        // צריך להזיז אותו שמאלה ב-offsetFromRight
        track.style.marginRight = `${Math.round(offsetFromRight)}px`;
        track.style.transform = '';
      } else {
        track.style.marginRight = '0';
        track.style.transform = '';
      }

      label.textContent = '';
      ticks.replaceChildren();

      if (!canvas) return;

      const dpr = window.devicePixelRatio || 1;
      canvas.width = Math.max(1, Math.floor(w * dpr));
      canvas.height = Math.max(1, Math.floor(h * dpr));
      const ctx = canvas.getContext('2d');
      if (!ctx) return;

      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx.clearRect(0, 0, w, h);

      const cmTotal = (pageOrientation === 'landscape') ? 29.7 : 21;
      const hasFraction = Math.abs(cmTotal - Math.round(cmTotal)) > 1e-6;
      const totalMm = Math.round(cmTotal * 10);
      const pxPerMm = totalMm > 0 ? (w / totalMm) : 0;
      if (pxPerMm <= 0) return;

      ctx.fillStyle = 'rgba(255, 255, 255, 0.95)';
      ctx.fillRect(0, 0, w, h);

      ctx.strokeStyle = 'rgba(17, 24, 39, 0.85)';
      ctx.lineWidth = 1;

      const topPad = 3;
      const majorH = 22;
      const halfH = 15;
      const minorH = 10;

      ctx.beginPath();
      for (let mm = 0; mm <= totalMm; mm++) {
        const x = Math.round(mm * pxPerMm) + 0.5;
        const isMajor = (mm % 10 === 0);
        const isHalf = (!isMajor && mm % 5 === 0);
        const tickH = isMajor ? majorH : (isHalf ? halfH : minorH);
        ctx.moveTo(x, topPad);
        ctx.lineTo(x, topPad + tickH);
      }
      ctx.stroke();

      ctx.fillStyle = 'rgba(17, 24, 39, 0.95)';
      ctx.font = '700 11px Arial';
      ctx.textBaseline = 'bottom';

      const labelY = h - 4;
      for (let mm = 0; mm <= totalMm; mm += 10) {
        if (hasFraction && mm === totalMm) continue;
        let x = mm * pxPerMm;
        const text = String(Math.round(mm / 10));

        if (mm === 0) {
          ctx.textAlign = 'left';
          x = 2;
        } else if (mm === totalMm) {
          ctx.textAlign = 'right';
          // להזיז את המספר האחרון מילימטר אחד שמאלה
          x = Math.max(0, w - 2 - pxPerMm);
        } else {
          ctx.textAlign = 'center';
        }

        ctx.fillText(text, x, labelY);
      }

      if (hasFraction) {
        ctx.textAlign = 'right';
        // גם את 29.7 להזיז מילימטר אחד שמאלה
        ctx.fillText(String(cmTotal), Math.max(0, w - 2 - pxPerMm), labelY);
      }
    }

    setTimeout(() => {
      const preview = document.getElementById('printPreview');
      if (preview) {
        preview.addEventListener('scroll', updatePrintRuler, { passive: true });
      }
      window.addEventListener('resize', updatePrintRuler);
    }, 0);

    async function captureElementToCanvas(element, options = {}) {
      const noPrintElements = element.querySelectorAll('.no-print');
      noPrintElements.forEach(el => el.style.display = 'none');

      // הסרת סימון הבחירה מכל האלמנטים לפני הצילום
      const selectedStickers = element.querySelectorAll('.sticker-container.selected');
      const selectedWords = element.querySelectorAll('.text-word.selected');
      const selectedImages = element.querySelectorAll('.sticker-image.selected');
      
      selectedStickers.forEach(el => el.classList.remove('selected'));
      selectedWords.forEach(el => el.classList.remove('selected'));
      selectedImages.forEach(el => el.classList.remove('selected'));

      // תיקון טקסט עם גרדיאנט לייצוא - html2canvas לא תומך ב-background-clip: text
      const gradientTextElements = element.querySelectorAll('.text-word');
      const gradientTextBackup = [];

      const extractFirstColorFromGradient = (value) => {
        if (!value || typeof value !== 'string') return null;
        const hex = value.match(/#[0-9a-fA-F]{3,8}/);
        if (hex) return hex[0];
        const rgb = value.match(/rgba?\([^\)]+\)/);
        if (rgb) return rgb[0];
        const hsl = value.match(/hsla?\([^\)]+\)/);
        if (hsl) return hsl[0];
        return null;
      };

      const escapeXml = (s) => String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');

      const splitGradientArgs = (s) => {
        const out = [];
        let current = '';
        let depth = 0;
        for (let i = 0; i < s.length; i++) {
          const ch = s[i];
          if (ch === '(') depth++;
          if (ch === ')') depth = Math.max(0, depth - 1);
          if (ch === ',' && depth === 0) {
            out.push(current.trim());
            current = '';
            continue;
          }
          current += ch;
        }
        if (current.trim()) out.push(current.trim());
        return out;
      };

      const directionToEndpoints = (dir) => {
        const v = String(dir || '').trim().toLowerCase();
        if (!v) return { x1: 0, y1: 0.5, x2: 1, y2: 0.5 };
        if (v.startsWith('to ')) {
          const t = v.slice(3).trim();
          if (t === 'left') return { x1: 1, y1: 0.5, x2: 0, y2: 0.5 };
          if (t === 'right') return { x1: 0, y1: 0.5, x2: 1, y2: 0.5 };
          if (t === 'top') return { x1: 0.5, y1: 1, x2: 0.5, y2: 0 };
          if (t === 'bottom') return { x1: 0.5, y1: 0, x2: 0.5, y2: 1 };
          if (t === 'top left' || t === 'left top') return { x1: 1, y1: 1, x2: 0, y2: 0 };
          if (t === 'top right' || t === 'right top') return { x1: 0, y1: 1, x2: 1, y2: 0 };
          if (t === 'bottom left' || t === 'left bottom') return { x1: 1, y1: 0, x2: 0, y2: 1 };
          if (t === 'bottom right' || t === 'right bottom') return { x1: 0, y1: 0, x2: 1, y2: 1 };
        }

        const degMatch = v.match(/(-?\d+(?:\.\d+)?)deg/);
        if (degMatch) {
          const deg = parseFloat(degMatch[1]);
          const rad = (deg * Math.PI) / 180;
          const dx = Math.sin(rad);
          const dy = -Math.cos(rad);
          return {
            x1: 0.5 - dx / 2,
            y1: 0.5 - dy / 2,
            x2: 0.5 + dx / 2,
            y2: 0.5 + dy / 2
          };
        }

        return { x1: 0, y1: 0.5, x2: 1, y2: 0.5 };
      };

      const extractGradientCallArgs = (bg, fnName) => {
        const hay = String(bg || '');
        const needle = String(fnName || '').toLowerCase();
        const idx = hay.toLowerCase().indexOf(needle);
        if (idx < 0) return null;
        let i = idx + needle.length;
        let depth = 1;
        let out = '';
        for (; i < hay.length; i++) {
          const ch = hay[i];
          if (ch === '(') depth++;
          if (ch === ')') {
            depth = Math.max(0, depth - 1);
            if (depth === 0) break;
          }
          out += ch;
        }
        if (depth !== 0) return null;
        return out.trim();
      };

      const buildGradientSvgDataUrl = (el, cs, computedBg, gradientIndex) => {
        const bg = String(computedBg || '').trim();
        const linearArgs = extractGradientCallArgs(bg, 'linear-gradient(');
        const radialArgs = extractGradientCallArgs(bg, 'radial-gradient(');
        if (!linearArgs && !radialArgs) return null;

        const boxW = Math.max(1, Math.round(el.offsetWidth || el.getBoundingClientRect().width || 1));
        const boxH = Math.max(1, Math.round(el.offsetHeight || el.getBoundingClientRect().height || 1));

        const fontFamily = cs.fontFamily || 'sans-serif';
        const fontSize = cs.fontSize || '16px';
        const fontWeight = cs.fontWeight || 'normal';
        const fontStyle = cs.fontStyle || 'normal';
        const letterSpacing = cs.letterSpacing && cs.letterSpacing !== 'normal' ? cs.letterSpacing : '';

        const lh = (cs.lineHeight && cs.lineHeight !== 'normal') ? parseFloat(cs.lineHeight) : (parseFloat(fontSize) * 1.1);

        const rawText = (el.textContent || '').replace(/\r\n/g, '\n');
        const lines = rawText.split('\n');

        const textAlign = (cs.textAlign || 'center').toLowerCase();
        const x = (textAlign === 'left' || textAlign === 'start') ? 0 : (textAlign === 'right' || textAlign === 'end') ? boxW : (boxW / 2);
        const anchor = (textAlign === 'left' || textAlign === 'start') ? 'start' : (textAlign === 'right' || textAlign === 'end') ? 'end' : 'middle';

        const startY = (boxH / 2) - ((lines.length - 1) * lh) / 2;

        let gradientDef = '';
        let fillUrl = '';
        const gradientId = `__gt_${gradientIndex}`;

        if (linearArgs) {
          const args = splitGradientArgs(linearArgs);
          const maybeDir = args[0] || '';
          const hasDir = /^to\s+/i.test(maybeDir) || /deg\s*$/i.test(maybeDir);
          const dir = hasDir ? maybeDir : '';
          const colors = hasDir ? args.slice(1) : args;
          const c1 = extractFirstColorFromGradient(colors[0]);
          const c2 = extractFirstColorFromGradient(colors[1]);
          if (!c1 || !c2) return null;
          const ep = directionToEndpoints(dir);
          gradientDef = `<linearGradient id="${gradientId}" gradientUnits="objectBoundingBox" x1="${ep.x1}" y1="${ep.y1}" x2="${ep.x2}" y2="${ep.y2}">` +
            `<stop offset="0%" stop-color="${escapeXml(c1)}"/>` +
            `<stop offset="100%" stop-color="${escapeXml(c2)}"/>` +
            `</linearGradient>`;
          fillUrl = `url(#${gradientId})`;
        } else {
          const args = splitGradientArgs(radialArgs);
          const colors = args.filter(a => !/circle|ellipse|closest-side|farthest-side|closest-corner|farthest-corner|at\s+/i.test(a));
          const c1 = extractFirstColorFromGradient(colors[0]);
          const c2 = extractFirstColorFromGradient(colors[1]);
          if (!c1 || !c2) return null;
          gradientDef = `<radialGradient id="${gradientId}" cx="0.5" cy="0.5" r="0.75">` +
            `<stop offset="0%" stop-color="${escapeXml(c1)}"/>` +
            `<stop offset="100%" stop-color="${escapeXml(c2)}"/>` +
            `</radialGradient>`;
          fillUrl = `url(#${gradientId})`;
        }

        const tspans = lines.map((line, idx) => {
          const safe = escapeXml(line);
          if (idx === 0) {
            return `<tspan x="${x}" y="${startY}">${safe}</tspan>`;
          }
          return `<tspan x="${x}" dy="${lh}">${safe}</tspan>`;
        }).join('');

        const svg = `<?xml version="1.0" encoding="UTF-8"?>` +
          `<svg xmlns="http://www.w3.org/2000/svg" width="${boxW}" height="${boxH}" viewBox="0 0 ${boxW} ${boxH}">` +
          `<defs>${gradientDef}</defs>` +
          `<text x="${x}" y="${boxH / 2}" dominant-baseline="middle" text-anchor="${anchor}"` +
            ` fill="${fillUrl}" font-family="${escapeXml(fontFamily)}" font-size="${escapeXml(fontSize)}" font-weight="${escapeXml(fontWeight)}" font-style="${escapeXml(fontStyle)}"` +
            (letterSpacing ? ` letter-spacing="${escapeXml(letterSpacing)}"` : '') +
          `>${tspans}</text>` +
          `</svg>`;

        return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
      };

      gradientTextElements.forEach(el => {
        const cs = window.getComputedStyle(el);
        const computedBg = cs.backgroundImage || cs.background || '';
        const hasGradient = computedBg.includes('gradient') && (
          cs.webkitBackgroundClip === 'text' ||
          cs.backgroundClip === 'text' ||
          cs.webkitTextFillColor === 'transparent' ||
          cs.webkitTextFillColor === 'rgba(0, 0, 0, 0)'
        );

        if (hasGradient) {
          gradientTextBackup.push({
            el: el,
            originalVisibility: el.style.visibility,
            originalBackground: el.style.background,
            originalBackgroundImage: el.style.backgroundImage,
            originalWebkitTextFillColor: el.style.webkitTextFillColor,
            originalBackgroundClip: el.style.backgroundClip,
            originalWebkitBackgroundClip: el.style.webkitBackgroundClip,
            originalColor: el.style.color,
            overlayImg: null
          });

          const backup = gradientTextBackup[gradientTextBackup.length - 1];
          const svgUrl = buildGradientSvgDataUrl(el, cs, computedBg, gradientTextBackup.length);
          if (svgUrl && el.parentNode) {
            const img = document.createElement('img');
            img.src = svgUrl;
            img.alt = '';
            img.style.position = cs.position || 'absolute';
            img.style.left = cs.left;
            img.style.top = cs.top;
            img.style.width = cs.width;
            img.style.height = cs.height;
            img.style.transform = cs.transform;
            img.style.transformOrigin = cs.transformOrigin;
            img.style.zIndex = cs.zIndex;
            img.style.opacity = cs.opacity;
            img.style.pointerEvents = 'none';

            el.parentNode.appendChild(img);
            backup.overlayImg = img;
            el.style.visibility = 'hidden';
          } else {
            const fallbackColor = extractFirstColorFromGradient(computedBg) || '#000000';

            el.style.background = 'none';
            el.style.backgroundImage = 'none';
            el.style.backgroundClip = 'border-box';
            el.style.webkitBackgroundClip = 'border-box';
            el.style.webkitTextFillColor = fallbackColor;
            el.style.color = fallbackColor;
          }
        }
      });

      const rootPreview = (element && element.id === 'printPreview')
        ? element
        : (element && element.closest ? element.closest('#printPreview') : null);

      const rootInner = rootPreview ? rootPreview.querySelector('#printPreviewInner') : null;
      const savedPreviewStyles = rootPreview ? {
        width: rootPreview.style.width,
        height: rootPreview.style.height
      } : null;
      const savedInnerStyles = rootInner ? {
        transform: rootInner.style.transform,
        transformOrigin: rootInner.style.transformOrigin
      } : null;

      if (rootPreview && rootInner) {
        rootPreview.style.width = '';
        rootPreview.style.height = '';
        rootInner.style.transform = '';
        rootInner.style.transformOrigin = 'top right';
      }

      let textElements = [];
      let stickerElements = [];
      let pageElements = [];
      let savedPageStyles = [];

      try {
        if (document.fonts && document.fonts.ready) {
          await document.fonts.ready;
        }

        pageElements = Array.from(element.querySelectorAll('.print-page'));
        savedPageStyles = pageElements.map(el => ({
          el,
          boxShadow: el.style.boxShadow,
          margin: el.style.margin
        }));
        pageElements.forEach(el => {
          el.style.boxShadow = 'none';
          el.style.margin = '0';
        });

        stickerElements = Array.from(element.querySelectorAll('.sticker-container'));
        stickerElements.forEach(el => {
          el.dataset.exportOriginalLeft = el.style.left || '';
          el.dataset.exportOriginalTop = el.style.top || '';
          el.dataset.exportOriginalWidth = el.style.width || '';
          el.dataset.exportOriginalHeight = el.style.height || '';

          const left = parseFloat(el.style.left);
          const top = parseFloat(el.style.top);
          const width = parseFloat(el.style.width);
          const height = parseFloat(el.style.height);

          if (!Number.isNaN(left)) el.style.left = `${Math.round(left)}px`;
          if (!Number.isNaN(top)) el.style.top = `${Math.round(top)}px`;
          if (!Number.isNaN(width)) el.style.width = `${Math.round(width)}px`;
          if (!Number.isNaN(height)) el.style.height = `${Math.round(height)}px`;
        });

        textElements = Array.from(element.querySelectorAll('.text-word'));
        textElements.forEach(el => {
          el.dataset.exportOriginalTop = el.style.top || '';
          const currentTop = parseFloat(el.style.top);
          if (!Number.isNaN(currentTop)) {
            const computedTransform = window.getComputedStyle(el).transform;
            const extraOffset = computedTransform && computedTransform !== 'none' ? 12 : 10;
            el.style.top = `${currentTop - extraOffset}px`;
          }
        });

        const dpr = (typeof options.dpr === 'number')
          ? options.dpr
          : ((typeof window !== 'undefined' && window.devicePixelRatio) ? window.devicePixelRatio : 1);
        return await html2canvas(element, {
          scale: (options.scale || 2) * dpr,
          useCORS: true,
          allowTaint: true,
          backgroundColor: '#ffffff',
          logging: false
        });
      } finally {
        noPrintElements.forEach(el => el.style.display = '');
        
        // שחזור טקסט עם גרדיאנט
        gradientTextBackup.forEach(backup => {
          backup.el.style.background = backup.originalBackground;
          backup.el.style.backgroundImage = backup.originalBackgroundImage;
          backup.el.style.webkitTextFillColor = backup.originalWebkitTextFillColor;
          backup.el.style.backgroundClip = backup.originalBackgroundClip;
          backup.el.style.webkitBackgroundClip = backup.originalWebkitBackgroundClip;
          backup.el.style.color = backup.originalColor;
          backup.el.style.visibility = backup.originalVisibility;
          if (backup.overlayImg && backup.overlayImg.parentNode) {
            backup.overlayImg.parentNode.removeChild(backup.overlayImg);
          }
        });

        // החזרת סימון הבחירה לאלמנטים שהיו מסומנים
        selectedStickers.forEach(el => el.classList.add('selected'));
        selectedWords.forEach(el => el.classList.add('selected'));
        selectedImages.forEach(el => el.classList.add('selected'));

        savedPageStyles.forEach(s => {
          s.el.style.boxShadow = s.boxShadow;
          s.el.style.margin = s.margin;
        });

        stickerElements.forEach(el => {
          if (el.dataset && Object.prototype.hasOwnProperty.call(el.dataset, 'exportOriginalLeft')) {
            el.style.left = el.dataset.exportOriginalLeft;
            el.style.top = el.dataset.exportOriginalTop;
            el.style.width = el.dataset.exportOriginalWidth;
            el.style.height = el.dataset.exportOriginalHeight;
            delete el.dataset.exportOriginalLeft;
            delete el.dataset.exportOriginalTop;
            delete el.dataset.exportOriginalWidth;
            delete el.dataset.exportOriginalHeight;
          }
        });

        textElements.forEach(el => {
          if (el.dataset && Object.prototype.hasOwnProperty.call(el.dataset, 'exportOriginalTop')) {
            el.style.top = el.dataset.exportOriginalTop;
            delete el.dataset.exportOriginalTop;
          }
        });

        if (rootPreview && savedPreviewStyles) {
          rootPreview.style.width = savedPreviewStyles.width;
          rootPreview.style.height = savedPreviewStyles.height;
        }

        if (rootInner && savedInnerStyles) {
          rootInner.style.transform = savedInnerStyles.transform;
          rootInner.style.transformOrigin = savedInnerStyles.transformOrigin;
        }

        applyPrintPreviewScale();
      }
    }

    function createImageElement(image, stickerIndex) {
      const wrapper = document.createElement('div');
      wrapper.className = 'sticker-image';
      wrapper.dataset.imageId = image.id;
      wrapper.dataset.stickerIndex = stickerIndex;
      wrapper.style.left = `${image.x}px`;
      wrapper.style.top = `${image.y}px`;
      wrapper.style.width = `${image.width}px`;
      wrapper.style.height = `${image.height}px`;
      
      // Apply opacity - default to 1 (fully opaque) if not set
      // This ensures images have their own opacity independent of sticker
      wrapper.style.opacity = (image.opacity !== undefined) ? image.opacity : 1;

      const img = document.createElement('img');
      img.src = image.dataUrl;
      img.style.width = '100%';
      img.style.height = '100%';
      img.style.objectFit = 'contain';
      img.style.pointerEvents = 'none';
      wrapper.appendChild(img);

      const deleteBtn = document.createElement('button');
      deleteBtn.type = 'button';
      deleteBtn.className = 'delete-image-btn no-print';
      deleteBtn.textContent = '×';
      deleteBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        deleteImage(stickerIndex, image.id);
      });
      wrapper.appendChild(deleteBtn);

      // הסרנו את כפתור הסרת הרקע - עכשיו הוא בסרגל העליון

      const resizeHandle = document.createElement('div');
      resizeHandle.className = 'resize-handle no-print';
      resizeHandle.addEventListener('mousedown', (e) => {
        e.stopPropagation();
        startImageResize(e, stickerIndex, image.id);
      });
      wrapper.appendChild(resizeHandle);

      wrapper.addEventListener('mousedown', (e) => {
        if (e.target === resizeHandle || e.target === deleteBtn) return;
        e.stopPropagation();
        startImageDrag(e, stickerIndex, image.id);
      });

      wrapper.addEventListener('click', (e) => {
        e.stopPropagation();
        selectImage(stickerIndex, image.id);
      });

      return wrapper;
    }

    function createWordElement(word, stickerIndex) {
      const el = document.createElement('div');
      el.className = 'text-word';
      el.dataset.wordId = word.id;
      el.dataset.stickerIndex = stickerIndex;
      
      // Apply curve if available using SVG text path
      if (word.curve && word.curve !== 0) {
        el.innerHTML = createCurvedText(word.text, word.curve, word);
      } else {
        // טקסט רגיל - שים ב-span נפרד לעדכון קל בזמן עריכה
        const textSpan = document.createElement('span');
        textSpan.className = 'word-text-content';
        textSpan.textContent = word.text;
        el.appendChild(textSpan);
      }
      
      el.style.left = `${word.x}px`;
      el.style.top = `${word.y}px`;
      el.style.color = word.color || '#000000';
      
      // Apply gradient if available
      if (word.isGradient && word.color) {
        el.style.background = word.color;
        el.style.webkitBackgroundClip = 'text';
        el.style.backgroundClip = 'text';
        el.style.webkitTextFillColor = 'transparent';
        el.style.color = 'transparent';
      }
      el.style.fontSize = `${word.fontSize || 12}px`;
      el.style.fontFamily = word.fontFamily || 'Arial';
      el.style.fontWeight = word.fontWeight || '700';
      
      // Apply rotation if available
      if (word.rotation) {
        el.style.transform = `rotate(${word.rotation}deg)`;
      }
      
      // Apply opacity - default to 1 (fully opaque) if not set
      // This ensures words have their own opacity independent of sticker
      el.style.opacity = (word.opacity !== undefined) ? word.opacity : 1;

      const deleteBtn = document.createElement('button');
      deleteBtn.type = 'button';
      deleteBtn.className = 'delete-word-btn no-print';
      deleteBtn.textContent = '×';
      deleteBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        deleteWord(stickerIndex, word.id);
      });
      el.appendChild(deleteBtn);

      const resizeHandle = document.createElement('div');
      resizeHandle.className = 'word-resize-handle no-print';
      resizeHandle.addEventListener('mousedown', (e) => {
        e.stopPropagation();
        startWordResize(e, stickerIndex, word.id);
      });
      el.appendChild(resizeHandle);
      
      // Rotation handle - appears below center of text
      const rotateHandle = document.createElement('div');
      rotateHandle.className = 'word-rotate-handle no-print';
      rotateHandle.innerHTML = '↻';
      rotateHandle.addEventListener('mousedown', (e) => {
        e.stopPropagation();
        startWordRotate(e, stickerIndex, word.id, el);
      });
      el.appendChild(rotateHandle);
      
      // Curve handle - appears below left side of text
      const curveHandle = document.createElement('div');
      curveHandle.className = 'word-curve-handle no-print';
      curveHandle.innerHTML = '⌒';
      curveHandle.addEventListener('mousedown', (e) => {
        e.stopPropagation();
        startWordCurve(e, stickerIndex, word.id, el);
      });
      el.appendChild(curveHandle);

      el.addEventListener('mousedown', (e) => {
        if (e.target === deleteBtn || e.target === resizeHandle || e.target === rotateHandle || e.target === curveHandle) return;
        e.stopPropagation();
        startWordDrag(e, stickerIndex, word.id);
      });

      el.addEventListener('click', (e) => {
        e.stopPropagation();
        
        // אם הטקסט כבר נבחר, פתח עריכה ישירה
        if (el.classList.contains('selected')) {
          startInlineTextEdit(el, stickerIndex, word.id);
        } else {
          // אחרת, בחר את הטקסט
          selectWord(stickerIndex, word.id);
        }
      });

      // Remove double-click event since we're using single click now
      // el.addEventListener('dblclick', (e) => {
      //   e.stopPropagation();
      //   startInlineTextEdit(el, stickerIndex, word.id);
      // });

      return el;
    }
    
    // פונקציה ליצירת טקסט מקומר באמצעות SVG
    function createCurvedText(text, curve, word) {
      const fontSize = word.fontSize || 12;
      const fontFamily = word.fontFamily || 'Arial';
      const fontWeight = word.fontWeight || '700';
      const color = word.color || '#000000';
      const isGradient = word.isGradient;
      
      // חישוב רוחב משוער של הטקסט - מינימלי יותר
      const absCurve = Math.abs(curve);
      const charWidth = fontSize * 0.6;
      const textWidth = text.length * charWidth;
      const svgWidth = Math.max(textWidth + 20, 80);
      
      // גובה מינימלי - רק מה שצריך לקימור
      const svgHeight = fontSize + absCurve * 0.8 + 10;
      
      // חישוב הקשת - קומפקטי יותר
      const pathPadding = 5;
      const startX = pathPadding;
      const endX = svgWidth - pathPadding;
      const midX = svgWidth / 2;
      
      // curve חיובי = קימור למטה (חיוך), curve שלילי = קימור למעלה
      const baseY = curve > 0 ? fontSize * 0.8 : svgHeight - fontSize * 0.3;
      const curveY = baseY + curve * 0.7;
      
      const pathId = `curve-${word.id}`;
      
      let fillStyle = '';
      let defsContent = '';
      
      if (isGradient && color) {
        // פירוק הגרדיאנט לצבעים
        const gradientId = `grad-${word.id}`;
        let color1 = '#FF0000';
        let color2 = '#0000FF';
        
        // ניסיון לחלץ צבעים מהגרדיאנט
        const colorMatch = color.match(/(#[0-9A-Fa-f]{6}|#[0-9A-Fa-f]{3}|rgb\([^)]+\)|rgba\([^)]+\))/g);
        if (colorMatch && colorMatch.length >= 2) {
          color1 = colorMatch[0];
          color2 = colorMatch[1];
        }
        
        // בדיקה אם זה גרדיאנט רדיאלי
        const isRadial = color.includes('radial');
        
        if (isRadial) {
          defsContent = `
            <radialGradient id="${gradientId}" cx="50%" cy="50%" r="50%">
              <stop offset="0%" style="stop-color:${color1}"/>
              <stop offset="100%" style="stop-color:${color2}"/>
            </radialGradient>
          `;
        } else {
          defsContent = `
            <linearGradient id="${gradientId}" x1="0%" y1="0%" x2="100%" y2="0%">
              <stop offset="0%" style="stop-color:${color1}"/>
              <stop offset="100%" style="stop-color:${color2}"/>
            </linearGradient>
          `;
        }
        fillStyle = `fill="url(#${gradientId})"`;
      } else {
        fillStyle = `fill="${color}"`;
      }
      
      return `
        <svg width="${svgWidth}" height="${svgHeight}" style="overflow: visible;">
          <defs>
            ${defsContent}
            <path id="${pathId}" d="M ${startX},${baseY} Q ${midX},${curveY} ${endX},${baseY}" fill="none"/>
          </defs>
          <text ${fillStyle} font-size="${fontSize}" font-family="${fontFamily}" font-weight="${fontWeight}">
            <textPath href="#${pathId}" startOffset="50%" text-anchor="middle">
              ${text}
            </textPath>
          </text>
        </svg>
      `;
    }
    
    // משתנה גלובלי לסיבוב
    let rotatingWord = null;
    
    // משתנה גלובלי לקימור
    let curvingWord = null;
    
    function startWordRotate(e, stickerIndex, wordId, wordEl) {
      e.preventDefault();
      
      const sticker = stickers[stickerIndex];
      if (!sticker) return;
      
      const word = sticker.words.find(w => w.id === wordId);
      if (!word) return;
      
      // חישוב מרכז הטקסט
      const rect = wordEl.getBoundingClientRect();
      const centerX = rect.left + rect.width / 2;
      const centerY = rect.top + rect.height / 2;
      
      rotatingWord = {
        stickerIndex,
        wordId,
        wordEl,
        word,
        centerX,
        centerY,
        startAngle: word.rotation || 0,
        startX: e.clientX
      };
      
      pushHistory();
      
      document.addEventListener('mousemove', rotateWord);
      document.addEventListener('mouseup', stopWordRotate);
    }
    
    function rotateWord(e) {
      if (!rotatingWord) return;
      
      const { word, wordEl, startAngle, startX, stickerIndex, wordId } = rotatingWord;
      
      // חישוב הסיבוב לפי תנועה אופקית - כל 2 פיקסלים = מעלה אחת
      const deltaX = e.clientX - startX;
      const newRotation = startAngle + (deltaX * 0.5);
      
      // עדכון הסיבוב
      word.rotation = newRotation;
      wordEl.style.transform = `rotate(${newRotation}deg)`;
      
      // If sync is enabled, rotate corresponding words in other stickers
      if (syncMoveEnabled) {
        const seriesId = word.seriesId;
        if (seriesId) {
          stickers.forEach((sticker, idx) => {
            if (idx === stickerIndex) return;
            const correspondingWord = (sticker.words || []).find(w => w.seriesId === seriesId);
            if (!correspondingWord) return;
            correspondingWord.rotation = newRotation;

            const correspondingEl = document.querySelector(`[data-sticker-index="${idx}"] [data-word-id="${correspondingWord.id}"]`);
            if (correspondingEl) {
              correspondingEl.style.transform = `rotate(${newRotation}deg)`;
            }
          });
        } else {
          const wordIndex = stickers[stickerIndex].words.findIndex(w => w.id === wordId);
          stickers.forEach((sticker, idx) => {
            if (idx !== stickerIndex && sticker.words[wordIndex]) {
              const correspondingWord = sticker.words[wordIndex];
              correspondingWord.rotation = newRotation;

              const correspondingEl = document.querySelector(`[data-sticker-index="${idx}"] [data-word-id="${correspondingWord.id}"]`);
              if (correspondingEl) {
                correspondingEl.style.transform = `rotate(${newRotation}deg)`;
              }
            }
          });
        }
      }
    }
    
    function stopWordRotate() {
      if (rotatingWord) {
        renderStickers();
      }
      rotatingWord = null;
      document.removeEventListener('mousemove', rotateWord);
      document.removeEventListener('mouseup', stopWordRotate);
    }
    
    function startWordCurve(e, stickerIndex, wordId, wordEl) {
      e.preventDefault();
      
      const sticker = stickers[stickerIndex];
      if (!sticker) return;
      
      const word = sticker.words.find(w => w.id === wordId);
      if (!word) return;
      
      curvingWord = {
        stickerIndex,
        wordId,
        wordEl,
        word,
        startCurve: word.curve || 0,
        startY: e.clientY
      };
      
      pushHistory();
      
      document.addEventListener('mousemove', curveWord);
      document.addEventListener('mouseup', stopWordCurve);
    }
    
    function curveWord(e) {
      if (!curvingWord) return;
      
      const { word, wordEl, startCurve, startY, stickerIndex, wordId } = curvingWord;
      
      // חישוב הקימור לפי תנועה אנכית - כל פיקסל = יחידת קימור
      const deltaY = e.clientY - startY;
      const newCurve = startCurve + deltaY;
      
      // הגבלת הקימור לטווח סביר
      word.curve = Math.max(-100, Math.min(100, newCurve));
      
      // עדכון התצוגה
      if (word.curve !== 0) {
        wordEl.innerHTML = createCurvedText(word.text, word.curve, word);
        // הוספה מחדש של הכפתורים
        reattachWordControls(wordEl, curvingWord.stickerIndex, word.id);
      } else {
        // טקסט רגיל - שים ב-span נפרד
        const textSpan = document.createElement('span');
        textSpan.className = 'word-text-content';
        textSpan.textContent = word.text;
        wordEl.innerHTML = '';
        wordEl.appendChild(textSpan);
        reattachWordControls(wordEl, curvingWord.stickerIndex, word.id);
      }
      
      // If sync is enabled, curve corresponding words in other stickers
      if (syncMoveEnabled) {
        const seriesId = word.seriesId;
        if (seriesId) {
          stickers.forEach((sticker, idx) => {
            if (idx === stickerIndex) return;
            const correspondingWord = (sticker.words || []).find(w => w.seriesId === seriesId);
            if (!correspondingWord) return;
            correspondingWord.curve = word.curve;

            const correspondingEl = document.querySelector(`[data-sticker-index="${idx}"] [data-word-id="${correspondingWord.id}"]`);
            if (correspondingEl) {
              if (correspondingWord.curve !== 0) {
                correspondingEl.innerHTML = createCurvedText(correspondingWord.text, correspondingWord.curve, correspondingWord);
                reattachWordControls(correspondingEl, idx, correspondingWord.id);
              } else {
                const textSpan = document.createElement('span');
                textSpan.className = 'word-text-content';
                textSpan.textContent = correspondingWord.text;
                correspondingEl.innerHTML = '';
                correspondingEl.appendChild(textSpan);
                reattachWordControls(correspondingEl, idx, correspondingWord.id);
              }
            }
          });
        } else {
          const wordIndex = stickers[stickerIndex].words.findIndex(w => w.id === wordId);
          stickers.forEach((sticker, idx) => {
            if (idx !== stickerIndex && sticker.words[wordIndex]) {
              const correspondingWord = sticker.words[wordIndex];
              correspondingWord.curve = word.curve;

              const correspondingEl = document.querySelector(`[data-sticker-index="${idx}"] [data-word-id="${correspondingWord.id}"]`);
              if (correspondingEl) {
                if (correspondingWord.curve !== 0) {
                  correspondingEl.innerHTML = createCurvedText(correspondingWord.text, correspondingWord.curve, correspondingWord);
                  reattachWordControls(correspondingEl, idx, correspondingWord.id);
                } else {
                  const textSpan = document.createElement('span');
                  textSpan.className = 'word-text-content';
                  textSpan.textContent = correspondingWord.text;
                  correspondingEl.innerHTML = '';
                  correspondingEl.appendChild(textSpan);
                  reattachWordControls(correspondingEl, idx, correspondingWord.id);
                }
              }
            }
          });
        }
      }
    }
    
    function reattachWordControls(wordEl, stickerIndex, wordId) {
      // הסרת כפתורים קיימים
      const existingDelete = wordEl.querySelector('.delete-word-btn');
      const existingResize = wordEl.querySelector('.word-resize-handle');
      const existingRotate = wordEl.querySelector('.word-rotate-handle');
      const existingCurve = wordEl.querySelector('.word-curve-handle');
      if (existingDelete) existingDelete.remove();
      if (existingResize) existingResize.remove();
      if (existingRotate) existingRotate.remove();
      if (existingCurve) existingCurve.remove();
      
      // הוספת כפתורים מחדש
      const deleteBtn = document.createElement('button');
      deleteBtn.type = 'button';
      deleteBtn.className = 'delete-word-btn no-print';
      deleteBtn.textContent = '×';
      deleteBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        deleteWord(stickerIndex, wordId);
      });
      wordEl.appendChild(deleteBtn);

      const resizeHandle = document.createElement('div');
      resizeHandle.className = 'word-resize-handle no-print';
      resizeHandle.addEventListener('mousedown', (e) => {
        e.stopPropagation();
        startWordResize(e, stickerIndex, wordId);
      });
      wordEl.appendChild(resizeHandle);
      
      const rotateHandle = document.createElement('div');
      rotateHandle.className = 'word-rotate-handle no-print';
      rotateHandle.innerHTML = '↻';
      rotateHandle.addEventListener('mousedown', (e) => {
        e.stopPropagation();
        startWordRotate(e, stickerIndex, wordId, wordEl);
      });
      wordEl.appendChild(rotateHandle);
      
      const curveHandle = document.createElement('div');
      curveHandle.className = 'word-curve-handle no-print';
      curveHandle.innerHTML = '⌒';
      curveHandle.addEventListener('mousedown', (e) => {
        e.stopPropagation();
        startWordCurve(e, stickerIndex, wordId, wordEl);
      });
      wordEl.appendChild(curveHandle);
    }
    
    function stopWordCurve() {
      if (curvingWord) {
        renderStickers();
      }
      curvingWord = null;
      document.removeEventListener('mousemove', curveWord);
      document.removeEventListener('mouseup', stopWordCurve);
    }

    function showStatus(message, isError = false) {
      const statusDiv = document.getElementById('statusMessage');
      statusDiv.textContent = message;
      statusDiv.className = `mt-4 text-center text-sm font-medium ${isError ? 'text-red-600' : 'text-green-600'}`;
      
      statusDiv.classList.remove('hidden');
      
      setTimeout(() => {
        statusDiv.classList.add('hidden');
      }, 3000);
    }

    function getStickerLayoutConfigFromUI() {
      const uploadLimitInput = document.getElementById('uploadLimitInput');
      const stickersPerRowInput = document.getElementById('stickersPerRowInput');
      const stickerSizeModeSelect = document.getElementById('stickerSizeModeSelect');
      const edgeMarginInput = document.getElementById('edgeMarginInput');
      const gapInput = document.getElementById('gapInput');

      if (uploadLimitInput) {
        const v = Number(uploadLimitInput.value);
        stickerLayoutConfig.uploadLimit = Number.isFinite(v) && v >= 0 ? Math.floor(v) : 0;
      }
      if (stickersPerRowInput) {
        const v = Number(stickersPerRowInput.value);
        stickerLayoutConfig.stickersPerRow = Math.max(1, Math.min(30, Number.isFinite(v) ? Math.floor(v) : 2));
      }
      if (stickerSizeModeSelect) {
        const m = String(stickerSizeModeSelect.value || '').toLowerCase();
        stickerLayoutConfig.sizeMode = (m === 'height') ? 'height' : 'width';
      }
      if (edgeMarginInput) {
        const v = Number(edgeMarginInput.value);
        stickerLayoutConfig.edgeMargin = Math.max(0, Math.min(400, Number.isFinite(v) ? v : 20));
      }
      if (gapInput) {
        const v = Number(gapInput.value);
        stickerLayoutConfig.gap = Math.max(0, Math.min(400, Number.isFinite(v) ? v : 20));
      }

      return stickerLayoutConfig;
    }

    function applyStickerLayoutConfigToUI() {
      const uploadLimitInput = document.getElementById('uploadLimitInput');
      const stickersPerRowInput = document.getElementById('stickersPerRowInput');
      const stickerSizeModeSelect = document.getElementById('stickerSizeModeSelect');
      const edgeMarginInput = document.getElementById('edgeMarginInput');
      const gapInput = document.getElementById('gapInput');

      if (uploadLimitInput) uploadLimitInput.value = String(stickerLayoutConfig.uploadLimit ?? 0);
      if (stickersPerRowInput) stickersPerRowInput.value = String(stickerLayoutConfig.stickersPerRow ?? 2);
      if (stickerSizeModeSelect) stickerSizeModeSelect.value = String(stickerLayoutConfig.sizeMode || 'width');
      if (edgeMarginInput) edgeMarginInput.value = String(stickerLayoutConfig.edgeMargin ?? 1);
      if (gapInput) gapInput.value = String(stickerLayoutConfig.gap ?? 1);

      updateStickerLayoutInfo();
    }

    function updateStickerLayoutInfo() {
      const info = document.getElementById('stickerLayoutInfo');
      if (!info) return;

      const cfg = stickerLayoutConfig;
      const pageWidth = pageOrientation === 'landscape' ? 297 * MM_TO_PX : 210 * MM_TO_PX;
      const pageHeight = pageOrientation === 'landscape' ? 210 * MM_TO_PX : 297 * MM_TO_PX;
      const contentWidth = Math.max(1, pageWidth - (cfg.edgeMargin * 2));
      const gap = Math.max(0, cfg.gap);
      const contentHeight = Math.max(1, pageHeight - (cfg.edgeMargin * 2));

      const mode = (cfg.sizeMode === 'height') ? 'height' : 'width';
      if (mode === 'height') {
        const rows = Math.max(1, cfg.stickersPerRow);
        const cellHeight = Math.max(1, (contentHeight - gap * (rows - 1)) / rows);
        const cellHeightCM = (cellHeight / MM_TO_PX / 10).toFixed(1);
        info.textContent = `יוצא בערך: ${rows} מדבקות לאורך העמוד (גובה כל מדבקה: ~${cellHeightCM} ס"מ, רוחב לפי יחס המדבקה)`;
      } else {
        const cols = Math.max(1, cfg.stickersPerRow);
        const cellWidth = Math.max(1, (contentWidth - gap * (cols - 1)) / cols);
        const cellWidthCM = (cellWidth / MM_TO_PX / 10).toFixed(1);
        info.textContent = `יוצא בערך: ${cols} מדבקות ברוחב (רוחב כל מדבקה: ~${cellWidthCM} ס"מ)`;
      }
    }
    
    // פונקציה חדשה לחישוב כמה מדבקות צריך להוסיף בפועל
    function calculateActualStickerCount(requestedCount) {
      const cfg = stickerLayoutConfig;
      const mode = (cfg.sizeMode === 'height') ? 'height' : 'width';
      
      // אם לא הוזן מספר ספציפי, מחזירים 0 (כלומר - כל הקבצים)
      if (!requestedCount || requestedCount <= 0) {
        return 0;
      }
      
      // במצב "לאורך" - המספר שהוזן הוא בדיוק מספר המדבקות שצריך להוסיף
      // (כי הן מסודרות בעמודה אחת)
      return requestedCount;
    }

    function applyStickerLayoutAndRender() {
      getStickerLayoutConfigFromUI();
      updateStickerLayoutInfo();
      // מחילים שינויי גודל ומיקום על כל המדבקות
      // תמיד קוראים ל-reflowStickers כדי לעדכן גדלים לפי המצב החדש
      if (stickers.length > 0) {
        reflowStickers();
      }
      renderStickers();
      updateFileCount();
    }

    // פונקציה חדשה שמחילה פריסה רק על מדבקות חדשות (RTL - מימין לשמאל)
    function applyLayoutToNewStickers(startIndex = 0) {
      const cfg = stickerLayoutConfig;
      const pageWidth = pageOrientation === 'landscape' ? 297 * MM_TO_PX : 210 * MM_TO_PX;
      const pageHeight = pageOrientation === 'landscape' ? 210 * MM_TO_PX : 297 * MM_TO_PX;
      const edge = Math.max(0, cfg.edgeMargin);
      const gap = Math.max(0, cfg.gap);

      const contentWidth = Math.max(1, pageWidth - (edge * 2));
      const contentHeight = Math.max(1, pageHeight - (edge * 2));
      const maxH = contentHeight;

      const mode = (cfg.sizeMode === 'height') ? 'height' : 'width';
      const cols = (mode === 'width')
        ? Math.max(1, cfg.stickersPerRow)
        : 1; // במצב לאורך - תמיד עמודה אחת
      const cellWidth = (mode === 'width')
        ? Math.max(1, (contentWidth - gap * (cols - 1)) / cols)
        : contentWidth; // במצב לאורך - הרוחב המלא הזמין
      const rows = (mode === 'height')
        ? Math.max(1, cfg.stickersPerRow)
        : 1;
      const cellHeight = (mode === 'height')
        ? Math.max(1, (contentHeight - gap * (rows - 1)) / rows)
        : contentHeight;

      // חישוב מיקום התחלתי בהתבסס על מדבקות קיימות (RTL)
      let page = 0;
      let colIndex = 0;
      let y = edge;
      let rowMaxHeight = 0;

      // מציאת המיקום הנוכחי בהתבסס על מדבקות קיימות
      if (startIndex > 0) {
        // מציאת העמוד והשורה האחרונים
        let lastPage = 0;
        let lastY = edge;
        let lastRowMaxHeight = 0;
        let itemsInCurrentRow = 0;
        
        for (let i = 0; i < startIndex && i < stickers.length; i++) {
          const sticker = stickers[i];
          if (sticker.page > lastPage) {
            lastPage = sticker.page;
            lastY = edge;
            lastRowMaxHeight = 0;
            itemsInCurrentRow = 0;
          }
          
          if (sticker.page === lastPage) {
            // בדיקה אם זו שורה חדשה
            if (Math.abs(sticker.y - lastY) > 1) {
              lastY = sticker.y;
              lastRowMaxHeight = sticker.height;
              itemsInCurrentRow = 1;
            } else {
              itemsInCurrentRow++;
              lastRowMaxHeight = Math.max(lastRowMaxHeight, sticker.height);
            }
          }
        }
        
        page = lastPage;
        y = lastY;
        colIndex = itemsInCurrentRow;
        rowMaxHeight = lastRowMaxHeight;
        
        // אם השורה מלאה, עוברים לשורה הבאה
        if (colIndex >= cols) {
          colIndex = 0;
          y = y + rowMaxHeight + gap;
          rowMaxHeight = 0;
        }
      }

      // החלת פריסה רק על מדבקות חדשות (RTL)
      for (let i = startIndex; i < stickers.length; i++) {
        const sticker = stickers[i];
        sticker.words = sticker.words || [];
        sticker.images = sticker.images || [];

        // מדבקות מדויקות - שמירה על הגודל המקורי במילימטרים (עם כיול אם קיים)
        if (sticker.precisionCut && sticker.precisionWidthMM && sticker.precisionHeightMM) {
          // החלת כיול מדפסת אם קיים
          let precisionW, precisionH;
          if (typeof getCalibratedPxFromMM === 'function') {
            const calibrated = getCalibratedPxFromMM(sticker.precisionWidthMM, sticker.precisionHeightMM, MM_TO_PX);
            precisionW = calibrated.widthPx;
            precisionH = calibrated.heightPx;
          } else {
            precisionW = sticker.precisionWidthMM * MM_TO_PX;
            precisionH = sticker.precisionHeightMM * MM_TO_PX;
          }
          
          if (colIndex >= cols) {
            colIndex = 0;
            y = y + rowMaxHeight + gap;
            rowMaxHeight = 0;
          }

          if (y + precisionH > pageHeight - edge) {
            page += 1;
            colIndex = 0;
            y = edge;
            rowMaxHeight = 0;
          }

          const xPos = pageWidth - edge - (colIndex * (cellWidth + gap)) - cellWidth + Math.max(0, (cellWidth - precisionW) / 2);

          sticker.page = page;
          sticker.x = xPos;
          sticker.y = y;
          sticker.width = precisionW;
          sticker.height = precisionH;
          sticker.originalWidth = precisionW;
          sticker.originalHeight = precisionH;

          rowMaxHeight = Math.max(rowMaxHeight, precisionH);
          colIndex += 1;
          continue; // מעבר למדבקה הבאה
        }

        if (!Number.isFinite(sticker.originalWidth) || !Number.isFinite(sticker.originalHeight) || sticker.originalWidth <= 0 || sticker.originalHeight <= 0) {
          const fallbackCell = cellWidth;
          sticker.originalWidth = fallbackCell;
          sticker.originalHeight = fallbackCell;
        }

        const aspectRatio = sticker.originalWidth / sticker.originalHeight;

        let newW;
        let newH;
        if (mode === 'height') {
          newH = cellHeight;
          newW = newH * aspectRatio;
          if (newW > cellWidth) {
            newW = cellWidth;
            newH = newW / aspectRatio;
          }
        } else {
          newW = cellWidth;
          newH = newW / aspectRatio;
          if (newH > maxH) {
            newH = maxH;
            newW = newH * aspectRatio;
          }
        }

        if (colIndex >= cols) {
          colIndex = 0;
          y = y + (mode === 'height' ? cellHeight : rowMaxHeight) + gap;
          rowMaxHeight = 0;
        }

        if (y + (mode === 'height' ? cellHeight : newH) > pageHeight - edge) {
          page += 1;
          colIndex = 0;
          y = edge;
          rowMaxHeight = 0;
        }

        // חישוב X מימין לשמאל (RTL)
        const xPos = pageWidth - edge - (colIndex * (cellWidth + gap)) - cellWidth + Math.max(0, (cellWidth - newW) / 2);

        sticker.page = page;
        sticker.x = xPos;
        sticker.y = y;
        sticker.width = newW;
        sticker.height = newH;

        rowMaxHeight = Math.max(rowMaxHeight, newH);
        colIndex += 1;
      }
    }

    function readImageFileAsDataUrlWithSize(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onerror = () => reject(new Error('Failed to read image'));
        reader.onload = (event) => {
          const dataUrl = event.target.result;
          const img = new Image();
          img.onload = () => resolve({
            dataUrl,
            fileName: file.name,
            originalWidth: img.width,
            originalHeight: img.height
          });
          img.onerror = () => reject(new Error('Failed to load image'));
          img.src = dataUrl;
        };
        reader.readAsDataURL(file);
      });
    }

    function reflowStickers() {
      const cfg = stickerLayoutConfig;
      
      console.log('=== reflowStickers START ===');
      console.log('Total stickers:', stickers.length);
      console.log('Config:', JSON.stringify(cfg));
      console.log('pageOrientation:', pageOrientation);
      
      const pageWidth = pageOrientation === 'landscape' ? 297 * MM_TO_PX : 210 * MM_TO_PX;
      const pageHeight = pageOrientation === 'landscape' ? 210 * MM_TO_PX : 297 * MM_TO_PX;
      const edge = Math.max(0, cfg.edgeMargin);
      const gap = Math.max(0, cfg.gap);

      const contentWidth = Math.max(1, pageWidth - (edge * 2));
      const contentHeight = Math.max(1, pageHeight - (edge * 2));
      const maxH = contentHeight;

      const mode = (cfg.sizeMode === 'height') ? 'height' : 'width';
      
      console.log('=== reflowStickers DEBUG ===');
      console.log('Mode:', mode);
      console.log('stickersPerRow:', cfg.stickersPerRow);
      console.log('pageWidth:', pageWidth, 'pageHeight:', pageHeight);
      console.log('contentWidth:', contentWidth, 'contentHeight:', contentHeight);
      
      const cols = (mode === 'width')
        ? Math.max(1, cfg.stickersPerRow)
        : 1; // במצב לאורך - תמיד עמודה אחת
      const cellWidth = (mode === 'width')
        ? Math.max(1, (contentWidth - gap * (cols - 1)) / cols)
        : contentWidth; // במצב לאורך - הרוחב המלא הזמין
      const rows = (mode === 'height')
        ? Math.max(1, cfg.stickersPerRow)
        : 1;
      const cellHeight = (mode === 'height')
        ? Math.max(1, (contentHeight - gap * (rows - 1)) / rows)
        : contentHeight;
        
      console.log('cols:', cols, 'rows:', rows);
      console.log('cellWidth:', cellWidth, 'cellHeight:', cellHeight);

      const derivedCellWidth = cellWidth;
      const derivedCols = cols;

      let page = 0;
      let colIndex = 0;
      let y = edge;
      let rowMaxHeight = 0;

      stickers.forEach((sticker) => {
        sticker.words = sticker.words || [];
        sticker.images = sticker.images || [];

        // מדבקות מדויקות - שמירה על הגודל המקורי במילימטרים (עם כיול אם קיים)
        if (sticker.precisionCut && sticker.precisionWidthMM && sticker.precisionHeightMM) {
          // החלת כיול מדפסת אם קיים
          let precisionW, precisionH;
          if (typeof getCalibratedPxFromMM === 'function') {
            const calibrated = getCalibratedPxFromMM(sticker.precisionWidthMM, sticker.precisionHeightMM, MM_TO_PX);
            precisionW = calibrated.widthPx;
            precisionH = calibrated.heightPx;
          } else {
            precisionW = sticker.precisionWidthMM * MM_TO_PX;
            precisionH = sticker.precisionHeightMM * MM_TO_PX;
          }
          
          // בדיקה אם צריך לעבור לעמוד חדש
          if (colIndex >= derivedCols) {
            colIndex = 0;
            y = y + rowMaxHeight + gap;
            rowMaxHeight = 0;
          }

          if (y + precisionH > pageHeight - edge) {
            page += 1;
            colIndex = 0;
            y = edge;
            rowMaxHeight = 0;
          }

          const usedCellWidth = (mode === 'height') ? derivedCellWidth : cellWidth;
          const xPos = pageWidth - edge - (colIndex * (usedCellWidth + gap)) - usedCellWidth + Math.max(0, (usedCellWidth - precisionW) / 2);

          sticker.page = page;
          sticker.x = xPos;
          sticker.y = y;
          // שמירה על הגודל המדויק - לא משנים!
          sticker.width = precisionW;
          sticker.height = precisionH;
          sticker.originalWidth = precisionW;
          sticker.originalHeight = precisionH;

          rowMaxHeight = Math.max(rowMaxHeight, precisionH);
          colIndex += 1;
          return; // מעבר למדבקה הבאה
        }

        if (!Number.isFinite(sticker.originalWidth) || !Number.isFinite(sticker.originalHeight) || sticker.originalWidth <= 0 || sticker.originalHeight <= 0) {
          const fallbackCell = (mode === 'height') ? cellHeight : cellWidth;
          const fallbackW = Number.isFinite(sticker.width) && sticker.width > 0 ? sticker.width : fallbackCell;
          const fallbackH = Number.isFinite(sticker.height) && sticker.height > 0 ? sticker.height : fallbackCell;
          sticker.originalWidth = fallbackW;
          sticker.originalHeight = fallbackH;
        }

        const fallbackCell = (mode === 'height') ? cellHeight : cellWidth;
        const prevW = Number.isFinite(sticker.width) && sticker.width > 0 ? sticker.width : fallbackCell;
        const prevH = Number.isFinite(sticker.height) && sticker.height > 0 ? sticker.height : fallbackCell;
        const aspectRatio = sticker.originalWidth / sticker.originalHeight;

        let newW;
        let newH;
        if (mode === 'height') {
          // במצב לאורך - קובעים את הגובה ומחשבים את הרוחב לפי יחס גובה-רוחב
          newH = cellHeight;
          newW = newH * aspectRatio;
          console.log('Height mode - cellHeight:', cellHeight, 'aspectRatio:', aspectRatio, 'newW:', newW, 'newH:', newH);
          // אם הרוחב גדול מדי - מקטינים את שניהם
          if (newW > contentWidth) {
            newW = contentWidth;
            newH = newW / aspectRatio;
            console.log('Width too large, adjusted to:', newW, newH);
          }
        } else {
          newW = cellWidth;
          newH = newW / aspectRatio;
          if (newH > maxH) {
            newH = maxH;
            newW = newH * aspectRatio;
          }
        }

        if (colIndex >= derivedCols) {
          colIndex = 0;
          y = y + (mode === 'height' ? cellHeight : rowMaxHeight) + gap;
          rowMaxHeight = 0;
        }

        if (y + (mode === 'height' ? cellHeight : newH) > pageHeight - edge) {
          page += 1;
          colIndex = 0;
          y = edge;
          rowMaxHeight = 0;
        }

        const scale = prevW > 0 ? (newW / prevW) : 1;

        // חישוב X מימין לשמאל (RTL)
        let xPos;
        if (mode === 'height') {
          // במצב לאורך - ממרכזים את המדבקה או מיישרים לימין
          xPos = pageWidth - edge - newW; // יישור לימין
        } else {
          const usedCellWidth = cellWidth;
          xPos = pageWidth - edge - (colIndex * (usedCellWidth + gap)) - usedCellWidth + Math.max(0, (usedCellWidth - newW) / 2);
        }

        sticker.page = page;
        sticker.x = xPos;
        sticker.y = y;
        sticker.width = newW;
        sticker.height = newH;

        sticker.words.forEach(w => {
          if (Number.isFinite(w.x)) w.x = w.x * scale;
          if (Number.isFinite(w.y)) w.y = w.y * scale;
          if (Number.isFinite(w.fontSize)) w.fontSize = w.fontSize * scale;
        });

        sticker.images.forEach(img => {
          if (Number.isFinite(img.x)) img.x = img.x * scale;
          if (Number.isFinite(img.y)) img.y = img.y * scale;
          if (Number.isFinite(img.width)) img.width = img.width * scale;
          if (Number.isFinite(img.height)) img.height = img.height * scale;
        });

        rowMaxHeight = Math.max(rowMaxHeight, newH);
        colIndex += 1;
      });
    }

    function reflowStickersPositionsOnly() {
      const cfg = stickerLayoutConfig;
      const pageWidth = pageOrientation === 'landscape' ? 297 * MM_TO_PX : 210 * MM_TO_PX;
      const pageHeight = pageOrientation === 'landscape' ? 210 * MM_TO_PX : 297 * MM_TO_PX;
      const edge = Math.max(0, cfg.edgeMargin);
      const gap = Math.max(0, cfg.gap);

      const contentWidth = Math.max(1, pageWidth - (edge * 2));
      const contentHeight = Math.max(1, pageHeight - (edge * 2));

      // וידוא שלכל מדבקה יש גודל תקין והתאמה לגבולות הדף
      stickers.forEach((sticker) => {
        if (!sticker) return;
        sticker.words = sticker.words || [];
        sticker.images = sticker.images || [];

        // מדבקות מדויקות - שמירה על הגודל המקורי במילימטרים (עם כיול אם קיים)
        if (sticker.precisionCut && sticker.precisionWidthMM && sticker.precisionHeightMM) {
          // החלת כיול מדפסת אם קיים
          if (typeof getCalibratedPxFromMM === 'function') {
            const calibrated = getCalibratedPxFromMM(sticker.precisionWidthMM, sticker.precisionHeightMM, MM_TO_PX);
            sticker.width = calibrated.widthPx;
            sticker.height = calibrated.heightPx;
          } else {
            sticker.width = sticker.precisionWidthMM * MM_TO_PX;
            sticker.height = sticker.precisionHeightMM * MM_TO_PX;
          }
          sticker.originalWidth = sticker.width;
          sticker.originalHeight = sticker.height;
          return; // לא משנים את הגודל של מדבקות מדויקות
        }

        if (!Number.isFinite(sticker.width) || sticker.width <= 1) {
          sticker.width = 100;
        }
        if (!Number.isFinite(sticker.height) || sticker.height <= 1) {
          sticker.height = 100;
        }
        
        // הקטנת מדבקה שגדולה מדי לרוחב הדף
        if (sticker.width > contentWidth) {
          const scale = contentWidth / sticker.width;
          sticker.width = contentWidth;
          sticker.height = sticker.height * scale;
        }
        
        // הקטנת מדבקה שגדולה מדי לגובה הדף
        if (sticker.height > contentHeight) {
          const scale = contentHeight / sticker.height;
          sticker.height = contentHeight;
          sticker.width = sticker.width * scale;
        }
      });

      // אלגוריתם סידור חכם - מימין לשמאל (RTL)
      let pages = [{ rows: [] }];

      stickers.forEach((sticker) => {
        if (!sticker) return;

        const w = sticker.width;
        const h = sticker.height;

        let placed = false;

        // חיפוש מקום בעמודים קיימים
        for (let pageIndex = 0; pageIndex < pages.length && !placed; pageIndex++) {
          const page = pages[pageIndex];

          // ניסיון למצוא מקום בשורה קיימת
          for (let rowIndex = 0; rowIndex < page.rows.length && !placed; rowIndex++) {
            const row = page.rows[rowIndex];

            // חישוב הרוחב התפוס בשורה
            let rowUsedWidth = 0;
            row.items.forEach((item, idx) => {
              rowUsedWidth += item.width;
              if (idx < row.items.length - 1) rowUsedWidth += gap;
            });

            // בדיקה אם יש מקום בשורה הזו (מימין לשמאל)
            const spaceNeeded = rowUsedWidth > 0 ? w + gap : w;
            const remainingWidth = contentWidth - rowUsedWidth - (rowUsedWidth > 0 ? gap : 0);
            
            // וידוא שהמדבקה לא תחרוג מגבול שמאל
            if (remainingWidth >= w) {
              // בדיקה אם הגובה מתאים
              const newRowHeight = Math.max(row.height, h);
              
              // חישוב גובה כל השורות עד לשורה הנוכחית כולל
              let totalHeightWithNewRow = 0;
              for (let i = 0; i < page.rows.length; i++) {
                if (i === rowIndex) {
                  totalHeightWithNewRow += newRowHeight;
                } else {
                  totalHeightWithNewRow += page.rows[i].height;
                }
                if (i < page.rows.length - 1) totalHeightWithNewRow += gap;
              }
              
              if (totalHeightWithNewRow <= contentHeight) {
                // מצאנו מקום! חישוב X מימין לשמאל
                const xPos = pageWidth - edge - rowUsedWidth - (rowUsedWidth > 0 ? gap : 0) - w;
                sticker.page = pageIndex;
                sticker.x = xPos;
                sticker.y = row.y;
                row.items.push(sticker);
                row.height = newRowHeight;
                placed = true;
              }
            }
          }

          // אם לא מצאנו מקום בשורות קיימות, ננסה ליצור שורה חדשה
          if (!placed) {
            let totalRowsHeight = 0;
            page.rows.forEach((row, idx) => {
              totalRowsHeight += row.height;
              if (idx < page.rows.length - 1) totalRowsHeight += gap;
            });

            const newRowY = edge + totalRowsHeight + (page.rows.length > 0 ? gap : 0);

            // בדיקה אם יש מקום לשורה חדשה בעמוד
            if (newRowY + h <= pageHeight - edge) {
              // מיקום X מימין לשמאל - מתחילים מימין
              const xPos = pageWidth - edge - w;
              const newRow = {
                y: newRowY,
                height: h,
                items: [sticker]
              };
              page.rows.push(newRow);
              sticker.page = pageIndex;
              sticker.x = xPos;
              sticker.y = newRowY;
              placed = true;
            }
          }
        }

        // אם לא מצאנו מקום באף עמוד, ניצור עמוד חדש
        if (!placed) {
          const newPageIndex = pages.length;
          // מיקום X מימין לשמאל - מתחילים מימין
          const xPos = pageWidth - edge - w;
          const newRow = {
            y: edge,
            height: h,
            items: [sticker]
          };
          pages.push({ rows: [newRow] });
          sticker.page = newPageIndex;
          sticker.x = xPos;
          sticker.y = edge;
        }
      });

      // עדכון מיקום Y של כל המדבקות לפי גובה השורות הסופי
      pages.forEach((page, pageIndex) => {
        let currentY = edge;
        page.rows.forEach((row) => {
          row.items.forEach((sticker) => {
            sticker.y = currentY;
          });
          currentY += row.height + gap;
        });
      });
    }

    function compactStickers() {
      const edgePadding = Math.max(0, stickerLayoutConfig.edgeMargin);
      const gapPadding = Math.max(0, stickerLayoutConfig.gap);
      const pageWidth = pageOrientation === 'landscape' ? 297 * MM_TO_PX : 210 * MM_TO_PX;
      const pageHeight = pageOrientation === 'landscape' ? 210 * MM_TO_PX : 297 * MM_TO_PX;

      for (let i = 0; i < stickers.length; i++) {
        const s = stickers[i];
        if (!s) continue;

        const w = Number.isFinite(s.width) && s.width > 0 ? s.width : 100;
        const h = Number.isFinite(s.height) && s.height > 0 ? s.height : 100;

        s.lockPosition = false;
        s.page = -1;

        const pos = findNextAvailablePosition(w, h, pageWidth, pageHeight, edgePadding, gapPadding, 0, edgePadding, edgePadding);
        if (!pos) continue;

        s.page = pos.page;
        s.x = pos.x;
        s.y = pos.y;
        s.lockPosition = true;
      }
    }

    function selectSticker(index) {
      selectedSticker = index;
      selectedWord = null;
      selectedImage = null;
      
      document.querySelectorAll('.sticker-container').forEach(s => s.classList.remove('selected'));
      document.querySelectorAll('.text-word').forEach(w => w.classList.remove('selected'));
      document.querySelectorAll('.sticker-image').forEach(i => i.classList.remove('selected'));
      
      const stickerEl = document.querySelector(`[data-sticker-index="${index}"]`);
      if (stickerEl) {
        stickerEl.classList.add('selected');
      }
      
      // Update opacity control
      updateOpacityControl('sticker', index);
      
      showStatus(`מדבקה ${index + 1} נבחרה`);
    }

    function deleteStickerByIndex(index) {
      const fileName = stickers[index].fileName || `מדבקה ${index + 1}`;

      pushHistory();
      
      stickers.splice(index, 1);
      selectedSticker = null;
      selectedWord = null;
      selectedImage = null;
      
      if (autoArrangeEnabled) {
        getStickerLayoutConfigFromUI();
        updateStickerLayoutInfo();
        reflowStickersPositionsOnly();
      }
      renderStickers();
      updateFileCount();
      showStatus(`המדבקה "${fileName}" נמחקה בהצלחה!`);
    }

    function duplicateSticker(index) {
      const originalSticker = stickers[index];

      pushHistory();
      
      const edgePadding = Math.max(0, stickerLayoutConfig.edgeMargin);
      const gapPadding = Math.max(0, stickerLayoutConfig.gap);
      const pageWidth = pageOrientation === 'landscape' ? 297 * MM_TO_PX : 210 * MM_TO_PX;
      const pageHeight = pageOrientation === 'landscape' ? 210 * MM_TO_PX : 297 * MM_TO_PX;
      
      // Find next available position that doesn't overlap
      const startPage = Number.isFinite(originalSticker.page) ? originalSticker.page : 0;
      const startX = (Number.isFinite(originalSticker.x) ? originalSticker.x : edgePadding) + (Number.isFinite(originalSticker.width) ? originalSticker.width : 0) + gapPadding;
      const startY = Number.isFinite(originalSticker.y) ? originalSticker.y : edgePadding;
      const newPosition = findNextAvailablePosition(originalSticker.width, originalSticker.height, pageWidth, pageHeight, edgePadding, gapPadding, startPage, startX, startY);
      
      if (!newPosition) {
        showStatus('אין מספיק מקום לשכפול המדבקה!', true);
        return;
      }
      
      // Create a deep copy of the sticker
      const duplicatedSticker = {
        id: `sticker-${Date.now()}-duplicate`,
        dataUrl: originalSticker.dataUrl,
        fileName: originalSticker.fileName ? `${originalSticker.fileName} (עותק)` : `מדבקה ${stickers.length + 1}`,
        page: newPosition.page,
        x: newPosition.x,
        y: newPosition.y,
        width: originalSticker.width,
        height: originalSticker.height,
        originalWidth: originalSticker.originalWidth,
        originalHeight: originalSticker.originalHeight,
        opacity: originalSticker.opacity,
        lockSize: !!originalSticker.lockSize,
        lockPosition: true,
        words: [],
        images: []
      };
      
      // Copy words with new IDs but keep them synchronized
      originalSticker.words.forEach(word => {
        const newWord = {
          id: `word-${++wordIdCounter}`,
          seriesId: word.seriesId,
          text: word.text,
          x: word.x,
          y: word.y,
          color: word.color,
          fontSize: word.fontSize,
          fontFamily: word.fontFamily,
          fontWeight: word.fontWeight,
          opacity: word.opacity,
          rotation: word.rotation,
          curve: word.curve,
          isGradient: word.isGradient
        };
        duplicatedSticker.words.push(newWord);
      });
      
      // Copy images with new IDs but keep them synchronized
      if (originalSticker.images) {
        originalSticker.images.forEach(image => {
          const newImage = {
            id: `image-${++imageIdCounter}`,
            seriesId: image.seriesId,
            dataUrl: image.dataUrl,
            x: image.x,
            y: image.y,
            width: image.width,
            height: image.height,
            originalWidth: image.originalWidth,
            originalHeight: image.originalHeight,
            opacity: image.opacity
          };
          duplicatedSticker.images.push(newImage);
        });
      }
      
      // Add to end of stickers array
      stickers.push(duplicatedSticker);

      if (autoArrangeEnabled) {
        getStickerLayoutConfigFromUI();
        updateStickerLayoutInfo();
        reflowStickersPositionsOnly();
      }
      renderStickers();
      updateFileCount();
      showStatus(`המדבקה שוכפלה!`);
    }

    function findNextAvailablePosition(stickerWidth, stickerHeight, pageWidth, pageHeight, edgePadding, gapPadding, startPage = 0, startX = edgePadding, startY = edgePadding) {
      const maxPages = 20; // Increase to 20 pages
      
      const startPageIndex = Math.max(0, Math.min(maxPages - 1, Number(startPage) || 0));
      for (let pageIndex = startPageIndex; pageIndex < maxPages; pageIndex++) {
        // Get all stickers on this page
        const stickersOnPage = stickers.filter(s => (s.page || 0) === pageIndex);
        
        // If this is an empty page, start from top-left
        if (stickersOnPage.length === 0) {
          // Check if sticker fits on page
          if (stickerWidth + edgePadding * 2 <= pageWidth && stickerHeight + edgePadding * 2 <= pageHeight) {
            return {
              page: pageIndex,
              x: edgePadding,
              y: edgePadding
            };
          } else {
            // Sticker too big for any page
            return null;
          }
        }
        
        // Try different starting positions on this page
        const gridSize = 20;

        const snapToGrid = (v) => Math.max(edgePadding, Math.ceil(v / gridSize) * gridSize);
        const initialY = pageIndex === startPageIndex ? snapToGrid(startY) : edgePadding;
        
        for (let testY = initialY; testY + stickerHeight + edgePadding <= pageHeight; testY += gridSize) {
          const rowStartX = (pageIndex === startPageIndex && testY === initialY) ? snapToGrid(startX) : edgePadding;
          for (let testX = rowStartX; testX + stickerWidth + edgePadding <= pageWidth; testX += gridSize) {
            const proposedRect = {
              x: testX,
              y: testY,
              width: stickerWidth,
              height: stickerHeight
            };
            
            // Check if this position overlaps with any existing sticker
            const hasOverlap = stickersOnPage.some(sticker => {
              return rectanglesOverlap(proposedRect, {
                x: sticker.x,
                y: sticker.y,
                width: sticker.width,
                height: sticker.height
              }, gapPadding);
            });
            
            if (!hasOverlap) {
              // Found a valid position!
              return {
                page: pageIndex,
                x: testX,
                y: testY
              };
            }
          }
        }
        
        // No space found on this page, continue to next page
      }
      
      // No position found on any page
      return null;
    }

    function rectanglesOverlap(rect1, rect2, padding) {
      // Add padding to the comparison
      return !(
        rect1.x + rect1.width + padding <= rect2.x ||
        rect2.x + rect2.width + padding <= rect1.x ||
        rect1.y + rect1.height + padding <= rect2.y ||
        rect2.y + rect2.height + padding <= rect1.y
      );
    }

    function startStickerResize(e, stickerIndex) {
      e.stopPropagation();

      pushHistory();
      
      const sticker = stickers[stickerIndex];
      const stickerEl = document.querySelector(`[data-sticker-index="${stickerIndex}"]`);
      
      resizingSticker = { 
        stickerIndex,
        startX: e.clientX,
        startY: e.clientY,
        startWidth: sticker.width,
        startHeight: sticker.height,
        startTop: sticker.y, // Save original Y position
        element: stickerEl
      };
      
      document.addEventListener('mousemove', resizeSticker);
      document.addEventListener('mouseup', stopStickerResize);
    }

    function resizeSticker(e) {
      if (!resizingSticker) return;
      
      const { stickerIndex, startX, startY, startWidth, startHeight, startTop } = resizingSticker;
      const sticker = stickers[stickerIndex];
      if (!sticker) return;
      
      const stickerEl = document.querySelector(`[data-sticker-index="${stickerIndex}"]`);
      if (!stickerEl) return;
      
      // Calculate delta from start position
      const deltaX = e.clientX - startX;
      
      // Calculate new width based on horizontal delta only
      const newWidth = Math.max(50, startWidth + deltaX);
      
      // Keep aspect ratio based on original dimensions
      const aspectRatio = startWidth / startHeight;
      const calculatedHeight = newWidth / aspectRatio;
      
      // Update sticker dimensions
      sticker.width = newWidth;
      sticker.height = calculatedHeight;

      sticker.lockSize = true;
      
      // CRITICAL: Lock Y position - never change it during resize
      sticker.y = startTop;
      
      // Update visual appearance
      stickerEl.style.width = `${newWidth}px`;
      stickerEl.style.height = `${calculatedHeight}px`;
      stickerEl.style.top = `${startTop}px`;
      stickerEl.style.left = `${sticker.x}px`;
      
      // עדכון תצוגת המידות בזמן אמת
      const widthLabel = stickerEl.querySelector('.sticker-dimension-width');
      const heightLabel = stickerEl.querySelector('.sticker-dimension-height');
      if (widthLabel) {
        const widthInCm = (newWidth / MM_TO_PX / 10).toFixed(1);
        widthLabel.textContent = `${widthInCm} ס"מ`;
      }
      if (heightLabel) {
        const heightInCm = (calculatedHeight / MM_TO_PX / 10).toFixed(1);
        heightLabel.textContent = `${heightInCm} ס"מ`;
      }
    }

    function stopStickerResize() {
      resizingSticker = null;
      document.removeEventListener('mousemove', resizeSticker);
      document.removeEventListener('mouseup', stopStickerResize);
    }

    function deleteSelectedSticker() {
      if (selectedSticker === null) {
        showStatus('בחר מדבקה למחיקה! לחץ על מדבקה כדי לבחור אותה.', true);
        return;
      }
      
      const stickerIndex = selectedSticker;
      const fileName = stickers[stickerIndex].fileName || `מדבקה ${stickerIndex + 1}`;

      pushHistory();
      
      stickers.splice(stickerIndex, 1);
      selectedSticker = null;
      selectedWord = null;
      selectedImage = null;
      
      if (autoArrangeEnabled) {
        getStickerLayoutConfigFromUI();
        updateStickerLayoutInfo();
        reflowStickersPositionsOnly();
      }
      renderStickers();
      updateFileCount();
      showStatus(`המדבקה "${fileName}" נמחקה בהצלחה!`);
    }

    function selectWord(stickerIndex, wordId) {
      selectedSticker = stickerIndex;
      selectedWord = wordId;
      selectedImage = null;
      
      document.querySelectorAll('.sticker-container').forEach(s => s.classList.remove('selected'));
      document.querySelectorAll('.text-word').forEach(w => w.classList.remove('selected'));
      document.querySelectorAll('.sticker-image').forEach(i => i.classList.remove('selected'));
      
      const wordEl = document.querySelector(`[data-word-id="${wordId}"]`);
      if (wordEl) {
        wordEl.classList.add('selected');
      }

      const baseSticker = stickers[stickerIndex];
      const baseWord = baseSticker ? (baseSticker.words || []).find(w => w.id === wordId) : null;
      const fontSizeInput = document.getElementById('fontSizeInput');
      if (baseWord && fontSizeInput && Number.isFinite(baseWord.fontSize)) {
        fontSizeInput.value = String(Math.round(baseWord.fontSize));
      }
      
      // עדכן את בחירת הפונט לפי הטקסט הנבחר
      const fontFamilyInput = document.getElementById('fontFamilyInput');
      if (baseWord && fontFamilyInput && baseWord.fontFamily) {
        fontFamilyInput.value = baseWord.fontFamily;
      }
      
      // עדכן את בחירת המשקל לפי הטקסט הנבחר
      const fontWeightInput = document.getElementById('fontWeightInput');
      if (baseWord && fontWeightInput && baseWord.fontWeight) {
        fontWeightInput.value = baseWord.fontWeight;
      }
      
      // Update opacity control
      updateOpacityControl('word', stickerIndex, wordId);
      
      // פתח את טאב הטקסט אוטומטית
      if (typeof setActiveToolsSection === 'function') {
        setActiveToolsSection('Text');
      }
      
      // הכנס את הטקסט של המילה לשדה הקלט
      const wordInput = document.getElementById('wordInput');
      if (baseWord && wordInput) {
        wordInput.value = baseWord.text || '';
        wordInput.focus();
        wordInput.select(); // בחר את כל הטקסט כדי שיהיה קל לערוך
      }
    }
    
    // פונקציה לעדכון צבע של טקסט נבחר
    function applyColorToSelectedWord(color, isGradient) {
      if (selectedWord === null || selectedSticker === null) return false;
      
      const sticker = stickers[selectedSticker];
      if (!sticker) return false;
      
      const word = sticker.words.find(w => w.id === selectedWord);
      if (!word) return false;
      
      pushHistory();
      word.color = color;
      word.isGradient = isGradient;
      renderStickers();
      return true;
    }
    
    // פונקציה לעדכון פונט של טקסט נבחר
    function applyFontToSelectedWord(fontFamily) {
      if (selectedWord === null || selectedSticker === null) return false;
      
      const sticker = stickers[selectedSticker];
      if (!sticker) return false;
      
      const word = sticker.words.find(w => w.id === selectedWord);
      if (!word) return false;
      
      pushHistory();
      word.fontFamily = fontFamily;
      renderStickers();
      return true;
    }
    
    // פונקציה לעדכון משקל פונט של טקסט נבחר
    function applyFontWeightToSelectedWord(fontWeight) {
      if (selectedWord === null || selectedSticker === null) return false;
      
      const sticker = stickers[selectedSticker];
      if (!sticker) return false;
      
      const word = sticker.words.find(w => w.id === selectedWord);
      if (!word) return false;
      
      pushHistory();
      word.fontWeight = fontWeight;
      renderStickers();
      return true;
    }

    function startInlineTextEdit(wordElement, stickerIndex, wordId) {
      // מנע עריכה כפולה
      if (wordElement.querySelector('.inline-text-edit')) return;
      
      const sticker = stickers[stickerIndex];
      if (!sticker) return;
      
      const word = sticker.words.find(w => w.id === wordId);
      if (!word) return;

      // שמור את הטקסט המקורי
      const originalText = word.text;
      const originalColor = word.color;
      const originalIsGradient = word.isGradient;
      
      // יצור שדה קלט שקוף לחלוטין - רק לקליטת הקלדה
      const input = document.createElement('input');
      input.type = 'text';
      input.value = originalText;
      input.className = 'inline-text-edit';
      
      // הוסף מסגרת עריכה לאלמנט הטקסט
      wordElement.classList.add('editing-mode');
      
      // סגנון שדה הקלט - שקוף לחלוטין, רק לקליטת הקלדה
      input.style.cssText = `
        position: absolute !important;
        top: 0 !important;
        left: 0 !important;
        width: 100% !important;
        height: 100% !important;
        border: none !important;
        background: transparent !important;
        color: transparent !important;
        caret-color: #3b82f6 !important;
        font-size: ${word.fontSize || 12}px !important;
        font-family: ${word.fontFamily || 'Arial'} !important;
        font-weight: ${word.fontWeight || '700'} !important;
        text-align: center !important;
        outline: none !important;
        z-index: 1001 !important;
        padding: 0 !important;
        margin: 0 !important;
        box-sizing: border-box !important;
        cursor: text !important;
      `;
      
      // הוסף את שדה הקלט
      wordElement.appendChild(input);
      
      // פוקוס ובחירת הטקסט
      setTimeout(() => {
        input.focus();
        input.select();
      }, 50);

      // פונקציה לעדכון הטקסט בזמן אמת
      const updateTextRealtime = () => {
        const newText = input.value;
        if (newText !== word.text) {
          word.text = newText;
          
          // עדכן את התצוגה בזמן אמת
          if (word.curve && word.curve !== 0) {
            // טקסט מקומר - עדכן את ה-SVG
            const svgContainer = wordElement.querySelector('svg');
            if (svgContainer) {
              const textPath = svgContainer.querySelector('textPath');
              if (textPath) {
                textPath.textContent = newText;
              }
            }
          } else {
            // טקסט רגיל - עדכן את ה-span
            const textSpan = wordElement.querySelector('.word-text-content');
            if (textSpan) {
              textSpan.textContent = newText;
            } else {
              // אם אין span, חפש טקסט ישיר
              const childNodes = Array.from(wordElement.childNodes);
              for (const node of childNodes) {
                if (node.nodeType === Node.TEXT_NODE) {
                  node.textContent = newText;
                  break;
                }
              }
            }
          }
        }
      };

      // עדכון בזמן הקלדה
      input.addEventListener('input', updateTextRealtime);

      // מאזין לשינויי צבע בזמן עריכה
      const handleColorChange = (e) => {
        if (e.target.classList.contains('color-palette-color') || 
            e.target.id === 'gradientPreview') {
          
          setTimeout(() => {
            const textColorSwatch = document.getElementById('textColorSwatch');
            const textColorPicker = document.getElementById('textColorPicker');
            
            if (textColorSwatch && textColorPicker) {
              const isGradient = textColorSwatch.dataset.isGradient === 'true';
              const gradientValue = textColorSwatch.dataset.gradientValue;
              const regularColor = textColorPicker.value;
              
              if (isGradient && gradientValue) {
                word.color = gradientValue;
                word.isGradient = true;
              } else {
                word.color = regularColor;
                word.isGradient = false;
              }
              
              // רנדר מחדש כדי לעדכן את הצבע
              renderStickers();
              // בחר מחדש את הטקסט לעריכה
              setTimeout(() => {
                const newWordEl = document.querySelector(`[data-sticker-index="${stickerIndex}"] [data-word-id="${wordId}"]`);
                if (newWordEl) {
                  startInlineTextEdit(newWordEl, stickerIndex, wordId);
                }
              }, 50);
            }
          }, 10);
        }
      };

      document.addEventListener('click', handleColorChange, true);
      
      // פונקציה לסיום העריכה
      const finishEdit = (save = true) => {
        if (!input.parentNode) return;
        
        document.removeEventListener('click', handleColorChange, true);
        wordElement.classList.remove('editing-mode');
        
        if (save && input.value.trim() !== '') {
          pushHistory();
          word.text = input.value.trim();
          
          const wordInput = document.getElementById('wordInput');
          if (wordInput) {
            wordInput.value = word.text;
          }
        } else if (!save) {
          // שחזר את הטקסט והצבע המקוריים
          word.text = originalText;
          word.color = originalColor;
          word.isGradient = originalIsGradient;
        }
        
        // הסר את שדה הקלט
        input.remove();
        
        // רנדר מחדש כדי לעדכן את כל המדבקות עם הטקסט והצבע החדשים
        renderStickers();
      };

      // אירועים לסיום העריכה
      input.addEventListener('blur', () => {
        setTimeout(() => finishEdit(true), 100); // עיכוב קטן כדי לאפשר לחיצה על צבעים
      });
      
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          finishEdit(true);
        } else if (e.key === 'Escape') {
          e.preventDefault();
          finishEdit(false);
        }
      });

      // מנע גרירה בזמן עריכה
      input.addEventListener('mousedown', (e) => e.stopPropagation());
      input.addEventListener('click', (e) => e.stopPropagation());
    }

    function updateWordElFontSize(stickerIndex, word) {
      const el = document.querySelector(`[data-sticker-index="${stickerIndex}"] [data-word-id="${word.id}"]`);
      if (el) {
        el.style.fontSize = `${word.fontSize || 12}px`;
      }
    }

    function applyWordFontSize(stickerIndex, wordId, next) {
      const baseSticker = stickers[stickerIndex];
      const baseWord = baseSticker ? (baseSticker.words || []).find(w => w.id === wordId) : null;
      if (!baseWord) return;

      baseWord.fontSize = next;
      updateWordElFontSize(stickerIndex, baseWord);

      const fontSizeInput = document.getElementById('fontSizeInput');
      if (fontSizeInput) fontSizeInput.value = String(Math.round(next));

      if (syncMoveEnabled) {
        const seriesId = baseWord.seriesId;
        if (seriesId) {
          stickers.forEach((sticker, idx) => {
            if (!sticker || idx === stickerIndex) return;
            const correspondingWord = (sticker.words || []).find(w => w.seriesId === seriesId);
            if (!correspondingWord) return;
            correspondingWord.fontSize = next;
            updateWordElFontSize(idx, correspondingWord);
          });
        } else {
          const wordIndex = (baseSticker.words || []).findIndex(w => w.id === wordId);
          stickers.forEach((sticker, idx) => {
            if (!sticker || idx === stickerIndex) return;
            const correspondingWord = (sticker.words || [])[wordIndex];
            if (!correspondingWord) return;
            correspondingWord.fontSize = next;
            updateWordElFontSize(idx, correspondingWord);
          });
        }
      }
    }

    function adjustSelectedWordFontSize(delta) {
      console.log('adjustSelectedWordFontSize called with delta:', delta);
      console.log('selectedSticker:', selectedSticker, 'selectedWord:', selectedWord);
      
      if (selectedSticker === null || selectedWord === null) {
        showStatus('בחר טקסט (לחץ על הטקסט במדבקה) כדי לשנות גודל', true);
        return;
      }

      const stickerIndex = selectedSticker;
      const baseSticker = stickers[stickerIndex];
      const baseWord = baseSticker ? (baseSticker.words || []).find(w => w.id === selectedWord) : null;
      if (!baseWord) {
        showStatus('שגיאה: טקסט לא נמצא', true);
        return;
      }

      const current = Number.isFinite(baseWord.fontSize) ? baseWord.fontSize : 12;
      const next = Math.max(8, Math.min(240, current + delta));
      console.log('Font size change:', current, '->', next);
      applyWordFontSize(stickerIndex, selectedWord, next);
    }

    function startWordResize(e, stickerIndex, wordId) {
      const sticker = stickers[stickerIndex];
      const word = sticker ? (sticker.words || []).find(w => w.id === wordId) : null;
      if (!word) return;

      pushHistory();

      selectWord(stickerIndex, wordId);

      const current = Number.isFinite(word.fontSize) ? word.fontSize : 12;
      resizingWord = {
        stickerIndex,
        wordId,
        startX: e.clientX,
        startY: e.clientY,
        startFontSize: current
      };

      document.addEventListener('mousemove', resizeWordFont);
      document.addEventListener('mouseup', stopWordResize);
    }

    function resizeWordFont(e) {
      if (!resizingWord) return;

      const { stickerIndex, wordId, startX, startY, startFontSize } = resizingWord;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      const delta = (dx + dy) / 8;

      const next = Math.max(8, Math.min(240, Math.round(startFontSize + delta)));
      applyWordFontSize(stickerIndex, wordId, next);
      
      // Show font size indicator
      showFontSizeIndicator(next, e.clientX, e.clientY);
    }

    function stopWordResize() {
      resizingWord = null;
      document.removeEventListener('mousemove', resizeWordFont);
      document.removeEventListener('mouseup', stopWordResize);
      
      // Hide font size indicator
      hideFontSizeIndicator();
    }
    
    // Font size indicator functions
    function showFontSizeIndicator(size, x, y) {
      let indicator = document.getElementById('fontSizeIndicator');
      if (!indicator) {
        indicator = document.createElement('div');
        indicator.id = 'fontSizeIndicator';
        indicator.style.cssText = `
          position: fixed;
          background: rgba(0, 0, 0, 0.8);
          color: white;
          padding: 8px 16px;
          border-radius: 8px;
          font-size: 18px;
          font-weight: bold;
          z-index: 10000;
          pointer-events: none;
          box-shadow: 0 4px 12px rgba(0,0,0,0.3);
        `;
        document.body.appendChild(indicator);
      }
      indicator.textContent = `${size}px`;
      indicator.style.left = `${x + 20}px`;
      indicator.style.top = `${y - 40}px`;
      indicator.style.display = 'block';
    }
    
    function hideFontSizeIndicator() {
      const indicator = document.getElementById('fontSizeIndicator');
      if (indicator) {
        indicator.style.display = 'none';
      }
    }

    function selectImage(stickerIndex, imageId) {
      selectedSticker = stickerIndex;
      selectedWord = null;
      selectedImage = imageId;
      
      document.querySelectorAll('.sticker-container').forEach(s => s.classList.remove('selected'));
      document.querySelectorAll('.text-word').forEach(w => w.classList.remove('selected'));
      document.querySelectorAll('.sticker-image').forEach(i => i.classList.remove('selected'));
      
      const imageEl = document.querySelector(`[data-image-id="${imageId}"]`);
      if (imageEl) {
        imageEl.classList.add('selected');
      }

      // Update opacity control
      updateOpacityControl('image', stickerIndex, imageId);

      focusWordInput();
    }

    // Opacity control functions
    let currentOpacityTarget = { type: null, stickerIndex: null, elementId: null };
    
    function updateOpacityControl(type, stickerIndex, elementId = null) {
      const wrapper = document.getElementById('opacityControlWrapper');
      const slider = document.getElementById('opacitySlider');
      const valueDisplay = document.getElementById('opacityValue');
      const removeBackgroundBtn = document.getElementById('removeBackgroundBtn');
      const eraserBtn = document.getElementById('eraserBtn');
      const opacitySection = document.getElementById('opacitySection');
      
      if (!wrapper || !slider || !valueDisplay || !removeBackgroundBtn) return;
      
      currentOpacityTarget = { type, stickerIndex, elementId };
      
      // Get current opacity value
      let currentOpacity = 100;
      
      if (type === 'sticker' && stickers[stickerIndex]) {
        currentOpacity = Math.round((stickers[stickerIndex].opacity ?? 1) * 100);
      } else if (type === 'word' && stickers[stickerIndex]) {
        const word = (stickers[stickerIndex].words || []).find(w => w.id === elementId);
        if (word) {
          currentOpacity = Math.round((word.opacity ?? 1) * 100);
        }
      } else if (type === 'image' && stickers[stickerIndex]) {
        const image = (stickers[stickerIndex].images || []).find(i => i.id === elementId);
        if (image) {
          currentOpacity = Math.round((image.opacity ?? 1) * 100);
        }
      }
      
      slider.value = currentOpacity;
      valueDisplay.textContent = `${currentOpacity}%`;
      
      // הפעלת הסליידר והסרת המצב המושבת
      slider.disabled = false;
      if (opacitySection) {
        opacitySection.classList.remove('opacity-section-disabled');
      }
      
      // הגדרת כפתור הסרת רקע
      removeBackgroundBtn.onclick = async () => {
        // הפעלת אנימציית טעינה
        const loader = document.getElementById('removeBackgroundLoader');
        const content = document.getElementById('removeBackgroundContent');
        
        if (loader && content) {
          loader.classList.remove('hidden');
          content.classList.add('opacity-0');
          removeBackgroundBtn.disabled = true;
        }
        
        // עצירת שמירה אוטומטית ל-25 שניות
        if (typeof GoogleDriveManager !== 'undefined' && GoogleDriveManager.setProcessingBackground) {
          GoogleDriveManager.setProcessingBackground(true);
          
          // ביטול אוטומטי אחרי 25 שניות
          setTimeout(() => {
            if (typeof GoogleDriveManager !== 'undefined' && GoogleDriveManager.setProcessingBackground) {
              GoogleDriveManager.setProcessingBackground(false);
              console.log('Auto-save resumed after 25 seconds');
            }
          }, 25000);
        }
        
        try {
          if (type === 'sticker') {
            await removeBackgroundFromSelected(stickerIndex, 'sticker');
          } else if (type === 'image') {
            await removeBackgroundFromSelected(stickerIndex, 'image', elementId);
          }
        } catch (error) {
          console.error('Background removal failed:', error);
        } finally {
          // הסתרת אנימציית טעינה
          if (loader && content) {
            loader.classList.add('hidden');
            content.classList.remove('opacity-0');
            removeBackgroundBtn.disabled = false;
          }
        }
      };
      
      // הפעלה/השבתה של כפתורים בהתאם לסוג האלמנט
      if (type === 'sticker' || type === 'image') {
        // הפעלת כפתור הסרת רקע
        removeBackgroundBtn.disabled = false;
        removeBackgroundBtn.className = 'px-4 py-2 bg-gradient-to-r from-red-500 to-pink-600 text-white font-bold rounded-lg shadow-md hover:from-red-600 hover:to-pink-700 transition-all flex items-center gap-2 relative overflow-hidden cursor-pointer';
      } else {
        // השבתת כפתור הסרת רקע
        removeBackgroundBtn.disabled = true;
        removeBackgroundBtn.className = 'px-4 py-2 bg-gray-300 text-gray-500 font-bold rounded-lg shadow-md cursor-not-allowed flex items-center gap-2 relative overflow-hidden';
      }
      
      // כפתור המחק פעיל רק עבור תמונות
      if (eraserBtn) {
        if (type === 'image') {
          // הפעלת כפתור המחק - הכפתור תמיד מופעל כשתמונה נבחרת
          // גם אם המחק כבר פעיל, המשתמש יכול ללחוץ עליו כדי לכבות אותו
          eraserBtn.disabled = false;
          eraserBtn.className = 'px-4 py-2 bg-gradient-to-r from-orange-500 to-amber-600 text-white font-bold rounded-lg shadow-md hover:from-orange-600 hover:to-amber-700 transition-all flex items-center gap-2 cursor-pointer';
        } else {
          // השבתת כפתור המחק עבור מדבקות וטקסט
          eraserBtn.disabled = true;
          eraserBtn.className = 'px-4 py-2 bg-gray-300 text-gray-500 font-bold rounded-lg shadow-md cursor-not-allowed flex items-center gap-2';
          
          // כיבוי המחק אם הוא פעיל
          if (typeof window.deactivateEraserTool === 'function') {
            window.deactivateEraserTool();
          }
        }
      }
    }
    
    function hideOpacityControl() {
      const slider = document.getElementById('opacitySlider');
      const removeBackgroundBtn = document.getElementById('removeBackgroundBtn');
      const eraserBtn = document.getElementById('eraserBtn');
      const opacitySection = document.getElementById('opacitySection');
      const eraserSizeControl = document.getElementById('eraserSizeControl');
      
      // כיבוי כלי המחק אם הוא פעיל
      if (typeof window.deactivateEraserTool === 'function') {
        window.deactivateEraserTool();
      }
      
      // הסתרת סקלת גודל המחק
      if (eraserSizeControl) {
        eraserSizeControl.classList.add('hidden');
      }
      
      // השבתת כל הכפתורים והסליידר
      if (slider) slider.disabled = true;
      
      // הוספת class למצב מושבת
      if (opacitySection) {
        opacitySection.classList.add('opacity-section-disabled');
      }
      
      if (removeBackgroundBtn) {
        removeBackgroundBtn.disabled = true;
        removeBackgroundBtn.className = 'px-4 py-2 bg-gray-300 text-gray-500 font-bold rounded-lg shadow-md cursor-not-allowed flex items-center gap-2 relative overflow-hidden';
      }
      
      if (eraserBtn) {
        eraserBtn.disabled = true;
        eraserBtn.className = 'px-4 py-2 bg-gray-300 text-gray-500 font-bold rounded-lg shadow-md cursor-not-allowed flex items-center gap-2';
      }
      
      currentOpacityTarget = { type: null, stickerIndex: null, elementId: null };
    }
    
    function applyOpacity(opacityPercent) {
      const { type, stickerIndex, elementId } = currentOpacityTarget;
      const opacityValue = opacityPercent / 100;
      
      if (type === 'sticker' && stickers[stickerIndex]) {
        pushHistory();
        stickers[stickerIndex].opacity = opacityValue;
        // Apply opacity only to the main sticker image, not the container
        const stickerEl = document.querySelector(`[data-sticker-index="${stickerIndex}"]`);
        if (stickerEl) {
          const mainImage = stickerEl.querySelector('.sticker-main-image');
          if (mainImage) {
            mainImage.style.opacity = opacityValue;
          }
        }
        
        // Sync opacity to all other stickers if sync is enabled
        if (syncMoveEnabled) {
          stickers.forEach((sticker, idx) => {
            if (idx === stickerIndex) return;
            sticker.opacity = opacityValue;
            const otherStickerEl = document.querySelector(`[data-sticker-index="${idx}"]`);
            if (otherStickerEl) {
              const otherMainImage = otherStickerEl.querySelector('.sticker-main-image');
              if (otherMainImage) {
                otherMainImage.style.opacity = opacityValue;
              }
            }
          });
        }
      } else if (type === 'word' && stickers[stickerIndex]) {
        const word = (stickers[stickerIndex].words || []).find(w => w.id === elementId);
        if (word) {
          pushHistory();
          word.opacity = opacityValue;
          const wordEl = document.querySelector(`[data-word-id="${elementId}"]`);
          if (wordEl) {
            wordEl.style.opacity = opacityValue;
          }
          
          // Sync opacity to corresponding words in other stickers if sync is enabled
          if (syncMoveEnabled && word.seriesId) {
            stickers.forEach((sticker, idx) => {
              if (idx === stickerIndex) return;
              const correspondingWord = (sticker.words || []).find(w => w.seriesId === word.seriesId);
              if (!correspondingWord) return;
              correspondingWord.opacity = opacityValue;
              const correspondingEl = document.querySelector(`[data-word-id="${correspondingWord.id}"]`);
              if (correspondingEl) {
                correspondingEl.style.opacity = opacityValue;
              }
            });
          }
        }
      } else if (type === 'image' && stickers[stickerIndex]) {
        const image = (stickers[stickerIndex].images || []).find(i => i.id === elementId);
        if (image) {
          pushHistory();
          image.opacity = opacityValue;
          const imageEl = document.querySelector(`[data-image-id="${elementId}"]`);
          if (imageEl) {
            imageEl.style.opacity = opacityValue;
          }
          
          // Sync opacity to corresponding images in other stickers if sync is enabled
          if (syncMoveEnabled && image.seriesId) {
            stickers.forEach((sticker, idx) => {
              if (idx === stickerIndex) return;
              const correspondingImage = (sticker.images || []).find(i => i.seriesId === image.seriesId);
              if (!correspondingImage) return;
              correspondingImage.opacity = opacityValue;
              const correspondingEl = document.querySelector(`[data-image-id="${correspondingImage.id}"]`);
              if (correspondingEl) {
                correspondingEl.style.opacity = opacityValue;
              }
            });
          }
        }
      }
    }

    function addWordToSelectedSticker() {
      if (selectedSticker === null) {
        showStatus('בחר מדבקה תחילה! לחץ על מדבקה כדי לבחור אותה.', true);
        return;
      }
      
      const wordInput = document.getElementById('wordInput');
      const text = wordInput.value.trim();
      
      if (!text) {
        showStatus('הזן מילה!', true);
        return;
      }
      
      const textColorSwatch = document.getElementById('textColorSwatch');
      const color = document.getElementById('textColorPicker').value;
      const isGradient = textColorSwatch && textColorSwatch.dataset.isGradient === 'true';
      const gradientValue = textColorSwatch ? textColorSwatch.dataset.gradientValue : null;
      const fontSize = parseInt(document.getElementById('fontSizeInput').value);
      const fontFamily = document.getElementById('fontFamilyInput').value;
      const fontWeight = document.getElementById('fontWeightInput').value;

      const sticker = stickers[selectedSticker];
      if (!sticker) {
        showStatus('שגיאה: מדבקה לא נמצאה', true);
        return;
      }

      pushHistory();
      sticker.words = sticker.words || [];
      const tempDiv = document.createElement('div');
      tempDiv.style.position = 'absolute';
      tempDiv.style.visibility = 'hidden';
      tempDiv.style.whiteSpace = 'nowrap';
      tempDiv.style.padding = '0';
      tempDiv.style.lineHeight = '1';
      tempDiv.style.fontSize = `${fontSize}px`;
      tempDiv.style.fontFamily = fontFamily;
      tempDiv.style.fontWeight = fontWeight;
      tempDiv.textContent = text;
      document.body.appendChild(tempDiv);
      const textW = tempDiv.offsetWidth;
      const textH = tempDiv.offsetHeight;
      document.body.removeChild(tempDiv);

      // מיקום במרכז המדבקה - פשוט וישיר
      const stickerW = sticker.width || 100;
      const stickerH = sticker.height || 100;
      const fixedX = (stickerW - textW) / 2;
      const fixedY = (stickerH - textH) / 2;

      const seriesId = `series-${++wordSeriesCounter}`;
      
      const word = {
        id: `word-${++wordIdCounter}`,
        seriesId,
        text: text,
        x: fixedX,
        y: fixedY,
        color: isGradient ? gradientValue : color,
        fontSize: fontSize,
        fontFamily: fontFamily,
        fontWeight: fontWeight,
        isGradient: isGradient
      };
      
      stickers[selectedSticker].words.push(word);
      
      renderStickers();
      showStatus(`המילה "${text}" נוספה למדבקה נבחרת!`);
      wordInput.value = '';
    }

    function addWordToAllStickers() {
      if (stickers.length === 0) {
        showStatus('העלה מדבקות תחילה!', true);
        return;
      }
      
      const wordInput = document.getElementById('wordInput');
      const text = wordInput.value.trim();
      
      if (!text) {
        showStatus('הזן מילה!', true);
        return;
      }
      
      const textColorSwatch = document.getElementById('textColorSwatch');
      const color = document.getElementById('textColorPicker').value;
      const isGradient = textColorSwatch && textColorSwatch.dataset.isGradient === 'true';
      const gradientValue = textColorSwatch ? textColorSwatch.dataset.gradientValue : null;
      const fontSize = parseInt(document.getElementById('fontSizeInput').value);
      const fontFamily = document.getElementById('fontFamilyInput').value;
      const fontWeight = document.getElementById('fontWeightInput').value;

      const seriesId = `series-${++wordSeriesCounter}`;

      pushHistory();

      const tempDiv = document.createElement('div');
      tempDiv.style.position = 'absolute';
      tempDiv.style.visibility = 'hidden';
      tempDiv.style.whiteSpace = 'nowrap';
      tempDiv.style.padding = '0';
      tempDiv.style.lineHeight = '1';
      tempDiv.style.fontSize = `${fontSize}px`;
      tempDiv.style.fontFamily = fontFamily;
      tempDiv.style.fontWeight = fontWeight;
      tempDiv.textContent = text;
      document.body.appendChild(tempDiv);
      const textW = tempDiv.offsetWidth;
      const textH = tempDiv.offsetHeight;
      document.body.removeChild(tempDiv);
      
      stickers.forEach(sticker => {
        sticker.words = sticker.words || [];
        // מיקום במרכז המדבקה - פשוט וישיר
        const stickerW = sticker.width || 100;
        const stickerH = sticker.height || 100;
        const fixedX = (stickerW - textW) / 2;
        const fixedY = (stickerH - textH) / 2;
        const word = {
          id: `word-${++wordIdCounter}`,
          seriesId,
          text: text,
          x: fixedX,
          y: fixedY,
          color: isGradient ? gradientValue : color,
          fontSize: fontSize,
          fontFamily: fontFamily,
          fontWeight: fontWeight,
          isGradient: isGradient
        };
        sticker.words.push(word);
      });
      
      renderStickers();
      showStatus(`המילה "${text}" נוספה לכל ${stickers.length} המדבקות!`);
      wordInput.value = '';
    }

    function replaceWordInAllStickers() {
      if (stickers.length === 0) {
        showStatus('אין מדבקות!', true);
        return;
      }
      
      const wordInput = document.getElementById('wordInput');
      const newText = wordInput.value.trim();
      
      if (!newText) {
        showStatus('הזן מילה חדשה!', true);
        return;
      }
      
      const textColorSwatch = document.getElementById('textColorSwatch');
      const color = document.getElementById('textColorPicker').value;
      const isGradient = textColorSwatch && textColorSwatch.dataset.isGradient === 'true';
      const gradientValue = textColorSwatch ? textColorSwatch.dataset.gradientValue : null;
      const fontSize = parseInt(document.getElementById('fontSizeInput').value);
      const fontFamily = document.getElementById('fontFamilyInput').value;
      const fontWeight = document.getElementById('fontWeightInput').value;
      
      let replacedCount = 0;

      pushHistory();
      
      // Create temporary element for measuring text dimensions
      const tempDiv = document.createElement('div');
      tempDiv.style.position = 'absolute';
      tempDiv.style.visibility = 'hidden';
      tempDiv.style.whiteSpace = 'nowrap';
      tempDiv.style.padding = '2px';
      document.body.appendChild(tempDiv);
      
      stickers.forEach((sticker, stickerIndex) => {
        if (sticker.words.length > 0) {
          const lastWord = sticker.words[sticker.words.length - 1];
          
          // Measure old word dimensions
          tempDiv.style.fontSize = `${lastWord.fontSize}px`;
          tempDiv.style.fontFamily = lastWord.fontFamily || 'Arial';
          tempDiv.style.fontWeight = lastWord.fontWeight || '700';
          tempDiv.textContent = lastWord.text;
          const oldWidth = tempDiv.offsetWidth;
          const oldHeight = tempDiv.offsetHeight;
          
          // Calculate center of old word
          const oldCenterX = lastWord.x + oldWidth / 2;
          const oldCenterY = lastWord.y + oldHeight / 2;
          
          // Measure new word dimensions
          tempDiv.style.fontSize = `${fontSize}px`;
          tempDiv.style.fontFamily = fontFamily;
          tempDiv.style.fontWeight = fontWeight;
          tempDiv.textContent = newText;
          const newWidth = tempDiv.offsetWidth;
          const newHeight = tempDiv.offsetHeight;
          
          // Calculate new position to center the new word on the old center
          const newX = oldCenterX - newWidth / 2;
          const newY = oldCenterY - newHeight / 2;
          
          // Update word properties
          lastWord.text = newText;
          lastWord.color = isGradient ? gradientValue : color;
          lastWord.fontSize = fontSize;
          lastWord.fontFamily = fontFamily;
          lastWord.fontWeight = fontWeight;
          lastWord.isGradient = isGradient;
          lastWord.x = Math.max(0, newX);
          lastWord.y = Math.max(0, newY);
          
          replacedCount++;
        }
      });
      
      document.body.removeChild(tempDiv);
      
      if (replacedCount === 0) {
        showStatus('אין מילות להחליף! הוסף מילות תחילה.', true);
        return;
      }
      
      renderStickers();
      showStatus(`המילה הוחלפה ל-"${newText}" ומרוכזה ב-${replacedCount} מדבקות!`);
      wordInput.value = '';
    }

    function deleteWord(stickerIndex, wordId) {
      pushHistory();
      if (syncDeleteEnabled) {
        const baseSticker = stickers[stickerIndex];
        const baseWord = baseSticker ? (baseSticker.words || []).find(w => w.id === wordId) : null;
        const seriesId = baseWord && baseWord.seriesId ? baseWord.seriesId : null;

        if (seriesId) {
          stickers.forEach(sticker => {
            sticker.words = (sticker.words || []).filter(w => w.seriesId !== seriesId);
          });
          renderStickers();
          showStatus(`המילה נמחקה מכל ${stickers.length} המדבקות! (סנכרון מחיקה דלוק)`);
          return;
        }

        // Fallback (older projects without seriesId)
        const wordIndex = baseSticker.words.findIndex(w => w.id === wordId);
        if (wordIndex !== -1) {
          stickers.forEach(sticker => {
            if (sticker.words[wordIndex]) {
              sticker.words.splice(wordIndex, 1);
            }
          });
          renderStickers();
          showStatus(`המילה נמחקה מכל ${stickers.length} המדבקות! (סנכרון מחיקה דלוק)`);
        }
      } else {
        // Delete only from current sticker
        stickers[stickerIndex].words = stickers[stickerIndex].words.filter(w => w.id !== wordId);
        renderStickers();
        showStatus('המילה נמחקה');
      }
      
      if (selectedSticker === stickerIndex) {
        selectSticker(stickerIndex);
      }
    }

    function deleteImage(stickerIndex, imageId) {
      pushHistory();
      if (syncDeleteEnabled) {
        // Find the index of this image in the current sticker
        const imageIndex = stickers[stickerIndex].images.findIndex(i => i.id === imageId);
        
        if (imageIndex !== -1) {
          // Delete the image at the same index from all stickers
          stickers.forEach(sticker => {
            if (sticker.images && sticker.images[imageIndex]) {
              sticker.images.splice(imageIndex, 1);
            }
          });
          renderStickers();
          showStatus(`התמונה נמחקה מכל ${stickers.length} המדבקות! (סנכרון מחיקה דלוק)`);
        }
      } else {
        // Delete only from current sticker
        stickers[stickerIndex].images = stickers[stickerIndex].images.filter(i => i.id !== imageId);
        renderStickers();
        showStatus('התמונה נמחקה');
      }
      
      if (selectedSticker === stickerIndex) {
        selectSticker(stickerIndex);
      }
    }

    function addImageToAllStickers(imageDataUrl, originalWidth, originalHeight) {
      if (stickers.length === 0) {
        showStatus('העלה מדבקות תחילה!', true);
        return;
      }

      pushHistory();
      
      // Calculate the image size relative to sticker height
      // We'll make it about 70% of the sticker height
      const firstSticker = stickers[0];
      const targetHeight = firstSticker.height * 0.7;
      const aspectRatio = originalWidth / originalHeight;
      const calculatedWidth = targetHeight * aspectRatio;
      
      // Create a series ID for syncing these images across stickers
      const seriesId = `image-series-${++imageSeriesCounter}`;
      
      stickers.forEach(sticker => {
        sticker.images = sticker.images || [];
        const x = Math.max(0, (sticker.width - calculatedWidth) / 2);
        const y = Math.max(0, (sticker.height - targetHeight) / 2);
        const image = {
          id: `image-${++imageIdCounter}`,
          seriesId: seriesId,
          dataUrl: imageDataUrl,
          x: x,
          y: y,
          width: calculatedWidth,
          height: targetHeight,
          originalWidth: originalWidth,
          originalHeight: originalHeight
        };
        sticker.images.push(image);
      });
      
      renderStickers();
      showStatus(`התמונה נוספה לכל ${stickers.length} המדבקות!`);
      focusWordInput();
    }

    function addImageToSelectedSticker(imageDataUrl, originalWidth, originalHeight) {
      if (selectedSticker === null) {
        showStatus('בחר מדבקה תחילה! לחץ על מדבקה כדי לבחור אותה.', true);
        return;
      }

      pushHistory();
      
      const sticker = stickers[selectedSticker];
      const targetHeight = sticker.height * 0.7;
      const aspectRatio = originalWidth / originalHeight;
      const calculatedWidth = targetHeight * aspectRatio;
      
      const x = Math.max(0, (sticker.width - calculatedWidth) / 2);
      const y = Math.max(0, (sticker.height - targetHeight) / 2);

      sticker.images = sticker.images || [];
      const image = {
        id: `image-${++imageIdCounter}`,
        dataUrl: imageDataUrl,
        x: x,
        y: y,
        width: calculatedWidth,
        height: targetHeight,
        originalWidth: originalWidth,
        originalHeight: originalHeight
      };
      sticker.images.push(image);
      
      renderStickers();
      showStatus('התמונה נוספה למדבקה נבחרת!');
      focusWordInput();
    }

    function addElementsToLibrary(files) {
      let addedCount = 0;
      
      Array.from(files).forEach(file => {
        if (file.type.startsWith('image/')) {
          const reader = new FileReader();
          reader.onload = function(event) {
            const img = new Image();
            img.onload = function() {
              elementsLibrary.push({
                id: `element-${Date.now()}-${addedCount}`,
                name: file.name,
                dataUrl: event.target.result,
                width: img.width,
                height: img.height
              });
              addedCount++;
              
              if (addedCount === files.length) {
                renderElementsGallery();
                showStatus(`${addedCount} אלמנטים נוספו למאגר!`);
              }
            };
            img.src = event.target.result;
          };
          reader.readAsDataURL(file);
        }
      });
    }

    function renderElementsGallery() {
      const gallery = document.getElementById('elementsGallery');
      const countSpan = document.getElementById('elementsCount');
      
      if (elementsLibrary.length === 0) {
        gallery.classList.add('hidden');
        countSpan.textContent = '';
        return;
      }
      
      gallery.classList.remove('hidden');
      countSpan.textContent = `${elementsLibrary.length} אלמנטים במאגר`;
      gallery.innerHTML = '';
      
      elementsLibrary.forEach((element, index) => {
        const elementDiv = document.createElement('div');
        elementDiv.className = 'relative group cursor-pointer border-2 border-gray-200 rounded-lg p-2 hover:border-green-500 hover:shadow-lg transition-all bg-white';
        elementDiv.title = element.name;
        
        const img = document.createElement('img');
        img.src = element.dataUrl;
        img.className = 'w-full h-16 object-contain';
        
        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'absolute top-0 right-0 w-5 h-5 bg-red-500 text-white rounded-full text-xs hidden group-hover:block';
        deleteBtn.textContent = '×';
        deleteBtn.onclick = (e) => {
          e.stopPropagation();
          elementsLibrary.splice(index, 1);
          renderElementsGallery();
          showStatus('אלמנט הוסר מהמאגר');
        };
        
        elementDiv.appendChild(img);
        elementDiv.appendChild(deleteBtn);
        
        // Click to add element
        elementDiv.addEventListener('click', () => {
          showElementAddMenu(element);
        });
        
        gallery.appendChild(elementDiv);
      });
    }

    function showElementAddMenu(element) {
      const choice = confirm(`הוסף את "${element.name}" לכל המדבקות?\n\nאישור = לכל המדבקות\nביטול = למדבקה נבחרת בלבד`);
      
      if (choice) {
        // Add to all stickers
        addImageToAllStickers(element.dataUrl, element.width, element.height);
      } else {
        // Add to selected sticker only
        addImageToSelectedSticker(element.dataUrl, element.width, element.height);
      }
    }

    function startWordDrag(e, stickerIndex, wordId) {
      pushHistory();
      const wordEl = e.currentTarget;
      const rect = wordEl.getBoundingClientRect();
      const parentRect = wordEl.parentElement.getBoundingClientRect();
      
      offsetX = e.clientX - rect.left;
      offsetY = e.clientY - rect.top;
      
      const word = stickers[stickerIndex].words.find(w => w.id === wordId);
      if (word) {
        initialDragPosition = { x: word.x, y: word.y };
      }
      
      draggedElement = { type: 'word', stickerIndex, wordId, wordEl, parentRect };
      
      document.addEventListener('mousemove', dragWord);
      document.addEventListener('mouseup', stopDrag);
    }

    function dragWord(e) {
      if (!draggedElement || draggedElement.type !== 'word') return;
      
      const { wordEl, parentRect, stickerIndex, wordId } = draggedElement;
      
      let newX = e.clientX - parentRect.left - offsetX;
      let newY = e.clientY - parentRect.top - offsetY;
      
      // Allow 30% movement beyond sticker boundaries
      const extraSpaceX = parentRect.width * 0.3;
      const extraSpaceY = parentRect.height * 0.3;
      
      newX = Math.max(-extraSpaceX, Math.min(newX, parentRect.width + extraSpaceX - wordEl.offsetWidth));
      newY = Math.max(-extraSpaceY, Math.min(newY, parentRect.height + extraSpaceY - wordEl.offsetHeight));
      
      const word = stickers[stickerIndex].words.find(w => w.id === wordId);
      if (word && initialDragPosition) {
        const deltaX = newX - initialDragPosition.x;
        const deltaY = newY - initialDragPosition.y;
        
        word.x = newX;
        word.y = newY;
        wordEl.style.left = `${newX}px`;
        wordEl.style.top = `${newY}px`;
        
        // If sync is enabled, move corresponding words in other stickers
        if (syncMoveEnabled) {
          const seriesId = word.seriesId;
          if (seriesId) {
            stickers.forEach((sticker, idx) => {
              if (idx === stickerIndex) return;
              const correspondingWord = (sticker.words || []).find(w => w.seriesId === seriesId);
              if (!correspondingWord) return;
              correspondingWord.x += deltaX;
              correspondingWord.y += deltaY;

              const correspondingEl = document.querySelector(`[data-sticker-index="${idx}"] [data-word-id="${correspondingWord.id}"]`);
              if (correspondingEl) {
                correspondingEl.style.left = `${correspondingWord.x}px`;
                correspondingEl.style.top = `${correspondingWord.y}px`;
              }
            });
          } else {
            const wordIndex = stickers[stickerIndex].words.findIndex(w => w.id === wordId);
            stickers.forEach((sticker, idx) => {
              if (idx !== stickerIndex && sticker.words[wordIndex]) {
                const correspondingWord = sticker.words[wordIndex];
                correspondingWord.x += deltaX;
                correspondingWord.y += deltaY;

                const correspondingEl = document.querySelector(`[data-sticker-index="${idx}"] [data-word-id="${correspondingWord.id}"]`);
                if (correspondingEl) {
                  correspondingEl.style.left = `${correspondingWord.x}px`;
                  correspondingEl.style.top = `${correspondingWord.y}px`;
                }
              }
            });
          }
          
          // Update initial position for next movement
          initialDragPosition.x = newX;
          initialDragPosition.y = newY;
        }
      }
    }

    function stopDrag() {
      if (draggedElement && draggedElement.type === 'sticker-swap') {
        const { stickerIndex, lastClientX, lastClientY, stickerEl, startLeft, startTop, targetStickerEl: prevTargetStickerEl } = draggedElement;

        if (prevTargetStickerEl) prevTargetStickerEl.classList.remove('swap-target');

        const prevPointerEvents = stickerEl ? stickerEl.style.pointerEvents : '';
        if (stickerEl) stickerEl.style.pointerEvents = 'none';
        const targetEl = document.elementFromPoint(lastClientX, lastClientY);
        if (stickerEl) stickerEl.style.pointerEvents = prevPointerEvents;

        const nextTargetStickerEl = targetEl ? targetEl.closest('.sticker-container') : null;
        const targetIndexRaw = nextTargetStickerEl ? nextTargetStickerEl.dataset.stickerIndex : null;
        const targetIndex = targetIndexRaw !== null ? parseInt(targetIndexRaw, 10) : NaN;

        if (Number.isFinite(targetIndex) && targetIndex !== stickerIndex && stickers[targetIndex]) {
          pushHistory();
          if (stickerEl) stickerEl.classList.remove('swap-dragging');
          const tmp = stickers[stickerIndex];
          stickers[stickerIndex] = stickers[targetIndex];
          stickers[targetIndex] = tmp;
          selectedSticker = targetIndex;
          getStickerLayoutConfigFromUI();
          updateStickerLayoutInfo();
          reflowStickersPositionsOnly();
          renderStickers();
          updateFileCount();
        } else {
          if (stickerEl) {
            stickerEl.classList.remove('swap-dragging');
            if (typeof startLeft === 'string') stickerEl.style.left = startLeft;
            if (typeof startTop === 'string') stickerEl.style.top = startTop;
            stickerEl.style.zIndex = '';
          }
        }
      }
      draggedElement = null;
      initialDragPosition = null;
      document.removeEventListener('mousemove', dragWord);
      document.removeEventListener('mousemove', dragImage);
      document.removeEventListener('mousemove', dragSticker);
      document.removeEventListener('mousemove', trackStickerSwap);
      document.removeEventListener('mouseup', stopDrag);
    }

    function startStickerDrag(e, stickerIndex) {
      const stickerEl = document.querySelector(`[data-sticker-index="${stickerIndex}"]`);
      if (!stickerEl) return;

      const pageEl = stickerEl.closest('.print-page');
      if (!pageEl) return;

      const rect = stickerEl.getBoundingClientRect();
      const parentRect = pageEl.getBoundingClientRect();

      offsetX = e.clientX - rect.left;
      offsetY = e.clientY - rect.top;

      const sticker = stickers[stickerIndex];
      if (sticker) {
        initialDragPosition = { x: sticker.x, y: sticker.y };
      }

      if (autoArrangeEnabled) {
        stickerEl.classList.add('swap-dragging');
        draggedElement = {
          type: 'sticker-swap',
          stickerIndex,
          stickerEl,
          parentRect,
          startLeft: stickerEl.style.left,
          startTop: stickerEl.style.top,
          startZIndex: stickerEl.style.zIndex,
          targetStickerEl: null,
          lastClientX: e.clientX,
          lastClientY: e.clientY
        };
        document.addEventListener('mousemove', trackStickerSwap);
        document.addEventListener('mouseup', stopDrag);
        return;
      }

      draggedElement = { type: 'sticker', stickerIndex, stickerEl, parentRect };

      document.addEventListener('mousemove', dragSticker);
      document.addEventListener('mouseup', stopDrag);
    }

    function trackStickerSwap(e) {
      if (!draggedElement || draggedElement.type !== 'sticker-swap') return;

      const { stickerEl, parentRect } = draggedElement;
      draggedElement.lastClientX = e.clientX;
      draggedElement.lastClientY = e.clientY;

      if (stickerEl && parentRect) {
        let newX = e.clientX - parentRect.left - offsetX;
        let newY = e.clientY - parentRect.top - offsetY;

        newX = Math.max(0, Math.min(newX, parentRect.width - stickerEl.offsetWidth));
        newY = Math.max(0, Math.min(newY, parentRect.height - stickerEl.offsetHeight));

        stickerEl.style.left = `${newX}px`;
        stickerEl.style.top = `${newY}px`;
      }

      const prevPointerEvents = stickerEl ? stickerEl.style.pointerEvents : '';
      if (stickerEl) stickerEl.style.pointerEvents = 'none';
      const targetEl = document.elementFromPoint(e.clientX, e.clientY);
      if (stickerEl) stickerEl.style.pointerEvents = prevPointerEvents;

      const nextTargetStickerEl = targetEl ? targetEl.closest('.sticker-container') : null;
      const nextTargetIndexRaw = nextTargetStickerEl ? nextTargetStickerEl.dataset.stickerIndex : null;
      const nextTargetIndex = nextTargetIndexRaw !== null ? parseInt(nextTargetIndexRaw, 10) : NaN;

      const prevTarget = draggedElement.targetStickerEl;
      if (prevTarget && prevTarget !== nextTargetStickerEl) prevTarget.classList.remove('swap-target');

      if (Number.isFinite(nextTargetIndex) && nextTargetIndex !== draggedElement.stickerIndex) {
        if (nextTargetStickerEl) nextTargetStickerEl.classList.add('swap-target');
        draggedElement.targetStickerEl = nextTargetStickerEl;
      } else {
        if (nextTargetStickerEl) nextTargetStickerEl.classList.remove('swap-target');
        draggedElement.targetStickerEl = null;
      }
    }

    function dragSticker(e) {
      if (!draggedElement || draggedElement.type !== 'sticker') return;

      const { stickerEl, parentRect, stickerIndex } = draggedElement;
      const sticker = stickers[stickerIndex];
      if (!sticker) return;

      let newX = e.clientX - parentRect.left - offsetX;
      let newY = e.clientY - parentRect.top - offsetY;

      // שימוש במידות A4 האמיתיות במקום במידות המוצגות
      const pageWidth = pageOrientation === 'landscape' ? 297 * MM_TO_PX : 210 * MM_TO_PX;
      const pageHeight = pageOrientation === 'landscape' ? 210 * MM_TO_PX : 297 * MM_TO_PX;
      
      // חישוב הסקייל של התצוגה
      const scale = parentRect.width / pageWidth;
      
      // המרה מקואורדינטות מסך לקואורדינטות דף אמיתיות
      newX = newX / scale;
      newY = newY / scale;

      // הגבלה לגבולות הדף האמיתיים
      newX = Math.max(0, Math.min(newX, pageWidth - sticker.width));
      newY = Math.max(0, Math.min(newY, pageHeight - sticker.height));

      sticker.x = newX;
      sticker.y = newY;
      stickerEl.style.left = `${newX}px`;
      stickerEl.style.top = `${newY}px`;
    }

    function startImageDrag(e, stickerIndex, imageId) {
      // בדיקה אם כלי המחק פעיל - אם כן, לא לגרור
      if (typeof window.isEraserToolActive === 'function' && window.isEraserToolActive()) {
        console.log('Eraser is active, preventing drag');
        return;
      }
      
      pushHistory();
      const imageEl = e.currentTarget;
      const rect = imageEl.getBoundingClientRect();
      const parentEl = imageEl.parentElement;
      const parentRect = parentEl.getBoundingClientRect();
      const scaleX = (parentEl && parentRect && parentEl.offsetWidth) ? (parentRect.width / parentEl.offsetWidth) : 1;
      const scaleY = (parentEl && parentRect && parentEl.offsetHeight) ? (parentRect.height / parentEl.offsetHeight) : 1;
      
      offsetX = (e.clientX - rect.left) / (scaleX || 1);
      offsetY = (e.clientY - rect.top) / (scaleY || 1);
      
      const image = stickers[stickerIndex].images.find(i => i.id === imageId);
      if (image) {
        initialDragPosition = { x: image.x, y: image.y };
      }
      
      draggedElement = { type: 'image', stickerIndex, imageId, imageEl, parentEl, scaleX: scaleX || 1, scaleY: scaleY || 1 };
      
      document.addEventListener('mousemove', dragImage);
      document.addEventListener('mouseup', stopDrag);
    }

    function dragImage(e) {
      if (!draggedElement || draggedElement.type !== 'image') return;
      
      const { imageEl, parentEl, scaleX, scaleY, stickerIndex, imageId } = draggedElement;
      const parentRect = parentEl ? parentEl.getBoundingClientRect() : null;
      if (!parentRect) return;
      
      const effectiveScaleX = scaleX || 1;
      const effectiveScaleY = scaleY || 1;
      let newX = ((e.clientX - parentRect.left) / effectiveScaleX) - offsetX;
      let newY = ((e.clientY - parentRect.top) / effectiveScaleY) - offsetY;
      
      const parentW = parentEl ? parentEl.offsetWidth : (parentRect.width / effectiveScaleX);
      const parentH = parentEl ? parentEl.offsetHeight : (parentRect.height / effectiveScaleY);
      const overflowX = parentW * 0.3;
      const overflowY = parentH * 0.3;
      newX = Math.max(-overflowX, Math.min(newX, (parentW - imageEl.offsetWidth) + overflowX));
      newY = Math.max(-overflowY, Math.min(newY, (parentH - imageEl.offsetHeight) + overflowY));
      
      const image = stickers[stickerIndex].images.find(i => i.id === imageId);
      if (image && initialDragPosition) {
        const deltaX = newX - initialDragPosition.x;
        const deltaY = newY - initialDragPosition.y;
        
        image.x = newX;
        image.y = newY;
        imageEl.style.left = `${newX}px`;
        imageEl.style.top = `${newY}px`;
        
        // If sync is enabled, move corresponding images in other stickers
        if (syncMoveEnabled) {
          const imageIndex = stickers[stickerIndex].images.findIndex(i => i.id === imageId);
          
          stickers.forEach((sticker, idx) => {
            if (idx !== stickerIndex && sticker.images && sticker.images[imageIndex]) {
              const correspondingImage = sticker.images[imageIndex];
              correspondingImage.x += deltaX;
              correspondingImage.y += deltaY;
              
              // Update visual position
              const correspondingEl = document.querySelector(`[data-sticker-index="${idx}"] [data-image-id="${correspondingImage.id}"]`);
              if (correspondingEl) {
                const correspondingParentEl = correspondingEl.parentElement;
                const corrParentW = correspondingParentEl ? correspondingParentEl.offsetWidth : parentW;
                const corrParentH = correspondingParentEl ? correspondingParentEl.offsetHeight : parentH;
                const corrOverflowX = corrParentW * 0.3;
                const corrOverflowY = corrParentH * 0.3;
                const corrMinX = -corrOverflowX;
                const corrMaxX = (corrParentW - correspondingEl.offsetWidth) + corrOverflowX;
                const corrMinY = -corrOverflowY;
                const corrMaxY = (corrParentH - correspondingEl.offsetHeight) + corrOverflowY;

                correspondingImage.x = Math.max(corrMinX, Math.min(correspondingImage.x, corrMaxX));
                correspondingImage.y = Math.max(corrMinY, Math.min(correspondingImage.y, corrMaxY));
                correspondingEl.style.left = `${correspondingImage.x}px`;
                correspondingEl.style.top = `${correspondingImage.y}px`;
              }
            }
          });
          
          // Update initial position for next movement
          initialDragPosition.x = newX;
          initialDragPosition.y = newY;
        }
      }
    }

    function startImageResize(e, stickerIndex, imageId) {
      e.stopPropagation();

      pushHistory();

      const imageEl = document.querySelector(`[data-sticker-index="${stickerIndex}"] [data-image-id="${imageId}"]`);
      const parentEl = imageEl ? imageEl.parentElement : null;
      const parentRect = parentEl ? parentEl.getBoundingClientRect() : null;
      const scaleX = (parentEl && parentRect && parentEl.offsetWidth)
        ? (parentRect.width / parentEl.offsetWidth)
        : 1;
      const scaleY = (parentEl && parentRect && parentEl.offsetHeight)
        ? (parentRect.height / parentEl.offsetHeight)
        : 1;

      resizingImage = {
        stickerIndex,
        imageId,
        startClientX: e.clientX,
        startClientY: e.clientY,
        startWidth: imageEl ? imageEl.offsetWidth : 0,
        startHeight: imageEl ? imageEl.offsetHeight : 0,
        scaleX: scaleX || 1,
        scaleY: scaleY || 1
      };
      
      document.addEventListener('mousemove', resizeImage);
      document.addEventListener('mouseup', stopImageResize);
    }

    function resizeImage(e) {
      if (!resizingImage) return;
      
      const { stickerIndex, imageId, startClientX, startClientY, startWidth, startHeight, scaleX, scaleY } = resizingImage;
      const image = stickers[stickerIndex].images.find(i => i.id === imageId);
      if (!image) return;
      
      const imageEl = document.querySelector(`[data-sticker-index="${stickerIndex}"] [data-image-id="${imageId}"]`);
      if (!imageEl) return;

      const effectiveScaleX = scaleX || 1;
      const effectiveScaleY = scaleY || 1;
      const deltaX = (e.clientX - startClientX) / effectiveScaleX;
      const deltaY = (e.clientY - startClientY) / effectiveScaleY;
      const resizeDelta = (Math.abs(deltaX) >= Math.abs(deltaY)) ? deltaX : deltaY;
      const newWidth = startWidth + resizeDelta;
      const newHeight = startHeight + resizeDelta;
      
      if (newWidth > 20 && newHeight > 20) {
        let aspectRatio;
        if (Number.isFinite(image.originalWidth) && Number.isFinite(image.originalHeight) && image.originalHeight > 0) {
          aspectRatio = image.originalWidth / image.originalHeight;
        } else if (Number.isFinite(image.width) && Number.isFinite(image.height) && image.height > 0) {
          aspectRatio = image.width / image.height;
        } else if (startHeight > 0) {
          aspectRatio = startWidth / startHeight;
        } else if (imageEl.offsetHeight > 0) {
          aspectRatio = imageEl.offsetWidth / imageEl.offsetHeight;
        } else {
          aspectRatio = 1;
        }

        if (!Number.isFinite(aspectRatio) || aspectRatio <= 0) aspectRatio = 1;
        const calculatedHeight = newWidth / aspectRatio;
        
        image.width = newWidth;
        image.height = calculatedHeight;
        
        imageEl.style.width = `${newWidth}px`;
        imageEl.style.height = `${calculatedHeight}px`;

        // התאמת מיקום לאחר שינוי גודל כדי שהאלמנט לא ייתקע "מחוץ" למדבקה
        const parentEl = imageEl.parentElement;
        const parentW = parentEl ? parentEl.offsetWidth : 0;
        const parentH = parentEl ? parentEl.offsetHeight : 0;
        if (parentW > 0 && parentH > 0) {
          const overflowX = parentW * 0.3;
          const overflowY = parentH * 0.3;
          const minX = -overflowX;
          const minY = -overflowY;
          const maxX = (parentW - imageEl.offsetWidth) + overflowX;
          const maxY = (parentH - imageEl.offsetHeight) + overflowY;

          image.x = Math.max(minX, Math.min(image.x, maxX));
          image.y = Math.max(minY, Math.min(image.y, maxY));
          imageEl.style.left = `${image.x}px`;
          imageEl.style.top = `${image.y}px`;
        }
        
        // If sync is enabled, resize corresponding images in other stickers
        if (syncMoveEnabled) {
          const imageIndex = stickers[stickerIndex].images.findIndex(i => i.id === imageId);
          
          stickers.forEach((sticker, idx) => {
            if (idx !== stickerIndex && sticker.images && sticker.images[imageIndex]) {
              const correspondingImage = sticker.images[imageIndex];
              correspondingImage.width = newWidth;
              correspondingImage.height = calculatedHeight;
              
              // Update visual size
              const correspondingEl = document.querySelector(`[data-sticker-index="${idx}"] [data-image-id="${correspondingImage.id}"]`);
              if (correspondingEl) {
                correspondingEl.style.width = `${newWidth}px`;
                correspondingEl.style.height = `${calculatedHeight}px`;

                const correspondingParentEl = correspondingEl.parentElement;
                const corrParentW = correspondingParentEl ? correspondingParentEl.offsetWidth : 0;
                const corrParentH = correspondingParentEl ? correspondingParentEl.offsetHeight : 0;
                if (corrParentW > 0 && corrParentH > 0) {
                  const corrOverflowX = corrParentW * 0.3;
                  const corrOverflowY = corrParentH * 0.3;
                  const corrMinX = -corrOverflowX;
                  const corrMinY = -corrOverflowY;
                  const corrMaxX = (corrParentW - correspondingEl.offsetWidth) + corrOverflowX;
                  const corrMaxY = (corrParentH - correspondingEl.offsetHeight) + corrOverflowY;

                  correspondingImage.x = Math.max(corrMinX, Math.min(correspondingImage.x, corrMaxX));
                  correspondingImage.y = Math.max(corrMinY, Math.min(correspondingImage.y, corrMaxY));
                  correspondingEl.style.left = `${correspondingImage.x}px`;
                  correspondingEl.style.top = `${correspondingImage.y}px`;
                }
              }
            }
          });
        }
      }
    }

    function stopImageResize() {
      resizingImage = null;
      document.removeEventListener('mousemove', resizeImage);
      document.removeEventListener('mouseup', stopImageResize);
    }

    async function removeBackgroundFromSelected(stickerIndex, type, elementId) {
      let target;
      if (type === 'sticker') {
        target = stickers[stickerIndex];
      } else if (type === 'image') {
        target = stickers[stickerIndex].images.find(img => img.id === elementId);
      }

      if (type === 'sticker') {
        const images = (stickers[stickerIndex] && stickers[stickerIndex].images) ? stickers[stickerIndex].images : [];
        const preferredImageId = (typeof selectedSticker !== 'undefined' && typeof selectedImage !== 'undefined' && selectedSticker === stickerIndex && selectedImage)
          ? selectedImage
          : ((images.length === 1 && images[0] && images[0].id) ? images[0].id : null);
        if (preferredImageId) {
          type = 'image';
          elementId = preferredImageId;
          target = images.find(img => img.id === preferredImageId);
        }
      }

      if (!target || !target.dataUrl) {
        showStatus('לא נמצאה תמונה לעיבוד', true);
        return;
      }

      // Don't process SVGs as they usually don't have backgrounds in the same way
      if (target.dataUrl.startsWith('data:image/svg+xml')) {
        showStatus('הסרת רקע אינה נתמכת בקבצי SVG', true);
        return;
      }

      // בדיקה אם הספרייה זמינה (מקומית או CDN)
      showStatus('טוען כלי הסרת רקע (AI)...');
      let removeFunc = window.imglyRemoveBackground || window.removeBackground;
      if (typeof removeFunc === 'undefined' && typeof window.ensureBackgroundRemovalLibraryLoaded === 'function') {
        try {
          removeFunc = await window.ensureBackgroundRemovalLibraryLoaded();
        } catch (e) {
          console.error('Failed to load background removal library:', e);
          showStatus('שגיאה בטעינת כלי הסרת הרקע. בדוק חיבור/חסימה לרשת.', true);
          return;
        }
      }
      if (typeof removeFunc === 'undefined') {
        showStatus('ספריית הסרת הרקע לא זמינה. אנא רענן את הדף ונסה שוב.', true);
        console.error('Background removal library not available');
        return;
      }

      showStatus('מעבד תמונה להסרת רקע... (עשוי לקחת כמה שניות)');
      
      try {
        // Convert dataURL to Blob/File for the library
        const response = await fetch(target.dataUrl);
        const blob = await response.blob();
        
        // Run AI removal - עובד גם עם CDN וגם עם קבצים מקומיים
        const config = {
          publicPath: window.IMGLY_BG_REMOVAL_PUBLIC_PATH,
          debug: false,
          progress: (key, current, total) => {
            const percent = Math.round((current / total) * 100);
            showStatus(`מסיר רקע... ${percent}%`);
          }
        };

        const resultBlob = await removeFunc(blob, config);

        // Convert back to dataURL
        const reader = new FileReader();
        reader.onloadend = () => {
          pushHistory();
          target.dataUrl = reader.result;
          renderStickers();
          showStatus('הרקע הוסר בהצלחה! (התוצאה שקופה - על רקע לבן זה יכול להיראות דומה)');
        };
        reader.readAsDataURL(resultBlob);

      } catch (error) {
        console.error('Background removal error:', error);
        
        let errorMessage = error.message;
        if (errorMessage.includes('fetch') || errorMessage.includes('model') || error.name === 'TypeError') {
          errorMessage = 'שגיאה בטעינת מודל ה-AI. ייתכן שהחיבור לאינטרנט איטי או שחסימת "נטפרי" מונעת את הורדת המודל.';
        }
        
        showStatus('שגיאה: ' + errorMessage, true);
      }
    }

    function updateFileCount() {
      const fileCount = document.getElementById('fileCount');
      if (stickers.length > 0) {
        fileCount.textContent = `${stickers.length} מדבקות`;
      } else {
        fileCount.textContent = '';
      }
    }

    function normalizeCanvasForPdf(sourceCanvas, pdfWidthMm, pdfHeightMm) {
      const dpi = (EXPORT_QUALITY && Number.isFinite(EXPORT_QUALITY.pdfDpi)) ? EXPORT_QUALITY.pdfDpi : 360;
      const targetW = Math.max(1, Math.round((pdfWidthMm / 25.4) * dpi));
      const targetH = Math.max(1, Math.round((pdfHeightMm / 25.4) * dpi));
      const target = document.createElement('canvas');
      target.width = targetW;
      target.height = targetH;
      const ctx = target.getContext('2d');
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(sourceCanvas, 0, 0, targetW, targetH);
      return target;
    }

    async function downloadAsPDF() {
      if (stickers.length === 0) {
        showStatus('אין מדבקות להורדה', true);
        return;
      }

      const preview = document.getElementById('printPreview');
      const { jsPDF } = window.jspdf;
      showStatus('מכין PDF...');

      try {
        const pages = preview.querySelectorAll('.print-page');
        const pdfOrientation = pageOrientation === 'landscape' ? 'l' : 'p';
        const pdf = new jsPDF({ orientation: pdfOrientation, unit: 'mm', format: 'a4', compress: true });
        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfHeight = pdf.internal.pageSize.getHeight();
        const pdfDpr = (typeof window !== 'undefined' && window.devicePixelRatio)
          ? Math.min(2, window.devicePixelRatio)
          : 1;
        const jpegQuality = (EXPORT_QUALITY && Number.isFinite(EXPORT_QUALITY.jpegQuality)) ? EXPORT_QUALITY.jpegQuality : 0.98;
        const pdfCompression = (EXPORT_QUALITY && EXPORT_QUALITY.pdfCompression) ? EXPORT_QUALITY.pdfCompression : 'NONE';

        if (pages.length === 0) {
          const canvas = await captureElementToCanvas(preview, { scale: EXPORT_QUALITY.pdfScale, dpr: pdfDpr });
          const normalized = normalizeCanvasForPdf(canvas, pdfWidth, pdfHeight);
          const imgData = normalized.toDataURL('image/jpeg', jpegQuality);
          pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight, undefined, pdfCompression);
        } else {
          for (let i = 0; i < pages.length; i++) {
            const canvas = await captureElementToCanvas(pages[i], { scale: EXPORT_QUALITY.pdfScale, dpr: pdfDpr });
            const normalized = normalizeCanvasForPdf(canvas, pdfWidth, pdfHeight);
            const imgData = normalized.toDataURL('image/jpeg', jpegQuality);
            if (i > 0) {
              pdf.addPage();
            }
            pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight, undefined, pdfCompression);
          }
        }

        pdf.save('מדבקות.pdf');
        showStatus('PDF הורד בהצלחה! ✓');
      } catch (error) {
        console.error('PDF Error:', error);
        showStatus('שגיאה בהורדת PDF', true);
      }
    }

    async function downloadAsImage() {
      if (stickers.length === 0) {
        showStatus('אין מדבקות להורדה', true);
        return;
      }

      const preview = document.getElementById('printPreview');
      showStatus('מכין תמונה...');

      try {
        await downloadPreviewAsZip(preview, 'מדבקות');
      } catch (error) {
        console.error('Image Error:', error);
        showStatus('שגיאה בהורדת תמונה', true);
      }
    }

    async function downloadPreviewAsZip(preview, zipBaseName) {
      if (typeof JSZip === 'undefined') {
        showStatus('שגיאה: JSZip לא נטען', true);
        return;
      }

      const pages = preview.querySelectorAll('.print-page');
      const nodes = pages.length ? Array.from(pages) : [preview];
      const zip = nodes.length > 1 ? new JSZip() : null;
      const maxBytes = EXPORT_QUALITY.zipMaxBytes;

      const canvasToPngBlob = (c) => new Promise((resolve) => {
        c.toBlob((blob) => resolve(blob), 'image/png');
      });

      const canvasToJpegBlob = (c, quality) => new Promise((resolve) => {
        c.toBlob((blob) => resolve(blob), 'image/jpeg', quality);
      });

      const downscaleCanvas = (sourceCanvas, scaleFactor) => {
        const target = document.createElement('canvas');
        target.width = Math.max(1, Math.round(sourceCanvas.width * scaleFactor));
        target.height = Math.max(1, Math.round(sourceCanvas.height * scaleFactor));
        const ctx = target.getContext('2d');
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';
        ctx.drawImage(sourceCanvas, 0, 0, target.width, target.height);
        return target;
      };

      for (let i = 0; i < nodes.length; i++) {
        showStatus(`מכין תמונה ${i + 1}/${nodes.length}...`);

        const baseCanvas = await captureElementToCanvas(nodes[i], { scale: EXPORT_QUALITY.imageScale });
        let workingCanvas = baseCanvas;

        let blob = await canvasToPngBlob(workingCanvas);
        let ext = 'png';

        if (!blob) {
          showStatus('שגיאה בהכנת תמונה', true);
          return;
        }

        if (blob.size > maxBytes) {
          ext = 'jpg';
          let quality = 0.98;

          for (let attempt = 0; attempt < 18; attempt++) {
            blob = await canvasToJpegBlob(workingCanvas, quality);

            if (!blob) {
              showStatus('שגיאה בהכנת תמונה', true);
              return;
            }

            if (blob.size > maxBytes) {
              if (quality > 0.88) {
                quality = Math.max(0.88, quality - 0.03);
                continue;
              }

              workingCanvas = downscaleCanvas(workingCanvas, 0.92);
              quality = 0.92;
            } else {
              break;
            }
          }
        }

        if (!blob || blob.size > maxBytes) {
          showStatus('התמונה גדולה מדי להורדה', true);
          return;
        }

        const fileName = nodes.length === 1
          ? `${zipBaseName}.${ext}`
          : `${zipBaseName}-עמוד-${i + 1}.${ext}`;

        if (nodes.length === 1) {
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = fileName;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
          showStatus('התמונה הורדה בהצלחה! ✓');
          return;
        }

        zip.file(fileName, blob);
      }

      showStatus('מכין קובץ ZIP...');
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${zipBaseName}.zip`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      showStatus('ZIP הורד בהצלחה! ✓');
    }

    function switchTab(mode) {
      currentMode = mode;

      const tabWords = document.getElementById('tabWords');
      const tabNumbers = document.getElementById('tabNumbers');
      const tabNamesLottery = document.getElementById('tabNamesLottery');
      const wordsContent = document.getElementById('wordsContent');
      const numbersContent = document.getElementById('numbersContent');
      const namesLotteryContent = document.getElementById('namesLotteryContent');
      const printPreviewSection = document.getElementById('printPreviewSection');
      
      // Get the buttons in the header section
      const downloadPdfBtn = document.getElementById('downloadPdfBtn');
      const downloadImageBtn = document.getElementById('downloadImageBtn');
      const printBtn = document.getElementById('printBtn');
      const saveProjectBtn = document.getElementById('saveProjectBtn');
      const saveToDriveBtn = document.getElementById('saveToDriveBtn');
      const loadProjectLabel = document.querySelector('label[for="loadProjectInput"]');

      // הסרת active מכל הטאבים
      tabWords.classList.remove('active');
      tabNumbers.classList.remove('active');
      if (tabNamesLottery) tabNamesLottery.classList.remove('active');
      wordsContent.classList.add('hidden');
      numbersContent.classList.add('hidden');
      if (namesLotteryContent) namesLotteryContent.classList.add('hidden');

      if (mode === 'words') {
        tabWords.classList.add('active');
        wordsContent.classList.remove('hidden');
        printPreviewSection.classList.remove('hidden');
        document.getElementById('pageTitle').textContent = 'עיצוב מדבקות';
        
        // Show word mode buttons
        if (downloadPdfBtn) downloadPdfBtn.classList.remove('hidden');
        if (downloadImageBtn) downloadImageBtn.classList.remove('hidden');
        if (printBtn) printBtn.classList.remove('hidden');
        if (saveProjectBtn) saveProjectBtn.classList.remove('hidden');
        if (loadProjectLabel) loadProjectLabel.classList.remove('hidden');
        // Show save to drive if signed in
        if (saveToDriveBtn && typeof GoogleDriveManager !== 'undefined' && GoogleDriveManager.isSignedIn()) {
          saveToDriveBtn.classList.remove('hidden');
        }
      } else if (mode === 'numbers') {
        tabNumbers.classList.add('active');
        numbersContent.classList.remove('hidden');
        printPreviewSection.classList.add('hidden');
        document.getElementById('pageTitle').textContent = 'הגרלת מספרים';
        
        // Hide word mode buttons
        if (downloadPdfBtn) downloadPdfBtn.classList.add('hidden');
        if (downloadImageBtn) downloadImageBtn.classList.add('hidden');
        if (printBtn) printBtn.classList.add('hidden');
        if (saveProjectBtn) saveProjectBtn.classList.add('hidden');
        if (loadProjectLabel) loadProjectLabel.classList.add('hidden');
        // Hide save to drive in numbers mode
        if (saveToDriveBtn) saveToDriveBtn.classList.add('hidden');
      } else if (mode === 'namesLottery') {
        if (tabNamesLottery) tabNamesLottery.classList.add('active');
        const namesLotteryContent = document.getElementById('namesLotteryContent');
        if (namesLotteryContent) namesLotteryContent.classList.remove('hidden');
        printPreviewSection.classList.add('hidden');
        document.getElementById('pageTitle').textContent = 'הגרלת שמות';
        
        // Hide most buttons
        if (downloadPdfBtn) downloadPdfBtn.classList.add('hidden');
        if (downloadImageBtn) downloadImageBtn.classList.add('hidden');
        if (printBtn) printBtn.classList.add('hidden');
        if (saveProjectBtn) saveProjectBtn.classList.add('hidden');
        if (loadProjectLabel) loadProjectLabel.classList.add('hidden');
        // Hide header save to drive (we have one in names section)
        if (saveToDriveBtn) saveToDriveBtn.classList.add('hidden');
        
        // Show names-specific save to drive button if signed in
        const saveToDriveNamesBtn = document.getElementById('saveToDriveNamesBtn');
        if (saveToDriveNamesBtn && typeof GoogleDriveManager !== 'undefined' && GoogleDriveManager.isSignedIn()) {
          saveToDriveNamesBtn.classList.remove('hidden');
        }
      }
    }
    
    function switchNumbersSubTab(subTab) {
      const numberingTab = document.getElementById('numbersSubTabNumbering');
      const lotteryTab = document.getElementById('numbersSubTabLottery');
      const numberingContent = document.getElementById('numberingSubContent');
      const lotteryContent = document.getElementById('lotterySubContent');
      
      // הסרת active מכל הטאבים הפנימיים
      numberingTab.classList.remove('active');
      lotteryTab.classList.remove('active');
      numberingContent.classList.add('hidden');
      lotteryContent.classList.add('hidden');
      
      if (subTab === 'numbering') {
        numberingTab.classList.add('active');
        numberingContent.classList.remove('hidden');
      } else {
        lotteryTab.classList.add('active');
        lotteryContent.classList.remove('hidden');
      }
    }

    function renderNumberedStickers() {
      const preview = document.getElementById('numbersPreview');
      const emptyState = document.getElementById('numbersEmptyState');
      const previewSection = document.getElementById('numbersPreviewSection');

      if (numberedStickers.length === 0) {
        preview.innerHTML = '';
        emptyState.classList.remove('hidden');
        previewSection.classList.add('hidden');
        return;
      }

      emptyState.classList.add('hidden');
      previewSection.classList.remove('hidden');

      preview.innerHTML = '';

      const fragment = document.createDocumentFragment();
      const maxPageIndex = numberedStickers.reduce((max, s) => Math.max(max, Number.isFinite(s.page) ? s.page : 0), 0);
      const pageCount = Math.max(1, maxPageIndex + 1);

      const pages = [];
      for (let p = 0; p < pageCount; p++) {
        const pageEl = document.createElement('div');
        pageEl.className = pageOrientation === 'landscape' ? 'print-page landscape' : 'print-page portrait';
        pageEl.dataset.pageIndex = p;
        fragment.appendChild(pageEl);
        pages.push(pageEl);
      }

      numberedStickers.forEach((sticker, index) => {
        const pageIndex = Number.isFinite(sticker.page) ? sticker.page : 0;
        const pageEl = pages[Math.max(0, Math.min(pageIndex, pages.length - 1))];

        const stickerDiv = document.createElement('div');
        stickerDiv.className = 'sticker-container';

        stickerDiv.style.left = `${sticker.x}px`;
        stickerDiv.style.top = `${sticker.y}px`;
        stickerDiv.style.width = `${sticker.width}px`;
        stickerDiv.style.height = `${sticker.height}px`;

        const img = document.createElement('img');
        img.src = sticker.dataUrl;
        img.style.width = '100%';
        img.style.height = '100%';
        img.style.objectFit = 'contain';
        img.style.pointerEvents = 'none';

        stickerDiv.appendChild(img);

        const numberEl = document.createElement('div');
        numberEl.className = 'text-word';
        numberEl.textContent = sticker.number;
        numberEl.style.left = `${sticker.numberX}px`;
        numberEl.style.top = `${sticker.numberY}px`;
        numberEl.style.color = sticker.numberColor;
        numberEl.style.fontSize = `${sticker.numberFontSize}px`;
        numberEl.style.fontFamily = sticker.numberFontFamily;
        numberEl.style.fontWeight = sticker.numberFontWeight;
        numberEl.style.transform = 'translate(-50%, -50%)';

        numberEl.addEventListener('mousedown', (e) => {
          e.stopPropagation();
          startNumberDrag(e);
        });

        stickerDiv.appendChild(numberEl);
        pageEl.appendChild(stickerDiv);
      });

      preview.replaceChildren(fragment);
    }

    let numberedStickersPerPage = 0;

    function generateNumberedStickers() {
      if (!singleStickerTemplate) {
        showStatus('העלה מדבקה אחת תחילה!', true);
        return;
      }

      const start = Number(document.getElementById('startNumber').value);
      const end = Number(document.getElementById('endNumber').value);
      if (!Number.isFinite(start) || !Number.isFinite(end) || start < 1 || end < 1 || end < start) {
        showStatus('טווח מספרים לא תקין', true);
        return;
      }

      const stickersPerRow = Math.max(1, Math.min(20, Number(document.getElementById('stickersPerRow').value) || 1));
      const spacing = Math.max(0, Number(document.getElementById('numberSpacing')?.value) || 0);
      const numberColor = document.getElementById('numberColorPicker').value;
      const numberFontFamily = document.getElementById('numberFontFamily').value;
      const numberFontWeight = document.getElementById('numberFontWeight').value;
      const numberFontSize = Math.max(8, Number(document.getElementById('numberFontSize').value) || 32);

      const padding = spacing;
      const pageWidth = 210 * 3.7795275591;
      const pageHeight = 297 * 3.7795275591;

      const maxStickerWidth = (pageWidth - (padding * (stickersPerRow + 1))) / stickersPerRow;
      const aspectRatio = singleStickerTemplate.width / singleStickerTemplate.height;
      const stickerHeight = maxStickerWidth / aspectRatio;

      const rowsPerPage = Math.max(1, Math.floor((pageHeight - padding) / (stickerHeight + padding)));
      numberedStickersPerPage = stickersPerRow * rowsPerPage;

      numberedStickers = [];
      for (let n = start; n <= end; n++) {
        const idx = numberedStickers.length;
        const page = Math.floor(idx / numberedStickersPerPage);
        const indexInPage = idx % numberedStickersPerPage;
        const col = indexInPage % stickersPerRow;
        const row = Math.floor(indexInPage / stickersPerRow);

        const x = padding + col * (maxStickerWidth + padding);
        const y = padding + row * (stickerHeight + padding);

        numberedStickers.push({
          id: `numbered-${n}`,
          dataUrl: singleStickerTemplate.dataUrl,
          fileName: singleStickerTemplate.fileName,
          page,
          x,
          y,
          width: maxStickerWidth,
          height: stickerHeight,
          number: String(n),
          numberX: maxStickerWidth / 2,
          numberY: stickerHeight / 2,
          numberColor,
          numberFontFamily,
          numberFontWeight,
          numberFontSize
        });
      }

      renderNumberedStickers();
      showStatus(`נוצרו ${numberedStickers.length} מדבקות ממוספרות!`);
    }

    function startNumberDrag(e) {
      const numberEl = e.currentTarget;
      const parentRect = numberEl.parentElement.getBoundingClientRect();
      const offsetX = e.clientX - parentRect.left;
      const offsetY = e.clientY - parentRect.top;

      numberDragStart = { offsetX, offsetY, parentRect };
      document.addEventListener('mousemove', dragNumber);
      document.addEventListener('mouseup', stopNumberDrag);
    }

    function dragNumber(e) {
      if (!numberDragStart) return;
      
      const { offsetX, offsetY, parentRect } = numberDragStart;
      
      const newX = e.clientX - parentRect.left;
      const newY = e.clientY - parentRect.top;

      numberedStickers.forEach(sticker => {
        sticker.numberX += newX - offsetX;
        sticker.numberY += newY - offsetY;
      });

      numberDragStart.offsetX = newX;
      numberDragStart.offsetY = newY;

      renderNumberedStickers();
    }

    function stopNumberDrag() {
      numberDragStart = null;
      document.removeEventListener('mousemove', dragNumber);
      document.removeEventListener('mouseup', stopNumberDrag);
    }

    function centerNumbers() {
      if (numberedStickers.length === 0) {
        showStatus('אין מדבקות!', true);
        return;
      }

      numberedStickers.forEach(sticker => {
        sticker.numberX = sticker.width / 2;
        sticker.numberY = sticker.height / 2;
      });

      renderNumberedStickers();
      showStatus('המספרים מורכזו בכל המדבקות!');
    }

    async function downloadNumbersAsPDF() {
      if (numberedStickers.length === 0) {
        showStatus('אין מדבקות להורדה', true);
        return;
      }

      const preview = document.getElementById('numbersPreview');
      const { jsPDF } = window.jspdf;
      showStatus('מכין PDF...');

      try {
        const pages = preview.querySelectorAll('.print-page');
        const pdf = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4', compress: true });
        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfHeight = pdf.internal.pageSize.getHeight();
        const pdfDpr = (typeof window !== 'undefined' && window.devicePixelRatio)
          ? Math.min(2, window.devicePixelRatio)
          : 1;
        const jpegQuality = (EXPORT_QUALITY && Number.isFinite(EXPORT_QUALITY.jpegQuality)) ? EXPORT_QUALITY.jpegQuality : 0.98;
        const pdfCompression = (EXPORT_QUALITY && EXPORT_QUALITY.pdfCompression) ? EXPORT_QUALITY.pdfCompression : 'NONE';

        if (pages.length === 0) {
          const canvas = await captureElementToCanvas(preview, { scale: EXPORT_QUALITY.pdfScale, dpr: pdfDpr });
          const normalized = normalizeCanvasForPdf(canvas, pdfWidth, pdfHeight);
          const imgData = normalized.toDataURL('image/jpeg', jpegQuality);
          pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight, undefined, pdfCompression);
        } else {
          for (let i = 0; i < pages.length; i++) {
            const canvas = await captureElementToCanvas(pages[i], { scale: EXPORT_QUALITY.pdfScale, dpr: pdfDpr });
            const normalized = normalizeCanvasForPdf(canvas, pdfWidth, pdfHeight);
            const imgData = normalized.toDataURL('image/jpeg', jpegQuality);
            if (i > 0) pdf.addPage();
            pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight, undefined, pdfCompression);
          }
        }

        pdf.save('מדבקות-ממוספרות.pdf');
        showStatus('PDF הורד בהצלחה! ✓');
      } catch (error) {
        console.error('PDF Error:', error);
        showStatus('שגיאה בהורדת PDF', true);
      }
    }

    async function downloadNumbersAsImage() {
      if (numberedStickers.length === 0) {
        showStatus('אין מדבקות להורדה', true);
        return;
      }

      const preview = document.getElementById('numbersPreview');
      showStatus('מכין תמונה...');

      try {
        await downloadPreviewAsZip(preview, 'מדבקות-ממוספרות');
      } catch (error) {
        console.error('Image Error:', error);
        showStatus('שגיאה בהורדת תמונה', true);
      }
    }

    function saveProject() {
      if (stickers.length === 0) {
        showStatus('אין פרויקט לשמור!', true);
        return;
      }
      
      const projectData = getProjectData();
      if (!projectData) return;
      
      const jsonString = JSON.stringify(projectData, null, 2);
      const blob = new Blob([jsonString], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = currentProjectFileName || `מדבקות-${Date.now()}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      
      showStatus('הפרויקט נשמר בהצלחה! ✓');
    }

    // Get project data for saving (used by both local save and Google Drive)
    function getProjectData() {
      if (stickers.length === 0) {
        return null;
      }
      
      return {
        version: '1.2',
        stickers: stickers,
        layout: stickerLayoutConfig,
        pageOrientation: pageOrientation,
        savedAt: new Date().toISOString()
      };
    }

    // Clear current project (used for new project)
    function clearProject() {
      pushHistory();
      stickers = [];
      selectedSticker = null;
      selectedWord = null;
      selectedImage = null;
      wordIdCounter = 0;
      wordSeriesCounter = 0;
      imageIdCounter = 0;
      imageSeriesCounter = 0;
      currentProjectFileName = null;
      
      renderStickers();
      updateFileCount();
      updateUndoRedoButtons();
    }

    // Load project data (used by both local load and Google Drive)
    function loadProjectData(projectData, fileName) {
      try {
        if (!projectData.stickers || !Array.isArray(projectData.stickers)) {
          showStatus('קובץ לא תקין!', true);
          return false;
        }
        
        if (fileName) {
          currentProjectFileName = fileName;
        }
        
        stickers = projectData.stickers;
        if (projectData.layout && typeof projectData.layout === 'object') {
          stickerLayoutConfig = {
            uploadLimit: Number.isFinite(projectData.layout.uploadLimit) ? projectData.layout.uploadLimit : 0,
            stickersPerRow: Number.isFinite(projectData.layout.stickersPerRow) ? projectData.layout.stickersPerRow : 2,
            edgeMargin: Number.isFinite(projectData.layout.edgeMargin) ? projectData.layout.edgeMargin : 1,
            gap: Number.isFinite(projectData.layout.gap) ? projectData.layout.gap : 1,
            sizeMode: (String(projectData.layout.sizeMode || 'width').toLowerCase() === 'height') ? 'height' : 'width'
          };
        }
        
        // Load page orientation
        if (projectData.pageOrientation === 'landscape' || projectData.pageOrientation === 'portrait') {
          pageOrientation = projectData.pageOrientation;
          const orientationDropdownIcon = document.getElementById('orientationDropdownIcon');
          const orientationDropdownText = document.getElementById('orientationDropdownText');
          if (orientationDropdownIcon && orientationDropdownText) {
            if (pageOrientation === 'portrait') {
              orientationDropdownIcon.textContent = '📄';
              orientationDropdownText.textContent = 'כיוון הדף: לאורך';
            } else {
              orientationDropdownIcon.textContent = '📃';
              orientationDropdownText.textContent = 'כיוון הדף: לרוחב';
            }
          }
        }

        stickers.forEach(s => {
          s.words = s.words || [];
          s.images = s.images || [];
          if (!Number.isFinite(s.originalWidth) || !Number.isFinite(s.originalHeight) || s.originalWidth <= 0 || s.originalHeight <= 0) {
            const fallbackW = Number.isFinite(s.width) && s.width > 0 ? s.width : 1;
            const fallbackH = Number.isFinite(s.height) && s.height > 0 ? s.height : 1;
            s.originalWidth = fallbackW;
            s.originalHeight = fallbackH;
          }
        });

        wordIdCounter = Math.max(...stickers.flatMap(s => 
          s.words.map(w => parseInt(w.id.split('-')[1]) || 0)
        ), 0);

        wordSeriesCounter = Math.max(0, ...stickers.flatMap(s =>
          (s.words || []).map(w => {
            const raw = (w && w.seriesId) ? String(w.seriesId) : '';
            const m = raw.match(/^series-(\d+)$/);
            return m ? (parseInt(m[1], 10) || 0) : 0;
          })
        ));

        imageIdCounter = Math.max(...stickers.flatMap(s => 
          (s.images || []).map(i => parseInt(i.id.split('-')[1]) || 0)
        ), 0);
        
        applyStickerLayoutConfigToUI();
        if (autoArrangeEnabled) {
          reflowStickersPositionsOnly();
        } else {
          reflowStickers();
        }
        renderStickers();
        
        // Force A4 dimensions after loading project - multiple calls to ensure it works
        applyPrintPreviewScale();
        setTimeout(() => {
          applyPrintPreviewScale();
        }, 50);
        setTimeout(() => {
          applyPrintPreviewScale();
        }, 200);
        
        updateFileCount();
        showStatus(`הפרויקט נטען בהצלחה! ${stickers.length} מדבקות`);
        return true;
      } catch (error) {
        console.error('Load Error:', error);
        showStatus('שגיאה בטעינת הפרויקט', true);
        return false;
      }
    }

    function loadProject(file) {
      if (!file) return;
      
      currentProjectFileName = file.name;
      
      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const projectData = JSON.parse(e.target.result);
          loadProjectData(projectData, file.name);
        } catch (error) {
          console.error('Load Error:', error);
          showStatus('שגיאה בטעינת הפרויקט', true);
        }
      };
      reader.readAsText(file);
    }

    function openGithubStickersPicker(onSelect) {
      const modal = document.getElementById('githubStickersModal');
      const grid = document.getElementById('githubStickersGrid');
      grid.innerHTML = '';

      if (!GITHUB_FILES || !GITHUB_FILES.stickers) {
        showStatus('לא נמצאו מדבקות במאגר', true);
        return;
      }

      const selectHandler = typeof onSelect === 'function' ? onSelect : addSingleStickerFromGithub;

      // Create sections for each category
      Object.keys(GITHUB_FILES.stickers).forEach(category => {
        const categoryFiles = GITHUB_FILES.stickers[category];
        if (!Array.isArray(categoryFiles) || categoryFiles.length === 0) return;

        // Create category header
        const categoryHeader = document.createElement('div');
        categoryHeader.className = 'category-header col-span-full text-lg font-bold text-gray-800 bg-gradient-to-r from-indigo-100 to-purple-100 px-4 py-2 rounded-lg border-2 border-indigo-200 mb-2';
        categoryHeader.textContent = t(categoryTranslations[category] || category);
        grid.appendChild(categoryHeader);

        // Create buttons for files in this category
        categoryFiles.forEach((fileName) => {
          const url = `${GITHUB_REPO.stickers}${encodeURIComponent(category)}/${encodeURIComponent(fileName)}`;

          const btn = document.createElement('button');
          btn.type = 'button';
          btn.className = 'group border-2 border-gray-200 rounded-xl overflow-hidden hover:border-indigo-400 hover:shadow-lg transition-all bg-white';

          const img = document.createElement('img');
          img.src = url;
          img.alt = fileName;
          img.loading = 'lazy';
          img.className = 'w-full h-32 object-contain bg-white';

          btn.appendChild(img);

          btn.addEventListener('click', async () => {
            await selectHandler(fileName, category);
            closeGithubStickersPicker();
          });

          grid.appendChild(btn);
        });
      });

      modal.classList.remove('hidden');
    }

    function closeGithubStickersPicker() {
      document.getElementById('githubStickersModal').classList.add('hidden');
    }

    function openGithubElementsPicker(addToAll) {
      const modal = document.getElementById('githubElementsModal');
      const grid = document.getElementById('githubElementsGrid');
      if (!modal || !grid) {
        showStatus('שגיאה: חלון האלמנטים לא נמצא', true);
        return;
      }

      if (!GITHUB_FILES || !GITHUB_FILES.elements) {
        showStatus('לא נמצאו אלמנטים במאגר', true);
        return;
      }

      if (addToAll) {
        if (stickers.length === 0) {
          showStatus('העלה מדבקות תחילה!', true);
          return;
        }
      } else {
        if (selectedSticker === null) {
          showStatus('בחר מדבקה תחילה! לחץ על מדבקה כדי לבחור אותה.', true);
          return;
        }
      }

      grid.innerHTML = '';

      // Create sections for each category
      Object.keys(GITHUB_FILES.elements).forEach(category => {
        const categoryFiles = GITHUB_FILES.elements[category];
        if (!Array.isArray(categoryFiles) || categoryFiles.length === 0) return;

        // Create category header
        const categoryHeader = document.createElement('div');
        categoryHeader.className = 'category-header col-span-full text-lg font-bold text-gray-800 bg-gradient-to-r from-pink-100 to-rose-100 px-4 py-2 rounded-lg border-2 border-pink-200 mb-2';
        categoryHeader.textContent = t(categoryTranslations[category] || category);
        grid.appendChild(categoryHeader);

        // Create buttons for files in this category
        categoryFiles.forEach((fileName) => {
          const url = `${GITHUB_REPO.elements}${encodeURIComponent(category)}/${encodeURIComponent(fileName)}`;

          const btn = document.createElement('button');
          btn.type = 'button';
          btn.className = 'group border-2 border-gray-200 rounded-xl overflow-hidden hover:border-indigo-400 hover:shadow-lg transition-all bg-white';

          const img = document.createElement('img');
          img.src = url;
          img.alt = fileName;
          img.loading = 'lazy';
          img.className = 'w-full h-32 object-contain bg-white';

          btn.appendChild(img);

          btn.addEventListener('click', async () => {
            await addElementFromGithub(fileName, addToAll, category);
            closeGithubElementsPicker();
          });

          grid.appendChild(btn);
        });
      });

      modal.classList.remove('hidden');
    }

    function closeGithubElementsPicker() {
      const modal = document.getElementById('githubElementsModal');
      if (modal) modal.classList.add('hidden');
    }

    async function addElementFromGithub(fileName, addToAll, category) {
      try {
        showStatus('טוען אלמנט מהמאגר...');

        const url = `${GITHUB_REPO.elements}${encodeURIComponent(category)}/${encodeURIComponent(fileName)}`;
        const dataUrl = await fetchImageAsDataUrl(url);

        const img = await new Promise((resolve, reject) => {
          const image = new Image();
          image.onload = () => resolve(image);
          image.onerror = () => reject(new Error(`Image failed to load: ${url}`));
          image.src = dataUrl;
        });

        if (addToAll) {
          addImageToAllStickers(dataUrl, img.width, img.height);
        } else {
          addImageToSelectedSticker(dataUrl, img.width, img.height);
        }

        showStatus('האלמנט נוסף בהצלחה!');
      } catch (error) {
        console.error('GitHub Element Error:', error);
        showStatus('שגיאה בטעינת אלמנט מהמאגר', true);
      }
    }

    async function addSingleStickerFromGithub(fileName, category) {
      try {
        showStatus('טוען מדבקה מהמאגר...');

        const url = `${GITHUB_REPO.stickers}${encodeURIComponent(category)}/${encodeURIComponent(fileName)}`;
        const dataUrl = await fetchImageAsDataUrl(url);

        const img = await new Promise((resolve, reject) => {
          const image = new Image();
          image.onload = () => resolve(image);
          image.onerror = () => reject(new Error(`Image failed to load: ${url}`));
          image.src = dataUrl;
        });

        const originalWidth = img.width;
        const originalHeight = img.height;

        const cfg = getStickerLayoutConfigFromUI();
        const desiredCount = Number.isFinite(cfg.uploadLimit) && cfg.uploadLimit > 0 ? cfg.uploadLimit : 1;
        const countToAdd = Math.max(1, desiredCount);
        const startIndex = stickers.length; // שמירת האינדקס של המדבקות החדשות

        pushHistory();

        for (let i = 0; i < countToAdd; i++) {
          stickers.push({
            id: `sticker-github-${Date.now()}-${i}`,
            dataUrl,
            fileName,
            category,
            page: 0,
            x: 0,
            y: 0,
            width: 1,
            height: 1,
            originalWidth,
            originalHeight,
            words: [],
            images: []
          });
        }

        // החלת פריסה רק על המדבקות החדשות
        applyLayoutToNewStickers(startIndex);
        renderStickers();
        updateFileCount();
        showStatus('המדבקה נוספה למסמך!');
      } catch (error) {
        console.error('GitHub Single Sticker Error:', error);
        showStatus('שגיאה בטעינת מדבקה מהמאגר', true);
      }
    }

    async function setNumbersTemplateFromRepo(fileName, category) {
      try {
        showStatus('טוען מדבקה מהמאגר...');

        const url = `${GITHUB_REPO.stickers}${encodeURIComponent(category)}/${encodeURIComponent(fileName)}`;
        const dataUrl = await fetchImageAsDataUrl(url);

        const img = await new Promise((resolve, reject) => {
          const image = new Image();
          image.onload = () => resolve(image);
          image.onerror = () => reject(new Error(`Image failed to load: ${url}`));
          image.src = dataUrl;
        });

        singleStickerTemplate = {
          dataUrl,
          fileName,
          category,
          width: img.width,
          height: img.height
        };

        document.getElementById('stickerFileName').textContent = `✓ ${fileName}`;
        showStatus(`המדבקה "${fileName}" הועלתה בהצלחה!`);
        checkNumberStickerAndUpdateButtons();
      } catch (error) {
        console.error('Numbers Repo Template Error:', error);
        showStatus('שגיאה בטעינת מדבקה מהמאגר', true);
      }
    }

    async function setNamesNoteTemplateFromRepo(fileName, category) {
      try {
        showStatus('טוען תמונת פתק מהמאגר...');

        const url = `${GITHUB_REPO.stickers}${encodeURIComponent(category)}/${encodeURIComponent(fileName)}`;
        const dataUrl = await fetchImageAsDataUrl(url);

        const img = await new Promise((resolve, reject) => {
          const image = new Image();
          image.onload = () => resolve(image);
          image.onerror = () => reject(new Error(`Image failed to load: ${url}`));
          image.src = dataUrl;
        });

        namesNoteTemplate = {
          dataUrl,
          fileName,
          category,
          width: img.width,
          height: img.height
        };

        document.getElementById('namesNoteFileName').textContent = `✓ ${fileName}`;
        showStatus(`תמונת הפתק "${fileName}" הועלתה בהצלחה!`);
        
        // שלב 2: הפעלת כל הכפתורים אחרי בחירת מדבקה מהמאגר
        enableNamesButtons();
      } catch (error) {
        console.error('Names Note Repo Template Error:', error);
        showStatus('שגיאה בטעינת תמונת פתק מהמאגר', true);
      }
    }

    // Event Listeners
    
    // Toggle stickers menu
    document.getElementById('uploadStickersBtn').addEventListener('click', function(e) {
      e.stopPropagation();
      const menu = document.getElementById('stickersMenu');
      menu.classList.toggle('hidden');
      
      // Close other menus
      document.getElementById('imageAllMenu').classList.add('hidden');
      document.getElementById('imageSingleMenu').classList.add('hidden');
      const numbersMenu = document.getElementById('numbersStickerMenu');
      if (numbersMenu) numbersMenu.classList.add('hidden');
    });
    
    // Toggle image menus
    document.getElementById('addImageAllBtn').addEventListener('click', function(e) {
      e.stopPropagation();
      const menu = document.getElementById('imageAllMenu');
      menu.classList.toggle('hidden');
      
      document.getElementById('stickersMenu').classList.add('hidden');
      document.getElementById('imageSingleMenu').classList.add('hidden');
      const numbersMenu = document.getElementById('numbersStickerMenu');
      if (numbersMenu) numbersMenu.classList.add('hidden');
    });
    
    document.getElementById('addImageSingleBtn').addEventListener('click', function(e) {
      e.stopPropagation();
      const menu = document.getElementById('imageSingleMenu');
      menu.classList.toggle('hidden');
      
      document.getElementById('stickersMenu').classList.add('hidden');
      document.getElementById('imageAllMenu').classList.add('hidden');
      const numbersMenu = document.getElementById('numbersStickerMenu');
      if (numbersMenu) numbersMenu.classList.add('hidden');
    });

    // Numbers tab sticker template menu
    document.getElementById('uploadNumberStickerBtn').addEventListener('click', function(e) {
      e.stopPropagation();
      const menu = document.getElementById('numbersStickerMenu');
      menu.classList.toggle('hidden');

      document.getElementById('stickersMenu').classList.add('hidden');
      document.getElementById('imageAllMenu').classList.add('hidden');
      document.getElementById('imageSingleMenu').classList.add('hidden');
    });
    
    // Names Note Image Menu
    document.getElementById('namesNoteImageBtn').addEventListener('click', function(e) {
      e.stopPropagation();
      const menu = document.getElementById('namesNoteImageMenu');
      menu.classList.toggle('hidden');
      
      // Close other menus
      document.getElementById('stickersMenu').classList.add('hidden');
      document.getElementById('imageAllMenu').classList.add('hidden');
      document.getElementById('imageSingleMenu').classList.add('hidden');
      const numbersMenu = document.getElementById('numbersStickerMenu');
      if (numbersMenu) numbersMenu.classList.add('hidden');
    });
    
    // Close menus when clicking outside
    document.addEventListener('click', function(e) {
      document.getElementById('stickersMenu').classList.add('hidden');
      document.getElementById('imageAllMenu').classList.add('hidden');
      document.getElementById('imageSingleMenu').classList.add('hidden');
      const numbersMenu = document.getElementById('numbersStickerMenu');
      if (numbersMenu) numbersMenu.classList.add('hidden');
      const namesNoteMenu = document.getElementById('namesNoteImageMenu');
      if (namesNoteMenu) namesNoteMenu.classList.add('hidden');
      
      // Hide opacity control when clicking on empty page area
      const printPreview = document.getElementById('printPreview');
      if (printPreview && printPreview.contains(e.target)) {
        const clickedOnSticker = e.target.closest('.sticker-container');
        const clickedOnWord = e.target.closest('.text-word');
        const clickedOnImage = e.target.closest('.sticker-image');
        
        if (!clickedOnSticker && !clickedOnWord && !clickedOnImage) {
          hideOpacityControl();
          selectedSticker = null;
          selectedWord = null;
          selectedImage = null;
          document.querySelectorAll('.sticker-container').forEach(s => s.classList.remove('selected'));
          document.querySelectorAll('.text-word').forEach(w => w.classList.remove('selected'));
          document.querySelectorAll('.sticker-image').forEach(i => i.classList.remove('selected'));
        }
      }
    });

    const TUTORIAL_BASE_URL = 'https://raw.githubusercontent.com/yoelyoel111/automatic-fishstick/main/';
    const TUTORIAL_VIDEOS = {
      Stickers: { file: 'מדבקות.webm', title: 'מדבקות' },
      Text: { file: 'טקסט.webm', title: 'הוספת טקסט' },
      Image: { file: 'תמונה.webm', title: 'הוספת תמונה' },
      Global: { file: 'הגדרות.webm', title: 'הגדרות כלליות' },
      NamesUpload: { file: 'הגרלתמספרים.webm', title: 'הגרלת מספרים' }
    };

    const tutorialPreloaders = {};
    function getTutorialVideoUrl(sectionKey) {
      const info = TUTORIAL_VIDEOS[sectionKey];
      if (!info || !info.file) return null;
      return `${TUTORIAL_BASE_URL}${encodeURIComponent(info.file)}`;
    }

    async function diagnoseTutorialVideoUrl(url) {
      try {
        const head = await fetch(url, { method: 'HEAD' });
        const contentType = head.headers.get('content-type') || '';
        const contentLength = head.headers.get('content-length') || '';
        console.log('Tutorial video HEAD:', { url, status: head.status, ok: head.ok, contentType, contentLength });
        if (!head.ok) {
          return { ok: false, reason: `HTTP ${head.status}` };
        }

        // Detect Git LFS pointer (raw returns small text file instead of binary video)
        const probe = await fetch(url, { headers: { Range: 'bytes=0-200' } });
        const text = await probe.text();
        if (text && text.startsWith('version https://git-lfs.github.com/spec/v1')) {
          return { ok: false, reason: 'GIT_LFS_POINTER' };
        }

        // If server says it's text/html or text/plain, it's suspicious (could be error page)
        if (contentType.includes('text/html')) {
          return { ok: false, reason: 'CONTENT_TYPE_HTML' };
        }

        return { ok: true, reason: 'OK' };
      } catch (e) {
        console.error('Tutorial video diagnose failed:', url, e);
        return { ok: false, reason: 'NETWORK_ERROR' };
      }
    }

    function preloadTutorialVideo(sectionKey) {
      const url = getTutorialVideoUrl(sectionKey);
      if (!url) return;
      if (tutorialPreloaders[sectionKey]) return;
      const v = document.createElement('video');
      v.preload = 'auto';
      v.muted = true;
      v.playsInline = true;
      v.src = url;
      try { v.load(); } catch (_) {}
      tutorialPreloaders[sectionKey] = v;
    }

    async function openTutorialVideoModal(sectionKey) {
      const url = getTutorialVideoUrl(sectionKey);
      if (!url) {
        showStatus('אין סרטון הדרכה למדור זה עדיין', true);
        return;
      }

      const modal = document.getElementById('tutorialVideoModal');
      const videoEl = document.getElementById('tutorialVideoEl');
      const titleEl = document.getElementById('tutorialVideoModalTitle');
      if (!modal || !videoEl) return;

      const info = TUTORIAL_VIDEOS[sectionKey];
      if (titleEl) titleEl.textContent = info && info.title ? `סרטון הדרכה - ${info.title}` : 'סרטון הדרכה';

      console.log('Tutorial video URL:', sectionKey, url);

      const diag = await diagnoseTutorialVideoUrl(url);
      if (!diag.ok) {
        if (diag.reason === 'GIT_LFS_POINTER') {
          showStatus('הסרטון עלה ל-GitHub כ-Git LFS. raw לא מחזיר וידאו אלא קובץ מצביע. העלה את הקובץ בלי LFS או השתמש בלינק ישיר מתאים.', true);
        } else if (diag.reason === 'CONTENT_TYPE_HTML') {
          showStatus('הקישור ל-raw מחזיר HTML במקום וידאו (כנראה דף שגיאה). בדוק שהקובץ באמת בתיקייה הראשית ושאפשר לפתוח את ה-raw בדפדפן.', true);
        } else if (diag.reason === 'HTTP 418') {
          showStatus('הסינון (NetFree) חוסם את raw.githubusercontent.com (HTTP 418). צריך לאשר את הדומיין/הקישור בסינון או לאחסן את הסרטונים במקום אחר (למשל Releases/Drive עם לינק ישיר).', true);
        } else if (diag.reason.startsWith('HTTP ')) {
          showStatus(`שגיאה בטעינת הסרטון: ${diag.reason}. בדוק שם קובץ/תיקייה ב-GitHub.`, true);
        } else {
          showStatus('שגיאת רשת בטעינת הסרטון. נסה לפתוח את קישור ה-raw בדפדפן.', true);
        }
        return;
      }

      videoEl.pause();
      videoEl.preload = 'auto';
      videoEl.src = url;
      try { videoEl.load(); } catch (_) {}

      if (!videoEl.dataset.tutorialHandlersBound) {
        videoEl.dataset.tutorialHandlersBound = '1';
        videoEl.addEventListener('error', () => {
          const err = videoEl.error;
          const code = err ? err.code : 0;
          const codeText = code === 1 ? 'MEDIA_ERR_ABORTED'
            : code === 2 ? 'MEDIA_ERR_NETWORK'
            : code === 3 ? 'MEDIA_ERR_DECODE'
            : code === 4 ? 'MEDIA_ERR_SRC_NOT_SUPPORTED'
            : 'UNKNOWN';
          console.error('Tutorial video failed to load:', { sectionKey, url, code, codeText, err });
          showStatus('שגיאה בטעינת סרטון הדרכה. בדוק שהקובץ קיים ב-GitHub וניתן לפתוח את קישור ה-raw בדפדפן', true);
        });
      }

      modal.classList.remove('hidden');
      const playPromise = videoEl.play();
      if (playPromise && typeof playPromise.catch === 'function') playPromise.catch(() => {});
    }

    function closeTutorialVideoModal() {
      const modal = document.getElementById('tutorialVideoModal');
      const videoEl = document.getElementById('tutorialVideoEl');
      if (videoEl) {
        try { videoEl.pause(); } catch (_) {}
        videoEl.removeAttribute('src');
        try { videoEl.load(); } catch (_) {}
      }
      if (modal) modal.classList.add('hidden');
    }

    function setActiveToolsSection(sectionKey) {
      const sections = [
        { key: 'Stickers', contentId: 'toolsSectionContentStickers', chevronId: 'toolsSectionChevronStickers', buttonId: 'toolsSectionBtnStickers' },
        { key: 'Text', contentId: 'toolsSectionContentText', chevronId: 'toolsSectionChevronText', buttonId: 'toolsSectionBtnText' },
        { key: 'Image', contentId: 'toolsSectionContentImage', chevronId: 'toolsSectionChevronImage', buttonId: 'toolsSectionBtnImage' },
        { key: 'Global', contentId: 'toolsSectionContentGlobal', chevronId: 'toolsSectionChevronGlobal', buttonId: 'toolsSectionBtnGlobal' }
      ];

      sections.forEach(s => {
        const content = document.getElementById(s.contentId);
        const chevron = document.getElementById(s.chevronId);
        const button = document.getElementById(s.buttonId);
        const isOpen = s.key === sectionKey;
        
        if (content) content.classList.toggle('hidden', !isOpen);
        if (chevron) chevron.textContent = isOpen ? '▾' : '▸';
        
        // הוספת/הסרת המחלקה collapsed לכפתור
        if (button) button.classList.toggle('collapsed', !isOpen);
      });

      // Initialize color palette when Text section is opened
      if (sectionKey === 'Text') {
        initializeColorPalette();
      }

      preloadTutorialVideo(sectionKey);
    }

    const toolsSectionBtnStickers = document.getElementById('toolsSectionBtnStickers');
    if (toolsSectionBtnStickers) toolsSectionBtnStickers.addEventListener('click', () => setActiveToolsSection('Stickers'));
    const toolsSectionBtnText = document.getElementById('toolsSectionBtnText');
    if (toolsSectionBtnText) toolsSectionBtnText.addEventListener('click', () => setActiveToolsSection('Text'));
    const toolsSectionBtnImage = document.getElementById('toolsSectionBtnImage');
    if (toolsSectionBtnImage) toolsSectionBtnImage.addEventListener('click', () => setActiveToolsSection('Image'));
    const toolsSectionBtnGlobal = document.getElementById('toolsSectionBtnGlobal');
    if (toolsSectionBtnGlobal) toolsSectionBtnGlobal.addEventListener('click', () => setActiveToolsSection('Global'));

    document.querySelectorAll('[data-tutorial-section]').forEach(el => {
      const sectionKey = el.dataset.tutorialSection;
      el.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        openTutorialVideoModal(sectionKey);
      });
      el.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          e.stopPropagation();
          openTutorialVideoModal(sectionKey);
        }
      });
    });

    setActiveToolsSection('Stickers');

    const tutorialVideoModalClose = document.getElementById('tutorialVideoModalClose');
    if (tutorialVideoModalClose) tutorialVideoModalClose.addEventListener('click', closeTutorialVideoModal);
    const tutorialVideoModalBackdrop = document.getElementById('tutorialVideoModalBackdrop');
    if (tutorialVideoModalBackdrop) tutorialVideoModalBackdrop.addEventListener('click', closeTutorialVideoModal);

    function openUserGuideModal() {
      const modal = document.getElementById('userGuideModal');
      if (modal) modal.classList.remove('hidden');
    }

    function closeUserGuideModal() {
      const modal = document.getElementById('userGuideModal');
      if (modal) modal.classList.add('hidden');
    }

    const userGuideBtn = document.getElementById('userGuideBtn');
    if (userGuideBtn) {
      userGuideBtn.addEventListener('click', function(e) {
        e.preventDefault();
        openUserGuideModal();
      });
    }

    const userGuideModalClose = document.getElementById('userGuideModalClose');
    if (userGuideModalClose) userGuideModalClose.addEventListener('click', closeUserGuideModal);
    const userGuideModalBackdrop = document.getElementById('userGuideModalBackdrop');
    if (userGuideModalBackdrop) userGuideModalBackdrop.addEventListener('click', closeUserGuideModal);

    function openContactModal() {
      const modal = document.getElementById('contactModal');
      if (modal) modal.classList.remove('hidden');
    }

    function closeContactModal() {
      const modal = document.getElementById('contactModal');
      if (modal) modal.classList.add('hidden');
    }

    const contactBtn = document.getElementById('contactBtn');
    if (contactBtn) {
      contactBtn.addEventListener('click', function(e) {
        e.preventDefault();
        openContactModal();
      });
    }

    const contactModalClose = document.getElementById('contactModalClose');
    if (contactModalClose) contactModalClose.addEventListener('click', closeContactModal);
    const contactModalBackdrop = document.getElementById('contactModalBackdrop');
    if (contactModalBackdrop) contactModalBackdrop.addEventListener('click', closeContactModal);

    const tutorialVideoBtn = document.getElementById('tutorialVideoBtn');
    if (tutorialVideoBtn) {
      tutorialVideoBtn.addEventListener('click', function(e) {
        e.preventDefault();
        window.open('https://drive.google.com/file/d/1S3Q6sMVRQip_zNzWuxISKr81EYzSktbK/view?usp=sharing', '_blank', 'noopener');
      });
    }

    // Repo loaders
    document.getElementById('loadFromGithubStickers').addEventListener('click', function() {
      openGithubStickersPicker();
      document.getElementById('stickersMenu').classList.add('hidden');
    });

    document.getElementById('loadFromRepoNumbersSticker').addEventListener('click', function() {
      openGithubStickersPicker(setNumbersTemplateFromRepo);
      document.getElementById('numbersStickerMenu').classList.add('hidden');
    });

    document.getElementById('loadNamesNoteFromGithub').addEventListener('click', function() {
      openGithubStickersPicker(setNamesNoteTemplateFromRepo);
      document.getElementById('namesNoteImageMenu').classList.add('hidden');
    });

    document.getElementById('githubStickersModalClose').addEventListener('click', closeGithubStickersPicker);
    document.getElementById('githubStickersModalBackdrop').addEventListener('click', closeGithubStickersPicker);

    document.getElementById('githubElementsModalClose').addEventListener('click', closeGithubElementsPicker);
    document.getElementById('githubElementsModalBackdrop').addEventListener('click', closeGithubElementsPicker);

    document.getElementById('loadImageFromGithubAll').addEventListener('click', function() {
      openGithubElementsPicker(true);
      document.getElementById('imageAllMenu').classList.add('hidden');
    });
    
    document.getElementById('loadImageFromGithubSingle').addEventListener('click', function() {
      openGithubElementsPicker(false);
      document.getElementById('imageSingleMenu').classList.add('hidden');
    });
    
    document.getElementById('addWordToSelectedBtn').addEventListener('click', function() {
      addWordToSelectedSticker();
    });

    const decreaseSelectedWordFontBtn = document.getElementById('decreaseSelectedWordFontBtn');
    if (decreaseSelectedWordFontBtn) decreaseSelectedWordFontBtn.addEventListener('click', function() {
      adjustSelectedWordFontSize(-2);
    });

    const increaseSelectedWordFontBtn = document.getElementById('increaseSelectedWordFontBtn');
    if (increaseSelectedWordFontBtn) increaseSelectedWordFontBtn.addEventListener('click', function() {
      adjustSelectedWordFontSize(2);
    });

    document.getElementById('addToAllBtn').addEventListener('click', function() {
      addWordToAllStickers();
    });

    document.getElementById('replaceAllBtn').addEventListener('click', function() {
      replaceWordInAllStickers();
    });
    
    // Event listener לשינוי פונט - מעדכן טקסט נבחר
    const fontFamilyInput = document.getElementById('fontFamilyInput');
    if (fontFamilyInput) {
      fontFamilyInput.addEventListener('change', function() {
        if (selectedWord !== null) {
          applyFontToSelectedWord(this.value);
        }
      });
    }
    
    // Event listener לשינוי משקל פונט - מעדכן טקסט נבחר
    const fontWeightInput = document.getElementById('fontWeightInput');
    if (fontWeightInput) {
      fontWeightInput.addEventListener('change', function() {
        if (selectedWord !== null) {
          applyFontWeightToSelectedWord(this.value);
        }
      });
    }

    const syncElementsBtn = document.getElementById('syncElementsBtn');
    if (syncElementsBtn) syncElementsBtn.addEventListener('click', function() {
      const next = !(syncMoveEnabled && syncDeleteEnabled);
      syncMoveEnabled = next;
      syncDeleteEnabled = next;

      if (next) {
        syncElementsBtn.classList.remove('border-gray-300', 'hover:bg-gray-100');
        syncElementsBtn.classList.add('border-green-400', 'text-green-700', 'bg-green-50', 'hover:bg-green-100');
        showStatus('סנכרון אלמנטים הופעל ✓');
      } else {
        syncElementsBtn.classList.remove('border-green-400', 'text-green-700', 'bg-green-50', 'hover:bg-green-100');
        syncElementsBtn.classList.add('border-gray-300', 'hover:bg-gray-100');
        showStatus('סנכרון אלמנטים כבוי');
      }
    });

    const undoBtn = document.getElementById('undoBtn');
    if (undoBtn) undoBtn.addEventListener('click', undoLastAction);
    const redoBtn = document.getElementById('redoBtn');
    if (redoBtn) redoBtn.addEventListener('click', redoLastAction);

    // Opacity slider event listener
    const opacitySlider = document.getElementById('opacitySlider');
    const opacityValue = document.getElementById('opacityValue');
    if (opacitySlider) {
      opacitySlider.addEventListener('input', function() {
        const value = parseInt(this.value);
        if (opacityValue) {
          opacityValue.textContent = `${value}%`;
        }
        applyOpacity(value);
      });
    }

    function setLayoutMode(isAutoArrange) {
      autoArrangeEnabled = !!isAutoArrange;

      const freeEditBtn = document.getElementById('freeEditBtn');
      const autoArrangeBtn = document.getElementById('autoArrangeBtn');

      if (freeEditBtn && autoArrangeBtn) {
        if (autoArrangeEnabled) {
          // Free edit inactive
          freeEditBtn.classList.remove('border-green-400', 'text-green-700', 'bg-green-50', 'hover:bg-green-100');
          freeEditBtn.classList.add('border-gray-300', 'hover:bg-gray-100');

          // Auto arrange active
          autoArrangeBtn.classList.remove('border-gray-300', 'hover:bg-gray-100');
          autoArrangeBtn.classList.add('border-green-400', 'text-green-700', 'bg-green-50', 'hover:bg-green-100');
        } else {
          // Auto arrange inactive
          autoArrangeBtn.classList.remove('border-green-400', 'text-green-700', 'bg-green-50', 'hover:bg-green-100');
          autoArrangeBtn.classList.add('border-gray-300', 'hover:bg-gray-100');

          // Free edit active
          freeEditBtn.classList.remove('border-gray-300', 'hover:bg-gray-100');
          freeEditBtn.classList.add('border-green-400', 'text-green-700', 'bg-green-50', 'hover:bg-green-100');
        }
      }

      stickers.forEach(s => {
        if (s) s.lockPosition = autoArrangeEnabled;
      });

      if (autoArrangeEnabled) {
        getStickerLayoutConfigFromUI();
        updateStickerLayoutInfo();
        reflowStickersPositionsOnly();
        renderStickers();
        updateFileCount();
      }
    }

    const freeEditBtn = document.getElementById('freeEditBtn');
    if (freeEditBtn) freeEditBtn.addEventListener('click', function() {
      setLayoutMode(false);
      showStatus('עריכה חופשית פעילה');
    });

    const autoArrangeBtn = document.getElementById('autoArrangeBtn');
    if (autoArrangeBtn) autoArrangeBtn.addEventListener('click', function() {
      setLayoutMode(true);
      if (stickers.length > 0) {
        pushHistory();
        reflowStickersPositionsOnly();
        renderStickers();
        showStatus('המדבקות סודרו אוטומטית לפי מקום פנוי ✓');
      } else {
        showStatus('סדר אוטומטי פעיל ✓');
      }
    });

    const applyLayoutToAllBtn = document.getElementById('applyLayoutToAllBtn');
    if (applyLayoutToAllBtn) applyLayoutToAllBtn.addEventListener('click', function() {
      if (stickers.length === 0) {
        showStatus('אין מדבקות להחלת פריסה', true);
        return;
      }
      
      const confirmed = confirm('האם אתה בטוח שברצונך להחיל את הגדרות הגודל הנוכחיות על כל המדבקות הקיימות? פעולה זו תשנה את גודל כל המדבקות.');
      if (confirmed) {
        pushHistory();
        getStickerLayoutConfigFromUI();
        if (autoArrangeEnabled) {
          reflowStickersPositionsOnly();
        } else {
          reflowStickers();
        }
        renderStickers();
        showStatus('הגדרות הגודל הוחלו על כל המדבקות ✓');
      }
    });

    setLayoutMode(true);

    // Page orientation dropdown
    const orientationDropdownBtn = document.getElementById('orientationDropdownBtn');
    const orientationDropdownMenu = document.getElementById('orientationDropdownMenu');
    const orientationDropdownIcon = document.getElementById('orientationDropdownIcon');
    const orientationDropdownText = document.getElementById('orientationDropdownText');
    const orientationPortraitBtn = document.getElementById('orientationPortraitBtn');
    const orientationLandscapeBtn = document.getElementById('orientationLandscapeBtn');
    
    // Toggle dropdown menu
    if (orientationDropdownBtn) {
      orientationDropdownBtn.addEventListener('click', function(e) {
        e.stopPropagation();
        if (orientationDropdownMenu) {
          orientationDropdownMenu.classList.toggle('hidden');
        }
      });
    }
    
    // Close dropdown when clicking outside
    document.addEventListener('click', function(e) {
      if (orientationDropdownMenu && !orientationDropdownMenu.classList.contains('hidden')) {
        if (!e.target.closest('#orientationDropdownBtn') && !e.target.closest('#orientationDropdownMenu')) {
          orientationDropdownMenu.classList.add('hidden');
        }
      }
    });
    
    function updateOrientationDropdown(orientation) {
      if (orientationDropdownIcon && orientationDropdownText) {
        if (orientation === 'portrait') {
          orientationDropdownIcon.textContent = '📄';
          orientationDropdownText.textContent = 'כיוון הדף: לאורך';
        } else {
          orientationDropdownIcon.textContent = '📃';
          orientationDropdownText.textContent = 'כיוון הדף: לרוחב';
        }
      }
    }
    
    function setPageOrientation(orientation) {
      pageOrientation = orientation;
      
      // Update dropdown display
      updateOrientationDropdown(orientation);
      
      // Close dropdown menu
      if (orientationDropdownMenu) {
        orientationDropdownMenu.classList.add('hidden');
      }
      
      // Reflow stickers if any exist
      if (stickers.length > 0) {
        pushHistory();
        if (autoArrangeEnabled) {
          // במצב סידור אוטומטי - מחשב מחדש מיקומים וגדלים
          reflowStickersPositionsOnly();
        } else {
          // במצב עריכה חופשית - רק מחליף את כיוון הדף בלי לשנות גדלים
          console.log('Free edit mode - only changing page orientation, keeping sticker sizes');
        }
        renderStickers();
      } else {
        // Show empty page with correct orientation
        showEmptyPage();
      }
      
      // Update sticker layout info to reflect new page dimensions
      updateStickerLayoutInfo();
      
      showStatus(orientation === 'portrait' ? 'כיוון הדף: לאורך' : 'כיוון הדף: לרוחב');

      requestAnimationFrame(() => {
        try {
          const preview = document.getElementById('printPreview');
          if (preview) {
            preview.scrollLeft = Math.max(0, preview.scrollWidth - preview.clientWidth);
          }
        } catch (_) {}
      });
    }
    
    function showEmptyPage() {
      const preview = document.getElementById('printPreview');
      const previewInner = document.getElementById('printPreviewInner') || preview;
      const printPreviewSection = document.getElementById('printPreviewSection');
      
      if (!preview || !previewInner || !printPreviewSection) return;
      
      // Always show the preview section
      printPreviewSection.classList.remove('hidden');
      
      // Create empty page if no stickers
      if (stickers.length === 0) {
        const pageEl = document.createElement('div');
        pageEl.className = pageOrientation === 'landscape' ? 'print-page landscape' : 'print-page portrait';
        pageEl.dataset.pageIndex = 0;
        previewInner.replaceChildren(pageEl);
        applyPrintPreviewScale();
      }
    }
    
    if (orientationPortraitBtn) {
      orientationPortraitBtn.addEventListener('click', function() {
        setPageOrientation('portrait');
      });
    }
    
    if (orientationLandscapeBtn) {
      orientationLandscapeBtn.addEventListener('click', function() {
        setPageOrientation('landscape');
      });
    }
    
    // Show empty page on load
    showEmptyPage();
    
    // Initialize orientation dropdown display
    updateOrientationDropdown(pageOrientation);

    document.getElementById('imageInput').addEventListener('change', function(e) {
      const file = e.target.files[0];
      if (!file) return;

      if (!file.type.startsWith('image/')) {
        showStatus('ניתן להעלות רק קבצי תמונה', true);
        e.target.value = '';
        return;
      }

      const reader = new FileReader();
      reader.onload = function(event) {
        const img = new Image();
        img.onload = function() {
          addImageToAllStickers(event.target.result, img.width, img.height);
        };
        img.src = event.target.result;
      };
      reader.readAsDataURL(file);
      e.target.value = '';
    });

    document.getElementById('singleImageInput').addEventListener('change', function(e) {
      const file = e.target.files[0];
      if (!file) return;

      if (!file.type.startsWith('image/')) {
        showStatus('ניתן להעלות רק קבצי תמונה', true);
        e.target.value = '';
        return;
      }

      const reader = new FileReader();
      reader.onload = function(event) {
        const img = new Image();
        img.onload = function() {
          addImageToSelectedSticker(event.target.result, img.width, img.height);
        };
        img.src = event.target.result;
      };
      reader.readAsDataURL(file);
      e.target.value = '';
    });

    document.getElementById('downloadPdfBtn').addEventListener('click', function() {
      downloadAsPDF();
    });

    document.getElementById('downloadImageBtn').addEventListener('click', function() {
      downloadAsImage();
    });

    document.getElementById('printBtn').addEventListener('click', async function() {
      await printAsImage();
    });

    // פונקציה להדפסה באמצעות תמונה
    async function printAsImage() {
      if (stickers.length === 0) {
        showStatus('אין מדבקות להדפסה', true);
        return;
      }

      const preview = document.getElementById('printPreview');
      showStatus('מכין להדפסה...');

      try {
        // החלת כיול מדפסת על המדבקות לפני ההדפסה
        const originalSizes = [];
        let calibrationApplied = false;
        
        if (typeof getCalibrationScale === 'function') {
          const scale = getCalibrationScale();
          if (scale.scaleX !== 1 || scale.scaleY !== 1) {
            calibrationApplied = true;
            console.log('Applying printer calibration:', scale);
            
            // שמירת הגדלים המקוריים והחלת הכיול
            stickers.forEach((sticker, index) => {
              originalSizes[index] = {
                width: sticker.width,
                height: sticker.height,
                x: sticker.x,
                y: sticker.y
              };
              
              // החלת הכיול
              sticker.width = sticker.width * scale.scaleX;
              sticker.height = sticker.height * scale.scaleY;
            });
            
            // רענון התצוגה עם הגדלים המכוילים
            renderStickers();
            
            // המתנה קצרה לעדכון ה-DOM
            await new Promise(resolve => setTimeout(resolve, 100));
          }
        }

        const pages = preview.querySelectorAll('.print-page');
        const nodes = pages.length ? Array.from(pages) : [preview];
        
        // יצירת חלון הדפסה חדש
        const printWindow = window.open('', '_blank');
        if (!printWindow) {
          // שחזור הגדלים המקוריים אם נכשל
          if (calibrationApplied) {
            stickers.forEach((sticker, index) => {
              if (originalSizes[index]) {
                sticker.width = originalSizes[index].width;
                sticker.height = originalSizes[index].height;
                sticker.x = originalSizes[index].x;
                sticker.y = originalSizes[index].y;
              }
            });
            renderStickers();
          }
          showStatus('נא לאפשר חלונות קופצים להדפסה', true);
          return;
        }

        // הכנת ה-HTML של חלון ההדפסה
        const pageOrientation = document.querySelector('.print-page.landscape') ? 'landscape' : 'portrait';
        printWindow.document.write(`
          <!DOCTYPE html>
          <html>
          <head>
            <title>הדפסת מדבקות</title>
            <style>
              @page {
                size: A4 ${pageOrientation};
                margin: 0;
              }
              * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
              }
              body {
                margin: 0;
                padding: 0;
                background: white;
              }
              .print-page-img {
                width: ${pageOrientation === 'landscape' ? '297mm' : '210mm'};
                height: ${pageOrientation === 'landscape' ? '210mm' : '297mm'};
                display: block;
                page-break-after: always;
                object-fit: contain;
              }
              .print-page-img:last-child {
                page-break-after: auto;
              }
              @media print {
                body { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
                .print-page-img { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
              }
            </style>
          </head>
          <body>
        `);

        // יצירת תמונות לכל עמוד
        for (let i = 0; i < nodes.length; i++) {
          showStatus(`מכין עמוד ${i + 1}/${nodes.length}...`);
          
          const canvas = await captureElementToCanvas(nodes[i], { scale: EXPORT_QUALITY.imageScale });
          const dataUrl = canvas.toDataURL('image/png');
          
          printWindow.document.write(`<img class="print-page-img" src="${dataUrl}" />`);
        }

        printWindow.document.write('</body></html>');
        printWindow.document.close();

        // שחזור הגדלים המקוריים אחרי יצירת התמונות
        if (calibrationApplied) {
          stickers.forEach((sticker, index) => {
            if (originalSizes[index]) {
              sticker.width = originalSizes[index].width;
              sticker.height = originalSizes[index].height;
              sticker.x = originalSizes[index].x;
              sticker.y = originalSizes[index].y;
            }
          });
          renderStickers();
        }

        // המתנה לטעינת התמונות ואז הדפסה
        printWindow.onload = function() {
          setTimeout(() => {
            printWindow.print();
            // סגירת החלון אחרי ההדפסה
            printWindow.onafterprint = function() {
              printWindow.close();
            };
            // גיבוי - סגירה אחרי 60 שניות אם המשתמש ביטל
            setTimeout(() => {
              if (!printWindow.closed) {
                printWindow.close();
              }
            }, 60000);
          }, 500);
        };

        showStatus('מדפיס...');
      } catch (error) {
        console.error('Print Error:', error);
        showStatus('שגיאה בהדפסה', true);
      }
    }

    document.getElementById('saveProjectBtn').addEventListener('click', function() {
      saveProject();
    });

    document.getElementById('loadProjectInput').addEventListener('change', function(e) {
      const file = e.target.files[0];
      if (file) {
        loadProject(file);
      }
      e.target.value = '';
    });

    document.getElementById('fileInput').addEventListener('change', function(e) {
      const files = e.target.files;
      if (files.length === 0) return;

      showStatus('טוען מדבקות...');

      const cfg = getStickerLayoutConfigFromUI();
      const desiredCount = Number.isFinite(cfg.uploadLimit) && cfg.uploadLimit > 0 ? cfg.uploadLimit : 0;
      const actualCount = calculateActualStickerCount(desiredCount);
      const imageFiles = Array.from(files).filter(f => f.type && f.type.startsWith('image/'));

      if (imageFiles.length === 0) {
        showStatus('לא נמצאו קבצי תמונה להעלאה', true);
        e.target.value = '';
        return;
      }

      const countToAdd = actualCount > 0 ? actualCount : imageFiles.length;
      const startIndex = stickers.length; // שמירת האינדקס של המדבקות החדשות

      pushHistory();

      (async () => {
        try {
          const loaded = [];
          for (let i = 0; i < imageFiles.length; i++) {
            loaded.push(await readImageFileAsDataUrlWithSize(imageFiles[i]));
          }

          for (let i = 0; i < countToAdd; i++) {
            const src = loaded[i % loaded.length];
            stickers.push({
              id: `sticker-${Date.now()}-${i}`,
              dataUrl: src.dataUrl,
              fileName: src.fileName,
              page: 0,
              x: 0,
              y: 0,
              width: src.originalWidth,
              height: src.originalHeight,
              originalWidth: src.originalWidth,
              originalHeight: src.originalHeight,
              words: [],
              images: []
            });
          }

          // החלת פריסה על המדבקות החדשות
          // במקום לנסות לחשב מיקומים חלקיים, פשוט נריץ את reflowStickers על הכל
          reflowStickers();
          renderStickers();
          updateFileCount();
          showStatus(`${countToAdd} מדבקות הועלו בהצלחה!`);
        } catch (err) {
          console.error('Upload Error:', err);
          showStatus('שגיאה בהעלאת מדבקות', true);
        }
      })();

      e.target.value = '';
    });

    document.getElementById('folderInput').addEventListener('change', function(e) {
      const files = e.target.files;
      if (files.length === 0) return;

      showStatus('טוען תיקיית מדבקות...');

      const cfg = getStickerLayoutConfigFromUI();
      const desiredCount = Number.isFinite(cfg.uploadLimit) && cfg.uploadLimit > 0 ? cfg.uploadLimit : 0;
      const actualCount = calculateActualStickerCount(desiredCount);
      const imageFiles = Array.from(files).filter(f => f.type && f.type.startsWith('image/'));

      if (imageFiles.length === 0) {
        showStatus('לא נמצאו קבצי תמונה להעלאה', true);
        e.target.value = '';
        return;
      }

      const countToAdd = actualCount > 0 ? actualCount : imageFiles.length;
      const startIndex = stickers.length; // שמירת האינדקס של המדבקות החדשות

      pushHistory();

      (async () => {
        try {
          const loaded = [];
          for (let i = 0; i < imageFiles.length; i++) {
            loaded.push(await readImageFileAsDataUrlWithSize(imageFiles[i]));
          }

          for (let i = 0; i < countToAdd; i++) {
            const src = loaded[i % loaded.length];
            stickers.push({
              id: `sticker-${Date.now()}-${i}`,
              dataUrl: src.dataUrl,
              fileName: src.fileName,
              page: 0,
              x: 0,
              y: 0,
              width: src.originalWidth,
              height: src.originalHeight,
              originalWidth: src.originalWidth,
              originalHeight: src.originalHeight,
              words: [],
              images: []
            });
          }

          // החלת פריסה על המדבקות החדשות
          // במקום לנסות לחשב מיקומים חלקיים, פשוט נריץ את reflowStickers על הכל
          reflowStickers();
          renderStickers();
          updateFileCount();
          showStatus(`${countToAdd} מדבקות הועלו מהתיקייה בהצלחה!`);
        } catch (err) {
          console.error('Folder Upload Error:', err);
          showStatus('שגיאה בהעלאת מדבקות', true);
        }
      })();

      e.target.value = '';
    });

    document.getElementById('tabWords').addEventListener('click', function() {
      switchTab('words');
    });

    document.getElementById('tabNumbers').addEventListener('click', function() {
      switchTab('numbers');
    });

    document.getElementById('tabNamesLottery').addEventListener('click', function() {
      switchTab('namesLottery');
    });

    // Numbers Sub-tabs
    document.getElementById('numbersSubTabNumbering').addEventListener('click', function() {
      switchNumbersSubTab('numbering');
    });

    document.getElementById('numbersSubTabLottery').addEventListener('click', function() {
      switchNumbersSubTab('lottery');
    });

    // Names Lottery Sub-tabs
    document.getElementById('namesSubTabNotes').addEventListener('click', function() {
      switchNamesSubTab('notes');
    });

    document.getElementById('namesSubTabLottery').addEventListener('click', function() {
      switchNamesSubTab('lottery');
    });

    // Names Upload Type change
    const namesUploadType = document.getElementById('namesUploadType');
    if (namesUploadType) {
      namesUploadType.addEventListener('change', function() {
        if (typeof updateNamesUploadInstructions === 'function') {
          updateNamesUploadInstructions();
        }
      });
    }

    // Names Excel Input
    document.getElementById('namesExcelInput').addEventListener('change', function(e) {
      handleNamesExcelUpload(e);
    });

    // Download Names Template
    const downloadNamesTemplateBtn = document.getElementById('downloadNamesTemplateBtn');
    if (downloadNamesTemplateBtn) {
      downloadNamesTemplateBtn.addEventListener('click', function() {
        downloadNamesTemplate();
      });
    }

    // Generate Names Lottery - All
    const generateNamesLotteryAllBtn = document.getElementById('generateNamesLotteryAllBtn');
    if (generateNamesLotteryAllBtn) {
      generateNamesLotteryAllBtn.addEventListener('click', function() {
        generateNamesLottery('all');
      });
    }

    // Generate Names Lottery - By Group
    const generateNamesLotteryByGroupBtn = document.getElementById('generateNamesLotteryByGroupBtn');
    if (generateNamesLotteryByGroupBtn) {
      generateNamesLotteryByGroupBtn.addEventListener('click', function() {
        generateNamesLottery('byGroup');
      });
    }


    // Element SDK Integration
    if (window.elementSdk) {
      window.elementSdk.init({
        defaultConfig: defaultConfig,
        onConfigChange: onConfigChange,
        mapToCapabilities: (config) => ({
          recolorables: [
            {
              get: () => config.background_color || defaultConfig.background_color,
              set: (value) => {
                config.background_color = value;
                window.elementSdk.setConfig({ background_color: value });
              }
            },
            {
              get: () => config.button_color || defaultConfig.button_color,
              set: (value) => {
                config.button_color = value;
                window.elementSdk.setConfig({ button_color: value });
              }
            },
            {
              get: () => config.text_color || defaultConfig.text_color,
              set: (value) => {
                config.text_color = value;
                window.elementSdk.setConfig({ text_color: value });
              }
            }
          ],
          borderables: [],
          fontEditable: {
            get: () => config.font_family || defaultConfig.font_family,
            set: (value) => {
              config.font_family = value;
              window.elementSdk.setConfig({ font_family: value });
            }
          },
          fontSizeable: {
            get: () => config.font_size || defaultConfig.font_size,
            set: (value) => {
              config.font_size = value;
              window.elementSdk.setConfig({ font_size: value });
            }
          }
        }),
        mapToEditPanelValues: (config) => new Map([
          ['page_title', config.page_title || defaultConfig.page_title],
          ['default_word', config.default_word || defaultConfig.default_word]
        ])
      });
    }
    
    // Initialize names upload instructions on page load
    if (typeof updateNamesUploadInstructions === 'function') {
      updateNamesUploadInstructions();
    }

    // ========== פונקציות לניהול מצב כפתורים ==========
    
    function enableWordsButtons() {
      // הפעלת כפתורי הסרגל העליון
      const topButtons = ['downloadPdfBtn', 'downloadImageBtn', 'printBtn', 'saveProjectBtn'];
      topButtons.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
          btn.classList.remove('btn-disabled');
        }
      });

      // הפעלת כפתורי כלי העריכה
      const editButtons = [
        'undoBtn', 'redoBtn', 'decreaseSelectedWordFontBtn', 'increaseSelectedWordFontBtn',
        'addWordToSelectedBtn', 'addToAllBtn', 'replaceAllBtn', 'addImageAllBtn', 
        'addImageSingleBtn', 'syncElementsBtn', 'freeEditBtn', 'autoArrangeBtn'
      ];
      editButtons.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
          btn.classList.remove('btn-disabled');
        }
      });
    }

    function disableWordsButtons() {
      // השבתת כפתורי הסרגל העליון (חוץ מטען פרויקט)
      const topButtons = ['downloadPdfBtn', 'downloadImageBtn', 'printBtn', 'saveProjectBtn'];
      topButtons.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
          btn.classList.add('btn-disabled');
        }
      });

      // השבתת כפתורי כלי העריכה
      const editButtons = [
        'undoBtn', 'redoBtn', 'decreaseSelectedWordFontBtn', 'increaseSelectedWordFontBtn',
        'addWordToSelectedBtn', 'addToAllBtn', 'replaceAllBtn', 'addImageAllBtn', 
        'addImageSingleBtn', 'syncElementsBtn', 'freeEditBtn', 'autoArrangeBtn'
      ];
      editButtons.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
          btn.classList.add('btn-disabled');
        }
      });
    }

    function enableNumbersButtons() {
      const numbersButtons = ['generateNumbersBtn', 'centerNumbersBtn', 'downloadNumbersPdfBtn', 'downloadNumbersImageBtn'];
      numbersButtons.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
          btn.classList.remove('btn-disabled');
        }
      });
    }

    function disableNumbersButtons() {
      const numbersButtons = ['generateNumbersBtn', 'centerNumbersBtn', 'downloadNumbersPdfBtn', 'downloadNumbersImageBtn'];
      numbersButtons.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
          btn.classList.add('btn-disabled');
        }
      });
    }

    function enableLotteryButtons() {
      const lotteryButtons = ['generateLotteryBtn', 'downloadLotteryPdfBtn'];
      lotteryButtons.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
          btn.classList.remove('btn-disabled');
        }
      });
    }

    function disableLotteryButtons() {
      const lotteryButtons = ['generateLotteryBtn', 'downloadLotteryPdfBtn'];
      lotteryButtons.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
          btn.classList.add('btn-disabled');
        }
      });
    }

    function enableNamesButtons() {
      // הפעלת כל הכפתורים בטאב הגרלת שמות
      const namesButtons = [
        'generateNamesNotesBtn', 'centerNamesNotesBtn', 'downloadNamesNotesPdfBtn', 'printNamesNotesBtn',
        'generateNamesLotteryAllBtn', 'generateNamesLotteryByGroupBtn', 'downloadNamesLotteryPdfBtn', 'printNamesLotteryBtn', 'namesNoteImageBtn'
      ];
      namesButtons.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
          btn.classList.remove('btn-disabled');
        }
      });

      // הפעלת הטאבים הפנימיים
      const namesSubTabNotes = document.getElementById('namesSubTabNotes');
      const namesSubTabLottery = document.getElementById('namesSubTabLottery');
      if (namesSubTabNotes) namesSubTabNotes.classList.remove('btn-disabled');
      if (namesSubTabLottery) namesSubTabLottery.classList.remove('btn-disabled');
    }

    // שלב 1: הפעלת רק כפתור בחר מדבקה והטאבים (אחרי העלאת אקסל)
    function enableNamesImageButtonOnly() {
      const namesNoteImageBtn = document.getElementById('namesNoteImageBtn');
      if (namesNoteImageBtn) {
        namesNoteImageBtn.classList.remove('btn-disabled');
      }

      // הפעלת הטאבים הפנימיים
      const namesSubTabNotes = document.getElementById('namesSubTabNotes');
      const namesSubTabLottery = document.getElementById('namesSubTabLottery');
      if (namesSubTabNotes) namesSubTabNotes.classList.remove('btn-disabled');
      if (namesSubTabLottery) namesSubTabLottery.classList.remove('btn-disabled');
    }

    function disableNamesButtons() {
      // השבתת כל הכפתורים בטאב הגרלת שמות
      const namesButtons = [
        'generateNamesNotesBtn', 'centerNamesNotesBtn', 'downloadNamesNotesPdfBtn', 'printNamesNotesBtn',
        'generateNamesLotteryAllBtn', 'generateNamesLotteryByGroupBtn', 'downloadNamesLotteryPdfBtn', 'printNamesLotteryBtn', 'namesNoteImageBtn'
      ];
      namesButtons.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
          btn.classList.add('btn-disabled');
        }
      });

      // השבתת הטאבים הפנימיים
      const namesSubTabNotes = document.getElementById('namesSubTabNotes');
      const namesSubTabLottery = document.getElementById('namesSubTabLottery');
      if (namesSubTabNotes) namesSubTabNotes.classList.add('btn-disabled');
      if (namesSubTabLottery) namesSubTabLottery.classList.add('btn-disabled');
    }

    // בדיקה אם יש מדבקות והפעלה/השבתה בהתאם
    function checkStickersAndUpdateButtons() {
      if (stickers && stickers.length > 0) {
        enableWordsButtons();
      } else {
        disableWordsButtons();
      }
    }

    // בדיקה אם יש מדבקה למיספור
    function checkNumberStickerAndUpdateButtons() {
      const stickerFileName = document.getElementById('stickerFileName');
      if (stickerFileName && stickerFileName.textContent.trim() !== '') {
        enableNumbersButtons();
      } else {
        disableNumbersButtons();
      }
    }

    // בדיקה אם יש נתוני שמות
    function checkNamesDataAndUpdateButtons() {
      if (namesData && namesData.length > 0) {
        enableNamesButtons();
      } else {
        disableNamesButtons();
      }
    }

    // קריאה לפונקציה בכל פעם שמשתנה מצב המדבקות
    const originalRenderStickers = renderStickers;
    renderStickers = function() {
      originalRenderStickers();
      checkStickersAndUpdateButtons();
    };

    // השבתת כפתורים בהתחלה
    disableWordsButtons();
    disableNumbersButtons();
    enableLotteryButtons(); // הגרלת מספרים תמיד פעילה
    disableNamesButtons();

  
    // Color palette initialization function
    function initializeColorPalette() {
      const textColorSwatch = document.getElementById('textColorSwatch');
      const textColorPicker = document.getElementById('textColorPicker');
      const textColorPalette = document.getElementById('textColorPalette');
      const textColorPaletteGrid = document.getElementById('textColorPaletteGrid');

      // Skip if already initialized or elements don't exist
      if (!textColorSwatch || !textColorPicker || !textColorPalette || !textColorPaletteGrid) {
        return;
      }

      // Skip if already initialized
      if (textColorSwatch.dataset.initialized === 'true') {
        return;
      }

      // Gradient state
      let selectedGradientType = 'horizontal';
      let gradientColor1 = '#FF0000';
      let gradientColor2 = '#0000FF';
      let gradientColorMode = 0; // 0 = regular color mode, 1 = first gradient color, 2 = second gradient color

      const setSwatch = (color, isGradient = false) => {
        if (isGradient) {
          textColorSwatch.style.background = color;
          textColorSwatch.dataset.isGradient = 'true';
          textColorSwatch.dataset.gradientValue = color;
        } else {
          textColorSwatch.style.background = color;
          textColorSwatch.dataset.isGradient = 'false';
          textColorPicker.value = color;
        }
      };

      setSwatch(textColorPicker.value || '#000000');

      const paletteColors = [
        '#000000', '#444444', '#777777', '#BBBBBB', '#FFFFFF', '#FF0000',
        '#FF7A00', '#FFD400', '#00C853', '#00B0FF', '#2962FF',
        '#651FFF', '#D500F9', '#FF4081', '#8D6E63', '#1B5E20',
        '#004D40', '#01579B', '#1A237E', '#311B92', '#B71C1C',
        '#E65100', '#F57F17', '#33691E', '#263238', '#795548',
        '#607D8B', '#9E9E9E', '#FF5722', '#FF9800'
      ];

      const updateGradientPreview = () => {
        const preview = document.getElementById('gradientPreview');
        if (!preview) return;
        
        let gradientCSS = '';
        switch (selectedGradientType) {
          case 'horizontal':
            gradientCSS = `linear-gradient(to left, ${gradientColor1}, ${gradientColor2})`;
            break;
          case 'vertical':
            gradientCSS = `linear-gradient(to bottom, ${gradientColor1}, ${gradientColor2})`;
            break;
          case 'diagonal-down':
            gradientCSS = `linear-gradient(135deg, ${gradientColor1}, ${gradientColor2})`;
            break;
          case 'diagonal-up':
            gradientCSS = `linear-gradient(45deg, ${gradientColor1}, ${gradientColor2})`;
            break;
          case 'radial':
            gradientCSS = `radial-gradient(circle, ${gradientColor1}, ${gradientColor2})`;
            break;
        }
        preview.style.background = gradientCSS;
        preview.dataset.gradientValue = gradientCSS;
      };

      const updateGradientColorButtons = () => {
        const color1Btn = document.getElementById('gradientColor1Btn');
        const color2Btn = document.getElementById('gradientColor2Btn');
        if (color1Btn) {
          color1Btn.style.background = gradientColor1;
          // Update active state
          if (gradientColorMode === 1) {
            color1Btn.style.border = '3px solid #3b82f6';
            color1Btn.style.boxShadow = '0 0 0 2px rgba(59, 130, 246, 0.2)';
          } else {
            color1Btn.style.border = '2px solid rgba(0,0,0,0.15)';
            color1Btn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
          }
        }
        if (color2Btn) {
          color2Btn.style.background = gradientColor2;
          // Update active state
          if (gradientColorMode === 2) {
            color2Btn.style.border = '3px solid #3b82f6';
            color2Btn.style.boxShadow = '0 0 0 2px rgba(59, 130, 246, 0.2)';
          } else {
            color2Btn.style.border = '2px solid rgba(0,0,0,0.15)';
            color2Btn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
          }
        }
        updateGradientPreview();
      };

      textColorPaletteGrid.innerHTML = '';
      
      // Regular colors
      paletteColors.forEach((c) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'color-palette-color';
        btn.style.background = c;
        btn.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          
          // Check if we're in gradient mode (gradientColorMode 1 or 2)
          if (gradientColorMode === 1 || gradientColorMode === 2) {
            // Update the selected gradient color slot
            if (gradientColorMode === 1) {
              gradientColor1 = c;
            } else {
              gradientColor2 = c;
            }
            updateGradientColorButtons();
          } else {
            // Regular color mode - set as solid color and close palette
            setSwatch(c);
            // עדכן טקסט נבחר אם יש
            applyColorToSelectedWord(c, false);
            textColorPalette.classList.add('hidden');
          }
        });
        textColorPaletteGrid.appendChild(btn);
      });

      // Add gradient section
      const gradientSection = document.createElement('div');
      gradientSection.className = 'gradient-section';
      gradientSection.id = 'gradientSection';
      
      // Add click listener to gradient section for empty space clicks
      gradientSection.addEventListener('click', (e) => {
        // Check if click was on empty space within gradient section
        const isEmptySpace = e.target === gradientSection ||
                            e.target.classList.contains('gradient-separator') ||
                            e.target.classList.contains('gradient-types') ||
                            e.target.classList.contains('gradient-controls') ||
                            e.target.classList.contains('gradient-color-pickers');
        
        if (isEmptySpace) {
          // Reset gradient mode - hide visual indicators
          const color1Btn = document.getElementById('gradientColor1Btn');
          const color2Btn = document.getElementById('gradientColor2Btn');
          if (color1Btn) {
            color1Btn.style.border = '2px solid rgba(0,0,0,0.15)';
            color1Btn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
          }
          if (color2Btn) {
            color2Btn.style.border = '2px solid rgba(0,0,0,0.15)';
            color2Btn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
          }
          // Reset gradient mode to 0 (regular color mode)
          gradientColorMode = 0;
        }
      });
      
      const gradientSeparator = document.createElement('div');
      gradientSeparator.className = 'gradient-separator';
      gradientSeparator.textContent = 'גרדיאנט';
      gradientSection.appendChild(gradientSeparator);

      // Gradient type buttons
      const gradientTypes = document.createElement('div');
      gradientTypes.className = 'gradient-types';
      
      const gradientOptions = [
        { type: 'horizontal', label: 'אופקי', bg: 'linear-gradient(to left, #FF6B6B, #4ECDC4)' },
        { type: 'vertical', label: 'אנכי', bg: 'linear-gradient(to bottom, #FF6B6B, #4ECDC4)' },
        { type: 'diagonal-down', label: '↘', bg: 'linear-gradient(135deg, #FF6B6B, #4ECDC4)' },
        { type: 'diagonal-up', label: '↙', bg: 'linear-gradient(45deg, #FF6B6B, #4ECDC4)' },
        { type: 'radial', label: 'עגול', bg: 'radial-gradient(circle, #FF6B6B, #4ECDC4)' }
      ];

      gradientOptions.forEach(option => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'gradient-type-btn';
        btn.textContent = option.label;
        btn.style.background = option.bg;
        btn.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          selectedGradientType = option.type;
          
          // Update active state
          gradientTypes.querySelectorAll('.gradient-type-btn').forEach(b => {
            b.style.borderColor = 'rgba(0,0,0,0.15)';
          });
          btn.style.borderColor = '#3b82f6';
          
          updateGradientPreview();
        });
        gradientTypes.appendChild(btn);
      });
      
      // Set default active state
      gradientTypes.children[0].style.borderColor = '#3b82f6';
      
      gradientSection.appendChild(gradientTypes);

      // Gradient controls
      const gradientControls = document.createElement('div');
      gradientControls.className = 'gradient-controls';

      const gradientColorPickers = document.createElement('div');
      gradientColorPickers.className = 'gradient-color-pickers';

      // Color 1 picker
      const color1Btn = document.createElement('button');
      color1Btn.type = 'button';
      color1Btn.id = 'gradientColor1Btn';
      color1Btn.className = 'gradient-color-btn';
      color1Btn.style.background = gradientColor1;
      color1Btn.style.border = '2px solid rgba(0,0,0,0.15)'; // Not active by default
      color1Btn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
      color1Btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        gradientColorMode = 1;
        // Update active state
        color1Btn.style.border = '3px solid #3b82f6';
        color1Btn.style.boxShadow = '0 0 0 2px rgba(59, 130, 246, 0.2)';
        color2Btn.style.border = '2px solid rgba(0,0,0,0.15)';
        color2Btn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
      });

      // Color 2 picker
      const color2Btn = document.createElement('button');
      color2Btn.type = 'button';
      color2Btn.id = 'gradientColor2Btn';
      color2Btn.className = 'gradient-color-btn';
      color2Btn.style.background = gradientColor2;
      color2Btn.style.border = '2px solid rgba(0,0,0,0.15)'; // Not active by default
      color2Btn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
      color2Btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        gradientColorMode = 2;
        // Update active state
        color2Btn.style.border = '3px solid #3b82f6';
        color2Btn.style.boxShadow = '0 0 0 2px rgba(59, 130, 246, 0.2)';
        color1Btn.style.border = '2px solid rgba(0,0,0,0.15)';
        color1Btn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
      });

      gradientColorPickers.appendChild(color1Btn);
      gradientColorPickers.appendChild(color2Btn);
      gradientControls.appendChild(gradientColorPickers);

      // Gradient preview
      const gradientPreview = document.createElement('button');
      gradientPreview.type = 'button';
      gradientPreview.id = 'gradientPreview';
      gradientPreview.className = 'gradient-preview';
      gradientPreview.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const gradientValue = gradientPreview.dataset.gradientValue;
        if (gradientValue) {
          setSwatch(gradientValue, true);
          // עדכן טקסט נבחר אם יש
          applyColorToSelectedWord(gradientValue, true);
          textColorPalette.classList.add('hidden');
        }
      });
      gradientControls.appendChild(gradientPreview);

      gradientSection.appendChild(gradientControls);
      textColorPaletteGrid.appendChild(gradientSection);

      // Initialize gradient preview
      updateGradientPreview();

      // Event listeners
      textColorSwatch.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        textColorPalette.classList.toggle('hidden');
      });

      // Add click listener to palette background to exit gradient mode
      textColorPalette.addEventListener('click', (e) => {
        // Check if click was on empty space (not on any interactive element)
        const isEmptySpace = e.target === textColorPalette || 
                            e.target === textColorPaletteGrid ||
                            e.target.classList.contains('gradient-section') ||
                            e.target.classList.contains('gradient-separator') ||
                            e.target.classList.contains('gradient-types') ||
                            e.target.classList.contains('gradient-controls') ||
                            e.target.classList.contains('gradient-color-pickers');
        
        if (isEmptySpace) {
          // Reset gradient mode - hide visual indicators
          const color1Btn = document.getElementById('gradientColor1Btn');
          const color2Btn = document.getElementById('gradientColor2Btn');
          if (color1Btn) {
            color1Btn.style.border = '2px solid rgba(0,0,0,0.15)';
            color1Btn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
          }
          if (color2Btn) {
            color2Btn.style.border = '2px solid rgba(0,0,0,0.15)';
            color2Btn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
          }
          // Reset gradient mode to 0 (regular color mode)
          gradientColorMode = 0;
        }
      });

      document.addEventListener('click', (e) => {
        if (!textColorPalette.contains(e.target) && !textColorSwatch.contains(e.target)) {
          textColorPalette.classList.add('hidden');
        }
      });

      // Mark as initialized
      textColorSwatch.dataset.initialized = 'true';
    }
    // Function to update button states based on stickers count
    function updateButtonStates() {
      const hasStickers = stickers && stickers.length > 0;
      
      // List of buttons that should be disabled when no stickers
      const buttonsToDisable = [
        'downloadPdfBtn',
        'downloadImageBtn', 
        'printBtn',
        'saveProjectBtn'
      ];
      
      buttonsToDisable.forEach(buttonId => {
        const button = document.getElementById(buttonId);
        if (button) {
          if (hasStickers) {
            // Enable button - remove disabled styles
            button.disabled = false;
            button.classList.remove('btn-disabled');
            button.style.pointerEvents = '';
            button.style.opacity = '';
          } else {
            // Disable button - add disabled styles
            button.disabled = true;
            button.classList.add('btn-disabled');
            button.style.pointerEvents = 'none';
            button.style.opacity = '0.5';
          }
        }
      });
      
      // Load project button should always be enabled
      const loadProjectLabel = document.querySelector('label[for="loadProjectInput"]');
      if (loadProjectLabel) {
        loadProjectLabel.classList.remove('btn-disabled');
        loadProjectLabel.style.pointerEvents = '';
        loadProjectLabel.style.opacity = '';
      }
    }


    // Expose functions globally for Google Drive integration
    window.getProjectData = getProjectData;
    window.loadProjectData = loadProjectData;
    window.clearProject = clearProject;
    window.currentProjectFileName = null;
    
    // Getter/setter for currentProjectFileName
    Object.defineProperty(window, 'currentProjectFileName', {
      get: function() { return currentProjectFileName; },
      set: function(val) { currentProjectFileName = val; }
    });
