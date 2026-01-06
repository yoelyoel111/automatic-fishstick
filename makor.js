
    const defaultConfig = {
      page_title: 'מחליף מילים במדבקות',
      default_word: 'שם',
      primary_color: '#3b82f6',
      secondary_color: '#8b5cf6',
      text_color: '#1f2937',
      background_color: '#f0f4ff',
      button_color: '#3b82f6',
      font_family: 'Arial',
      font_size: 16
    };

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
      stickers: 'https://raw.githubusercontent.com/yoelyoel111/automatic-fishstick/main/%D7%9E%D7%93%D7%91%D7%A7%D7%95%D7%AA/',
      elements: 'https://raw.githubusercontent.com/yoelyoel111/automatic-fishstick/main/%D7%90%D7%9C%D7%9E%D7%A0%D7%98%D7%99%D7%9D/'
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
    
    // Available files in GitHub repo
    const GITHUB_FILES = {
      stickers: [
        '16.png',
        '26.png',
        '27.png',
        '28.png',
        '29.png',
        '30.png',
        '31.png',
        '32.png',
        '35.png',
        '36.png',
        '37.png',
        '40.png',
        'aaa.png'
      ],
      elements: [
        '17.png',
        '18.png',
        '19.png',
        '20.png',
        '21.png',
        '22.png',
        '23.png',
        '24.png',
        '25.png'
      ]
    };

    function cloneHistoryState() {
      return {
        stickers: JSON.parse(JSON.stringify(stickers)),
        stickerLayoutConfig: JSON.parse(JSON.stringify(stickerLayoutConfig)),
        selectedSticker,
        selectedWord,
        selectedImage,
        syncMoveEnabled,
        syncDeleteEnabled,
        autoArrangeEnabled,
        wordIdCounter,
        wordSeriesCounter,
        imageIdCounter
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
    }

    function restoreFromHistoryState(state) {
      if (!state) return;

      stickers = Array.isArray(state.stickers) ? state.stickers : [];
      stickerLayoutConfig = state.stickerLayoutConfig || stickerLayoutConfig;
      selectedSticker = (state.selectedSticker === null || Number.isFinite(state.selectedSticker)) ? state.selectedSticker : null;
      selectedWord = state.selectedWord ?? null;
      selectedImage = state.selectedImage ?? null;
      syncMoveEnabled = !!state.syncMoveEnabled;
      syncDeleteEnabled = !!state.syncDeleteEnabled;
      autoArrangeEnabled = !!state.autoArrangeEnabled;
      wordIdCounter = Number.isFinite(state.wordIdCounter) ? state.wordIdCounter : wordIdCounter;
      wordSeriesCounter = Number.isFinite(state.wordSeriesCounter) ? state.wordSeriesCounter : wordSeriesCounter;
      imageIdCounter = Number.isFinite(state.imageIdCounter) ? state.imageIdCounter : imageIdCounter;

      applyStickerLayoutConfigToUI();
      renderStickers();
      updateFileCount();
      updateUndoRedoButtons();
    }

    function undoLastAction() {
      if (undoStack.length === 0) {
        showStatus('אין פעולות לביטול', true);
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
        showStatus('אין פעולות לביצוע מחדש', true);
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
        showStatus('טוען מדבקות מהמאגר...');
        const files = (GITHUB_FILES && Array.isArray(GITHUB_FILES.stickers)) ? GITHUB_FILES.stickers : [];
        if (files.length === 0) {
          showStatus('לא נמצאו קבצים לטעינה מהמאגר', true);
          return;
        }

        const cfg = getStickerLayoutConfigFromUI();
        const desiredCount = Number.isFinite(cfg.uploadLimit) && cfg.uploadLimit > 0 ? cfg.uploadLimit : 0;
        const filesToLoad = desiredCount > 0 ? desiredCount : files.length;

        pushHistory();

        for (let i = 0; i < filesToLoad; i++) {
          const fileName = files[i % files.length];
          const url = `${GITHUB_REPO.stickers}${encodeURIComponent(fileName)}`;
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
            fileName,
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

        if (autoArrangeEnabled) {
          reflowStickersPositionsOnly();
        } else {
          reflowStickers();
        }
        renderStickers();
        updateFileCount();
        showStatus(`${filesToLoad} מדבקות נטענו מהמאגר בהצלחה!`);
      } catch (error) {
        console.error('GitHub Stickers Error:', error);
        showStatus('שגיאה בטעינת מדבקות מהמאגר', true);
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

    // ========== Names Lottery Functions ==========
    
    let namesData = [];
    let namesNoteTemplate = null;
    let namesNotes = [];
    let namesLotteryResults = [];
    
    function handleNamesNoteImageUpload(e) {
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
          namesNoteTemplate = {
            dataUrl: event.target.result,
            fileName: file.name,
            width: img.width,
            height: img.height
          };

          document.getElementById('namesNoteFileName').textContent = `✓ ${file.name}`;
          showStatus(`תמונת הפתק "${file.name}" הועלתה בהצלחה!`);
          
          // שלב 2: הפעלת כל הכפתורים אחרי בחירת מדבקה
          enableNamesButtons();
        };
        img.src = event.target.result;
      };
      reader.readAsDataURL(file);
      e.target.value = '';
    }

    function generateNamesNotes() {
      if (namesData.length === 0) {
        showStatus('יש להעלות קובץ אקסל עם שמות קודם', true);
        return;
      }

      if (!namesNoteTemplate) {
        showStatus('יש להעלות תמונת פתק קודם', true);
        return;
      }

      const notesPerRow = Math.max(1, Math.min(10, Number(document.getElementById('namesNotesPerRow')?.value) || 4));
      const spacing = Math.max(0, Number(document.getElementById('namesNotesSpacing')?.value) || 5);
      const textColor = document.getElementById('namesNotesColor')?.value || '#000000';
      const fontSize = Math.max(8, Number(document.getElementById('namesNotesFontSize')?.value) || 18);
      const uploadType = document.getElementById('namesUploadType').value;

      namesNotes = [];
      
      // Calculate note dimensions
      const pageWidth = 210 * MM_TO_PX; // A4 width in pixels
      const pageHeight = 297 * MM_TO_PX; // A4 height in pixels
      const margin = 20;
      const availableWidth = pageWidth - (2 * margin);
      const noteWidth = (availableWidth - ((notesPerRow - 1) * spacing)) / notesPerRow;
      const noteHeight = (noteWidth * namesNoteTemplate.height) / namesNoteTemplate.width;

      let currentPage = 0;
      let currentRow = 0;
      let currentCol = 0;

      namesData.forEach((item, index) => {
        const x = margin + (currentCol * (noteWidth + spacing));
        const y = margin + (currentRow * (noteHeight + spacing));

        // Check if we need a new page
        if (y + noteHeight > pageHeight - margin) {
          currentPage++;
          currentRow = 0;
          currentCol = 0;
        }

        let displayText = item.name;
        let ticketText = '';
        
        if (uploadType === 'groups' && item.group) {
          displayText += `\n${item.group}`;
        } else if (uploadType === 'tickets' && item.totalTickets > 1) {
          ticketText = `${item.ticketNumber}/${item.totalTickets}`;
        } else if (uploadType === 'groupsTickets' && item.group) {
          displayText += `\n${item.group}`;
          if (item.totalTickets > 1) {
            ticketText = `${item.ticketNumber}/${item.totalTickets}`;
          }
        }

        const note = {
          id: `note_${index}`,
          x: margin + (currentCol * (noteWidth + spacing)),
          y: margin + (currentRow * (noteHeight + spacing)),
          width: noteWidth,
          height: noteHeight,
          page: currentPage,
          backgroundImage: namesNoteTemplate.dataUrl,
          text: displayText,
          ticketText: ticketText, // מספר הכרטיס בנפרד
          textColor: textColor,
          fontSize: fontSize,
          textX: noteWidth / 2, // Center horizontally
          textY: noteHeight / 2  // 50% לתצוגה באתר
        };

        namesNotes.push(note);

        currentCol++;
        if (currentCol >= notesPerRow) {
          currentCol = 0;
          currentRow++;
        }
      });

      renderNamesNotes();
      showStatus(`נוצרו ${namesNotes.length} פתקים!`);
    }

    function renderNamesNotes() {
      const emptyState = document.getElementById('namesNotesEmptyState');
      const previewSection = document.getElementById('namesNotesPreviewSection');
      const preview = document.getElementById('namesNotesPreview');

      if (namesNotes.length === 0) {
        preview.innerHTML = '';
        emptyState.classList.remove('hidden');
        previewSection.classList.add('hidden');
        return;
      }

      emptyState.classList.add('hidden');
      previewSection.classList.remove('hidden');
      preview.innerHTML = '';

      // Group notes by page
      const pageGroups = {};
      namesNotes.forEach(note => {
        if (!pageGroups[note.page]) pageGroups[note.page] = [];
        pageGroups[note.page].push(note);
      });

      // Render each page
      Object.keys(pageGroups).forEach(pageNum => {
        const pageDiv = document.createElement('div');
        pageDiv.className = 'print-page';
        pageDiv.style.position = 'relative';

        pageGroups[pageNum].forEach(note => {
          const noteDiv = document.createElement('div');
          noteDiv.className = 'sticker-container';
          noteDiv.style.position = 'absolute';
          noteDiv.style.left = `${note.x}px`;
          noteDiv.style.top = `${note.y}px`;
          noteDiv.style.width = `${note.width}px`;
          noteDiv.style.height = `${note.height}px`;
          noteDiv.style.backgroundImage = `url(${note.backgroundImage})`;
          noteDiv.style.backgroundSize = 'cover';
          noteDiv.style.backgroundPosition = 'center';

          const textDiv = document.createElement('div');
          textDiv.className = 'text-word';
          textDiv.style.position = 'absolute';
          textDiv.style.left = '50%';
          textDiv.style.top = '50%'; // 50% לתצוגה באתר
          textDiv.style.fontSize = `${note.fontSize}px`;
          textDiv.style.color = note.textColor;
          textDiv.style.fontWeight = 'bold';
          textDiv.style.textAlign = 'center';
          textDiv.style.transform = 'translate(-50%, -50%)';
          textDiv.style.whiteSpace = 'pre-line';
          textDiv.style.lineHeight = '1.1';
          textDiv.style.width = `${note.width * 0.85}px`;
          textDiv.textContent = note.text;

          // הוסף מספר כרטיס בפינה השמאלית אם קיים
          if (note.ticketText) {
            const ticketDiv = document.createElement('div');
            ticketDiv.style.position = 'absolute';
            ticketDiv.style.left = '5px';
            ticketDiv.style.top = '5px';
            ticketDiv.style.fontSize = `${Math.max(8, note.fontSize * 0.4)}px`;
            ticketDiv.style.color = note.textColor;
            ticketDiv.style.fontWeight = 'bold';
            ticketDiv.style.backgroundColor = 'rgba(255, 255, 255, 0.8)';
            ticketDiv.style.padding = '2px 4px';
            ticketDiv.style.borderRadius = '3px';
            ticketDiv.style.lineHeight = '1';
            ticketDiv.textContent = note.ticketText;
            noteDiv.appendChild(ticketDiv);
          }

          // Add drag functionality
          textDiv.draggable = true;
          textDiv.addEventListener('dragstart', (e) => {
            draggedElement = textDiv;
            const rect = textDiv.getBoundingClientRect();
            offsetX = e.clientX - rect.left;
            offsetY = e.clientY - rect.top;
          });

          textDiv.addEventListener('dragend', () => {
            draggedElement = null;
          });

          noteDiv.addEventListener('dragover', (e) => {
            e.preventDefault();
          });

          noteDiv.addEventListener('drop', (e) => {
            e.preventDefault();
            if (draggedElement && draggedElement.parentElement === noteDiv) {
              const rect = noteDiv.getBoundingClientRect();
              const newX = e.clientX - rect.left - offsetX;
              const newY = e.clientY - rect.top - offsetY;
              
              draggedElement.style.left = `${newX}px`;
              draggedElement.style.top = `${newY}px`;
              
              // Update note data
              note.textX = newX;
              note.textY = newY;
              
              // Sync movement to all notes
              syncNamesNotesTextPosition(newX, newY);
            }
          });

          noteDiv.appendChild(textDiv);
          pageDiv.appendChild(noteDiv);
        });

        preview.appendChild(pageDiv);
      });
    }

    function syncNamesNotesTextPosition(newX, newY) {
      // Convert pixel position to percentage for consistent positioning
      const firstNote = namesNotes[0];
      if (!firstNote) return;
      
      const percentX = (newX / firstNote.width) * 100;
      const percentY = (newY / firstNote.height) * 100;
      
      namesNotes.forEach(note => {
        note.textX = newX;
        note.textY = newY;
      });
      
      // Update all text elements in the DOM with percentage positioning
      document.querySelectorAll('#namesNotesPreview .text-word').forEach(textEl => {
        textEl.style.left = `${percentX}%`;
        textEl.style.top = `${percentY}%`;
      });
    }

    function centerNamesNotes() {
      if (namesNotes.length === 0) {
        showStatus('אין פתקים לריכוז', true);
        return;
      }

      namesNotes.forEach(note => {
        note.textX = note.width / 2;
        note.textY = note.height / 2; // 50% לתצוגה באתר
      });

      renderNamesNotes();
      showStatus('הטקסט ברוכז בכל הפתקים!');
    }

    async function downloadNamesNotesAsPDF() {
      if (namesNotes.length === 0) {
        showStatus('אין פתקים להורדה', true);
        return;
      }

      const { jsPDF } = window.jspdf;
      showStatus('מכין PDF...');

      try {
        const pdf = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4', compress: true });
        
        // Group notes by page
        const pageGroups = {};
        namesNotes.forEach(note => {
          if (!pageGroups[note.page]) pageGroups[note.page] = [];
          pageGroups[note.page].push(note);
        });

        let isFirstPage = true;
        
        for (const pageNum of Object.keys(pageGroups)) {
          if (!isFirstPage) {
            pdf.addPage();
          }
          isFirstPage = false;

          const pageDiv = document.createElement('div');
          pageDiv.style.position = 'fixed';
          pageDiv.style.left = '-99999px';
          pageDiv.style.top = '0';
          pageDiv.style.width = '210mm';
          pageDiv.style.height = '297mm';
          pageDiv.style.background = '#ffffff';
          pageDiv.className = 'print-page';

          pageGroups[pageNum].forEach(note => {
            const noteDiv = document.createElement('div');
            noteDiv.style.position = 'absolute';
            noteDiv.style.left = `${note.x}px`;
            noteDiv.style.top = `${note.y}px`;
            noteDiv.style.width = `${note.width}px`;
            noteDiv.style.height = `${note.height}px`;
            noteDiv.style.backgroundImage = `url(${note.backgroundImage})`;
            noteDiv.style.backgroundSize = 'cover';
            noteDiv.style.backgroundPosition = 'center';

            const textDiv = document.createElement('div');
            textDiv.style.position = 'absolute';
            textDiv.style.left = '50%';
            textDiv.style.top = '40%'; // עוד יותר למעלה
            textDiv.style.fontSize = `${note.fontSize}px`;
            textDiv.style.color = note.textColor;
            textDiv.style.fontWeight = 'bold';
            textDiv.style.textAlign = 'center';
            textDiv.style.transform = 'translate(-50%, -50%)';
            textDiv.style.whiteSpace = 'pre-line';
            textDiv.style.lineHeight = '1.1';
            textDiv.style.width = `${note.width * 0.85}px`; // קצת יותר צר
            textDiv.textContent = note.text;

            // הוסף מספר כרטיס בפינה השמאלית אם קיים
            if (note.ticketText) {
              const ticketDiv = document.createElement('div');
              ticketDiv.style.position = 'absolute';
              ticketDiv.style.left = '5px';
              ticketDiv.style.top = '5px';
              ticketDiv.style.fontSize = `${Math.max(8, note.fontSize * 0.4)}px`;
              ticketDiv.style.color = note.textColor;
              ticketDiv.style.fontWeight = 'bold';
              ticketDiv.style.backgroundColor = 'rgba(255, 255, 255, 0.8)';
              ticketDiv.style.padding = '2px 4px';
              ticketDiv.style.borderRadius = '3px';
              ticketDiv.style.lineHeight = '1';
              ticketDiv.textContent = note.ticketText;
              noteDiv.appendChild(ticketDiv);
            }

            noteDiv.appendChild(textDiv);
            pageDiv.appendChild(noteDiv);
          });

          document.body.appendChild(pageDiv);

          const canvas = await captureElementToCanvas(pageDiv, { scale: EXPORT_QUALITY.pdfScale });
          document.body.removeChild(pageDiv);

          const pdfWidth = pdf.internal.pageSize.getWidth();
          const pdfHeight = pdf.internal.pageSize.getHeight();
          const imgData = canvas.toDataURL('image/jpeg', EXPORT_QUALITY.jpegQuality);
          
          pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight, undefined, EXPORT_QUALITY.pdfCompression);
        }

        pdf.save('פתקי_שמות.pdf');
        showStatus('PDF הורד בהצלחה!');
      } catch (error) {
        console.error('PDF generation error:', error);
        showStatus('שגיאה ביצירת PDF', true);
      }
    }

    function printNamesNotes() {
      if (namesNotes.length === 0) {
        showStatus('אין פתקים להדפסה', true);
        return;
      }

      // Create a print-friendly version with actual img elements instead of background images
      const printContainer = document.createElement('div');
      printContainer.style.width = '100%';
      printContainer.style.background = 'white';
      printContainer.style.padding = '0';
      printContainer.style.margin = '0';

      // Group notes by page
      const pageGroups = {};
      namesNotes.forEach(note => {
        if (!pageGroups[note.page]) pageGroups[note.page] = [];
        pageGroups[note.page].push(note);
      });

      // Create each page
      Object.keys(pageGroups).forEach(pageNum => {
        const pageDiv = document.createElement('div');
        pageDiv.className = 'print-page';
        pageDiv.style.width = '210mm';
        pageDiv.style.height = '297mm';
        pageDiv.style.position = 'relative';
        pageDiv.style.background = 'white';
        pageDiv.style.pageBreakAfter = 'always';
        pageDiv.style.margin = '0';
        pageDiv.style.padding = '0';

        pageGroups[pageNum].forEach(note => {
          const noteDiv = document.createElement('div');
          noteDiv.style.position = 'absolute';
          noteDiv.style.left = `${note.x}px`;
          noteDiv.style.top = `${note.y}px`;
          noteDiv.style.width = `${note.width}px`;
          noteDiv.style.height = `${note.height}px`;

          // Use actual img element instead of background-image for better print support
          const bgImg = document.createElement('img');
          bgImg.src = note.backgroundImage;
          bgImg.style.width = '100%';
          bgImg.style.height = '100%';
          bgImg.style.objectFit = 'cover';
          bgImg.style.position = 'absolute';
          bgImg.style.top = '0';
          bgImg.style.left = '0';
          bgImg.style.zIndex = '1';

          const textDiv = document.createElement('div');
          textDiv.style.position = 'absolute';
          textDiv.style.left = '50%';
          textDiv.style.top = '50%';
          textDiv.style.fontSize = `${note.fontSize}px`;
          textDiv.style.color = note.textColor;
          textDiv.style.fontWeight = 'bold';
          textDiv.style.textAlign = 'center';
          textDiv.style.transform = 'translate(-50%, -50%)';
          textDiv.style.whiteSpace = 'pre-line';
          textDiv.style.lineHeight = '1.1';
          textDiv.style.width = `${note.width * 0.85}px`;
          textDiv.style.zIndex = '2';
          textDiv.style.pointerEvents = 'none';
          textDiv.textContent = note.text;

          // הוסף מספר כרטיס בפינה השמאלית אם קיים
          if (note.ticketText) {
            const ticketDiv = document.createElement('div');
            ticketDiv.style.position = 'absolute';
            ticketDiv.style.left = '5px';
            ticketDiv.style.top = '5px';
            ticketDiv.style.fontSize = `${Math.max(8, note.fontSize * 0.4)}px`;
            ticketDiv.style.color = note.textColor;
            ticketDiv.style.fontWeight = 'bold';
            ticketDiv.style.backgroundColor = 'rgba(255, 255, 255, 0.8)';
            ticketDiv.style.padding = '2px 4px';
            ticketDiv.style.borderRadius = '3px';
            ticketDiv.style.lineHeight = '1';
            ticketDiv.style.zIndex = '3';
            ticketDiv.style.pointerEvents = 'none';
            ticketDiv.textContent = note.ticketText;
            noteDiv.appendChild(ticketDiv);
          }

          noteDiv.appendChild(bgImg);
          noteDiv.appendChild(textDiv);
          pageDiv.appendChild(noteDiv);
        });

        printContainer.appendChild(pageDiv);
      });

      // Replace body content temporarily
      const originalContent = document.body.innerHTML;
      document.body.innerHTML = '';
      document.body.appendChild(printContainer);
      
      // Print
      window.print();
      
      // Restore original content
      document.body.innerHTML = originalContent;
      
      // Re-initialize after restoring content
      location.reload();
    }

    async function downloadNamesLotteryAsPDF() {
      if (namesLotteryResults.length === 0) {
        showStatus('אין תוצאות הגרלה להורדה', true);
        return;
      }

      showStatus('מכין PDF...');

      try {
        // Create a temporary container for PDF generation
        const container = document.createElement('div');
        container.style.position = 'fixed';
        container.style.left = '-99999px';
        container.style.top = '0';
        container.style.width = '210mm';
        container.style.background = '#ffffff';
        container.style.padding = '15mm';
        container.style.boxSizing = 'border-box';
        container.style.fontFamily = 'Arial, sans-serif';
        container.style.fontSize = '14px';
        container.style.lineHeight = '1.6';
        container.style.direction = 'rtl';
        container.style.textAlign = 'right';

        // Add title
        const title = document.createElement('div');
        title.textContent = 'תוצאות הגרלת השמות';
        title.style.fontSize = '24px';
        title.style.fontWeight = 'bold';
        title.style.textAlign = 'center';
        title.style.marginBottom = '10mm';
        title.style.borderBottom = '2px solid #333';
        title.style.paddingBottom = '5mm';
        container.appendChild(title);

        // Add date
        const dateDiv = document.createElement('div');
        const date = new Date().toLocaleDateString('he-IL');
        dateDiv.textContent = `תאריך: ${date}`;
        dateDiv.style.textAlign = 'center';
        dateDiv.style.marginBottom = '10mm';
        dateDiv.style.fontSize = '12px';
        dateDiv.style.color = '#666';
        container.appendChild(dateDiv);

        // Check if we have group-based results
        const hasGroups = namesLotteryResults.some(item => item.groupName);
        
        if (hasGroups) {
          // Render by groups
          const groups = {};
          namesLotteryResults.forEach(item => {
            const groupName = item.groupName || 'ללא קבוצה';
            if (!groups[groupName]) {
              groups[groupName] = [];
            }
            groups[groupName].push(item);
          });
          
          Object.keys(groups).sort().forEach((groupName, groupIndex) => {
            const groupItems = groups[groupName];
            
            // Group header
            const groupHeader = document.createElement('div');
            groupHeader.textContent = `${groupName} (${groupItems.length} זוכים)`;
            groupHeader.style.fontSize = '18px';
            groupHeader.style.fontWeight = 'bold';
            groupHeader.style.backgroundColor = '#f3f4f6';
            groupHeader.style.padding = '8px 12px';
            groupHeader.style.marginTop = groupIndex > 0 ? '15px' : '0';
            groupHeader.style.marginBottom = '10px';
            groupHeader.style.borderRadius = '5px';
            groupHeader.style.textAlign = 'center';
            container.appendChild(groupHeader);
            
            // Group items
            const groupList = document.createElement('div');
            groupList.style.marginBottom = '10px';
            
            groupItems.forEach((item) => {
              const itemDiv = document.createElement('div');
              itemDiv.textContent = `${item.number}. ${item.name}`;
              itemDiv.style.padding = '4px 8px';
              itemDiv.style.borderBottom = '1px solid #eee';
              itemDiv.style.fontSize = '14px';
              groupList.appendChild(itemDiv);
            });
            
            container.appendChild(groupList);
          });
          
        } else {
          // Render as single list
          const resultsList = document.createElement('div');
          
          namesLotteryResults.forEach((item) => {
            const itemDiv = document.createElement('div');
            let text = `${item.number}. ${item.name}`;
            if (item.group) {
              text += ` (${item.group})`;
            }
            itemDiv.textContent = text;
            itemDiv.style.padding = '4px 8px';
            itemDiv.style.borderBottom = '1px solid #eee';
            itemDiv.style.fontSize = '14px';
            resultsList.appendChild(itemDiv);
          });
          
          container.appendChild(resultsList);
        }
        
        // Add footer
        const footer = document.createElement('div');
        footer.textContent = `סה"כ ${namesLotteryResults.length} תוצאות`;
        footer.style.textAlign = 'center';
        footer.style.marginTop = '15mm';
        footer.style.fontSize = '12px';
        footer.style.fontStyle = 'italic';
        footer.style.color = '#666';
        footer.style.borderTop = '1px solid #ccc';
        footer.style.paddingTop = '5mm';
        container.appendChild(footer);

        document.body.appendChild(container);

        // Use html2canvas to capture the content
        const canvas = await html2canvas(container, {
          scale: 2,
          useCORS: true,
          allowTaint: true,
          backgroundColor: '#ffffff',
          width: container.offsetWidth,
          height: container.offsetHeight
        });

        document.body.removeChild(container);

        // Create PDF with multiple pages if needed
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('p', 'mm', 'a4');
        
        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfHeight = pdf.internal.pageSize.getHeight();
        
        const canvasWidth = canvas.width;
        const canvasHeight = canvas.height;
        
        // Calculate how many pages we need
        const ratio = canvasWidth / pdfWidth;
        const scaledHeight = canvasHeight / ratio;
        
        if (scaledHeight <= pdfHeight) {
          // Single page
          const imgData = canvas.toDataURL('image/png');
          pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, scaledHeight);
        } else {
          // Multiple pages
          const pageHeight = pdfHeight * ratio;
          let yPosition = 0;
          let pageNumber = 0;
          
          while (yPosition < canvasHeight) {
            if (pageNumber > 0) {
              pdf.addPage();
            }
            
            // Create canvas for this page
            const pageCanvas = document.createElement('canvas');
            pageCanvas.width = canvasWidth;
            pageCanvas.height = Math.min(pageHeight, canvasHeight - yPosition);
            
            const pageCtx = pageCanvas.getContext('2d');
            pageCtx.drawImage(
              canvas,
              0, yPosition, canvasWidth, pageCanvas.height,
              0, 0, canvasWidth, pageCanvas.height
            );
            
            const pageImgData = pageCanvas.toDataURL('image/png');
            pdf.addImage(pageImgData, 'PNG', 0, 0, pdfWidth, pageCanvas.height / ratio);
            
            yPosition += pageHeight;
            pageNumber++;
          }
        }
        
        pdf.save('הגרלת_שמות.pdf');
        showStatus('PDF הורד בהצלחה!');
        
      } catch (error) {
        console.error('PDF generation error:', error);
        showStatus('שגיאה ביצירת PDF', true);
      }
    }

    function printNamesLottery() {
      if (namesLotteryResults.length === 0) {
        showStatus('אין תוצאות הגרלה להדפסה', true);
        return;
      }

      const originalContent = document.body.innerHTML;
      const resultsSection = document.getElementById('namesLotteryResultsSection');
      
      if (!resultsSection) {
        showStatus('שגיאה בהכנת ההדפסה', true);
        return;
      }

      document.body.innerHTML = resultsSection.outerHTML;
      window.print();
      document.body.innerHTML = originalContent;
      
      // Re-initialize after restoring content
      location.reload();
    }
    
    function switchNamesSubTab(subTab) {
      const notesTab = document.getElementById('namesSubTabNotes');
      const lotteryTab = document.getElementById('namesSubTabLottery');
      const notesContent = document.getElementById('namesNotesContent');
      const lotteryContent = document.getElementById('namesLotterySubContent');
      
      // הסרת active מכל הטאבים הפנימיים
      notesTab.classList.remove('active');
      lotteryTab.classList.remove('active');
      notesContent.classList.add('hidden');
      lotteryContent.classList.add('hidden');
      
      if (subTab === 'notes') {
        notesTab.classList.add('active');
        notesContent.classList.remove('hidden');
      } else {
        lotteryTab.classList.add('active');
        lotteryContent.classList.remove('hidden');
      }
    }
    
    function updateNamesUploadInstructions() {
      const type = document.getElementById('namesUploadType').value;
      const instructions = document.getElementById('namesUploadInstructions');
      
      if (type === 'simple') {
        instructions.innerHTML = '<p class="text-sm text-blue-700"><strong>שמות בלבד:</strong> הזן שמות בעמודה A מתא A2 ומטה (שורה 1 לכותרת)</p>';
      } else if (type === 'groups') {
        instructions.innerHTML = `
          <p class="text-sm text-blue-700 mb-2"><strong>שמות לפי קבוצות:</strong></p>
          <ul class="text-sm text-blue-700 list-disc list-inside space-y-1">
            <li>שמות הכיתות/קבוצות בשורה 1: C1, D1, E1 וכן הלאה</li>
            <li>שמות התלמידים מתחת לכל כיתה: C2, C3, C4... D2, D3, D4... וכן הלאה</li>
            <li>עמודות A ו-B יכולות להישאר ריקות</li>
            <li>בפתקים יופיע: שם התלמיד + שם הכיתה</li>
          </ul>
        `;
      } else if (type === 'tickets') {
        instructions.innerHTML = `
          <p class="text-sm text-blue-700 mb-2"><strong>שמות וכמות כרטיסים:</strong></p>
          <ul class="text-sm text-blue-700 list-disc list-inside space-y-1">
            <li>שמות התלמידים בעמודה C: C2, C3, C4 וכן הלאה</li>
            <li>כמות הכרטיסים בעמודה D: D2, D3, D4 וכן הלאה</li>
            <li>עמודות A ו-B יכולות להישאר ריקות</li>
            <li>לכל תלמיד יופיעו כרטיסים לפי הכמות שהוזנה</li>
          </ul>
        `;
      } else if (type === 'groupsTickets') {
        instructions.innerHTML = `
          <p class="text-sm text-blue-700 mb-2"><strong>שמות קבוצות וכרטיסים:</strong></p>
          <ul class="text-sm text-blue-700 list-disc list-inside space-y-1">
            <li>שמות הקבוצות בשורה 1: C1, E1, G1 וכן הלאה (כל קבוצה תופסת 2 עמודות)</li>
            <li>שמות התלמידים מתחת לכל קבוצה: C2, C3, C4... E2, E3, E4... וכן הלאה</li>
            <li>כמות הכרטיסים לכל תלמיד: D2, D3, D4... F2, F3, F4... וכן הלאה</li>
            <li>עמודות A ו-B יכולות להישאר ריקות</li>
            <li>לכל תלמיד יופיעו כרטיסים לפי הכמות שהוזנה + שם הקבוצה</li>
          </ul>
        `;
      }
    }
    
    function downloadNamesTemplate() {
      const uploadType = document.getElementById('namesUploadType').value;
      
      let csvContent = '';
      if (uploadType === 'simple') {
        csvContent = '\uFEFFשם\nישראל ישראלי\nמשה כהן\nדוד לוי\nשרה אברהם\nרחל יעקב';
      } else if (uploadType === 'groups') {
        csvContent = '\uFEFFבדיקה,,כיתה א,כיתה ב,כיתה ג\n,,ישראל ישראלי,משה כהן,דוד לוי\n,,שרה אברהם,רחל יעקב,לאה יצחק\n,,יוסף שמעון,אהרון בנימין,נפתלי גד\n,,אברהם יצחק,זבולון דן,יהודה ראובן\n,,שמעון לוי,,';
      } else if (uploadType === 'tickets') {
        csvContent = '\uFEFFבדיקה,,שם,כמות כרטיסים\n,,ישראל ישראלי,5\n,,משה כהן,3\n,,דוד לוי,7\n,,שרה אברהם,2\n,,רחל יעקב,4';
      } else if (uploadType === 'groupsTickets') {
        csvContent = '\uFEFFבדיקה,,קבוצה א,כמות,קבוצה ב,כמות,קבוצה ג,כמות\n,,ישראל ישראלי,5,משה כהן,3,דוד לוי,7\n,,שרה אברהם,2,רחל יעקב,4,לאה יצחק,6\n,,יוסף שמעון,3,אהרון בנימין,5,נפתלי גד,2\n,,אברהם יצחק,4,זבולון דן,3,יהודה ראובן,5\n,,שמעון לוי,2,,,';
      }
      
      // Create blob with UTF-8 BOM for Hebrew support
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8-sig;' });
      const link = document.createElement('a');
      const fileName = uploadType === 'simple' ? 'תבנית-שמות-בלבד.csv' : 
                       uploadType === 'groups' ? 'תבנית-שמות-לפי-קבוצות.csv' : 
                       uploadType === 'tickets' ? 'תבנית-שמות-וכרטיסים.csv' :
                       'תבנית-שמות-קבוצות-וכרטיסים.csv';
      link.download = fileName;
      link.href = URL.createObjectURL(blob);
      link.click();
      URL.revokeObjectURL(link.href);
      
      showStatus('תבנית הורדה בהצלחה!');
    }
    
    function handleNamesExcelUpload(e) {
      const file = e.target.files[0];
      if (!file) return;
      
      // Try to detect file encoding and read accordingly
      const reader = new FileReader();
      reader.onload = function(event) {
        try {
          const text = event.target.result;
          handleCSVText(text, file.name);
        } catch (err) {
          console.error('CSV parse error:', err);
          showStatus('שגיאה בקריאת הקובץ - נסה קובץ CSV', true);
        }
      };
      
      // Read file with UTF-8 encoding to support Hebrew
      reader.readAsText(file, 'UTF-8');
      
      // If UTF-8 fails, try with different encoding
      reader.onerror = function() {
        const reader2 = new FileReader();
        reader2.onload = function(event) {
          try {
            let text = event.target.result;
            // Process the same way as above
            handleCSVText(text, file.name);
          } catch (err) {
            showStatus('שגיאה בקריאת הקובץ - בדוק שהקובץ תקין', true);
          }
        };
        reader2.readAsText(file, 'windows-1255'); // Hebrew encoding
      };
    }
    
    function handleCSVText(text, fileName) {
      // Remove UTF-8 BOM if present
      if (text.charCodeAt(0) === 0xFEFF) {
        text = text.slice(1);
      }
      
      // Clean up text - remove extra whitespace and normalize line endings
      text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
      
      const lines = text.split('\n').filter(line => line.trim());
      const uploadType = document.getElementById('namesUploadType').value;
      
      namesData = [];
      
      if (uploadType === 'simple') {
        // אפשרות 1: שמות בלבד - קריאה מעמודה A בלבד
        for (let i = 1; i < lines.length; i++) {
          const cells = lines[i].split(',').map(cell => cell.trim().replace(/^"|"$/g, ''));
          const studentName = cells[0] ? cells[0].trim() : '';
          
          // הוסף שם אם הוא קיים ולא ריק (מתעלם מתאים ריקים)
          if (studentName && studentName.length > 0) {
            namesData.push({ name: studentName });
          }
        }
      } else if (uploadType === 'tickets') {
        // אפשרות 3: שמות וכמות כרטיסים - C=שם, D=כמות
        for (let i = 1; i < lines.length; i++) {
          const cells = lines[i].split(',').map(cell => cell.trim().replace(/^"|"$/g, ''));
          const studentName = cells[2] ? cells[2].trim() : ''; // עמודה C
          const ticketCount = cells[3] ? parseInt(cells[3].trim()) : 1; // עמודה D
          
          if (studentName && studentName.length > 0) {
            // צור כמה רשומות לפי כמות הכרטיסים
            for (let t = 0; t < Math.max(1, ticketCount); t++) {
              namesData.push({ 
                name: studentName,
                ticketNumber: t + 1,
                totalTickets: Math.max(1, ticketCount)
              });
            }
          }
        }
      } else if (uploadType === 'groupsTickets') {
        // אפשרות 4: שמות קבוצות וכרטיסים - C,D=קבוצה1, E,F=קבוצה2, וכן הלאה
        const headers = lines[0] ? lines[0].split(',').map(cell => cell.trim().replace(/^"|"$/g, '')) : [];
        
        // התחל מעמודה C (אינדקס 2) ועבור על כל זוג עמודות (שם קבוצה, כמות)
        for (let col = 2; col < headers.length; col += 2) {
          const groupName = headers[col] ? headers[col].trim() : '';
          
          // דלג על קבוצות ריקות
          if (!groupName || groupName.length === 0) continue;
          
          // קרא את כל השמות וכמויות תחת הקבוצה הזו (משורה 2 ואילך)
          for (let row = 1; row < lines.length; row++) {
            const cells = lines[row].split(',').map(cell => cell.trim().replace(/^"|"$/g, ''));
            const studentName = cells[col] ? cells[col].trim() : ''; // עמודת השם
            const ticketCount = cells[col + 1] ? parseInt(cells[col + 1].trim()) : 1; // עמודת הכמות
            
            // הוסף תלמיד אם השם קיים ולא ריק
            if (studentName && studentName.length > 0) {
              // צור כמה רשומות לפי כמות הכרטיסים
              for (let t = 0; t < Math.max(1, ticketCount); t++) {
                namesData.push({ 
                  name: studentName, 
                  group: groupName,
                  ticketNumber: t + 1,
                  totalTickets: Math.max(1, ticketCount)
                });
              }
            }
          }
        }
      } else {
        // אפשרות 2: שמות לפי קבוצות
        // C1, D1, E1... = שמות הקבוצות
        // C2, C3, C4... = השמות תחת כל קבוצה
        const headers = lines[0] ? lines[0].split(',').map(cell => cell.trim().replace(/^"|"$/g, '')) : [];
        
        // התחל מעמודה C (אינדקס 2) ועבור על כל העמודות שיש בהן שמות קבוצות
        for (let col = 2; col < headers.length; col++) {
          const groupName = headers[col] ? headers[col].trim() : '';
          
          // דלג על קבוצות ריקות
          if (!groupName || groupName.length === 0) continue;
          
          // קרא את כל השמות תחת הקבוצה הזו (משורה 2 ואילך)
          for (let row = 1; row < lines.length; row++) {
            const cells = lines[row].split(',').map(cell => cell.trim().replace(/^"|"$/g, ''));
            const studentName = cells[col] ? cells[col].trim() : '';
            
            // הוסף תלמיד אם השם קיים ולא ריק (מתעלם מתאים ריקים)
            if (studentName && studentName.length > 0) {
              namesData.push({ 
                name: studentName, 
                group: groupName 
              });
            }
          }
        }
      }
      
      if (namesData.length === 0) {
        showStatus('לא נמצאו שמות בקובץ', true);
        return;
      }
      
      // Update UI
      document.getElementById('namesFileInfo').textContent = `✓ ${fileName} (${namesData.length} שמות)`;
      
      // Show preview
      const previewSection = document.getElementById('namesPreviewSection');
      const previewList = document.getElementById('namesPreviewList');
      previewSection.classList.remove('hidden');
      
      let previewHtml = '';
      if (uploadType === 'groups') {
        const groups = {};
        namesData.forEach(item => {
          if (!groups[item.group]) groups[item.group] = [];
          groups[item.group].push(item.name);
        });
        for (const [group, names] of Object.entries(groups)) {
          previewHtml += `<div class="mb-4 p-3 bg-gray-50 rounded-lg border border-gray-200">`;
          previewHtml += `<div class="font-bold text-indigo-800 text-base mb-2">${group} (${names.length} תלמידים)</div>`;
          previewHtml += `<div class="text-gray-700 text-sm leading-relaxed">${names.join(', ')}</div>`;
          previewHtml += `</div>`;
        }
      } else if (uploadType === 'tickets') {
        const students = {};
        namesData.forEach(item => {
          if (!students[item.name]) {
            students[item.name] = item.totalTickets;
          }
        });
        previewHtml += `<div class="mb-4 p-3 bg-gray-50 rounded-lg border border-gray-200">`;
        previewHtml += `<div class="font-bold text-indigo-800 text-base mb-2">שמות וכמות כרטיסים</div>`;
        previewHtml += `<div class="text-gray-700 text-sm leading-relaxed">`;
        for (const [name, count] of Object.entries(students)) {
          previewHtml += `${name} (${count} כרטיסים), `;
        }
        previewHtml = previewHtml.slice(0, -2); // הסר את הפסיק האחרון
        previewHtml += `</div></div>`;
      } else if (uploadType === 'groupsTickets') {
        const groups = {};
        namesData.forEach(item => {
          if (!groups[item.group]) groups[item.group] = {};
          if (!groups[item.group][item.name]) {
            groups[item.group][item.name] = item.totalTickets;
          }
        });
        for (const [group, students] of Object.entries(groups)) {
          previewHtml += `<div class="mb-4 p-3 bg-gray-50 rounded-lg border border-gray-200">`;
          previewHtml += `<div class="font-bold text-indigo-800 text-base mb-2">${group}</div>`;
          previewHtml += `<div class="text-gray-700 text-sm leading-relaxed">`;
          for (const [name, count] of Object.entries(students)) {
            previewHtml += `${name} (${count} כרטיסים), `;
          }
          previewHtml = previewHtml.slice(0, -2); // הסר את הפסיק האחרון
          previewHtml += `</div></div>`;
        }
      } else {
        previewHtml = `<div class="text-gray-700">${namesData.map(item => item.name).join(', ')}</div>`;
      }
      previewList.innerHTML = previewHtml;
      
      showStatus(`נטענו ${namesData.length} שמות בהצלחה!`);
      
      // שלב 1: הפעלת רק כפתור בחר מדבקה והטאבים
      enableNamesImageButtonOnly();
    }
    
    function generateNamesLottery(mode = 'all') {
      if (namesData.length === 0) {
        showStatus('יש להעלות קובץ אקסל עם שמות קודם', true);
        return;
      }
      
      const winnersCount = Math.max(1, Number(document.getElementById('namesLotteryWinners')?.value) || 5);
      const startNum = Math.max(1, Number(document.getElementById('namesLotteryStartNum')?.value) || 1);
      const uploadType = document.getElementById('namesUploadType').value;
      
      if (mode === 'all') {
        // הגרלה אחת מכל השמות
        generateAllNamesLottery(winnersCount, startNum);
      } else if (mode === 'byGroup' && (uploadType === 'groups' || uploadType === 'groupsTickets')) {
        // הגרלה נפרדת לכל קבוצה
        generateGroupNamesLottery(winnersCount, startNum);
      } else if (mode === 'byGroup' && (uploadType === 'simple' || uploadType === 'tickets')) {
        showStatus('הגרלה לפי קבוצה זמינה רק כשמעלים קובץ עם קבוצות', true);
        return;
      }
      
      renderNamesLotteryResults();
    }
    
    function generateAllNamesLottery(winnersCount, startNum) {
      // Create array of numbers
      const numbers = [];
      for (let i = 0; i < winnersCount; i++) {
        numbers.push(startNum + i);
      }
      
      // Shuffle names array
      const shuffledNames = [...namesData];
      for (let i = shuffledNames.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [shuffledNames[i], shuffledNames[j]] = [shuffledNames[j], shuffledNames[i]];
      }
      
      // Take only the winners and assign numbers
      namesLotteryResults = shuffledNames.slice(0, winnersCount).map((item, index) => ({
        ...item,
        number: numbers[index],
        isWinner: true
      }));
      
      // Sort by number
      namesLotteryResults.sort((a, b) => a.number - b.number);
      
      showStatus(`הוגרלו ${winnersCount} זוכים מתוך ${namesData.length} שמות!`);
    }
    
    function generateGroupNamesLottery(winnersCount, startNum) {
      // Group names by group
      const groups = {};
      namesData.forEach(item => {
        const group = item.group || 'ללא קבוצה';
        if (!groups[group]) {
          groups[group] = [];
        }
        groups[group].push(item);
      });
      
      namesLotteryResults = [];
      let currentNumber = startNum;
      
      // Generate lottery for each group - winnersCount from EACH group
      Object.keys(groups).sort().forEach(groupName => {
        const groupMembers = groups[groupName];
        const groupWinners = Math.min(winnersCount, groupMembers.length); // X תוצאות מכל קבוצה
        
        // Shuffle group members
        const shuffledGroup = [...groupMembers];
        for (let i = shuffledGroup.length - 1; i > 0; i--) {
          const j = Math.floor(Math.random() * (i + 1));
          [shuffledGroup[i], shuffledGroup[j]] = [shuffledGroup[j], shuffledGroup[i]];
        }
        
        // Take winners from this group
        const groupResults = shuffledGroup.slice(0, groupWinners).map((item, index) => ({
          ...item,
          number: currentNumber + index,
          isWinner: true,
          groupName: groupName
        }));
        
        namesLotteryResults.push(...groupResults);
        currentNumber += groupWinners;
      });
      
      // Sort by group name, then by number
      namesLotteryResults.sort((a, b) => {
        if (a.groupName !== b.groupName) {
          return a.groupName.localeCompare(b.groupName);
        }
        return a.number - b.number;
      });
      
      const totalWinners = namesLotteryResults.length;
      const groupCount = Object.keys(groups).length;
      showStatus(`הוגרלו ${winnersCount} תוצאות מכל קבוצה - סה"כ ${totalWinners} זוכים מ-${groupCount} קבוצות!`);
    }
    
    function renderNamesLotteryResults() {
      const emptyState = document.getElementById('namesLotteryEmptyState');
      const resultsSection = document.getElementById('namesLotteryResultsSection');
      const results = document.getElementById('namesLotteryResults');
      
      if (namesLotteryResults.length === 0) {
        results.innerHTML = '';
        emptyState.classList.remove('hidden');
        resultsSection.classList.add('hidden');
        return;
      }
      
      emptyState.classList.add('hidden');
      resultsSection.classList.remove('hidden');
      results.innerHTML = '';
      
      const perRow = Math.max(1, Math.min(10, Number(document.getElementById('namesLotteryPerRow')?.value) || 5));
      const fontSize = Math.max(12, Number(document.getElementById('namesLotteryFontSize')?.value) || 18);
      const color = document.getElementById('namesLotteryColor')?.value || '#000000';
      const uploadType = document.getElementById('namesUploadType').value;
      
      // Check if we have group-based results
      const hasGroups = namesLotteryResults.some(item => item.groupName);
      
      if (hasGroups) {
        // Render by groups
        renderNamesLotteryByGroups(perRow, fontSize, color);
      } else {
        // Render as single list
        renderNamesLotterySingle(perRow, fontSize, color, uploadType);
      }
    }
    
    function renderNamesLotterySingle(perRow, fontSize, color, uploadType) {
      const results = document.getElementById('namesLotteryResults');
      
      const table = document.createElement('table');
      table.style.width = '100%';
      table.style.borderCollapse = 'separate';
      table.style.borderSpacing = '0';
      table.style.tableLayout = 'fixed';
      table.style.borderRadius = '14px';
      table.style.border = '2px solid rgba(0,0,0,0.08)';
      table.style.overflow = 'hidden';
      
      const tbody = document.createElement('tbody');
      const rowCount = Math.ceil(namesLotteryResults.length / perRow);
      
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
          
          if (idx < namesLotteryResults.length) {
            const item = namesLotteryResults[idx];
            const el = document.createElement('div');
            
            let displayText = `${item.number}. ${item.name}`;
            if (uploadType === 'groups' && item.group) {
              displayText += ` (${item.group})`;
            }
            
            el.textContent = displayText;
            el.style.fontSize = `${fontSize}px`;
            el.style.fontWeight = '600';
            el.style.color = color;
            el.style.lineHeight = '1.3';
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
    
    function renderNamesLotteryByGroups(perRow, fontSize, color) {
      const results = document.getElementById('namesLotteryResults');
      
      // Group results by groupName
      const groups = {};
      namesLotteryResults.forEach(item => {
        const groupName = item.groupName || 'ללא קבוצה';
        if (!groups[groupName]) {
          groups[groupName] = [];
        }
        groups[groupName].push(item);
      });
      
      // Render each group
      Object.keys(groups).sort().forEach((groupName, groupIndex) => {
        const groupItems = groups[groupName];
        
        // Group header
        const groupHeader = document.createElement('div');
        groupHeader.style.marginTop = groupIndex > 0 ? '24px' : '0';
        groupHeader.style.marginBottom = '12px';
        groupHeader.style.padding = '12px 16px';
        groupHeader.style.backgroundColor = '#f3f4f6';
        groupHeader.style.borderRadius = '8px';
        groupHeader.style.fontWeight = 'bold';
        groupHeader.style.fontSize = `${Math.max(fontSize, 16)}px`;
        groupHeader.style.color = '#374151';
        groupHeader.style.textAlign = 'center';
        groupHeader.textContent = `${groupName} (${groupItems.length} זוכים)`;
        results.appendChild(groupHeader);
        
        // Group table
        const table = document.createElement('table');
        table.style.width = '100%';
        table.style.borderCollapse = 'separate';
        table.style.borderSpacing = '0';
        table.style.tableLayout = 'fixed';
        table.style.borderRadius = '8px';
        table.style.border = '1px solid rgba(0,0,0,0.08)';
        table.style.overflow = 'hidden';
        table.style.marginBottom = '8px';
        
        const tbody = document.createElement('tbody');
        const rowCount = Math.ceil(groupItems.length / perRow);
        
        for (let r = 0; r < rowCount; r++) {
          const tr = document.createElement('tr');
          tr.style.background = (r % 2 === 0) ? '#ffffff' : '#f9fafb';
          
          for (let c = 0; c < perRow; c++) {
            const idx = r * perRow + c;
            const td = document.createElement('td');
            td.style.padding = '10px 8px';
            td.style.textAlign = 'center';
            td.style.borderBottom = '1px solid rgba(0,0,0,0.04)';
            td.style.borderLeft = c === 0 ? 'none' : '1px solid rgba(0,0,0,0.04)';
            
            if (idx < groupItems.length) {
              const item = groupItems[idx];
              const el = document.createElement('div');
              el.textContent = `${item.number}. ${item.name}`;
              el.style.fontSize = `${fontSize}px`;
              el.style.fontWeight = '600';
              el.style.color = color;
              el.style.lineHeight = '1.3';
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
      });
    }
    
    // ========== End Names Lottery Functions ==========

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
      const emptyState = document.getElementById('emptyState');
      const printPreviewSection = document.getElementById('printPreviewSection');
      
      if (!preview || !emptyState || !printPreviewSection) {
        console.error('renderStickers: Required DOM elements not found');
        return;
      }

      if (stickers.length === 0) {
        previewInner.innerHTML = '';
        emptyState.classList.remove('hidden');
        printPreviewSection.classList.add('hidden');
        return;
      }
      
      emptyState.classList.add('hidden');
      printPreviewSection.classList.remove('hidden');

      try {
        const fragment = document.createDocumentFragment();

        const maxPageIndex = stickers.reduce((max, s) => Math.max(max, Number.isFinite(s.page) ? s.page : 0), 0);
        const pageCount = Math.max(1, maxPageIndex + 1);

        const pages = [];
        for (let p = 0; p < pageCount; p++) {
          const pageEl = document.createElement('div');
          pageEl.className = 'print-page';
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
          img.style.width = '100%';
          img.style.height = '100%';
          img.style.objectFit = 'contain';
          img.style.pointerEvents = 'none';
          
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
          stickerDiv.appendChild(controlsDiv);
          
          const resizeHandle = document.createElement('div');
          resizeHandle.className = 'sticker-resize-handle no-print';
          resizeHandle.addEventListener('mousedown', (e) => {
            e.stopPropagation();
            startStickerResize(e, index);
          });
          stickerDiv.appendChild(resizeHandle);
          
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
        return;
      }

      inner.style.transform = '';
      preview.style.width = '';
      preview.style.height = '';

      const naturalWidth = inner.scrollWidth;
      const naturalHeight = inner.scrollHeight;
      const maxScale = 1.35;
      const scale = naturalWidth > 0 ? Math.min(maxScale, section.clientWidth / naturalWidth) : 1;

      inner.style.transformOrigin = 'top right';
      const translateX = naturalWidth * (scale - 1);
      inner.style.transform = `translateX(${translateX}px) scale(${scale})`;
      preview.style.width = `${Math.ceil(naturalWidth * scale)}px`;
      preview.style.height = `${Math.ceil(naturalHeight * scale)}px`;
    }

    async function captureElementToCanvas(element, options = {}) {
      const noPrintElements = element.querySelectorAll('.no-print');
      noPrintElements.forEach(el => el.style.display = 'none');

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
      el.textContent = word.text;
      el.style.left = `${word.x}px`;
      el.style.top = `${word.y}px`;
      el.style.color = word.color || '#000000';
      el.style.fontSize = `${word.fontSize || 12}px`;
      el.style.fontFamily = word.fontFamily || 'Arial';
      el.style.fontWeight = word.fontWeight || '700';

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

      el.addEventListener('mousedown', (e) => {
        if (e.target === deleteBtn || e.target === resizeHandle) return;
        e.stopPropagation();
        startWordDrag(e, stickerIndex, word.id);
      });

      el.addEventListener('click', (e) => {
        e.stopPropagation();
        selectWord(stickerIndex, word.id);
      });

      return el;
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
      const pageWidth = 210 * MM_TO_PX;
      const pageHeight = 297 * MM_TO_PX;
      const contentWidth = Math.max(1, pageWidth - (cfg.edgeMargin * 2));
      const gap = Math.max(0, cfg.gap);
      const contentHeight = Math.max(1, pageHeight - (cfg.edgeMargin * 2));

      const mode = (cfg.sizeMode === 'height') ? 'height' : 'width';
      if (mode === 'height') {
        const rows = Math.max(1, cfg.stickersPerRow);
        const cellHeight = Math.max(1, (contentHeight - gap * (rows - 1)) / rows);
        const approxCols = Math.max(1, Math.floor((contentWidth + gap) / (cellHeight + gap)));
        info.textContent = `יוצא בערך: ${rows} מדבקות באורך (גובה תא ≈ ${Math.round(cellHeight)}px) | ~${approxCols} מדבקות ברוחב (תלוי ברוחב המדבקות)`;
      } else {
        const cols = Math.max(1, cfg.stickersPerRow);
        const cellWidth = Math.max(1, (contentWidth - gap * (cols - 1)) / cols);
        info.textContent = `יוצא בערך: ${cols} מדבקות ברוחב (רוחב תא ≈ ${Math.round(cellWidth)}px)`;
      }
    }

    function applyStickerLayoutAndRender() {
      getStickerLayoutConfigFromUI();
      updateStickerLayoutInfo();
      if (autoArrangeEnabled) {
        reflowStickersPositionsOnly();
      } else {
        reflowStickers();
      }
      renderStickers();
      updateFileCount();
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
      const pageWidth = 210 * MM_TO_PX;
      const pageHeight = 297 * MM_TO_PX;
      const edge = Math.max(0, cfg.edgeMargin);
      const gap = Math.max(0, cfg.gap);

      const contentWidth = Math.max(1, pageWidth - (edge * 2));
      const contentHeight = Math.max(1, pageHeight - (edge * 2));
      const maxH = contentHeight;

      const mode = (cfg.sizeMode === 'height') ? 'height' : 'width';
      const cols = (mode === 'width')
        ? Math.max(1, cfg.stickersPerRow)
        : 0;
      const cellWidth = (mode === 'width')
        ? Math.max(1, (contentWidth - gap * (cols - 1)) / cols)
        : 0;
      const rows = (mode === 'height')
        ? Math.max(1, cfg.stickersPerRow)
        : 0;
      const cellHeight = (mode === 'height')
        ? Math.max(1, (contentHeight - gap * (rows - 1)) / rows)
        : 0;
      let cellWidthForHeightMode = cellWidth;
      if (mode === 'height') {
        // Compute a safe column width for this fixed height so stickers don't overlap
        let maxW = 1;
        stickers.forEach((sticker) => {
          if (!sticker) return;
          if (!Number.isFinite(sticker.originalWidth) || !Number.isFinite(sticker.originalHeight) || sticker.originalWidth <= 0 || sticker.originalHeight <= 0) {
            const fallbackW = Number.isFinite(sticker.width) && sticker.width > 0 ? sticker.width : 1;
            const fallbackH = Number.isFinite(sticker.height) && sticker.height > 0 ? sticker.height : 1;
            sticker.originalWidth = fallbackW;
            sticker.originalHeight = fallbackH;
          }
          const ar = sticker.originalHeight > 0 ? (sticker.originalWidth / sticker.originalHeight) : 1;
          const wAtH = cellHeight * ar;
          if (Number.isFinite(wAtH) && wAtH > 0) maxW = Math.max(maxW, wAtH);
        });
        cellWidthForHeightMode = Math.max(1, Math.min(contentWidth, maxW));
      }

      const derivedCellWidth = (mode === 'height') ? cellWidthForHeightMode : cellWidth;
      const derivedCols = (mode === 'height')
        ? Math.max(1, Math.floor((contentWidth + gap) / (derivedCellWidth + gap)))
        : cols;

      let page = 0;
      let colIndex = 0;
      let x = edge;
      let y = edge;
      let rowMaxHeight = 0;

      stickers.forEach((sticker) => {
        sticker.words = sticker.words || [];
        sticker.images = sticker.images || [];

        if (!Number.isFinite(sticker.originalWidth) || !Number.isFinite(sticker.originalHeight) || sticker.originalWidth <= 0 || sticker.originalHeight <= 0) {
          const fallbackCell = (mode === 'height') ? derivedCellWidth : cellWidth;
          const fallbackW = Number.isFinite(sticker.width) && sticker.width > 0 ? sticker.width : fallbackCell;
          const fallbackH = Number.isFinite(sticker.height) && sticker.height > 0 ? sticker.height : fallbackCell;
          sticker.originalWidth = fallbackW;
          sticker.originalHeight = fallbackH;
        }

        const fallbackCell = (mode === 'height') ? derivedCellWidth : cellWidth;
        const prevW = Number.isFinite(sticker.width) && sticker.width > 0 ? sticker.width : fallbackCell;
        const prevH = Number.isFinite(sticker.height) && sticker.height > 0 ? sticker.height : (fallbackCell * (sticker.originalHeight / sticker.originalWidth));
        const aspectRatio = sticker.originalWidth / sticker.originalHeight;

        let newW;
        let newH;
        if (mode === 'height') {
          newH = cellHeight;
          newW = newH * aspectRatio;
          if (newW > derivedCellWidth) {
            newW = derivedCellWidth;
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

        if (colIndex >= derivedCols) {
          colIndex = 0;
          x = edge;
          y = y + (mode === 'height' ? cellHeight : rowMaxHeight) + gap;
          rowMaxHeight = 0;
        }

        if (y + (mode === 'height' ? cellHeight : newH) > pageHeight - edge) {
          page += 1;
          colIndex = 0;
          x = edge;
          y = edge;
          rowMaxHeight = 0;
        }

        const scale = prevW > 0 ? (newW / prevW) : 1;

        sticker.page = page;
        sticker.x = x + Math.max(0, ((mode === 'height' ? derivedCellWidth : cellWidth) - newW) / 2);
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
        x = x + (mode === 'height' ? derivedCellWidth : cellWidth) + gap;
        colIndex += 1;
      });
    }

    function reflowStickersPositionsOnly() {
      const cfg = stickerLayoutConfig;
      const pageWidth = 210 * MM_TO_PX;
      const pageHeight = 297 * MM_TO_PX;
      const edge = Math.max(0, cfg.edgeMargin);
      const gap = Math.max(0, cfg.gap);

      const contentWidth = Math.max(1, pageWidth - (edge * 2));
      const contentHeight = Math.max(1, pageHeight - (edge * 2));

      const mode = (cfg.sizeMode === 'height') ? 'height' : 'width';
      const cols = (mode === 'width') ? Math.max(1, cfg.stickersPerRow) : 0;
      const cellWidth = (mode === 'width')
        ? Math.max(1, (contentWidth - gap * (cols - 1)) / cols)
        : 0;
      const rows = (mode === 'height') ? Math.max(1, cfg.stickersPerRow) : 0;
      const cellHeight = (mode === 'height')
        ? Math.max(1, (contentHeight - gap * (rows - 1)) / rows)
        : 0;

      let cellWidthForHeightMode = cellWidth;
      if (mode === 'height') {
        let maxW = 1;
        stickers.forEach((sticker) => {
          if (!sticker) return;
          if (!Number.isFinite(sticker.originalWidth) || !Number.isFinite(sticker.originalHeight) || sticker.originalWidth <= 0 || sticker.originalHeight <= 0) {
            const fallbackW = Number.isFinite(sticker.width) && sticker.width > 0 ? sticker.width : 1;
            const fallbackH = Number.isFinite(sticker.height) && sticker.height > 0 ? sticker.height : 1;
            sticker.originalWidth = fallbackW;
            sticker.originalHeight = fallbackH;
          }
          const ar = sticker.originalHeight > 0 ? (sticker.originalWidth / sticker.originalHeight) : 1;
          const wAtH = cellHeight * ar;
          if (Number.isFinite(wAtH) && wAtH > 0) maxW = Math.max(maxW, wAtH);
        });
        cellWidthForHeightMode = Math.max(1, Math.min(contentWidth, maxW));
      }

      const derivedCellWidth = (mode === 'height') ? cellWidthForHeightMode : cellWidth;
      const derivedCols = (mode === 'height')
        ? Math.max(1, Math.floor((contentWidth + gap) / (derivedCellWidth + gap)))
        : cols;

      let page = 0;
      let colIndex = 0;
      let x = edge;
      let y = edge;
      let rowMaxHeight = 0;

      stickers.forEach((sticker) => {
        if (!sticker) return;

        sticker.words = sticker.words || [];
        sticker.images = sticker.images || [];

        if (!Number.isFinite(sticker.originalWidth) || !Number.isFinite(sticker.originalHeight) || sticker.originalWidth <= 0 || sticker.originalHeight <= 0) {
          const fallbackW = Number.isFinite(sticker.width) && sticker.width > 0 ? sticker.width : 1;
          const fallbackH = Number.isFinite(sticker.height) && sticker.height > 0 ? sticker.height : 1;
          sticker.originalWidth = fallbackW;
          sticker.originalHeight = fallbackH;
        }

        const aspectRatio = sticker.originalHeight > 0 ? (sticker.originalWidth / sticker.originalHeight) : 1;

        const hasValidSize = Number.isFinite(sticker.width) && sticker.width > 1 && Number.isFinite(sticker.height) && sticker.height > 1;
        if (!hasValidSize) {
          let baseW;
          let baseH;
          if (mode === 'height') {
            baseH = cellHeight;
            baseW = baseH * aspectRatio;
            if (baseW > derivedCellWidth) {
              baseW = derivedCellWidth;
              baseH = baseW / aspectRatio;
            }
          } else {
            baseW = cellWidth;
            baseH = baseW / aspectRatio;
          }
          sticker.width = baseW;
          sticker.height = baseH;
        }

        const w = Number.isFinite(sticker.width) && sticker.width > 0 ? sticker.width : 1;
        const h = Number.isFinite(sticker.height) && sticker.height > 0 ? sticker.height : 1;

        if (colIndex >= derivedCols) {
          colIndex = 0;
          x = edge;
          y = y + rowMaxHeight + gap;
          rowMaxHeight = 0;
        }

        if (y + h > pageHeight - edge) {
          page += 1;
          colIndex = 0;
          x = edge;
          y = edge;
          rowMaxHeight = 0;
        }

        sticker.page = page;
        sticker.x = x + Math.max(0, ((mode === 'height' ? derivedCellWidth : cellWidth) - w) / 2);
        sticker.y = y;

        rowMaxHeight = Math.max(rowMaxHeight, h);
        x = x + (mode === 'height' ? derivedCellWidth : cellWidth) + gap;
        colIndex += 1;
      });
    }

    function compactStickers() {
      const edgePadding = Math.max(0, stickerLayoutConfig.edgeMargin);
      const gapPadding = Math.max(0, stickerLayoutConfig.gap);
      const pageWidth = 210 * MM_TO_PX;
      const pageHeight = 297 * MM_TO_PX;

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
      const pageWidth = 210 * MM_TO_PX;
      const pageHeight = 297 * MM_TO_PX;
      
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
          fontWeight: word.fontWeight
        };
        duplicatedSticker.words.push(newWord);
      });
      
      // Copy images with new IDs but keep them synchronized
      if (originalSticker.images) {
        originalSticker.images.forEach(image => {
          const newImage = {
            id: `image-${++imageIdCounter}`,
            dataUrl: image.dataUrl,
            x: image.x,
            y: image.y,
            width: image.width,
            height: image.height,
            originalWidth: image.originalWidth,
            originalHeight: image.originalHeight
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
    }

    function stopWordResize() {
      resizingWord = null;
      document.removeEventListener('mousemove', resizeWordFont);
      document.removeEventListener('mouseup', stopWordResize);
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

      focusWordInput();
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
      
      const color = document.getElementById('textColorPicker').value;
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

      const fixedX = Math.max(0, (sticker.width - textW) / 2);
      const fixedY = Math.max(0, (sticker.height - textH) / 2);

      const seriesId = `series-${++wordSeriesCounter}`;
      
      const word = {
        id: `word-${++wordIdCounter}`,
        seriesId,
        text: text,
        x: fixedX,
        y: fixedY,
        color: color,
        fontSize: fontSize,
        fontFamily: fontFamily,
        fontWeight: fontWeight
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
      
      const color = document.getElementById('textColorPicker').value;
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
        const fixedX = Math.max(0, (sticker.width - textW) / 2);
        const fixedY = Math.max(0, (sticker.height - textH) / 2);
        const word = {
          id: `word-${++wordIdCounter}`,
          seriesId,
          text: text,
          x: fixedX,
          y: fixedY,
          color: color,
          fontSize: fontSize,
          fontFamily: fontFamily,
          fontWeight: fontWeight
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
      
      const color = document.getElementById('textColorPicker').value;
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
          lastWord.color = color;
          lastWord.fontSize = fontSize;
          lastWord.fontFamily = fontFamily;
          lastWord.fontWeight = fontWeight;
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
      
      stickers.forEach(sticker => {
        sticker.images = sticker.images || [];
        const x = Math.max(0, (sticker.width - calculatedWidth) / 2);
        const y = Math.max(0, (sticker.height - targetHeight) / 2);
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
      
      newX = Math.max(0, Math.min(newX, parentRect.width - wordEl.offsetWidth));
      newY = Math.max(0, Math.min(newY, parentRect.height - wordEl.offsetHeight));
      
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

      newX = Math.max(0, Math.min(newX, parentRect.width - stickerEl.offsetWidth));
      newY = Math.max(0, Math.min(newY, parentRect.height - stickerEl.offsetHeight));

      sticker.x = newX;
      sticker.y = newY;
      stickerEl.style.left = `${newX}px`;
      stickerEl.style.top = `${newY}px`;
    }

    function startImageDrag(e, stickerIndex, imageId) {
      pushHistory();
      const imageEl = e.currentTarget;
      const rect = imageEl.getBoundingClientRect();
      const parentEl = imageEl.parentElement;
      const parentRect = parentEl.getBoundingClientRect();
      const scale = parentEl && parentEl.offsetWidth ? (parentRect.width / parentEl.offsetWidth) : 1;
      
      offsetX = (e.clientX - rect.left) / (scale || 1);
      offsetY = (e.clientY - rect.top) / (scale || 1);
      
      const image = stickers[stickerIndex].images.find(i => i.id === imageId);
      if (image) {
        initialDragPosition = { x: image.x, y: image.y };
      }
      
      draggedElement = { type: 'image', stickerIndex, imageId, imageEl, parentRect, parentEl, scale };
      
      document.addEventListener('mousemove', dragImage);
      document.addEventListener('mouseup', stopDrag);
    }

    function dragImage(e) {
      if (!draggedElement || draggedElement.type !== 'image') return;
      
      const { imageEl, parentRect, parentEl, scale, stickerIndex, imageId } = draggedElement;
      
      const effectiveScale = scale || 1;
      let newX = ((e.clientX - parentRect.left) / effectiveScale) - offsetX;
      let newY = ((e.clientY - parentRect.top) / effectiveScale) - offsetY;
      
      const parentW = parentEl ? parentEl.offsetWidth : (parentRect.width / effectiveScale);
      const parentH = parentEl ? parentEl.offsetHeight : (parentRect.height / effectiveScale);
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
      
      resizingImage = { stickerIndex, imageId };
      
      document.addEventListener('mousemove', resizeImage);
      document.addEventListener('mouseup', stopImageResize);
    }

    function resizeImage(e) {
      if (!resizingImage) return;
      
      const { stickerIndex, imageId } = resizingImage;
      const image = stickers[stickerIndex].images.find(i => i.id === imageId);
      if (!image) return;
      
      const imageEl = document.querySelector(`[data-sticker-index="${stickerIndex}"] [data-image-id="${imageId}"]`);
      if (!imageEl) return;
      
      const parentRect = imageEl.parentElement.getBoundingClientRect();
      const rect = imageEl.getBoundingClientRect();
      
      const newWidth = e.clientX - rect.left;
      const newHeight = e.clientY - rect.top;
      
      if (newWidth > 20 && newHeight > 20) {
        const aspectRatio = image.originalWidth / image.originalHeight;
        const calculatedHeight = newWidth / aspectRatio;
        
        image.width = newWidth;
        image.height = calculatedHeight;
        
        imageEl.style.width = `${newWidth}px`;
        imageEl.style.height = `${calculatedHeight}px`;
        
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
      } else if (mode === 'namesLottery') {
        if (tabNamesLottery) tabNamesLottery.classList.add('active');
        const namesLotteryContent = document.getElementById('namesLotteryContent');
        if (namesLotteryContent) namesLotteryContent.classList.remove('hidden');
        printPreviewSection.classList.add('hidden');
        document.getElementById('pageTitle').textContent = 'הגרלת שמות';
        
        // Hide word mode buttons
        if (downloadPdfBtn) downloadPdfBtn.classList.add('hidden');
        if (downloadImageBtn) downloadImageBtn.classList.add('hidden');
        if (printBtn) printBtn.classList.add('hidden');
        if (saveProjectBtn) saveProjectBtn.classList.add('hidden');
        if (loadProjectLabel) loadProjectLabel.classList.add('hidden');
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
        pageEl.className = 'print-page';
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
      
      const projectData = {
        version: '1.1',
        stickers: stickers,
        layout: stickerLayoutConfig,
        savedAt: new Date().toISOString()
      };
      
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

    function loadProject(file) {
      if (!file) return;
      
      currentProjectFileName = file.name;
      
      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const projectData = JSON.parse(e.target.result);
          
          if (!projectData.stickers || !Array.isArray(projectData.stickers)) {
            showStatus('קובץ לא תקין!', true);
            return;
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
          updateFileCount();
          showStatus(`הפרויקט "${file.name}" נטען בהצלחה! ${stickers.length} מדבקות`);
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

      const files = (GITHUB_FILES && Array.isArray(GITHUB_FILES.stickers)) ? GITHUB_FILES.stickers : [];
      if (files.length === 0) {
        showStatus('לא נמצאו מדבקות במאגר', true);
        return;
      }

      const selectHandler = typeof onSelect === 'function' ? onSelect : addSingleStickerFromGithub;


      files.forEach((fileName) => {
        const url = `${GITHUB_REPO.stickers}${encodeURIComponent(fileName)}`;

        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'group border-2 border-gray-200 rounded-xl overflow-hidden hover:border-indigo-400 hover:shadow-lg transition-all bg-white';

        const img = document.createElement('img');
        img.src = url;
        img.alt = fileName;
        img.loading = 'lazy';
        img.className = 'w-full h-32 object-contain bg-white';

        const label = document.createElement('div');
        label.className = 'px-2 py-2 text-xs font-bold text-gray-700 bg-gray-50 group-hover:bg-indigo-50 truncate';
        label.textContent = fileName;

        btn.appendChild(img);
        btn.appendChild(label);

        btn.addEventListener('click', async () => {
          await selectHandler(fileName);
          closeGithubStickersPicker();
        });

        grid.appendChild(btn);
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

      const files = (GITHUB_FILES && Array.isArray(GITHUB_FILES.elements)) ? GITHUB_FILES.elements : [];
      if (files.length === 0) {
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

      files.forEach((fileName) => {
        const url = `${GITHUB_REPO.elements}${encodeURIComponent(fileName)}`;

        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'group border-2 border-gray-200 rounded-xl overflow-hidden hover:border-indigo-400 hover:shadow-lg transition-all bg-white';

        const img = document.createElement('img');
        img.src = url;
        img.alt = fileName;
        img.loading = 'lazy';
        img.className = 'w-full h-32 object-contain bg-white';

        const label = document.createElement('div');
        label.className = 'px-2 py-2 text-xs font-bold text-gray-700 bg-gray-50 group-hover:bg-indigo-50 truncate';
        label.textContent = fileName;

        btn.appendChild(img);
        btn.appendChild(label);

        btn.addEventListener('click', async () => {
          await addElementFromGithub(fileName, addToAll);
          closeGithubElementsPicker();
        });

        grid.appendChild(btn);
      });

      modal.classList.remove('hidden');
    }

    function closeGithubElementsPicker() {
      const modal = document.getElementById('githubElementsModal');
      if (modal) modal.classList.add('hidden');
    }

    async function addElementFromGithub(fileName, addToAll) {
      try {
        showStatus('טוען אלמנט מהמאגר...');

        const url = `${GITHUB_REPO.elements}${encodeURIComponent(fileName)}`;
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

    async function addSingleStickerFromGithub(fileName) {
      try {
        showStatus('טוען מדבקה מהמאגר...');

        const url = `${GITHUB_REPO.stickers}${encodeURIComponent(fileName)}`;
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

        pushHistory();

        for (let i = 0; i < countToAdd; i++) {
          stickers.push({
            id: `sticker-github-${Date.now()}-${i}`,
            dataUrl,
            fileName,
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

        if (autoArrangeEnabled) {
          reflowStickersPositionsOnly();
        } else {
          reflowStickers();
        }
        renderStickers();
        updateFileCount();
        showStatus('המדבקה נוספה למסמך!');
      } catch (error) {
        console.error('GitHub Single Sticker Error:', error);
        showStatus('שגיאה בטעינת מדבקה מהמאגר', true);
      }
    }

    async function setNumbersTemplateFromRepo(fileName) {
      try {
        showStatus('טוען מדבקה מהמאגר...');

        const url = `${GITHUB_REPO.stickers}${encodeURIComponent(fileName)}`;
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

    async function setNamesNoteTemplateFromRepo(fileName) {
      try {
        showStatus('טוען תמונת פתק מהמאגר...');

        const url = `${GITHUB_REPO.stickers}${encodeURIComponent(fileName)}`;
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
    document.addEventListener('click', function() {
      document.getElementById('stickersMenu').classList.add('hidden');
      document.getElementById('imageAllMenu').classList.add('hidden');
      document.getElementById('imageSingleMenu').classList.add('hidden');
      const numbersMenu = document.getElementById('numbersStickerMenu');
      if (numbersMenu) numbersMenu.classList.add('hidden');
      const namesNoteMenu = document.getElementById('namesNoteImageMenu');
      if (namesNoteMenu) namesNoteMenu.classList.add('hidden');
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

    const syncElementsBtn = document.getElementById('syncElementsBtn');
    if (syncElementsBtn) syncElementsBtn.addEventListener('click', function() {
      const next = !(syncMoveEnabled && syncDeleteEnabled);
      syncMoveEnabled = next;
      syncDeleteEnabled = next;

      if (next) {
        syncElementsBtn.classList.remove('from-gray-500', 'to-slate-600', 'hover:from-gray-600', 'hover:to-slate-700');
        syncElementsBtn.classList.add('from-green-500', 'to-emerald-600', 'hover:from-green-600', 'hover:to-emerald-700');
        showStatus('סנכרון אלמנטים הופעל ✓');
      } else {
        syncElementsBtn.classList.remove('from-green-500', 'to-emerald-600', 'hover:from-green-600', 'hover:to-emerald-700');
        syncElementsBtn.classList.add('from-gray-500', 'to-slate-600', 'hover:from-gray-600', 'hover:to-slate-700');
        showStatus('סנכרון אלמנטים כבוי');
      }
    });

    const undoBtn = document.getElementById('undoBtn');
    if (undoBtn) undoBtn.addEventListener('click', undoLastAction);
    const redoBtn = document.getElementById('redoBtn');
    if (redoBtn) redoBtn.addEventListener('click', redoLastAction);

    function setLayoutMode(isAutoArrange) {
      autoArrangeEnabled = !!isAutoArrange;

      const freeEditBtn = document.getElementById('freeEditBtn');
      const autoArrangeBtn = document.getElementById('autoArrangeBtn');

      if (freeEditBtn && autoArrangeBtn) {
        if (autoArrangeEnabled) {
          freeEditBtn.classList.remove('from-green-500', 'to-emerald-600', 'hover:from-green-600', 'hover:to-emerald-700');
          freeEditBtn.classList.add('from-gray-500', 'to-slate-600', 'hover:from-gray-600', 'hover:to-slate-700');

          autoArrangeBtn.classList.remove('from-gray-500', 'to-slate-600', 'hover:from-gray-600', 'hover:to-slate-700');
          autoArrangeBtn.classList.add('from-green-500', 'to-emerald-600', 'hover:from-green-600', 'hover:to-emerald-700');
        } else {
          autoArrangeBtn.classList.remove('from-green-500', 'to-emerald-600', 'hover:from-green-600', 'hover:to-emerald-700');
          autoArrangeBtn.classList.add('from-gray-500', 'to-slate-600', 'hover:from-gray-600', 'hover:to-slate-700');

          freeEditBtn.classList.remove('from-gray-500', 'to-slate-600', 'hover:from-gray-600', 'hover:to-slate-700');
          freeEditBtn.classList.add('from-green-500', 'to-emerald-600', 'hover:from-green-600', 'hover:to-emerald-700');
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
      showStatus('סדר אוטומטי פעיל ✓');
    });

    setLayoutMode(true);

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

    document.getElementById('printBtn').addEventListener('click', function() {
      window.print();
    });

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
      const imageFiles = Array.from(files).filter(f => f.type && f.type.startsWith('image/'));

      if (imageFiles.length === 0) {
        showStatus('לא נמצאו קבצי תמונה להעלאה', true);
        e.target.value = '';
        return;
      }

      const countToAdd = desiredCount > 0 ? desiredCount : imageFiles.length;

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
              width: 1,
              height: 1,
              originalWidth: src.originalWidth,
              originalHeight: src.originalHeight,
              words: [],
              images: []
            });
          }

          if (autoArrangeEnabled) {
            reflowStickersPositionsOnly();
          } else {
            reflowStickers();
          }
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
      const imageFiles = Array.from(files).filter(f => f.type && f.type.startsWith('image/'));

      if (imageFiles.length === 0) {
        showStatus('לא נמצאו קבצי תמונה להעלאה', true);
        e.target.value = '';
        return;
      }

      const countToAdd = desiredCount > 0 ? desiredCount : imageFiles.length;

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
              width: 1,
              height: 1,
              originalWidth: src.originalWidth,
              originalHeight: src.originalHeight,
              words: [],
              images: []
            });
          }

          if (autoArrangeEnabled) {
            reflowStickersPositionsOnly();
          } else {
            reflowStickers();
          }
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
    document.getElementById('namesUploadType').addEventListener('change', function() {
      updateNamesUploadInstructions();
    });

    // Names Excel Input
    document.getElementById('namesExcelInput').addEventListener('change', function(e) {
      handleNamesExcelUpload(e);
    });

    // Download Names Template
    document.getElementById('downloadNamesTemplateBtn').addEventListener('click', function() {
      downloadNamesTemplate();
    });

    // Generate Names Lottery - All
    document.getElementById('generateNamesLotteryAllBtn').addEventListener('click', function() {
      generateNamesLottery('all');
    });

    // Generate Names Lottery - By Group
    document.getElementById('generateNamesLotteryByGroupBtn').addEventListener('click', function() {
      generateNamesLottery('byGroup');
    });

    // Names Lottery PDF Download
    document.getElementById('downloadNamesLotteryPdfBtn').addEventListener('click', function() {
      downloadNamesLotteryAsPDF();
    });

    // Names Lottery Print
    document.getElementById('printNamesLotteryBtn').addEventListener('click', function() {
      printNamesLottery();
    });

    // Names Notes PDF Download
    document.getElementById('downloadNamesNotesPdfBtn').addEventListener('click', function() {
      downloadNamesNotesAsPDF();
    });

    // Names Notes Print
    document.getElementById('printNamesNotesBtn').addEventListener('click', function() {
      printNamesNotes();
    });

    // Generate Names Notes
    document.getElementById('generateNamesNotesBtn').addEventListener('click', function() {
      generateNamesNotes();
    });

    // Center Names Notes
    document.getElementById('centerNamesNotesBtn').addEventListener('click', function() {
      centerNamesNotes();
    });

    // Names Note Image Input
    document.getElementById('namesNoteImageInput').addEventListener('change', function(e) {
      handleNamesNoteImageUpload(e);
    });

    document.getElementById('generateLotteryBtn').addEventListener('click', function() {
      generateLottery();
    });

    document.getElementById('downloadLotteryPdfBtn').addEventListener('click', function() {
      downloadLotteryAsPDF();
    });

    document.getElementById('centerNumbersBtn').addEventListener('click', function() {
      centerNumbers();
    });

    document.getElementById('generateNumbersBtn').addEventListener('click', function() {
      generateNumberedStickers();
    });

    document.getElementById('singleStickerInput').addEventListener('change', function(e) {
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
          singleStickerTemplate = {
            dataUrl: event.target.result,
            fileName: file.name,
            width: img.width,
            height: img.height
          };

          document.getElementById('stickerFileName').textContent = `✓ ${file.name}`;
          showStatus(`המדבקה "${file.name}" הועלתה בהצלחה!`);
          checkNumberStickerAndUpdateButtons();
        };
        img.src = event.target.result;
      };
      reader.readAsDataURL(file);
      e.target.value = '';
    });

    document.getElementById('downloadNumbersPdfBtn').addEventListener('click', function() {
      downloadNumbersAsPDF();
    });

    document.getElementById('downloadNumbersImageBtn').addEventListener('click', function() {
      downloadNumbersAsImage();
    });

    const uploadLimitInput = document.getElementById('uploadLimitInput');
    if (uploadLimitInput) uploadLimitInput.addEventListener('input', () => {
      getStickerLayoutConfigFromUI();
      updateStickerLayoutInfo();
    });
    const stickersPerRowInput = document.getElementById('stickersPerRowInput');
    if (stickersPerRowInput) stickersPerRowInput.addEventListener('input', applyStickerLayoutAndRender);
    const stickerSizeModeSelect = document.getElementById('stickerSizeModeSelect');
    if (stickerSizeModeSelect) stickerSizeModeSelect.addEventListener('change', applyStickerLayoutAndRender);
    const edgeMarginInput = document.getElementById('edgeMarginInput');
    if (edgeMarginInput) edgeMarginInput.addEventListener('input', applyStickerLayoutAndRender);
    const gapInput = document.getElementById('gapInput');
    if (gapInput) gapInput.addEventListener('input', applyStickerLayoutAndRender);

    // Initialize
    applyStickerLayoutConfigToUI();
    updateUndoRedoButtons();
    const textColorSwatch = document.getElementById('textColorSwatch');
    const textColorPicker = document.getElementById('textColorPicker');
    const textColorPalette = document.getElementById('textColorPalette');
    const textColorPaletteGrid = document.getElementById('textColorPaletteGrid');

    if (textColorSwatch && textColorPicker) {
      const setSwatch = (color) => {
        textColorSwatch.style.background = color;
        textColorPicker.value = color;
      };

      setSwatch(textColorPicker.value || '#000000');

      const paletteColors = [
        '#000000', '#444444', '#777777', '#BBBBBB', '#FFFFFF',
        '#FF0000', '#FF7A00', '#FFD400', '#00C853', '#00B0FF',
        '#2962FF', '#651FFF', '#D500F9', '#FF4081', '#8D6E63',
        '#1B5E20', '#004D40', '#01579B', '#1A237E', '#311B92',
        '#B71C1C', '#E65100', '#F57F17', '#33691E', '#263238'
      ];

      if (textColorPaletteGrid) {
        textColorPaletteGrid.innerHTML = '';
        paletteColors.forEach((c) => {
          const btn = document.createElement('button');
          btn.type = 'button';
          btn.className = 'color-palette-color';
          btn.style.background = c;
          btn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            setSwatch(c);
            if (textColorPalette) textColorPalette.classList.add('hidden');
          });
          textColorPaletteGrid.appendChild(btn);
        });
      }

      textColorSwatch.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (!textColorPalette) return;
        textColorPalette.classList.toggle('hidden');
      });

      document.addEventListener('click', () => {
        if (textColorPalette) textColorPalette.classList.add('hidden');
      });
    }
    renderStickers();
    updateFileCount();

    window.addEventListener('resize', () => {
      applyPrintPreviewScale();
    });

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
    updateNamesUploadInstructions();

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
      
      // הפעלת כפתורי ההגרלה בטאב הגרלת שמות (מיד אחרי העלאת אקסל)
      const lotteryButtons = ['generateNamesLotteryAllBtn', 'generateNamesLotteryByGroupBtn', 'downloadNamesLotteryPdfBtn', 'printNamesLotteryBtn'];
      lotteryButtons.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) {
          btn.classList.remove('btn-disabled');
        }
      });
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

  
