    // Names Lottery PDF Download
    const downloadNamesLotteryPdfBtn = document.getElementById('downloadNamesLotteryPdfBtn');
    if (downloadNamesLotteryPdfBtn) {
      downloadNamesLotteryPdfBtn.addEventListener('click', function() {
        downloadNamesLotteryAsPDF();
      });
    }

    // Names Lottery Print
    const printNamesLotteryBtn = document.getElementById('printNamesLotteryBtn');
    if (printNamesLotteryBtn) {
      printNamesLotteryBtn.addEventListener('click', function() {
        printNamesLottery();
      });
    }

    // Names Notes PDF Download
    const downloadNamesNotesPdfBtn = document.getElementById('downloadNamesNotesPdfBtn');
    if (downloadNamesNotesPdfBtn) {
      downloadNamesNotesPdfBtn.addEventListener('click', function() {
        downloadNamesNotesAsPDF();
      });
    }

    // Names Notes Print
    const printNamesNotesBtn = document.getElementById('printNamesNotesBtn');
    if (printNamesNotesBtn) {
      printNamesNotesBtn.addEventListener('click', function() {
        printNamesNotes();
      });
    }

    // Generate Names Notes
    const generateNamesNotesBtn = document.getElementById('generateNamesNotesBtn');
    if (generateNamesNotesBtn) {
      generateNamesNotesBtn.addEventListener('click', function() {
        generateNamesNotes();
      });
    }

    // Center Names Notes
    const centerNamesNotesBtn = document.getElementById('centerNamesNotesBtn');
    if (centerNamesNotesBtn) {
      centerNamesNotesBtn.addEventListener('click', function() {
        centerNamesNotes();
      });
    }

    // Names Note Image Input
    const namesNoteImageInput = document.getElementById('namesNoteImageInput');
    if (namesNoteImageInput) {
      namesNoteImageInput.addEventListener('change', function(e) {
        handleNamesNoteImageUpload(e);
      });
    }

    const generateLotteryBtn = document.getElementById('generateLotteryBtn');
    if (generateLotteryBtn) {
      generateLotteryBtn.addEventListener('click', function() {
        generateLottery();
      });
    }

    const downloadLotteryPdfBtn = document.getElementById('downloadLotteryPdfBtn');
    if (downloadLotteryPdfBtn) {
      downloadLotteryPdfBtn.addEventListener('click', function() {
        downloadLotteryAsPDF();
      });
    }

    const centerNumbersBtn = document.getElementById('centerNumbersBtn');
    if (centerNumbersBtn) {
      centerNumbersBtn.addEventListener('click', function() {
        centerNumbers();
      });
    }

    const generateNumbersBtn = document.getElementById('generateNumbersBtn');
    if (generateNumbersBtn) {
      generateNumbersBtn.addEventListener('click', function() {
        generateNumberedStickers();
      });
    }

    const singleStickerInput = document.getElementById('singleStickerInput');
    if (singleStickerInput) {
      singleStickerInput.addEventListener('change', function(e) {
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
    }

    const downloadNumbersPdfBtn = document.getElementById('downloadNumbersPdfBtn');
    if (downloadNumbersPdfBtn) {
      downloadNumbersPdfBtn.addEventListener('click', function() {
        downloadNumbersAsPDF();
      });
    }

    const downloadNumbersImageBtn = document.getElementById('downloadNumbersImageBtn');
    if (downloadNumbersImageBtn) {
      downloadNumbersImageBtn.addEventListener('click', function() {
        downloadNumbersAsImage();
      });
    }

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
    
    // Color palette will be initialized when Text section is opened
    renderStickers();

    // Ensure preview scroller is positioned at the right edge in RTL layouts
    requestAnimationFrame(() => {
      try {
        const preview = document.getElementById('printPreview');
        if (preview) {
          preview.scrollLeft = Math.max(0, preview.scrollWidth - preview.clientWidth);
        }
      } catch (_) {}
    });
    updateFileCount();

    window.addEventListener('resize', () => {
      applyPrintPreviewScale();
    });