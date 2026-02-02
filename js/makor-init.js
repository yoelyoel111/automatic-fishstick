    // Names Lottery PDF Download
    const downloadNamesLotteryPdfBtn = document.getElementById('downloadNamesLotteryPdfBtn');
    if (downloadNamesLotteryPdfBtn) {
      downloadNamesLotteryPdfBtn.addEventListener('click', function() {
        downloadNamesLotteryAsPDF();
      });
    }

    // Tools Tabs Event Listeners
    const toolsTabStickers = document.getElementById('toolsTabStickers');
    if (toolsTabStickers) {
      console.log('Attaching click listener to toolsTabStickers');
      toolsTabStickers.addEventListener('click', () => {
        console.log('Stickers tab clicked');
        setActiveToolsSection('Stickers');
      });
    } else {
      console.error('toolsTabStickers not found!');
    }
    
    const toolsTabText = document.getElementById('toolsTabText');
    if (toolsTabText) {
      console.log('Attaching click listener to toolsTabText');
      toolsTabText.addEventListener('click', () => {
        console.log('Text tab clicked');
        setActiveToolsSection('Text');
      });
    } else {
      console.error('toolsTabText not found!');
    }
    
    const toolsTabImage = document.getElementById('toolsTabImage');
    if (toolsTabImage) {
      console.log('Attaching click listener to toolsTabImage');
      toolsTabImage.addEventListener('click', () => {
        console.log('Image tab clicked');
        setActiveToolsSection('Image');
      });
    } else {
      console.error('toolsTabImage not found!');
    }

    // Initialize with Stickers tab active
    console.log('Initializing with Stickers tab');
    setActiveToolsSection('Stickers');

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
          showStatus(t('onlyImageFiles'), true);
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
            showStatus(t('stickerUploadedSuccess', { name: file.name }));
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
    if (stickersPerRowInput) stickersPerRowInput.addEventListener('input', () => {
      // עדכון רק של המידע בממשק, לא של המדבקות הקיימות
      getStickerLayoutConfigFromUI();
      updateStickerLayoutInfo();
    });
    const stickerSizeModeSelect = document.getElementById('stickerSizeModeSelect');
    if (stickerSizeModeSelect) stickerSizeModeSelect.addEventListener('change', () => {
      // עדכון רק של המידע בממשק, לא של המדבקות הקיימות
      getStickerLayoutConfigFromUI();
      updateStickerLayoutInfo();
    });
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

    // Google Drive Integration
    const googleSignInBtn = document.getElementById('googleSignInBtn');
    if (googleSignInBtn) {
      googleSignInBtn.addEventListener('click', function() {
        GoogleDriveManager.signIn();
      });
    }

    const googleSignOutBtn = document.getElementById('googleSignOutBtn');
    if (googleSignOutBtn) {
      googleSignOutBtn.addEventListener('click', function() {
        GoogleDriveManager.signOut();
      });
    }

    // Save to Drive button (in header)
    const saveToDriveBtn = document.getElementById('saveToDriveBtn');
    if (saveToDriveBtn) {
      saveToDriveBtn.addEventListener('click', function() {
        if (typeof getProjectData === 'function') {
          const projectData = getProjectData();
          if (projectData) {
            GoogleDriveManager.showSaveToDriverDialog(projectData);
          } else {
            showStatus(t('noProjectToSave'), true);
          }
        }
      });
    }

    // Save to Drive button (in names section)
    const saveToDriveNamesBtn = document.getElementById('saveToDriveNamesBtn');
    if (saveToDriveNamesBtn) {
      saveToDriveNamesBtn.addEventListener('click', function() {
        // For names, we need to get names project data
        if (typeof getNamesProjectData === 'function') {
          const projectData = getNamesProjectData();
          if (projectData) {
            GoogleDriveManager.showSaveToDriverDialog(projectData);
          } else {
            showStatus(t('noDataToSave'), true);
          }
        } else {
          showStatus(t('saveFuncNotAvailable'), true);
        }
      });
    }

    // Refresh projects
    const refreshProjectsBtn = document.getElementById('refreshProjectsBtn');
    if (refreshProjectsBtn) {
      refreshProjectsBtn.addEventListener('click', function() {
        GoogleDriveManager.refreshProjectsBanner();
      });
    }

    // New project from banner
    const newProjectBtn = document.getElementById('newProjectBtn');
    if (newProjectBtn) {
      newProjectBtn.addEventListener('click', function() {
        GoogleDriveManager.showNewProjectDialog();
      });
    }