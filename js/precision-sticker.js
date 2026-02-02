/**
 * Precision Sticker Tool - כלי הכנת מדבקה מדויקת
 * גרסה פשוטה שתעבוד בוודאות
 */

// קבועי כיוונון לדיוק הקואורדינטות
const X_OFFSET_CORRECTION = 13; // תיקון לציר X (פיקסלים) - הקטנתי מ-15 ל-13
const Y_OFFSET_CORRECTION = 0;  // תיקון לציר Y (פיקסלים)

console.log('Precision Sticker script loaded - VERSION 20250115-01');
console.log('X_OFFSET_CORRECTION:', X_OFFSET_CORRECTION);
console.log('Y_OFFSET_CORRECTION:', Y_OFFSET_CORRECTION);

// ========== כיול מדפסת ==========
// מידות מדבקת הכיול (במ"מ) - מלבן ארוך מגלה סטיות טוב יותר
const CALIBRATION_TARGET_WIDTH = 200;  // 20 ס"מ
const CALIBRATION_TARGET_HEIGHT = 50;  // 5 ס"מ

// טעינת נתוני כיול מ-localStorage
function loadCalibrationData() {
  try {
    const saved = localStorage.getItem('printerCalibration');
    if (saved) {
      const data = JSON.parse(saved);
      console.log('Loaded calibration data:', data);
      return data;
    }
  } catch (e) {
    console.error('Error loading calibration data:', e);
  }
  return { scaleX: 1, scaleY: 1, calibrated: false };
}

// שמירת נתוני כיול ל-localStorage
function saveCalibrationData(data) {
  try {
    localStorage.setItem('printerCalibration', JSON.stringify(data));
    console.log('Saved calibration data:', data);
  } catch (e) {
    console.error('Error saving calibration data:', e);
  }
}

// קבלת יחסי הכיול הנוכחיים
function getCalibrationScale() {
  const data = loadCalibrationData();
  return { scaleX: data.scaleX || 1, scaleY: data.scaleY || 1 };
}

// פתיחת מודל כיול מדפסת
function openCalibrationModal() {
  console.log('Opening calibration modal');
  const modal = document.getElementById('printerCalibrationModal');
  if (modal) {
    // איפוס השדות
    const measuredWidth = document.getElementById('calibrationMeasuredWidth');
    const measuredHeight = document.getElementById('calibrationMeasuredHeight');
    if (measuredWidth) measuredWidth.value = '';
    if (measuredHeight) measuredHeight.value = '';
    
    // הצגת המידות הנוכחיות
    updateCalibrationStatus();
    
    modal.classList.remove('hidden');
  } else {
    console.error('Calibration modal not found');
  }
}

// סגירת מודל כיול
function closeCalibrationModal() {
  const modal = document.getElementById('printerCalibrationModal');
  if (modal) {
    modal.classList.add('hidden');
  }
}

// עדכון סטטוס הכיול בממשק
function updateCalibrationStatus() {
  const data = loadCalibrationData();
  const statusEl = document.getElementById('calibrationStatus');
  
  if (statusEl) {
    if (data.calibrated) {
      statusEl.innerHTML = `
        <span class="text-green-600 font-bold">✓ ${t('calibrated')}</span>
        <span class="text-sm text-gray-600 mr-2">
          (ratio: ${data.scaleX.toFixed(3)})
        </span>
      `;
    } else {
      statusEl.innerHTML = `<span class="text-orange-500 font-bold">⚠ ${t('notCalibrated')}</span>`;
    }
  }
}

// הדפסת מדבקת כיול - מוסיף מדבקה אמיתית לעורך ומדפיס
function printCalibrationSticker() {
  console.log('Creating calibration sticker');
  
  // סגירת מודל הכיול זמנית
  closeCalibrationModal();
  
  // יצירת מדבקת כיול כתמונה (canvas)
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  
  // גודל הקנבס בפיקסלים (ברזולוציה גבוהה)
  const scale = 3;
  const widthPx = CALIBRATION_TARGET_WIDTH * 3.78 * scale;
  const heightPx = CALIBRATION_TARGET_HEIGHT * 3.78 * scale;
  
  canvas.width = widthPx;
  canvas.height = heightPx;
  
  // רקע לבן
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, widthPx, heightPx);
  
  // מסגרת שחורה
  ctx.strokeStyle = '#000000';
  ctx.lineWidth = 4 * scale;
  ctx.strokeRect(2 * scale, 2 * scale, widthPx - 4 * scale, heightPx - 4 * scale);
  
  // מסגרת פנימית מקווקוות
  ctx.strokeStyle = '#666666';
  ctx.lineWidth = 2 * scale;
  ctx.setLineDash([10 * scale, 5 * scale]);
  const innerMargin = 15 * scale;
  ctx.strokeRect(innerMargin, innerMargin, widthPx - 2 * innerMargin, heightPx - 2 * innerMargin);
  ctx.setLineDash([]);
  
  // טקסט במרכז
  ctx.fillStyle = '#333333';
  ctx.font = `bold ${14 * scale}px Arial`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('מדבקת כיול', widthPx / 2, heightPx / 2 - 12 * scale);
  ctx.font = `${12 * scale}px Arial`;
  ctx.fillText(`${CALIBRATION_TARGET_WIDTH}×${CALIBRATION_TARGET_HEIGHT} מ"מ`, widthPx / 2, heightPx / 2 + 12 * scale);
  
  // המרה ל-dataURL
  const dataUrl = canvas.toDataURL('image/png');
  
  // שמירת המדבקות הקיימות
  const savedStickers = typeof stickers !== 'undefined' ? [...stickers] : [];
  const savedOrientation = typeof pageOrientation !== 'undefined' ? pageOrientation : 'portrait';
  
  // ניקוי המדבקות והוספת מדבקת הכיול בלבד
  if (typeof stickers !== 'undefined') {
    stickers.length = 0;
  }
  
  // שינוי לרוחב אם צריך (200 מ"מ רוחב)
  if (typeof pageOrientation !== 'undefined' && CALIBRATION_TARGET_WIDTH > 200) {
    pageOrientation = 'landscape';
  }
  
  // הוספת מדבקת הכיול
  const calibrationSticker = {
    id: `sticker-calibration-${Date.now()}`,
    dataUrl: dataUrl,
    fileName: 'מדבקת-כיול.png',
    page: 0,
    x: 20,
    y: 20,
    width: CALIBRATION_TARGET_WIDTH * 3.78,
    height: CALIBRATION_TARGET_HEIGHT * 3.78,
    originalWidth: CALIBRATION_TARGET_WIDTH * 3.78,
    originalHeight: CALIBRATION_TARGET_HEIGHT * 3.78,
    words: [],
    images: [],
    precisionCut: true,
    precisionWidthMM: CALIBRATION_TARGET_WIDTH,
    precisionHeightMM: CALIBRATION_TARGET_HEIGHT,
    isCalibrationSticker: true
  };
  
  if (typeof stickers !== 'undefined') {
    stickers.push(calibrationSticker);
  }
  
  // רענון התצוגה
  if (typeof renderStickers === 'function') {
    renderStickers();
  }
  if (typeof updateFileCount === 'function') {
    updateFileCount();
  }
  
  // הודעה למשתמש
  const message = `מדבקת כיול נוספה לעורך!\n\nלחץ על כפתור ההדפסה הרגיל להדפסה.\nאחרי ההדפסה, מדוד את המלבן והזן את המידות.\n\nהמידה הצפויה: ${CALIBRATION_TARGET_WIDTH}×${CALIBRATION_TARGET_HEIGHT} מ"מ`;
  
  // שמירת המדבקות המקוריות לשחזור אחר כך
  window._calibrationBackup = {
    stickers: savedStickers,
    orientation: savedOrientation
  };
  
  alert(message);
  
  // הצגת כפתור לשחזור המדבקות המקוריות
  showCalibrationRestoreButton();
}

// הצגת כפתור לשחזור המדבקות המקוריות אחרי כיול
function showCalibrationRestoreButton() {
  // בדיקה אם כבר קיים
  let restoreBtn = document.getElementById('calibrationRestoreBtn');
  if (restoreBtn) return;
  
  // יצירת הכפתור
  restoreBtn = document.createElement('button');
  restoreBtn.id = 'calibrationRestoreBtn';
  restoreBtn.className = 'fixed bottom-4 left-4 z-[9999] px-4 py-3 bg-gradient-to-r from-orange-500 to-amber-600 text-white font-bold rounded-lg shadow-lg hover:from-orange-600 hover:to-amber-700 transition-all';
  restoreBtn.innerHTML = '🔄 סיימתי כיול - שחזר מדבקות';
  restoreBtn.onclick = restoreStickersAfterCalibration;
  
  document.body.appendChild(restoreBtn);
}

// שחזור המדבקות המקוריות אחרי כיול
function restoreStickersAfterCalibration() {
  if (!window._calibrationBackup) {
    alert('אין מדבקות לשחזור');
    return;
  }
  
  // שחזור המדבקות
  if (typeof stickers !== 'undefined') {
    stickers.length = 0;
    window._calibrationBackup.stickers.forEach(s => stickers.push(s));
  }
  
  // שחזור הכיוון
  if (typeof pageOrientation !== 'undefined') {
    pageOrientation = window._calibrationBackup.orientation;
  }
  
  // רענון התצוגה
  if (typeof renderStickers === 'function') {
    renderStickers();
  }
  if (typeof updateFileCount === 'function') {
    updateFileCount();
  }
  
  // הסרת הכפתור
  const restoreBtn = document.getElementById('calibrationRestoreBtn');
  if (restoreBtn) {
    restoreBtn.remove();
  }
  
  // ניקוי הגיבוי
  delete window._calibrationBackup;
  
  // פתיחת מודל הכיול להזנת המידות
  openCalibrationModal();
  
  if (typeof showStatus === 'function') {
    showStatus(t('sizeDisplay', { w: (CALIBRATION_TARGET_WIDTH / 10), h: (CALIBRATION_TARGET_HEIGHT / 10) }));
  }
}

// החלת כיול על סמך המידות שהוזנו
function applyCalibration() {
  const measuredWidthInput = document.getElementById('calibrationMeasuredWidth');
  
  if (!measuredWidthInput) {
    alert('שגיאה: שדה המידה לא נמצא');
    return;
  }
  
  const measuredWidthCM = parseFloat(measuredWidthInput.value.replace(',', '.'));
  
  if (!measuredWidthCM || measuredWidthCM <= 0) {
    alert(t('invalidNumber'));
    return;
  }
  
  // המרה מס"מ למ"מ
  const measuredWidthMM = measuredWidthCM * 10;
  const targetWidthCM = CALIBRATION_TARGET_WIDTH / 10; // 200mm = 20cm
  
  // חישוב יחס הכיול - אותו יחס לשני הצירים
  // אם המדפסת הדפיסה 19.9 ס"מ במקום 20 ס"מ, צריך להגדיל ב-20/19.9 = 1.005
  const scale = CALIBRATION_TARGET_WIDTH / measuredWidthMM;
  
  console.log('Calibration calculated:', {
    targetWidthCM,
    targetWidthMM: CALIBRATION_TARGET_WIDTH,
    measuredWidthCM,
    measuredWidthMM,
    scale
  });
  
  // בדיקת סבירות - יחס הכיול צריך להיות קרוב ל-1
  if (scale < 0.8 || scale > 1.2) {
    const confirm = window.confirm(`יחס הכיול שחושב (${scale.toFixed(3)}) נראה קיצוני.\n\nהאם הזנת את המידה בס"מ? (צפוי: ${targetWidthCM} ס"מ)\n\nלחץ אישור להמשיך בכל זאת, או ביטול לתיקון.`);
    if (!confirm) return;
  }
  
  // שמירת הכיול
  saveCalibrationData({
    scaleX: scale,
    scaleY: scale,
    calibrated: true,
    measuredWidthCM,
    measuredWidthMM,
    targetWidthMM: CALIBRATION_TARGET_WIDTH,
    calibrationDate: new Date().toISOString()
  });
  
  // עדכון הסטטוס
  updateCalibrationStatus();
  
  // הודעת הצלחה
  const percentChange = ((scale - 1) * 100).toFixed(1);
  const direction = scale > 1 ? 'הגדלה' : 'הקטנה';
  alert(`הכיול הוחל בהצלחה!\n\nיחס כיול: ${scale.toFixed(4)}\n(${direction} של ${Math.abs(percentChange)}%)\n\nמעכשיו כל ההדפסות יותאמו אוטומטית.`);
  
  // סגירת המודל
  closeCalibrationModal();
}

// איפוס כיול
function resetCalibration() {
  if (confirm(t('confirmResetCalibration'))) {
    saveCalibrationData({ scaleX: 1, scaleY: 1, calibrated: false });
    updateCalibrationStatus();
    alert(t('calibrationReset'));
  }
}

// החלת כיול על מידות (להשתמש בפונקציות ההדפסה)
function applyCalibratedSize(widthMM, heightMM) {
  const scale = getCalibrationScale();
  return {
    width: widthMM * scale.scaleX,
    height: heightMM * scale.scaleY
  };
}

// החלת כיול על מידות בפיקסלים (MM_TO_PX = 3.7795275591)
function getCalibratedPxFromMM(widthMM, heightMM, mmToPx) {
  const scale = getCalibrationScale();
  const calibratedWidth = widthMM * scale.scaleX;
  const calibratedHeight = heightMM * scale.scaleY;
  return {
    widthPx: calibratedWidth * mmToPx,
    heightPx: calibratedHeight * mmToPx
  };
}

// פונקציה פשוטה לפתיחת המודל
function openPrecisionModal(e) {
  if (e) e.stopPropagation(); // מניעת התפשטות האירוע
  console.log('openPrecisionModal called');
  
  // עיכוב קטן כדי לוודא שהמודל לא ייסגר מיד
  setTimeout(function() {
    const modal = document.getElementById('precisionStickerModal');
    if (modal) {
      modal.classList.remove('hidden');
      console.log('Modal opened');
    } else {
      console.error('Modal not found');
      alert('שגיאה: המודל לא נמצא');
    }
  }, 10);
}

// פונקציה לסגירת המודל
function closePrecisionModal() {
  console.log('closePrecisionModal called');
  const modal = document.getElementById('precisionStickerModal');
  if (modal) {
    modal.classList.add('hidden');
    console.log('Modal closed');
  }
}

let precisionProportion = 1;

function getPrecisionProportion() {
  const slider = document.getElementById('precisionProportion');
  const value = slider ? parseFloat(slider.value) : precisionProportion;
  if (!isFinite(value) || value <= 0) return 1;
  return value;
}

function setPrecisionProportion(value) {
  if (!isFinite(value) || value <= 0) value = 1;
  precisionProportion = value;
  const slider = document.getElementById('precisionProportion');
  if (slider) slider.value = String(value);
  const numberInput = document.getElementById('precisionProportionNumber');
  if (numberInput) numberInput.value = String(value);
  const label = document.getElementById('precisionProportionValue');
  if (label) label.textContent = value.toFixed(3) + '×';
}

function precisionPxToMm(px) {
  return px / (3.78 * getPrecisionProportion());
}

function precisionMmToPx(mm) {
  return mm * 3.78 * getPrecisionProportion();
}

function syncMmInputsFromCropPx() {
  const cropArea = document.getElementById('precisionCropShape');
  const widthInput = document.getElementById('precisionWidth');
  const heightInput = document.getElementById('precisionHeight');
  const selectedShape = document.querySelector('input[name="precisionShape"]:checked');

  if (!cropArea || !widthInput || !heightInput) return;

  const wPx = cropArea.offsetWidth;
  const hPx = cropArea.offsetHeight;
  if (!wPx || !hPx) return;

  let wMm = precisionPxToMm(wPx);
  let hMm = precisionPxToMm(hPx);

  if (selectedShape?.value === 'square' || selectedShape?.value === 'circle') {
    hMm = wMm;
  }

  // המרה ממ"מ לס"מ (חלוקה ב-10)
  const wCm = wMm / 10;
  const hCm = hMm / 10;

  widthInput.value = wCm.toFixed(1);
  heightInput.value = hCm.toFixed(1);

  updateSizeDisplayFromPx(wPx, hPx, selectedShape?.value || 'rectangle');
}

// הוספת event listeners כשהדף נטען
document.addEventListener('DOMContentLoaded', function() {
  console.log('DOM loaded, setting up precision sticker');
  
  // עדכון סטטוס כיול מדפסת
  updateCalibrationStatus();
  
  // סגירה עם מקש ESC
  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
      const modal = document.getElementById('precisionStickerModal');
      if (modal && !modal.classList.contains('hidden')) {
        closePrecisionModal();
      }
      // סגירת מודל כיול גם כן
      const calibrationModal = document.getElementById('printerCalibrationModal');
      if (calibrationModal && !calibrationModal.classList.contains('hidden')) {
        closeCalibrationModal();
      }
    }
  });
  
  // כפתור סגירה
  const closeBtn = document.getElementById('precisionStickerModalClose');
  if (closeBtn) {
    closeBtn.onclick = closePrecisionModal;
    console.log('Close button listener added');
  }
  
  // רקע המודל - רק אם לוחצים על הרקע עצמו
  const backdrop = document.getElementById('precisionStickerModalBackdrop');
  if (backdrop) {
    backdrop.onclick = function(e) {
      if (e.target === backdrop) { // רק אם לוחצים על הרקע עצמו
        closePrecisionModal();
      }
    };
    console.log('Backdrop listener added');
  }
  
  // העלאת קובץ
  const fileInput = document.getElementById('precisionFileInput');
  if (fileInput) {
    fileInput.onchange = function(e) {
      const file = e.target.files[0];
      if (file) {
        console.log('File selected:', file.name);
        const reader = new FileReader();
        reader.onload = function(event) {
          loadPrecisionImage(event.target.result, file.name);
        };
        reader.readAsDataURL(file);
      }
    };
    console.log('File input listener added');
  }
  
  // תפריט העלאה
  const uploadBtn = document.getElementById('precisionUploadBtn');
  const uploadMenu = document.getElementById('precisionUploadMenu');
  
  if (uploadBtn) {
    uploadBtn.onclick = function(e) {
      e.stopPropagation();
      console.log('Upload button clicked');
      if (uploadMenu) {
        uploadMenu.classList.toggle('hidden');
        console.log('Upload menu toggled');
      }
    };
    console.log('Upload button listener added');
  }
  
  // סגירת התפריט בלחיצה מחוץ לו
  document.addEventListener('click', function(e) {
    if (uploadMenu && uploadBtn) {
      if (!uploadBtn.contains(e.target) && !uploadMenu.contains(e.target)) {
        uploadMenu.classList.add('hidden');
      }
    }
  });
  
  // כפתור החלה
  const applyBtn = document.getElementById('precisionApplyBtn');
  if (applyBtn) {
    applyBtn.onclick = function() {
      console.log('Apply button clicked');
      applyPrecisionCrop();
    };
    console.log('Apply button listener added');
  }
  
  // כפתור תצוגה מקדימה
  const previewBtn = document.getElementById('precisionPreviewBtn');
  if (previewBtn) {
    previewBtn.onclick = function() {
      console.log('Preview button clicked');
      showPrecisionPreview();
    };
    console.log('Preview button listener added');
  }
  
  // כפתור איפוס
  const resetBtn = document.getElementById('precisionResetBtn');
  if (resetBtn) {
    resetBtn.onclick = function() {
      console.log('Reset button clicked');
      resetPrecisionEditor();
    };
    console.log('Reset button listener added');
  }
  
  // שדות מידות
  const widthInput = document.getElementById('precisionWidth');
  const heightInput = document.getElementById('precisionHeight');
  const proportionSlider = document.getElementById('precisionProportion');
  const proportionNumber = document.getElementById('precisionProportionNumber');
  
  if (widthInput) {
    widthInput.oninput = function() {
      updateCropAreaFromInputs();
    };
  }
  
  if (heightInput) {
    heightInput.oninput = function() {
      updateCropAreaFromInputs();
    };
  }

  if (proportionSlider) {
    setPrecisionProportion(parseFloat(proportionSlider.value) || 1);
    proportionSlider.oninput = function() {
      const next = parseFloat(proportionSlider.value) || 1;
      setPrecisionProportion(next);
      updateCropAreaFromInputs();
    };
  }

  if (proportionNumber) {
    if (proportionSlider) {
      proportionNumber.value = proportionSlider.value;
    } else {
      proportionNumber.value = String(getPrecisionProportion());
    }

    proportionNumber.oninput = function() {
      const raw = String(proportionNumber.value || '').trim();
      if (!raw) return;

      // מאפשר גם פסיק וגם נקודה כמפריד עשרוני
      const normalized = raw.replace(',', '.');

      // בזמן הקלדה (למשל "1."), לא נכריח עדכון/עיגול עד שיש מספר מלא
      if (normalized === '.' || normalized === '-' || normalized.endsWith('.')) return;

      let next = parseFloat(normalized);
      if (!isFinite(next) || next <= 0) return;

      if (proportionSlider) {
        const min = parseFloat(proportionSlider.min);
        const max = parseFloat(proportionSlider.max);
        if (isFinite(min)) next = Math.max(min, next);
        if (isFinite(max)) next = Math.min(max, next);
        proportionSlider.value = String(next);
      }

      setPrecisionProportion(next);
      updateCropAreaFromInputs();
    };

    proportionNumber.onblur = function() {
      const raw = String(proportionNumber.value || '').trim();
      if (!raw) return;
      const normalized = raw.replace(',', '.');
      let next = parseFloat(normalized);
      if (!isFinite(next) || next <= 0) return;

      if (proportionSlider) {
        const min = parseFloat(proportionSlider.min);
        const max = parseFloat(proportionSlider.max);
        if (isFinite(min)) next = Math.max(min, next);
        if (isFinite(max)) next = Math.min(max, next);
        proportionSlider.value = String(next);
      }

      setPrecisionProportion(next);
      updateCropAreaFromInputs();
    };
  }
  
  // רדיו כפתורי צורות
  const shapeRadios = document.querySelectorAll('input[name="precisionShape"]');
  shapeRadios.forEach(radio => {
    radio.onchange = function() {
      updateCropAreaShape();
    };
  });
  
  // כפתורי הגדלה/הקטנה (0.5 ס"מ בכל לחיצה)
  const increaseBtn = document.getElementById('precisionIncrease');
  const decreaseBtn = document.getElementById('precisionDecrease');
  
  if (increaseBtn) {
    increaseBtn.onclick = function() {
      adjustCropAreaSize(0.5);
    };
  }
  
  if (decreaseBtn) {
    decreaseBtn.onclick = function() {
      adjustCropAreaSize(-0.5);
    };
  }

  // כפתורי מקסימום רוחב/גובה
  const maxWidthBtn = document.getElementById('precisionMaxWidth');
  const maxHeightBtn = document.getElementById('precisionMaxHeight');
  
  if (maxWidthBtn) {
    maxWidthBtn.onclick = function() {
      fitCropToMaxWidth();
    };
  }
  
  if (maxHeightBtn) {
    maxHeightBtn.onclick = function() {
      fitCropToMaxHeight();
    };
  }
});

function loadPrecisionImage(dataUrl, fileName) {
  console.log('Loading image:', fileName);
  
  // שמירת התמונה המקורית
  originalImageData = dataUrl;
  
  // איפוס הפרופורציה ל-1
  setPrecisionProportion(1);
  
  // עדכון שם הקובץ
  const fileNameEl = document.getElementById('precisionFileName');
  if (fileNameEl) {
    fileNameEl.textContent = fileName;
  }
  
  // הצגת השלבים
  const steps = ['precisionStep2', 'precisionStep3', 'precisionEditorArea'];
  steps.forEach(stepId => {
    const step = document.getElementById(stepId);
    if (step) {
      step.classList.remove('hidden');
    }
  });
  
  // טעינת התמונה
  const img = document.getElementById('precisionSourceImage');
  if (img) {
    img.src = dataUrl;
    img.onload = function() {
      console.log('Image loaded successfully');

      // קביעת זום התחלתי כך שהתמונה תיכנס במלואה לאזור העריכה
      // משתמשים בגודל האמיתי של אזור העריכה, לא בכל המסך
      // אם התמונה קטנה - לא נגדיל מעבר ל-100%
      try {
        const container = document.getElementById('precisionEditorContainer');
        const viewportHeight = window.innerHeight || 800;
        const viewportWidth = window.innerWidth || 1200;

        let targetHeight = viewportHeight * 0.8;
        let targetWidth = viewportWidth * 0.8;

        if (container) {
          const rect = container.getBoundingClientRect();
          if (rect.height > 0) targetHeight = rect.height * 0.9;
          if (rect.width > 0) targetWidth = rect.width * 0.9;
        }

        if (img.naturalHeight > 0 && img.naturalWidth > 0) {
          const scaleByHeight = targetHeight / img.naturalHeight;
          const scaleByWidth = targetWidth / img.naturalWidth;
          // בוחרים סקייל שמוודא שהתמונה כולה נכנסת גם לגובה וגם לרוחב
          const fitScale = Math.min(scaleByHeight, scaleByWidth);
          // לא מגדילים תמונה מעבר ל-100% – רק מקטינים אם היא ענקית
          const initialScale = Math.min(1, fitScale);
          let initialZoom = Math.round(initialScale * 100);

          if (!isFinite(initialZoom) || initialZoom <= 0) {
            initialZoom = 100;
          }

          // הגבלת טווח הזום כללי
          initialZoom = Math.max(5, Math.min(200, initialZoom));
          imageZoom = initialZoom;

          const zoomSlider = document.getElementById('precisionZoomSlider');
          const zoomValue = document.getElementById('precisionZoomValue');
          if (zoomSlider) zoomSlider.value = initialZoom;
          if (zoomValue) zoomValue.textContent = initialZoom + '%';

          console.log('Initial zoom set to:', initialZoom + '%');
        } else {
          imageZoom = 100;
        }
      } catch (zoomError) {
        console.error('Error calculating initial zoom:', zoomError);
        imageZoom = 100;
      }

      // החלת הזום הראשוני ולאחר מכן הגדרת מסגרת החיתוך על פי הפריסה בפועל
      setupImageZoom();
      setTimeout(function() {
        setupCropArea();
      }, 50);
    };
  }

}

function setupCropArea() {
  const cropArea = document.getElementById('precisionCropShape');
  const img = document.getElementById('precisionSourceImage');
  const container = document.getElementById('precisionEditorContainer');

  if (!cropArea || !img || !container) return;

  // משתמשים בגודל ובמיקום בפועל בדום אחרי הזום הראשוני
  const visibleWidth = img.offsetWidth;
  const visibleHeight = img.offsetHeight;
  const imgOffsetX = img.offsetLeft;
  const imgOffsetY = img.offsetTop;

  if (!visibleWidth || !visibleHeight) {
    console.warn('Image not laid out yet for crop area');
    return;
  }

  // גודל התחלתי של המסגרת - 5×5 ס"מ עם פרופורציה 1
  // 5 ס"מ = 50 מ"מ, 50 מ"מ * 3.78 = 189 פיקסלים
  const initialSizeCM = 5;
  const initialSizeMM = initialSizeCM * 10;
  const initialSizePX = initialSizeMM * 3.78; // פרופורציה 1
  
  // וידוא שהגודל ההתחלתי לא גדול מהתמונה
  const size = Math.min(initialSizePX, visibleWidth * 0.8, visibleHeight * 0.8);
  
  cropArea.style.width = size + 'px';
  cropArea.style.height = size + 'px';

  // מיקום התחלתי - במרכז התמונה הנראית
  const startX = imgOffsetX + (visibleWidth - size) / 2;
  const startY = imgOffsetY + (visibleHeight - size) / 2;

  cropArea.style.left = startX + 'px';
  cropArea.style.top = startY + 'px';
  cropArea.style.display = 'block';

  console.log('Crop area set up with initial 5×5 cm:', {
    visibleWidth,
    visibleHeight,
    imgOffsetX,
    imgOffsetY,
    startX,
    startY,
    size,
    initialSizeCM
  });

  // הוספת יכולת גרירה ושינוי גודל למסגרת
  setupDragging(cropArea, img);
  setupResizing(cropArea);

  // הגדרת כיוון וצורה התחלתיים
  updateContainerOrientation();
  updateCropAreaShape();

  // עדכון שדות המידות לפי הגודל ההתחלתי
  syncMmInputsFromCropPx();
}

function setupDragging(cropArea, img) {
  let isDragging = false;
  let startX, startY, startLeft, startTop;

  if (cropArea._precisionDraggingSetup) return;
  cropArea._precisionDraggingSetup = true;

  const supportsPointerEvents = typeof window !== 'undefined' && 'PointerEvent' in window;
  function getEventPoint(e) {
    if (e && e.touches && e.touches.length) return e.touches[0];
    if (e && e.changedTouches && e.changedTouches.length) return e.changedTouches[0];
    return e;
  }
  function getEventClientX(e) {
    const p = getEventPoint(e);
    return p ? p.clientX : 0;
  }
  function getEventClientY(e) {
    const p = getEventPoint(e);
    return p ? p.clientY : 0;
  }

  let activePointerId = null;

  function onStart(e) {
    // אם לוחצים על ידית השינוי גודל, לא נתחיל גרירה
    if (e.target.id === 'precisionResizeHandle') return;

    if (typeof e.isPrimary === 'boolean' && !e.isPrimary) return;
    if (typeof e.button === 'number' && e.button !== 0) return;
    
    isDragging = true;
    startX = getEventClientX(e);
    startY = getEventClientY(e);
    startLeft = parseInt(cropArea.style.left) || 0;
    startTop = parseInt(cropArea.style.top) || 0;

    if (e.pointerId != null) {
      activePointerId = e.pointerId;
      if (typeof cropArea.setPointerCapture === 'function') {
        try {
          cropArea.setPointerCapture(e.pointerId);
        } catch (_) {}
      }
    } else {
      activePointerId = null;
    }
    
    cropArea.style.cursor = 'grabbing';

    if (e.cancelable) e.preventDefault();
    
    console.log('Started dragging');
  }
  
  function onMove(e) {
    if (!isDragging) return;

    if (activePointerId != null && e.pointerId != null && e.pointerId !== activePointerId) return;

    if (e.cancelable) e.preventDefault();
    
    const deltaX = getEventClientX(e) - startX;
    const deltaY = getEventClientY(e) - startY;
    
    let newLeft = startLeft + deltaX;
    let newTop = startTop + deltaY;
    
    // הגבלה לגבולות התמונה הגלויה בפועל
    if (img) {
      const imgOffsetX = img.offsetLeft;
      const imgOffsetY = img.offsetTop;
      const scaledWidth = img.offsetWidth;
      const scaledHeight = img.offsetHeight;

      const maxLeft = imgOffsetX + scaledWidth - cropArea.offsetWidth;
      const maxTop = imgOffsetY + scaledHeight - cropArea.offsetHeight;

      newLeft = Math.max(imgOffsetX, Math.min(newLeft, maxLeft));
      newTop = Math.max(imgOffsetY, Math.min(newTop, maxTop));

      console.log('Dragging constrained to visible image:', {
        imgOffsetX,
        imgOffsetY,
        scaledWidth,
        scaledHeight,
        maxLeft,
        maxTop,
        newLeft,
        newTop
      });
    }
    
    cropArea.style.left = newLeft + 'px';
    cropArea.style.top = newTop + 'px';
  }
  
  function onEnd(e) {
    if (activePointerId != null && e && e.pointerId != null && e.pointerId !== activePointerId) return;
    if (isDragging) {
      isDragging = false;
      activePointerId = null;
      cropArea.style.cursor = 'move';
      console.log('Stopped dragging');
    }
  }

  const startOptions = { passive: false };
  if (supportsPointerEvents) {
    cropArea.addEventListener('pointerdown', onStart, startOptions);
    document.addEventListener('pointermove', onMove, startOptions);
    document.addEventListener('pointerup', onEnd, startOptions);
    document.addEventListener('pointercancel', onEnd, startOptions);
  }

  cropArea.addEventListener('mousedown', onStart, startOptions);
  document.addEventListener('mousemove', onMove, startOptions);
  document.addEventListener('mouseup', onEnd, startOptions);

  cropArea.addEventListener('touchstart', onStart, startOptions);
  document.addEventListener('touchmove', onMove, startOptions);
  document.addEventListener('touchend', onEnd, startOptions);
  document.addEventListener('touchcancel', onEnd, startOptions);
}

function setupResizing(cropArea) {
  const resizeHandle = document.getElementById('precisionResizeHandle');
  if (!resizeHandle) return;

  if (resizeHandle._precisionResizingSetup) return;
  resizeHandle._precisionResizingSetup = true;
  
  let isResizing = false;
  let startX, startY, startWidth, startHeight;

  const supportsPointerEvents = typeof window !== 'undefined' && 'PointerEvent' in window;
  function getEventPoint(e) {
    if (e && e.touches && e.touches.length) return e.touches[0];
    if (e && e.changedTouches && e.changedTouches.length) return e.changedTouches[0];
    return e;
  }
  function getEventClientX(e) {
    const p = getEventPoint(e);
    return p ? p.clientX : 0;
  }
  function getEventClientY(e) {
    const p = getEventPoint(e);
    return p ? p.clientY : 0;
  }

  let activePointerId = null;
  
  function onStart(e) {
    if (typeof e.isPrimary === 'boolean' && !e.isPrimary) return;
    if (typeof e.button === 'number' && e.button !== 0) return;
    isResizing = true;
    startX = getEventClientX(e);
    startY = getEventClientY(e);
    startWidth = cropArea.offsetWidth;
    startHeight = cropArea.offsetHeight;

    if (e.pointerId != null) {
      activePointerId = e.pointerId;
      if (typeof resizeHandle.setPointerCapture === 'function') {
        try {
          resizeHandle.setPointerCapture(e.pointerId);
        } catch (_) {}
      }
    } else {
      activePointerId = null;
    }
    
    if (e.cancelable) e.preventDefault();
    e.stopPropagation();
    
    console.log('Started resizing');
  }
  
  function onMove(e) {
    if (!isResizing) return;

    if (activePointerId != null && e.pointerId != null && e.pointerId !== activePointerId) return;

    if (e.cancelable) e.preventDefault();
    
    const deltaX = getEventClientX(e) - startX;
    const deltaY = getEventClientY(e) - startY;
    
    const selectedShape = document.querySelector('input[name="precisionShape"]:checked');
    const shape = selectedShape?.value || 'rectangle';
    
    let newWidth, newHeight;
    
    if (shape === 'square' || shape === 'circle') {
      // עבור ריבוע ועיגול - שמירה על יחס 1:1 מושלם
      const delta = Math.max(deltaX, deltaY);
      newWidth = Math.max(20, startWidth + delta);
      newHeight = newWidth; // תמיד שווה לרוחב
    } else {
      // עבור מלבן - רוחב וגובה נפרדים
      newWidth = Math.max(20, startWidth + deltaX);
      newHeight = Math.max(20, startHeight + deltaY);
    }
    
    // הגבלה לגבולות התמונה הגלויה בפועל
    const img = document.getElementById('precisionSourceImage');
    if (img) {
      const imgOffsetX = img.offsetLeft;
      const imgOffsetY = img.offsetTop;
      const scaledWidth = img.offsetWidth;
      const scaledHeight = img.offsetHeight;

      const currentLeft = parseInt(cropArea.style.left) || 0;
      const currentTop = parseInt(cropArea.style.top) || 0;

      const maxWidth = imgOffsetX + scaledWidth - currentLeft;
      const maxHeight = imgOffsetY + scaledHeight - currentTop;

      if (shape === 'square' || shape === 'circle') {
        // עבור עיגול וריבוע - הגבלה לפי הקטן מבין הרוחב והגובה
        const maxSize = Math.min(maxWidth, maxHeight);
        newWidth = Math.min(newWidth, maxSize);
        newHeight = newWidth; // שמירה על יחס 1:1 מושלם
      } else {
        // עבור מלבן - הגבלה נפרדת
        newWidth = Math.min(newWidth, maxWidth);
        newHeight = Math.min(newHeight, maxHeight);
      }

      console.log('Resize constrained to visible image:', { 
        scaledWidth, scaledHeight,
        imgOffsetX, imgOffsetY,
        maxWidth, maxHeight, 
        shape,
        newWidth, newHeight 
      });
    }
    
    cropArea.style.width = newWidth + 'px';
    cropArea.style.height = newHeight + 'px';
    
    // עדכון השדות - צריך לסנכרן מהפיקסלים בחזרה למ"מ ולס"מ
    syncMmInputsFromCropPx();
    
    // עדכון תצוגת המידות
    updateSizeDisplayFromPx(newWidth, newHeight, shape);
  }
  
  function onEnd(e) {
    if (activePointerId != null && e && e.pointerId != null && e.pointerId !== activePointerId) return;
    if (isResizing) {
      isResizing = false;
      activePointerId = null;
      console.log('Stopped resizing');
    }
  }

  const startOptions = { passive: false };
  if (supportsPointerEvents) {
    resizeHandle.addEventListener('pointerdown', onStart, startOptions);
    document.addEventListener('pointermove', onMove, startOptions);
    document.addEventListener('pointerup', onEnd, startOptions);
    document.addEventListener('pointercancel', onEnd, startOptions);
  }

  resizeHandle.addEventListener('mousedown', onStart, startOptions);
  document.addEventListener('mousemove', onMove, startOptions);
  document.addEventListener('mouseup', onEnd, startOptions);

  resizeHandle.addEventListener('touchstart', onStart, startOptions);
  document.addEventListener('touchmove', onMove, startOptions);
  document.addEventListener('touchend', onEnd, startOptions);
  document.addEventListener('touchcancel', onEnd, startOptions);
}

function updateSizeDisplayFromPx(widthPx, heightPx, shape) {
  const sizeLabel = document.getElementById('precisionShapeSize');
  if (sizeLabel) {
    const widthMM = precisionPxToMm(widthPx).toFixed(1);
    const heightMM = precisionPxToMm(heightPx).toFixed(1);
    
    if (shape === 'circle') {
      sizeLabel.textContent = `קוטר: ${widthMM} מ"מ`;
    } else {
      sizeLabel.textContent = `${widthMM}×${heightMM} מ"מ`;
    }
  }
}

function resetPrecisionEditor() {
  console.log('Resetting editor');
  
  // איפוס נתונים
  originalImageData = null;
  imageZoom = 100;
  
  // איפוס הפרופורציה ל-1
  setPrecisionProportion(1);
  
  // איפוס שם הקובץ
  const fileNameEl = document.getElementById('precisionFileName');
  if (fileNameEl) {
    fileNameEl.textContent = '';
  }
  
  // הסתרת השלבים
  const steps = ['precisionStep2', 'precisionStep3', 'precisionEditorArea'];
  steps.forEach(stepId => {
    const step = document.getElementById(stepId);
    if (step) {
      step.classList.add('hidden');
    }
  });
  
  // איפוס התמונה
  const img = document.getElementById('precisionSourceImage');
  if (img) {
    img.src = '';
    img.style.transform = 'scale(1)';
  }
  
  // איפוס input
  const fileInput = document.getElementById('precisionFileInput');
  if (fileInput) {
    fileInput.value = '';
  }
  
  // איפוס זום
  const zoomSlider = document.getElementById('precisionZoomSlider');
  const zoomValue = document.getElementById('precisionZoomValue');
  if (zoomSlider) zoomSlider.value = 100;
  if (zoomValue) zoomValue.textContent = '100%';
  
  // איפוס הסליידר של הפרופורציה
  const proportionSlider = document.getElementById('precisionProportion');
  if (proportionSlider) {
    proportionSlider.max = '5'; // איפוס ל-max המקורי
  }
}

// בדיקה שהכל נטען
setTimeout(function() {
  const btn = document.getElementById('precisionStickerBtn');
  const modal = document.getElementById('precisionStickerModal');
  
  console.log('Button exists:', !!btn);
  console.log('Modal exists:', !!modal);
  
  if (btn && !btn.onclick) {
    console.log('Adding onclick to button as fallback');
    btn.onclick = openPrecisionModal;
  }
}, 2000);

function updateCropAreaFromInputs() {
  const cropArea = document.getElementById('precisionCropShape');
  const widthInput = document.getElementById('precisionWidth');
  const heightInput = document.getElementById('precisionHeight');
  const selectedShape = document.querySelector('input[name="precisionShape"]:checked');
  
  if (!cropArea || !widthInput || !heightInput) return;
  
  // הקלט בס"מ - ממירים למ"מ (כופלים ב-10)
  const rawWidth = String(widthInput.value || '').replace(',', '.');
  const rawHeight = String(heightInput.value || '').replace(',', '.');
  let widthCM = parseFloat(rawWidth) || 5;
  let heightCM = parseFloat(rawHeight) || 5;
  
  // עבור עיגול - תמיד שמירה על יחס 1:1
  if (selectedShape && selectedShape.value === 'circle') {
    heightCM = widthCM;
    heightInput.value = widthCM.toFixed(1);
  } else if (selectedShape && selectedShape.value === 'square') {
    heightCM = widthCM;
    heightInput.value = widthCM.toFixed(1);
  }
  
  // המרה מס"מ למ"מ ואז לפיקסלים
  const widthMM = widthCM * 10;
  const heightMM = heightCM * 10;
  const widthPX = precisionMmToPx(widthMM);
  const heightPX = precisionMmToPx(heightMM);
  
  cropArea.style.width = widthPX + 'px';
  cropArea.style.height = heightPX + 'px';
  
  // הגבלה דינמית - התאמת מיקום הריבוע כך שיישאר בתוך גבולות התמונה
  const maxProportion = constrainCropAreaToImage();
  
  // הגבלת הסליידר לפרופורציה המקסימלית
  if (maxProportion !== null && isFinite(maxProportion)) {
    const slider = document.getElementById('precisionProportion');
    const currentProportion = getPrecisionProportion();
    
    if (slider) {
      // עדכון ה-max של הסליידר לפרופורציה המקסימלית (עם מרווח קטן)
      const newMax = Math.max(1, maxProportion * 1.01); // 1% מרווח
      slider.max = String(newMax);
      
      // אם הפרופורציה הנוכחית גדולה מהמקסימום, נעדכן אותה
      if (currentProportion > maxProportion) {
        setPrecisionProportion(maxProportion);
        console.log('Proportion capped at maximum:', maxProportion.toFixed(3));
      }
    }
  }
  
  updateSizeDisplayFromPx(widthPX, heightPX, selectedShape?.value || 'rectangle');
  
  console.log('Updated crop area from inputs:', widthCM, 'x', heightCM, 'cm', 'shape:', selectedShape?.value);
}

// פונקציה להגבלה דינמית של ריבוע הבחירה בתוך גבולות התמונה
// מחזירה את הפרופורציה המקסימלית שאפשר להגיע אליה
function constrainCropAreaToImage() {
  const cropArea = document.getElementById('precisionCropShape');
  const img = document.getElementById('precisionSourceImage');
  const widthInput = document.getElementById('precisionWidth');
  const heightInput = document.getElementById('precisionHeight');
  
  if (!cropArea || !img || !widthInput || !heightInput) return null;
  
  const imgOffsetX = img.offsetLeft;
  const imgOffsetY = img.offsetTop;
  const imgWidth = img.offsetWidth;
  const imgHeight = img.offsetHeight;
  
  const cropWidth = cropArea.offsetWidth;
  const cropHeight = cropArea.offsetHeight;
  
  // קבלת המיקום הנוכחי
  let currentLeft = parseInt(cropArea.style.left) || 0;
  let currentTop = parseInt(cropArea.style.top) || 0;
  
  // קבלת המידות המבוקשות בס"מ
  const rawWidth = String(widthInput.value || '').replace(',', '.');
  const rawHeight = String(heightInput.value || '').replace(',', '.');
  const widthCM = parseFloat(rawWidth) || 5;
  const heightCM = parseFloat(rawHeight) || 5;
  const widthMM = widthCM * 10;
  const heightMM = heightCM * 10;
  
  // חישוב הפרופורציה המקסימלית האפשרית
  // הפרופורציה המקסימלית היא כזו שבה הריבוע ממלא את התמונה ללא חריגה
  const maxProportionByWidth = imgWidth / (widthMM * 3.78);
  const maxProportionByHeight = imgHeight / (heightMM * 3.78);
  const maxProportion = Math.min(maxProportionByWidth, maxProportionByHeight);
  
  // בדיקה אם הריבוע גדול מהתמונה
  if (cropWidth > imgWidth || cropHeight > imgHeight) {
    // אם הריבוע גדול מהתמונה, נצמצם אותו למקסימום האפשרי
    
    // שמירה על יחס הפרופורציה המבוקש (מהשדות קלט), לא מהמידות הנוכחיות
    const desiredAspectRatio = widthMM / heightMM;
    
    let newWidth, newHeight;
    if (cropWidth > imgWidth && cropHeight > imgHeight) {
      // שני הממדים גדולים מדי - נבחר את ההגבלה המחמירה יותר
      if (imgWidth / desiredAspectRatio <= imgHeight) {
        newWidth = imgWidth;
        newHeight = imgWidth / desiredAspectRatio;
      } else {
        newHeight = imgHeight;
        newWidth = imgHeight * desiredAspectRatio;
      }
    } else if (cropWidth > imgWidth) {
      // רק הרוחב גדול מדי
      newWidth = imgWidth;
      newHeight = imgWidth / desiredAspectRatio;
    } else {
      // רק הגובה גדול מדי
      newHeight = imgHeight;
      newWidth = imgHeight * desiredAspectRatio;
    }
    
    cropArea.style.width = newWidth + 'px';
    cropArea.style.height = newHeight + 'px';
    
    // כפיית reflow כדי שהדפדפן יעדכן את המידות
    void cropArea.offsetHeight;
    
    // מרכוז הריבוע בתמונה
    currentLeft = imgOffsetX + (imgWidth - newWidth) / 2;
    currentTop = imgOffsetY + (imgHeight - newHeight) / 2;
    
    console.log('Crop area was too large, resized to fit:', { newWidth, newHeight, desiredAspectRatio });
    
    // החזרת הפרופורציה המקסימלית
    return maxProportion;
  } else {
    // הריבוע נכנס בתמונה - נוודא שהוא לא חורג מהגבולות
    
    // חישוב הגבולות המקסימליים
    const minLeft = imgOffsetX;
    const maxLeft = imgOffsetX + imgWidth - cropWidth;
    const minTop = imgOffsetY;
    const maxTop = imgOffsetY + imgHeight - cropHeight;
    
    // התאמת המיקום אם חורג
    if (currentLeft < minLeft) {
      currentLeft = minLeft;
    } else if (currentLeft > maxLeft) {
      currentLeft = maxLeft;
    }
    
    if (currentTop < minTop) {
      currentTop = minTop;
    } else if (currentTop > maxTop) {
      currentTop = maxTop;
    }
  }
  
  // עדכון המיקום
  cropArea.style.left = currentLeft + 'px';
  cropArea.style.top = currentTop + 'px';
  
  console.log('Crop area constrained to image bounds:', {
    left: currentLeft,
    top: currentTop,
    width: cropArea.offsetWidth,
    height: cropArea.offsetHeight,
    imgBounds: { x: imgOffsetX, y: imgOffsetY, width: imgWidth, height: imgHeight },
    maxProportion: maxProportion
  });
  
  return maxProportion;
}

function adjustCropAreaSize(deltaCM) {
  const widthInput = document.getElementById('precisionWidth');
  const heightInput = document.getElementById('precisionHeight');
  const selectedShape = document.querySelector('input[name="precisionShape"]:checked');
  
  if (!widthInput || !heightInput) return;
  
  const rawWidth = String(widthInput.value || '').replace(',', '.');
  const rawHeight = String(heightInput.value || '').replace(',', '.');
  const currentWidth = parseFloat(rawWidth) || 5;
  const currentHeight = parseFloat(rawHeight) || 5;
  
  const newWidth = Math.max(0.5, currentWidth + deltaCM);
  let newHeight = Math.max(0.5, currentHeight + deltaCM);
  
  // עבור עיגול - תמיד שמירה על יחס 1:1
  if (selectedShape && selectedShape.value === 'circle') {
    newHeight = newWidth;
  } else if (selectedShape && selectedShape.value === 'square') {
    newHeight = newWidth;
  }
  
  widthInput.value = newWidth.toFixed(1);
  heightInput.value = newHeight.toFixed(1);
  
  updateCropAreaFromInputs();
  
  console.log('Adjusted crop area size by', deltaCM, 'cm', 'shape:', selectedShape?.value);
}

// התאמת הפרופורציה כך שאזור החיתוך יגיע לרוחב מקסימלי של התמונה
function fitCropToMaxWidth() {
  const img = document.getElementById('precisionSourceImage');
  const widthInput = document.getElementById('precisionWidth');
  
  if (!img || !widthInput) return;
  
  // קבלת רוחב התמונה המוצגת
  const imgWidth = img.offsetWidth;
  if (imgWidth <= 0) return;
  
  // קבלת הרוחב הנוכחי בס"מ וממירים למ"מ
  const rawWidth = String(widthInput.value || '').replace(',', '.');
  const widthCM = parseFloat(rawWidth) || 5;
  const widthMM = widthCM * 10;
  
  // חישוב הפרופורציה הנדרשת:
  // רוחב בפיקסלים = רוחב במ"מ * 3.78 * פרופורציה
  // פרופורציה = רוחב התמונה / (רוחב במ"מ * 3.78)
  const newProportion = imgWidth / (widthMM * 3.78);
  
  // הגבלה לטווח הסליידר
  const slider = document.getElementById('precisionProportion');
  const min = slider ? parseFloat(slider.min) : 0.1;
  const max = slider ? parseFloat(slider.max) : 20;
  const clampedProportion = Math.max(min, Math.min(max, newProportion));
  
  // עדכון הפרופורציה
  setPrecisionProportion(clampedProportion);
  updateCropAreaFromInputs();
  
  console.log('Fit proportion to max width:', clampedProportion.toFixed(3));
}

// התאמת הפרופורציה כך שאזור החיתוך יגיע לגובה מקסימלי של התמונה
function fitCropToMaxHeight() {
  const img = document.getElementById('precisionSourceImage');
  const heightInput = document.getElementById('precisionHeight');
  
  if (!img || !heightInput) return;
  
  // קבלת גובה התמונה המוצגת
  const imgHeight = img.offsetHeight;
  if (imgHeight <= 0) return;
  
  // קבלת הגובה הנוכחי בס"מ וממירים למ"מ
  const rawHeight = String(heightInput.value || '').replace(',', '.');
  const heightCM = parseFloat(rawHeight) || 5;
  const heightMM = heightCM * 10;
  
  // חישוב הפרופורציה הנדרשת:
  // גובה בפיקסלים = גובה במ"מ * 3.78 * פרופורציה
  // פרופורציה = גובה התמונה / (גובה במ"מ * 3.78)
  const newProportion = imgHeight / (heightMM * 3.78);
  
  // הגבלה לטווח הסליידר
  const slider = document.getElementById('precisionProportion');
  const min = slider ? parseFloat(slider.min) : 0.1;
  const max = slider ? parseFloat(slider.max) : 20;
  const clampedProportion = Math.max(min, Math.min(max, newProportion));
  
  // עדכון הפרופורציה
  setPrecisionProportion(clampedProportion);
  updateCropAreaFromInputs();
  
  console.log('Fit proportion to max height:', clampedProportion.toFixed(3));
}

// משתנה לשמירת זום התמונה
let imageZoom = 100;

function updateCropAreaShape() {
  const cropArea = document.getElementById('precisionCropShape');
  const selectedShape = document.querySelector('input[name="precisionShape"]:checked');
  const heightInput = document.getElementById('precisionHeight');
  const heightWrapper = document.getElementById('precisionHeightWrapper');
  
  if (!cropArea || !selectedShape) return;
  
  const shape = selectedShape.value;
  
  // עדכון עיצוב הצורה
  if (shape === 'circle') {
    cropArea.style.borderRadius = '50%';
  } else {
    cropArea.style.borderRadius = '0';
  }
  
  // עדכון שדות המידות
  if (shape === 'square' || shape === 'circle') {
    // עבור ריבוע ועיגול - הגובה שווה לרוחב
    if (heightWrapper) heightWrapper.style.opacity = '0.5';
    if (heightInput) {
      const widthInput = document.getElementById('precisionWidth');
      heightInput.value = widthInput?.value || 50;
      heightInput.disabled = true;
    }
  } else {
    // עבור מלבן - גובה ורוחב נפרדים
    if (heightWrapper) heightWrapper.style.opacity = '1';
    if (heightInput) {
      heightInput.disabled = false;
    }
  }
  
  updateCropAreaFromInputs();
  console.log('Shape updated to:', shape);
}

function setupImageZoom() {
  console.log('setupImageZoom called');
  const img = document.getElementById('precisionSourceImage');
  const container = document.getElementById('precisionEditorContainer');
  
  if (!img || !container) {
    console.error('Missing img or container elements for zoom setup');
    return;
  }
  
  if (!img.naturalWidth || !img.naturalHeight) {
    console.error('Image not loaded yet, natural size:', img.naturalWidth, 'x', img.naturalHeight);
    setTimeout(() => {
      console.log('Retrying zoom setup...');
      setupImageZoom();
    }, 200);
    return;
  }

  container.style.minWidth = '200px';
  container.style.minHeight = '300px';

  img.style.removeProperty('width');
  img.style.removeProperty('height');
  img.style.removeProperty('transform');

  img.style.display = 'block';
  img.style.maxWidth = '100%';
  img.style.maxHeight = '100%';
  img.style.width = 'auto';
  img.style.height = 'auto';
  img.style.objectFit = 'contain';

  console.log('Image set to contain inside editor container', {
    naturalWidth: img.naturalWidth,
    naturalHeight: img.naturalHeight,
    containerWidth: container.offsetWidth,
    containerHeight: container.offsetHeight
  });
}

// משתנה לשמירת התמונה המקורית
let originalImageData = null;

function applyPrecisionCrop() {
  const cropArea = document.getElementById('precisionCropShape');
  const img = document.getElementById('precisionSourceImage');
  const container = document.getElementById('precisionEditorContainer');
  
  if (!cropArea || !img || !originalImageData || !container) {
    alert(t('uploadImageFirst'));
    return;
  }
  
  console.log('Starting precision crop...');
  
  try {
    // קבלת מידות אזור החיתוך יחסית לקונטיינר
    const cropRect = {
      x: parseInt(cropArea.style.left) || 0,
      y: parseInt(cropArea.style.top) || 0,
      width: cropArea.offsetWidth,
      height: cropArea.offsetHeight
    };
    
    // קבלת מיקום התמונה בקונטיינר
    const imgRect = img.getBoundingClientRect();
    const containerRect = container.getBoundingClientRect();

    const containerStyle = window.getComputedStyle(container);
    const containerPaddingLeft = parseFloat(containerStyle.paddingLeft) || 0;
    const containerPaddingTop = parseFloat(containerStyle.paddingTop) || 0;
    const containerBorderLeft = parseFloat(containerStyle.borderLeftWidth) || 0;
    const containerBorderTop = parseFloat(containerStyle.borderTopWidth) || 0;
    
    // חישוב המיקום האמיתי של התמונה יחסית לקונטיינר
    // גישה חלופית - חישוב מיקום התמונה ישירות מה-offset
    const imgOffsetX = img.offsetLeft;
    const imgOffsetY = img.offsetTop;
    
    const cropBoxRect = cropArea.getBoundingClientRect();
    const imgBoxRect = img.getBoundingClientRect();

    const relX = cropBoxRect.left - imgBoxRect.left;
    const relY = cropBoxRect.top - imgBoxRect.top;
    const relW = cropBoxRect.width;
    const relH = cropBoxRect.height;

    const imgBoxWidth = imgBoxRect.width;
    const imgBoxHeight = imgBoxRect.height;

    // יצירת canvas לחיתוך
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    
    // הגדרת גודל הפלט (רזולוציה גבוהה)
    const scale = 3;
    canvas.width = cropRect.width * scale;
    canvas.height = cropRect.height * scale;
    
    // טעינת התמונה המקורית
    const originalImg = new Image();
    originalImg.onload = function() {
      console.log('Original image loaded for cropping');
      console.log('Original image size:', originalImg.naturalWidth, 'x', originalImg.naturalHeight);
      console.log('Displayed image size:', img.offsetWidth, 'x', img.offsetHeight);

      const scaleX = originalImg.naturalWidth / imgBoxWidth;
      const scaleY = originalImg.naturalHeight / imgBoxHeight;

      console.log('Scale factors:', { scaleX, scaleY });

      const srcX = Math.max(0, relX * scaleX);
      const srcY = Math.max(0, relY * scaleY);
      const srcWidth = relW * scaleX;
      const srcHeight = relH * scaleY;

      console.log('Source coordinates before bounds check:', { srcX, srcY, srcWidth, srcHeight });

      // הגבלה לגבולות התמונה המקורית
      const finalSrcX = Math.max(0, Math.min(srcX, originalImg.naturalWidth - 1));
      const finalSrcY = Math.max(0, Math.min(srcY, originalImg.naturalHeight - 1));
      const finalSrcWidth = Math.min(srcWidth, originalImg.naturalWidth - finalSrcX);
      const finalSrcHeight = Math.min(srcHeight, originalImg.naturalHeight - finalSrcY);

      console.log('Final source coordinates:', { 
        finalSrcX, finalSrcY, finalSrcWidth, finalSrcHeight 
      });
      
      // בדיקת צורה לחיתוך עיגול
      const selectedShape = document.querySelector('input[name="precisionShape"]:checked');
      if (selectedShape?.value === 'circle') {
        // יצירת מסכה עיגולית
        ctx.beginPath();
        ctx.arc(canvas.width / 2, canvas.height / 2, Math.min(canvas.width, canvas.height) / 2, 0, Math.PI * 2);
        ctx.closePath();
        ctx.clip();
      }
      
      // ציור התמונה החתוכה
      ctx.drawImage(
        originalImg,
        finalSrcX, finalSrcY, finalSrcWidth, finalSrcHeight,
        0, 0, canvas.width, canvas.height
      );

      // המרה ל-dataURL
      const croppedDataUrl = canvas.toDataURL('image/png', 0.95);

      // הוספה לעורך הראשי - הקלט בס"מ, ממירים למ"מ
      const widthInput = document.getElementById('precisionWidth');
      const heightInput = document.getElementById('precisionHeight');

      const rawWidth = String(widthInput ? widthInput.value : '').replace(',', '.');
      const rawHeight = String(heightInput ? heightInput.value : '').replace(',', '.');
      let outWidthCm = parseFloat(rawWidth);
      let outHeightCm = parseFloat(rawHeight);

      if (!isFinite(outWidthCm) || outWidthCm <= 0) outWidthCm = 5;
      if (!isFinite(outHeightCm) || outHeightCm <= 0) outHeightCm = 5;

      if (selectedShape?.value === 'square' || selectedShape?.value === 'circle') {
        outHeightCm = outWidthCm;
      }

      // המרה מס"מ למ"מ ואז לפיקסלים (מ"מ * 3.78 = פיקסלים)
      const outWidthMm = outWidthCm * 10;
      const outHeightMm = outHeightCm * 10;

      // הוספה לעורך הראשי
      if (typeof stickers !== 'undefined') {
        if (typeof pushHistory === 'function') {
          pushHistory();
        }

        const widthPx = outWidthMm * 3.78;
        const heightPx = outHeightMm * 3.78;
        const newSticker = {
          id: `precision-sticker-${Date.now()}-${Math.random().toString(36).slice(2)}`,
          dataUrl: croppedDataUrl,
          fileName: `מדבקה מדויקת ${outWidthMm.toFixed(0)}×${outHeightMm.toFixed(0)} מ"מ`,
          page: 0,
          x: 20,
          y: 20,
          width: widthPx,
          height: heightPx,
          originalWidth: widthPx,
          originalHeight: heightPx,
          words: [],
          images: [],
          precisionCut: true,
          precisionWidthMM: outWidthMm,
          precisionHeightMM: outHeightMm
        };

        stickers.push(newSticker);

        if (typeof renderStickers === 'function') {
          renderStickers();
        }
        if (typeof scrollToAddedElement === 'function') {
          scrollToAddedElement(`[data-sticker-index="${stickers.length - 1}"]`);
        }
        if (typeof updateFileCount === 'function') {
          updateFileCount();
        }
      }

      // סגירת המודל
      closePrecisionModal();

      console.log('Precision crop completed successfully');

      // הודעת הצלחה
      if (typeof showStatus === 'function') {
        showStatus(t('precisionStickerAdded'));
      } else {
        alert(t('precisionStickerAdded'));
      }
    };
    
    originalImg.onerror = function() {
      console.error('Failed to load original image for cropping');
      alert(t('imageProcessError'));
    };
    
    originalImg.src = originalImageData;
    
  } catch (error) {
    console.error('Error in precision crop:', error);
    alert(t('imageProcessError') + ': ' + error.message);
  }
}

function addPrecisionStickerToMain(dataUrl, widthPx, heightPx) {
  console.log('Adding precision sticker to main editor');
  
  // המרה למילימטרים לתיעוד
  const widthMM = (widthPx / 3.78).toFixed(1);
  const heightMM = (heightPx / 3.78).toFixed(1);
  
  const cropArea = document.getElementById('precisionCropShape');
  const img = document.getElementById('precisionSourceImage');
  const container = document.getElementById('precisionEditorContainer');
  const previewArea = document.getElementById('precisionPreviewArea');
  const previewCanvas = document.getElementById('precisionPreviewCanvas');
  
  console.log('Preview elements check:', {
    cropArea: !!cropArea,
    img: !!img,
    container: !!container,
    previewArea: !!previewArea,
    previewCanvas: !!previewCanvas,
    originalImageData: !!originalImageData
  });
  
  if (!cropArea || !img || !originalImageData || !container || !previewCanvas) {
    console.error('Missing elements for preview:', {
      cropArea: !!cropArea,
      img: !!img,
      originalImageData: !!originalImageData,
      container: !!container,
      previewCanvas: !!previewCanvas
    });
    alert('יש להעלות תמונה ולהגדיר אזור חיתוך קודם');
    return;
  }
  
  console.log('Showing precision preview...');
  
  try {
    // קבלת מידות אזור החיתוך יחסית לקונטיינר
    const cropRect = {
      x: parseInt(cropArea.style.left) || 0,
      y: parseInt(cropArea.style.top) || 0,
      width: cropArea.offsetWidth,
      height: cropArea.offsetHeight
    };
    
    // קבלת מיקום התמונה בקונטיינר
    const imgRect = img.getBoundingClientRect();
    const containerRect = container.getBoundingClientRect();

    const containerStyle = window.getComputedStyle(container);
    const containerPaddingLeft = parseFloat(containerStyle.paddingLeft) || 0;
    const containerPaddingTop = parseFloat(containerStyle.paddingTop) || 0;
    const containerBorderLeft = parseFloat(containerStyle.borderLeftWidth) || 0;
    const containerBorderTop = parseFloat(containerStyle.borderTopWidth) || 0;
    
    // חישוב המיקום האמיתי של התמונה יחסית לקונטיינר
    // גישה חלופית - חישוב מיקום התמונה ישירות מה-offset
    const imgOffsetX = img.offsetLeft;
    const imgOffsetY = img.offsetTop;
    
    const cropBoxRect = cropArea.getBoundingClientRect();
    const imgBoxRect = img.getBoundingClientRect();

    const relX = cropBoxRect.left - imgBoxRect.left;
    const relY = cropBoxRect.top - imgBoxRect.top;
    const relW = cropBoxRect.width;
    const relH = cropBoxRect.height;

    const imgBoxWidth = imgBoxRect.width;
    const imgBoxHeight = imgBoxRect.height;

    // הגדרת גודל התצוגה המקדימה
    const previewSize = 200;
    const aspectRatio = cropRect.width / cropRect.height;
    let canvasWidth, canvasHeight;
    
    if (aspectRatio > 1) {
      canvasWidth = previewSize;
      canvasHeight = previewSize / aspectRatio;
    } else {
      canvasWidth = previewSize * aspectRatio;
      canvasHeight = previewSize;
    }
    
    previewCanvas.width = canvasWidth;
    previewCanvas.height = canvasHeight;
    previewCanvas.style.width = canvasWidth + 'px';
    previewCanvas.style.height = canvasHeight + 'px';
    
    const ctx = previewCanvas.getContext('2d');
    
    // טעינת התמונה המקורית
    const originalImg = new Image();
    originalImg.onload = function() {
      console.log('Original image loaded for preview');

      const scaleX = originalImg.naturalWidth / imgBoxWidth;
      const scaleY = originalImg.naturalHeight / imgBoxHeight;

      console.log('Preview - scale factors:', { scaleX, scaleY });

      const srcX = Math.max(0, relX * scaleX);
      const srcY = Math.max(0, relY * scaleY);
      const srcWidth = relW * scaleX;
      const srcHeight = relH * scaleY;

      // הגבלה לגבולות התמונה המקורית
      const finalSrcX = Math.max(0, Math.min(srcX, originalImg.naturalWidth - 1));
      const finalSrcY = Math.max(0, Math.min(srcY, originalImg.naturalHeight - 1));
      const finalSrcWidth = Math.min(srcWidth, originalImg.naturalWidth - finalSrcX);
      const finalSrcHeight = Math.min(srcHeight, originalImg.naturalHeight - finalSrcY);

      console.log('Preview source coordinates:', { 
        finalSrcX, finalSrcY, finalSrcWidth, finalSrcHeight 
      });
      
      // בדיקת צורה לחיתוך עיגול
      const selectedShape = document.querySelector('input[name="precisionShape"]:checked');
      console.log('Selected shape for preview:', selectedShape?.value);
      
      if (selectedShape?.value === 'circle') {
        // יצירת מסכה עיגולית
        ctx.save(); // שמירת מצב הקנבס
        ctx.beginPath();
        ctx.arc(canvasWidth / 2, canvasHeight / 2, Math.min(canvasWidth, canvasHeight) / 2, 0, Math.PI * 2);
        ctx.closePath();
        ctx.clip();
      }
      
      // ציור התמונה החתוכה
      console.log('Drawing preview image with params:', {
        srcX: finalSrcX, 
        srcY: finalSrcY, 
        srcWidth: finalSrcWidth, 
        srcHeight: finalSrcHeight,
        destWidth: canvasWidth,
        destHeight: canvasHeight
      });
      
      try {
        ctx.drawImage(
          originalImg,
          finalSrcX, finalSrcY, finalSrcWidth, finalSrcHeight,
          0, 0, canvasWidth, canvasHeight
        );
        console.log('Preview image drawn successfully');
      } catch (drawError) {
        console.error('Error drawing preview image:', drawError);
        alert('שגיאה בציור התצוגה המקדימה: ' + drawError.message);
        return;
      }
      
      if (selectedShape?.value === 'circle') {
        ctx.restore(); // שחזור מצב הקנבס
      }
      
      // הצגת התצוגה המקדימה
      if (previewArea) {
        previewArea.classList.remove('hidden');
        console.log('Preview area shown');
      } else {
        console.error('Preview area not found!');
      }
      
      console.log('Preview generated successfully');
    };
    
    originalImg.onerror = function() {
      console.error('Failed to load original image for preview');
      alert('שגיאה בטעינת התמונה לתצוגה מקדימה');
    };
    
    originalImg.src = originalImageData;
    
  } catch (error) {
    console.error('Error in precision preview:', error);
    alert('שגיאה בתצוגה מקדימה: ' + error.message);
  }
}