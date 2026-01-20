/**
 * Precision Image Tool - כלי הכנת תמונה מדויקת
 * מבוסס על precision-sticker.js אבל מוסיף תמונה למדבקה
 */

(function() {
  'use strict';
  
  console.log('Precision Image script loaded');

  // כל פונקציות הכיול כבר קיימות מ-precision-sticker.js
  // משתמשים בהן ישירות

  // המרות בין יחידות
  const MM_TO_PX_IMG = 3.7795275591;

  function precisionImageImageMmToPx(mm) {
    const scale = typeof getCalibrationScale === 'function' ? getCalibrationScale() : { scaleX: 1, scaleY: 1 };
    return mm * MM_TO_PX_IMG * scale.scaleX;
  }

  function precisionImageImagePxToMm(px) {
    const scale = typeof getCalibrationScale === 'function' ? getCalibrationScale() : { scaleX: 1, scaleY: 1 };
    return px / MM_TO_PX_IMG / scale.scaleX;
  }

  // משתנים מקומיים לתמונה מדויקת (לא גלובליים - בתוך IIFE)
  let originalImageData = null;
  let imageZoom = 100;

  // פונקציה פשוטה לפתיחת המודל
  function openPrecisionImageModal(e) {
    if (e) e.stopPropagation(); // מניעת התפשטות האירוע
  console.log('openPrecisionImageModal called');
  
  // עיכוב קטן כדי לוודא שהמודל לא ייסגר מיד
  setTimeout(function() {
    const modal = document.getElementById('precisionImageImageModal');
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
function closePrecisionImageModal() {
  console.log('closePrecisionImageModal called');
  const modal = document.getElementById('precisionImageModal');
  if (modal) {
    modal.classList.add('hidden');
    console.log('Modal closed');
  }
}

let precisionImageProportion = 1;

function getprecisionImageProportion() {
  const slider = document.getElementById('precisionImageProportion');
  const value = slider ? parseFloat(slider.value) : precisionImageProportion;
  if (!isFinite(value) || value <= 0) return 1;
  return value;
}

function setprecisionImageProportion(value) {
  if (!isFinite(value) || value <= 0) value = 1;
  precisionImageProportion = value;
  const slider = document.getElementById('precisionImageProportion');
  if (slider) slider.value = String(value);
  const numberInput = document.getElementById('precisionImageProportionNumber');
  if (numberInput) numberInput.value = String(value);
  const label = document.getElementById('precisionImageProportionValue');
  if (label) label.textContent = value.toFixed(3) + '×';
}

function precisionImagePxToMm(px) {
  return px / (3.78 * getprecisionImageProportion());
}

function precisionImageMmToPx(mm) {
  return mm * 3.78 * getprecisionImageProportion();
}

function syncMmInputsFromCropPx() {
  const cropArea = document.getElementById('precisionImageCropShape');
  const widthInput = document.getElementById('precisionImageWidth');
  const heightInput = document.getElementById('precisionImageHeight');
  const selectedShape = document.querySelector('input[name="precisionImageShape"]:checked');

  if (!cropArea || !widthInput || !heightInput) return;

  const wPx = cropArea.offsetWidth;
  const hPx = cropArea.offsetHeight;
  if (!wPx || !hPx) return;

  let wMm = precisionImagePxToMm(wPx);
  let hMm = precisionImagePxToMm(hPx);

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
      const modal = document.getElementById('precisionImageImageModal');
      if (modal && !modal.classList.contains('hidden')) {
        closePrecisionImageModal();
      }
      // סגירת מודל כיול גם כן
      const calibrationModal = document.getElementById('printerCalibrationModal');
      if (calibrationModal && !calibrationModal.classList.contains('hidden')) {
        closeCalibrationModal();
      }
    }
  });
  
  // כפתור סגירה
  const closeBtn = document.getElementById('precisionImageImageModalClose');
  if (closeBtn) {
    closeBtn.onclick = closePrecisionImageModal;
    console.log('Close button listener added');
  }
  
  // רקע המודל - רק אם לוחצים על הרקע עצמו
  const backdrop = document.getElementById('precisionImageImageModalBackdrop');
  if (backdrop) {
    backdrop.onclick = function(e) {
      if (e.target === backdrop) { // רק אם לוחצים על הרקע עצמו
        closePrecisionImageModal();
      }
    };
    console.log('Backdrop listener added');
  }
  
  // העלאת קובץ
  const fileInput = document.getElementById('precisionImageFileInput');
  if (fileInput) {
    fileInput.onchange = function(e) {
      const file = e.target.files[0];
      if (file) {
        console.log('File selected:', file.name);
        const reader = new FileReader();
        reader.onload = function(event) {
          loadprecisionImageImage(event.target.result, file.name);
        };
        reader.readAsDataURL(file);
      }
    };
    console.log('File input listener added');
  }
  
  // תפריט העלאה
  const uploadBtn = document.getElementById('precisionImageUploadBtn');
  const uploadMenu = document.getElementById('precisionImageUploadMenu');
  
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
  const applyBtn = document.getElementById('precisionImageApplyBtn');
  if (applyBtn) {
    applyBtn.onclick = function() {
      console.log('Apply button clicked');
      applyprecisionImageCrop();
    };
    console.log('Apply button listener added');
  }
  
  // כפתור תצוגה מקדימה
  const previewBtn = document.getElementById('precisionImagePreviewBtn');
  if (previewBtn) {
    previewBtn.onclick = function() {
      console.log('Preview button clicked');
      showprecisionImagePreview();
    };
    console.log('Preview button listener added');
  }
  
  // כפתור איפוס
  const resetBtn = document.getElementById('precisionImageResetBtn');
  if (resetBtn) {
    resetBtn.onclick = function() {
      console.log('Reset button clicked');
      resetprecisionImageEditor();
    };
    console.log('Reset button listener added');
  }
  
  // שדות מידות
  const widthInput = document.getElementById('precisionImageWidth');
  const heightInput = document.getElementById('precisionImageHeight');
  const proportionSlider = document.getElementById('precisionImageProportion');
  const proportionNumber = document.getElementById('precisionImageProportionNumber');
  
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
    setprecisionImageProportion(parseFloat(proportionSlider.value) || 1);
    proportionSlider.oninput = function() {
      const next = parseFloat(proportionSlider.value) || 1;
      setprecisionImageProportion(next);
      updateCropAreaFromInputs();
    };
  }

  if (proportionNumber) {
    if (proportionSlider) {
      proportionNumber.value = proportionSlider.value;
    } else {
      proportionNumber.value = String(getprecisionImageProportion());
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

      setprecisionImageProportion(next);
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

      setprecisionImageProportion(next);
      updateCropAreaFromInputs();
    };
  }
  
  // רדיו כפתורי צורות
  const shapeRadios = document.querySelectorAll('input[name="precisionImageShape"]');
  shapeRadios.forEach(radio => {
    radio.onchange = function() {
      updateCropAreaShape();
    };
  });
  
  // כפתורי הגדלה/הקטנה (0.5 ס"מ בכל לחיצה)
  const increaseBtn = document.getElementById('precisionImageIncrease');
  const decreaseBtn = document.getElementById('precisionImageDecrease');
  
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
  const maxWidthBtn = document.getElementById('precisionImageMaxWidth');
  const maxHeightBtn = document.getElementById('precisionImageMaxHeight');
  
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

function loadprecisionImageImage(dataUrl, fileName) {
  console.log('Loading image:', fileName);
  
  // שמירת התמונה המקורית
  originalImageData = dataUrl;
  
  // איפוס הפרופורציה ל-1
  setprecisionImageProportion(1);
  
  // עדכון שם הקובץ
  const fileNameEl = document.getElementById('precisionImageFileName');
  if (fileNameEl) {
    fileNameEl.textContent = fileName;
  }
  
  // הצגת השלבים
  const steps = ['precisionImageStep2', 'precisionImageStep3', 'precisionImageEditorArea'];
  steps.forEach(stepId => {
    const step = document.getElementById(stepId);
    if (step) {
      step.classList.remove('hidden');
    }
  });
  
  // טעינת התמונה
  const img = document.getElementById('precisionImageSourceImage');
  if (img) {
    img.src = dataUrl;
    img.onload = function() {
      console.log('Image loaded successfully');

      // קביעת זום התחלתי כך שהתמונה תיכנס במלואה לאזור העריכה
      // משתמשים בגודל האמיתי של אזור העריכה, לא בכל המסך
      // אם התמונה קטנה - לא נגדיל מעבר ל-100%
      try {
        const container = document.getElementById('precisionImageEditorContainer');
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

          const zoomSlider = document.getElementById('precisionImageZoomSlider');
          const zoomValue = document.getElementById('precisionImageZoomValue');
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
  const cropArea = document.getElementById('precisionImageCropShape');
  const img = document.getElementById('precisionImageSourceImage');
  const container = document.getElementById('precisionImageEditorContainer');

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

  cropArea.addEventListener('mousedown', function(e) {
    // אם לוחצים על ידית השינוי גודל, לא נתחיל גרירה
    if (e.target.id === 'precisionImageResizeHandle') return;
    
    isDragging = true;
    startX = e.clientX;
    startY = e.clientY;
    startLeft = parseInt(cropArea.style.left) || 0;
    startTop = parseInt(cropArea.style.top) || 0;
    
    cropArea.style.cursor = 'grabbing';
    e.preventDefault();
    
    console.log('Started dragging');
  });
  
  document.addEventListener('mousemove', function(e) {
    if (!isDragging) return;
    
    const deltaX = e.clientX - startX;
    const deltaY = e.clientY - startY;
    
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
  });
  
  document.addEventListener('mouseup', function() {
    if (isDragging) {
      isDragging = false;
      cropArea.style.cursor = 'move';
      console.log('Stopped dragging');
    }
  });
}

function setupResizing(cropArea) {
  const resizeHandle = document.getElementById('precisionImageResizeHandle');
  if (!resizeHandle) return;
  
  let isResizing = false;
  let startX, startY, startWidth, startHeight;
  
  resizeHandle.addEventListener('mousedown', function(e) {
    isResizing = true;
    startX = e.clientX;
    startY = e.clientY;
    startWidth = cropArea.offsetWidth;
    startHeight = cropArea.offsetHeight;
    
    e.preventDefault();
    e.stopPropagation();
    
    console.log('Started resizing');
  });
  
  document.addEventListener('mousemove', function(e) {
    if (!isResizing) return;
    
    const deltaX = e.clientX - startX;
    const deltaY = e.clientY - startY;
    
    const selectedShape = document.querySelector('input[name="precisionImageShape"]:checked');
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
    const img = document.getElementById('precisionImageSourceImage');
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
  });
  
  document.addEventListener('mouseup', function() {
    if (isResizing) {
      isResizing = false;
      console.log('Stopped resizing');
    }
  });
}

function updateSizeDisplayFromPx(widthPx, heightPx, shape) {
  const sizeLabel = document.getElementById('precisionImageShapeSize');
  if (sizeLabel) {
    const widthMM = precisionImagePxToMm(widthPx).toFixed(1);
    const heightMM = precisionImagePxToMm(heightPx).toFixed(1);
    
    if (shape === 'circle') {
      sizeLabel.textContent = `קוטר: ${widthMM} מ"מ`;
    } else {
      sizeLabel.textContent = `${widthMM}×${heightMM} מ"מ`;
    }
  }
}

function resetprecisionImageEditor() {
  console.log('Resetting editor');
  
  // איפוס נתונים
  originalImageData = null;
  imageZoom = 100;
  
  // איפוס הפרופורציה ל-1
  setprecisionImageProportion(1);
  
  // איפוס שם הקובץ
  const fileNameEl = document.getElementById('precisionImageFileName');
  if (fileNameEl) {
    fileNameEl.textContent = '';
  }
  
  // הסתרת השלבים
  const steps = ['precisionImageStep2', 'precisionImageStep3', 'precisionImageEditorArea'];
  steps.forEach(stepId => {
    const step = document.getElementById(stepId);
    if (step) {
      step.classList.add('hidden');
    }
  });
  
  // איפוס התמונה
  const img = document.getElementById('precisionImageSourceImage');
  if (img) {
    img.src = '';
    img.style.transform = 'scale(1)';
  }
  
  // איפוס input
  const fileInput = document.getElementById('precisionImageFileInput');
  if (fileInput) {
    fileInput.value = '';
  }
  
  // איפוס זום
  const zoomSlider = document.getElementById('precisionImageZoomSlider');
  const zoomValue = document.getElementById('precisionImageZoomValue');
  if (zoomSlider) zoomSlider.value = 100;
  if (zoomValue) zoomValue.textContent = '100%';
  
  // איפוס הסליידר של הפרופורציה
  const proportionSlider = document.getElementById('precisionImageProportion');
  if (proportionSlider) {
    proportionSlider.max = '5'; // איפוס ל-max המקורי
  }
}

// בדיקה שהכל נטען
setTimeout(function() {
  const btn = document.getElementById('precisionImageImageBtn');
  const modal = document.getElementById('precisionImageImageModal');
  
  console.log('Button exists:', !!btn);
  console.log('Modal exists:', !!modal);
  
  if (btn && !btn.onclick) {
    console.log('Adding onclick to button as fallback');
    btn.onclick = openPrecisionImageModal;
  }
}, 2000);

function updateCropAreaFromInputs() {
  const cropArea = document.getElementById('precisionImageCropShape');
  const widthInput = document.getElementById('precisionImageWidth');
  const heightInput = document.getElementById('precisionImageHeight');
  const selectedShape = document.querySelector('input[name="precisionImageShape"]:checked');
  
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
  const widthPX = precisionImageMmToPx(widthMM);
  const heightPX = precisionImageMmToPx(heightMM);
  
  cropArea.style.width = widthPX + 'px';
  cropArea.style.height = heightPX + 'px';
  
  // הגבלה דינמית - התאמת מיקום הריבוע כך שיישאר בתוך גבולות התמונה
  const maxProportion = constrainCropAreaToImage();
  
  // הגבלת הסליידר לפרופורציה המקסימלית
  if (maxProportion !== null && isFinite(maxProportion)) {
    const slider = document.getElementById('precisionImageProportion');
    const currentProportion = getprecisionImageProportion();
    
    if (slider) {
      // עדכון ה-max של הסליידר לפרופורציה המקסימלית (עם מרווח קטן)
      const newMax = Math.max(1, maxProportion * 1.01); // 1% מרווח
      slider.max = String(newMax);
      
      // אם הפרופורציה הנוכחית גדולה מהמקסימום, נעדכן אותה
      if (currentProportion > maxProportion) {
        setprecisionImageProportion(maxProportion);
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
  const cropArea = document.getElementById('precisionImageCropShape');
  const img = document.getElementById('precisionImageSourceImage');
  const widthInput = document.getElementById('precisionImageWidth');
  const heightInput = document.getElementById('precisionImageHeight');
  
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
    const maxWidth = imgWidth;
    const maxHeight = imgHeight;
    
    // שמירה על יחס הפרופורציה המקורי
    const aspectRatio = cropWidth / cropHeight;
    
    let newWidth, newHeight;
    if (cropWidth > imgWidth && cropHeight > imgHeight) {
      // שני הממדים גדולים מדי - נבחר את ההגבלה המחמירה יותר
      if (imgWidth / aspectRatio <= imgHeight) {
        newWidth = imgWidth;
        newHeight = imgWidth / aspectRatio;
      } else {
        newHeight = imgHeight;
        newWidth = imgHeight * aspectRatio;
      }
    } else if (cropWidth > imgWidth) {
      // רק הרוחב גדול מדי
      newWidth = imgWidth;
      newHeight = imgWidth / aspectRatio;
    } else {
      // רק הגובה גדול מדי
      newHeight = imgHeight;
      newWidth = imgHeight * aspectRatio;
    }
    
    cropArea.style.width = newWidth + 'px';
    cropArea.style.height = newHeight + 'px';
    
    // מרכוז הריבוע בתמונה
    currentLeft = imgOffsetX + (imgWidth - newWidth) / 2;
    currentTop = imgOffsetY + (imgHeight - newHeight) / 2;
    
    console.log('Crop area was too large, resized to fit:', { newWidth, newHeight });
    
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
  const widthInput = document.getElementById('precisionImageWidth');
  const heightInput = document.getElementById('precisionImageHeight');
  const selectedShape = document.querySelector('input[name="precisionImageShape"]:checked');
  
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
  const img = document.getElementById('precisionImageSourceImage');
  const widthInput = document.getElementById('precisionImageWidth');
  
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
  const slider = document.getElementById('precisionImageProportion');
  const min = slider ? parseFloat(slider.min) : 0.1;
  const max = slider ? parseFloat(slider.max) : 20;
  const clampedProportion = Math.max(min, Math.min(max, newProportion));
  
  // עדכון הפרופורציה
  setprecisionImageProportion(clampedProportion);
  updateCropAreaFromInputs();
  
  console.log('Fit proportion to max width:', clampedProportion.toFixed(3));
}

// התאמת הפרופורציה כך שאזור החיתוך יגיע לגובה מקסימלי של התמונה
function fitCropToMaxHeight() {
  const img = document.getElementById('precisionImageSourceImage');
  const heightInput = document.getElementById('precisionImageHeight');
  
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
  const slider = document.getElementById('precisionImageProportion');
  const min = slider ? parseFloat(slider.min) : 0.1;
  const max = slider ? parseFloat(slider.max) : 20;
  const clampedProportion = Math.max(min, Math.min(max, newProportion));
  
  // עדכון הפרופורציה
  setprecisionImageProportion(clampedProportion);
  updateCropAreaFromInputs();
  
  console.log('Fit proportion to max height:', clampedProportion.toFixed(3));
}

// משתנה לשמירת זום התמונה (כבר מוגדר למעלה)
// let imageZoom = 100;

function updateCropAreaShape() {
  const cropArea = document.getElementById('precisionImageCropShape');
  const selectedShape = document.querySelector('input[name="precisionImageShape"]:checked');
  const heightInput = document.getElementById('precisionImageHeight');
  const heightWrapper = document.getElementById('precisionImageHeightWrapper');
  
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
      const widthInput = document.getElementById('precisionImageWidth');
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
  const img = document.getElementById('precisionImageSourceImage');
  const container = document.getElementById('precisionImageEditorContainer');
  
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

// משתנה לשמירת התמונה המקורית (כבר מוגדר למעלה)
// let originalImageData = null;

function applyprecisionImageCrop() {
  const cropArea = document.getElementById('precisionImageCropShape');
  const img = document.getElementById('precisionImageSourceImage');
  const container = document.getElementById('precisionImageEditorContainer');
  
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
      const selectedShape = document.querySelector('input[name="precisionImageShape"]:checked');
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
      const widthInput = document.getElementById('precisionImageWidth');
      const heightInput = document.getElementById('precisionImageHeight');

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
      addprecisionImageImageToMain(croppedDataUrl, outWidthMm * 3.78, outHeightMm * 3.78);

      // סגירת המודל
      closePrecisionImageModal();

      console.log('Precision crop completed successfully');

      // הודעת הצלחה
      if (typeof showStatus === 'function') {
        showStatus(t('precisionImageImageAdded'));
      } else {
        alert(t('precisionImageImageAdded'));
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

function addprecisionImageImageToMain(dataUrl, widthPx, heightPx) {
  console.log('Adding Precision Image to selected sticker');
  
  // בדיקה שיש מדבקה נבחרת
  if (typeof selectedSticker === 'undefined' || selectedSticker === null) {
    alert('יש לבחור מדבקה לפני הוספת תמונה מדויקת');
    return;
  }
  
  // המרה למילימטרים לתיעוד
  const widthMM = (widthPx / 3.78).toFixed(1);
  const heightMM = (heightPx / 3.78).toFixed(1);
  
  // הוספה למדבקה הנבחרת
  if (typeof stickers !== 'undefined' && stickers[selectedSticker]) {
    if (typeof pushHistory === 'function') {
      pushHistory();
    }
    
    const sticker = stickers[selectedSticker];
    if (!sticker.images) sticker.images = [];
    
    // יצירת ID ייחודי
    const imageId = `precision-img-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
    
    const newImage = {
      id: imageId,
      dataUrl: dataUrl,
      x: sticker.width / 2 - widthPx / 2,
      y: sticker.height / 2 - heightPx / 2,
      width: widthPx,
      height: heightPx,
      opacity: 1,
      precisionImageCut: true,
      precisionImageWidthMM: parseFloat(widthMM),
      precisionImageHeightMM: parseFloat(heightMM)
    };
    
    sticker.images.push(newImage);
    
    // רענון התצוגה
    if (typeof renderStickers === 'function') {
      renderStickers();
    }
    
    console.log('Precision Image added to sticker successfully:', newImage);
    
    // הודעת הצלחה
    if (typeof showStatus === 'function') {
      showStatus(`תמונה מדויקת ${widthMM}×${heightMM} מ"מ נוספה למדבקה!`);
    }
  } else {
    console.warn('Selected sticker not found');
    alert('המדבקה הנבחרת לא נמצאה');
  }
}

function showprecisionImagePreview() {
  const cropArea = document.getElementById('precisionImageCropShape');
  const img = document.getElementById('precisionImageSourceImage');
  const container = document.getElementById('precisionImageEditorContainer');
  const previewArea = document.getElementById('precisionImagePreviewArea');
  const previewCanvas = document.getElementById('precisionImagePreviewCanvas');
  
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
      const selectedShape = document.querySelector('input[name="precisionImageShape"]:checked');
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


// הפיכת הפונקציות לגלובליות
window.openPrecisionImageModal = openPrecisionImageModal;
window.closePrecisionImageModal = closePrecisionImageModal;

console.log('Precision Image tool - functions are now global');

})(); // סגירת IIFE

