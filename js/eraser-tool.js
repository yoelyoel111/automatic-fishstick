/**
 * Eraser Tool - כלי מחק לתמונות
 */

console.log('Eraser tool loaded');

let eraserActive = false;
let eraserSize = 20;
let currentErasingImage = null;
let currentErasingSticker = null;
let eraserCanvas = null;
let eraserCtx = null;
let isErasing = false;

// דגל גלובלי שיאפשר לקוד הראשי לדעת שהמחק פעיל
window.isEraserToolActive = function() {
  return eraserActive;
};

// פונקציה גלובלית לכיבוי המחק
window.deactivateEraserTool = function() {
  if (!eraserActive) return;
  
  eraserActive = false;
  const eraserBtn = document.getElementById('eraserBtn');
  const eraserSizeControl = document.getElementById('eraserSizeControl');
  
  if (eraserBtn) {
    eraserBtn.classList.remove('ring-4', 'ring-orange-300');
    eraserBtn.style.transform = 'scale(1)';
  }
  if (eraserSizeControl) {
    eraserSizeControl.classList.add('hidden');
  }
  
  // הסרת מאזיני אירועים
  detachEraserFromImages();
  
  console.log('Eraser deactivated externally');
};

// אתחול כלי המחק
document.addEventListener('DOMContentLoaded', function() {
  const eraserBtn = document.getElementById('eraserBtn');
  const eraserSizeSlider = document.getElementById('eraserSizeSlider');
  const eraserSizeValue = document.getElementById('eraserSizeValue');
  const eraserSizeControl = document.getElementById('eraserSizeControl');
  
  if (eraserBtn) {
    eraserBtn.onclick = function() {
      toggleEraser();
    };
  }
  
  if (eraserSizeSlider) {
    eraserSizeSlider.oninput = function() {
      eraserSize = parseInt(eraserSizeSlider.value);
      if (eraserSizeValue) {
        eraserSizeValue.textContent = eraserSize;
      }
      updateEraserCursor();
    };
  }
  
  console.log('Eraser tool initialized');
});

// הפעלה/כיבוי של כלי המחק
function toggleEraser() {
  eraserActive = !eraserActive;
  const eraserBtn = document.getElementById('eraserBtn');
  const eraserSizeControl = document.getElementById('eraserSizeControl');
  
  if (eraserActive) {
    // הפעלת המחק
    if (eraserBtn) {
      eraserBtn.classList.add('ring-4', 'ring-orange-300');
      eraserBtn.style.transform = 'scale(1.05)';
    }
    if (eraserSizeControl) {
      eraserSizeControl.classList.remove('hidden');
    }
    
    // הוספת מאזיני אירועים לתמונות
    attachEraserToImages();
    
    console.log('Eraser activated');
  } else {
    // כיבוי המחק
    if (eraserBtn) {
      eraserBtn.classList.remove('ring-4', 'ring-orange-300');
      eraserBtn.style.transform = 'scale(1)';
    }
    if (eraserSizeControl) {
      eraserSizeControl.classList.add('hidden');
    }
    
    // הסרת מאזיני אירועים
    detachEraserFromImages();
    
    console.log('Eraser deactivated');
  }
}

// הוספת יכולת מחיקה לכל התמונות
function attachEraserToImages() {
  const printPreview = document.getElementById('printPreviewInner');
  if (!printPreview) return;
  
  // מציאת כל התמונות במדבקות
  const imageWrappers = printPreview.querySelectorAll('.sticker-image');
  imageWrappers.forEach(wrapper => {
    // הוספת listener רגיל
    wrapper.addEventListener('mousedown', startErasing);
  });
  
  // עדכון הסמן
  updateEraserCursor();
}

// הסרת יכולת מחיקה מהתמונות
function detachEraserFromImages() {
  const printPreview = document.getElementById('printPreviewInner');
  if (!printPreview) return;
  
  const imageWrappers = printPreview.querySelectorAll('.sticker-image');
  imageWrappers.forEach(wrapper => {
    wrapper.style.cursor = '';
    wrapper.removeEventListener('mousedown', startErasing);
  });
  
  // ניקוי canvas אם קיים
  cleanupEraserCanvas();
}

// ניקוי canvas
function cleanupEraserCanvas() {
  if (eraserCanvas && eraserCanvas.parentNode) {
    eraserCanvas.parentNode.removeChild(eraserCanvas);
  }
  eraserCanvas = null;
  eraserCtx = null;
  // לא מאפסים את currentErasingImage ו-currentErasingSticker כאן
  // כי אנחנו עדיין צריכים אותם ב-createEraserCanvas
}

// התחלת מחיקה
function startErasing(e) {
  if (!eraserActive) {
    console.log('Eraser not active');
    return;
  }
  
  // בדיקה אם לחצו על כפתור מחיקה או שינוי גודל
  if (e.target.classList.contains('delete-image-btn') || 
      e.target.classList.contains('resize-handle')) {
    console.log('Clicked on button or handle, ignoring');
    return;
  }
  
  // בדיקה אם כבר יש canvas פעיל - אם כן, זה אירוע על ה-canvas עצמו
  if (e.target.classList.contains('eraser-canvas-active')) {
    console.log('Clicked on active eraser canvas - already handled by canvas listeners');
    return;
  }
  
  // מניעת גרירת התמונה - חשוב שזה יהיה אחרי הבדיקות
  e.preventDefault();
  e.stopPropagation();
  
  const wrapper = e.currentTarget;
  const img = wrapper.querySelector('img');
  
  console.log('Start erasing - wrapper:', wrapper, 'img:', img);
  
  if (!img) {
    console.error('Could not find img element');
    return;
  }
  
  // מציאת המדבקה והתמונה
  const stickerIndex = parseInt(wrapper.dataset.stickerIndex);
  const imageId = wrapper.dataset.imageId;
  
  console.log('Sticker index:', stickerIndex, 'Image ID:', imageId);
  
  if (typeof stickers === 'undefined' || !stickers[stickerIndex]) {
    console.error('Could not find sticker, stickers:', typeof stickers, 'index:', stickerIndex);
    return;
  }
  
  currentErasingSticker = stickers[stickerIndex];
  currentErasingImage = currentErasingSticker.images.find(img => img.id === imageId);
  
  console.log('Found sticker:', currentErasingSticker, 'Found image:', currentErasingImage);
  
  if (!currentErasingImage) {
    console.error('Could not find image in sticker');
    return;
  }
  
  // יצירת canvas למחיקה
  createEraserCanvas(wrapper, img);
  
  // המתנה קצרה כדי לוודא שה-canvas נוסף ל-DOM
  setTimeout(() => {
    // התחלת מחיקה
    isErasing = true;
    
    // מחיקה בנקודה הנוכחית
    if (eraserCanvas && eraserCtx) {
      console.log('Starting immediate erase at click position');
      eraseAt(e);
    }
    
    // הוספת listeners גלובליים למעקב אחרי תנועת העכבר
    document.addEventListener('mousemove', handleGlobalMouseMove);
    document.addEventListener('mouseup', handleGlobalMouseUp);
  }, 10);
  
  console.log('Canvas created, will start erasing in 10ms');
}

// טיפול בתנועת עכבר גלובלית
function handleGlobalMouseMove(e) {
  if (isErasing && eraserCanvas && eraserCtx) {
    eraseAt(e);
  }
}

// טיפול בשחרור עכבר גלובלי
function handleGlobalMouseUp(e) {
  if (isErasing) {
    document.removeEventListener('mousemove', handleGlobalMouseMove);
    document.removeEventListener('mouseup', handleGlobalMouseUp);
    stopErasing();
  }
}

// עצירת מחיקה
function stopErasing() {
  if (!isErasing) return;
  
  console.log('Stopping erasing...');
  
  isErasing = false;
  
  // שמירת השינויים בחזרה לתמונה
  if (eraserCanvas && currentErasingImage) {
    console.log('Saving erased image...');
    currentErasingImage.dataUrl = eraserCanvas.toDataURL('image/png');
    
    // עדכון התצוגה
    if (typeof renderStickers === 'function') {
      renderStickers();
    }
    
    // שמירה בהיסטוריה
    if (typeof pushHistory === 'function') {
      pushHistory();
    }
    
    console.log('Image saved');
  }
  
  // ניקוי
  cleanupEraserCanvas();
  
  // איפוס המשתנים אחרי השמירה
  currentErasingImage = null;
  currentErasingSticker = null;
  
  // הפעלה מחדש של המאזינים
  attachEraserToImages();
  
  console.log('Stopped erasing');
}

// מחיקה בנקודה מסוימת
function eraseAt(e) {
  if (!isErasing || !eraserCanvas || !eraserCtx) {
    console.log('Cannot erase - isErasing:', isErasing, 'canvas:', !!eraserCanvas, 'ctx:', !!eraserCtx);
    return;
  }
  
  const rect = eraserCanvas.getBoundingClientRect();
  const x = e.clientX - rect.left;
  const y = e.clientY - rect.top;
  
  // בדיקה שהקואורדינטות בתוך גבולות ה-canvas
  if (x < 0 || y < 0 || x > rect.width || y > rect.height) {
    console.log('Coordinates outside canvas bounds:', x, y);
    return;
  }
  
  console.log('Erasing at:', x, y, 'eraserSize:', eraserSize);
  
  // מחיקה עם עיגול
  eraserCtx.globalCompositeOperation = 'destination-out';
  eraserCtx.beginPath();
  eraserCtx.arc(x, y, eraserSize / 2, 0, Math.PI * 2);
  eraserCtx.fill();
}

// יצירת canvas למחיקה
function createEraserCanvas(wrapper, img) {
  console.log('Creating eraser canvas...');
  
  // ניקוי canvas קודם אם קיים
  cleanupEraserCanvas();
  
  // יצירת canvas חדש
  eraserCanvas = document.createElement('canvas');
  const rect = wrapper.getBoundingClientRect();
  
  console.log('Wrapper rect:', rect);
  
  eraserCanvas.width = rect.width;
  eraserCanvas.height = rect.height;
  eraserCanvas.style.position = 'absolute';
  eraserCanvas.style.left = '0';
  eraserCanvas.style.top = '0';
  eraserCanvas.style.width = '100%';
  eraserCanvas.style.height = '100%';
  eraserCanvas.style.pointerEvents = 'auto';
  eraserCanvas.className = 'eraser-canvas-active';
  
  // הגדרת הסמן המותאם אישית ל-canvas - בגודל המחיקה המדויק
  const cursorSize = eraserSize; // הגודל המדויק של המחק
  const cursorCanvas = document.createElement('canvas');
  cursorCanvas.width = cursorSize;
  cursorCanvas.height = cursorSize;
  const cursorCtx = cursorCanvas.getContext('2d');
  
  // רקע שקוף
  cursorCtx.clearRect(0, 0, cursorSize, cursorSize);
  
  // ציור עיגול חיצוני (מתאר) - בדיוק בגודל המחיקה
  cursorCtx.strokeStyle = '#ff6b35';
  cursorCtx.lineWidth = 2;
  cursorCtx.beginPath();
  cursorCtx.arc(cursorSize / 2, cursorSize / 2, (cursorSize / 2) - 1, 0, Math.PI * 2);
  cursorCtx.stroke();
  
  // ציור עיגול פנימי קטן (נקודה במרכז)
  cursorCtx.fillStyle = '#ff6b35';
  cursorCtx.beginPath();
  cursorCtx.arc(cursorSize / 2, cursorSize / 2, 1, 0, Math.PI * 2);
  cursorCtx.fill();
  
  // המרה ל-data URL והגדרת הסמן
  const cursorUrl = cursorCanvas.toDataURL();
  eraserCanvas.style.cursor = `url(${cursorUrl}) ${cursorSize / 2} ${cursorSize / 2}, crosshair`;
  
  eraserCtx = eraserCanvas.getContext('2d');
  
  console.log('Canvas created:', eraserCanvas.width, 'x', eraserCanvas.height);
  console.log('Loading image:', currentErasingImage.dataUrl.substring(0, 50) + '...');
  
  // הוספת ה-canvas ל-wrapper מיד
  wrapper.appendChild(eraserCanvas);
  console.log('Canvas appended to wrapper');
  
  // הסתרת התמונה המקורית
  img.style.visibility = 'hidden';
  
  // טעינת התמונה המקורית ל-canvas
  const tempImg = new Image();
  tempImg.crossOrigin = 'anonymous';
  tempImg.onload = function() {
    console.log('Image loaded successfully, drawing to canvas...');
    eraserCtx.drawImage(tempImg, 0, 0, eraserCanvas.width, eraserCanvas.height);
    console.log('Eraser canvas ready!');
  };
  tempImg.onerror = function(err) {
    console.error('Failed to load image for eraser:', err);
  };
  tempImg.src = currentErasingImage.dataUrl;
  
  // הוספת event listeners ישירות ל-canvas - מיד, לא בתוך onload
  eraserCanvas.addEventListener('mousedown', function(e) {
    e.preventDefault();
    e.stopPropagation();
    isErasing = true;
    eraseAt(e);
  });
  
  eraserCanvas.addEventListener('mousemove', function(e) {
    if (isErasing) {
      eraseAt(e);
    }
  });
  
  eraserCanvas.addEventListener('mouseup', function(e) {
    stopErasing();
  });
  
  eraserCanvas.addEventListener('mouseleave', function(e) {
    if (isErasing) {
      stopErasing();
    }
  });
  
  console.log('Event listeners added to canvas');
}

// עדכון סמן המחק
function updateEraserCursor() {
  if (!eraserActive) return;
  
  // יצירת סמן מותאם אישית - עיגול בגודל המחיקה המדויק
  const cursorSize = eraserSize; // הגודל המדויק של המחק
  const canvas = document.createElement('canvas');
  canvas.width = cursorSize;
  canvas.height = cursorSize;
  const ctx = canvas.getContext('2d');
  
  // רקע שקוף
  ctx.clearRect(0, 0, cursorSize, cursorSize);
  
  // ציור עיגול חיצוני (מתאר) - בדיוק בגודל המחיקה
  ctx.strokeStyle = '#ff6b35';
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.arc(cursorSize / 2, cursorSize / 2, (cursorSize / 2) - 1, 0, Math.PI * 2);
  ctx.stroke();
  
  // ציור עיגול פנימי קטן (נקודה במרכז)
  ctx.fillStyle = '#ff6b35';
  ctx.beginPath();
  ctx.arc(cursorSize / 2, cursorSize / 2, 1, 0, Math.PI * 2);
  ctx.fill();
  
  // המרה ל-data URL
  const cursorUrl = canvas.toDataURL();
  
  // עדכון הסמן
  const imageWrappers = document.querySelectorAll('.sticker-image');
  imageWrappers.forEach(wrapper => {
    if (eraserActive) {
      wrapper.style.cursor = `url(${cursorUrl}) ${cursorSize / 2} ${cursorSize / 2}, crosshair`;
    } else {
      wrapper.style.cursor = '';
    }
  });
  
  console.log('Cursor updated, size:', cursorSize);
}
