/**
 * Background Removal Tool - כלי הסרת רקע
 * משתמש ב-@imgly/background-removal
 */

console.log('Background Removal script loaded');

const IMGLY_BG_REMOVAL_VERSION = window.IMGLY_BG_REMOVAL_VERSION || '1.5.8';
window.IMGLY_BG_REMOVAL_VERSION = IMGLY_BG_REMOVAL_VERSION;
if (!window.IMGLY_BG_REMOVAL_PUBLIC_PATH) {
  window.IMGLY_BG_REMOVAL_PUBLIC_PATH = `https://staticimgly.com/@imgly/background-removal-data/${IMGLY_BG_REMOVAL_VERSION}/dist/`;
}

let _bgRemovalImportPromise = null;

async function ensureBackgroundRemovalLibraryLoaded() {
  if (typeof window.imglyRemoveBackground === 'function' || typeof window.removeBackground === 'function') {
    return window.imglyRemoveBackground || window.removeBackground;
  }

  if (_bgRemovalImportPromise) {
    return _bgRemovalImportPromise;
  }

  _bgRemovalImportPromise = (async () => {
    const candidates = [
      `https://esm.sh/@imgly/background-removal@${IMGLY_BG_REMOVAL_VERSION}?bundle&target=es2020`,
      `https://cdn.skypack.dev/@imgly/background-removal@${IMGLY_BG_REMOVAL_VERSION}`,
      `https://unpkg.com/@imgly/background-removal@${IMGLY_BG_REMOVAL_VERSION}/dist/index.mjs`,
      `https://cdn.jsdelivr.net/npm/@imgly/background-removal@${IMGLY_BG_REMOVAL_VERSION}/dist/index.mjs`
    ];

    let mod = null;
    let lastError = null;
    for (const url of candidates) {
      try {
        mod = await import(url);
        break;
      } catch (e) {
        lastError = e;
      }
    }
    if (!mod) {
      const msg = (lastError && lastError.message) ? lastError.message : String(lastError || 'unknown error');
      throw new Error(`טעינת ספריית הסרת הרקע נכשלה: ${msg}`);
    }

    const fn = mod?.default || mod?.removeBackground || mod?.imglyRemoveBackground;

    if (typeof fn !== 'function') {
      throw new Error('טעינת ספריית הסרת הרקע נכשלה');
    }

    window.imglyRemoveBackground = fn;
    if (typeof mod?.preload === 'function') {
      window.imglyBackgroundRemovalPreload = mod.preload;
    }
    return fn;
  })();

  return _bgRemovalImportPromise;
}

window.ensureBackgroundRemovalLibraryLoaded = ensureBackgroundRemovalLibraryLoaded;

// פונקציה להסרת רקע מתמונה
async function removeImageBackground(imageElement, onProgress = null) {
  try {
    await ensureBackgroundRemovalLibraryLoaded();
    const removeFunc = window.imglyRemoveBackground || window.removeBackground;
    if (typeof removeFunc === 'undefined') {
      throw new Error('ספריית הסרת הרקע לא נטענה');
    }
    
    // המרת התמונה ל-blob
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    canvas.width = imageElement.naturalWidth || imageElement.width;
    canvas.height = imageElement.naturalHeight || imageElement.height;
    ctx.drawImage(imageElement, 0, 0);
    
    return new Promise((resolve, reject) => {
      canvas.toBlob(async (blob) => {
        try {
          console.log('Starting background removal...');
          
          // הסרת הרקע
          const resultBlob = await removeFunc(blob, {
            publicPath: window.IMGLY_BG_REMOVAL_PUBLIC_PATH,
            progress: onProgress || ((key, current, total) => {
              console.log(`${key}: ${current}/${total}`);
            })
          });
          
          // המרה ל-dataURL
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result);
          reader.onerror = reject;
          reader.readAsDataURL(resultBlob);
          
        } catch (error) {
          console.error('Background removal failed:', error);
          reject(error);
        }
      }, 'image/png');
    });
    
  } catch (error) {
    console.error('Background removal error:', error);
    throw error;
  }
}

// פונקציה להסרת רקע ממדבקה
async function removeStickerBackground(stickerIndex) {
  if (!stickers[stickerIndex]) {
    alert('מדבקה לא נמצאה');
    return;
  }
  
  const sticker = stickers[stickerIndex];
  
  try {
    // הצגת הודעת טעינה
    if (typeof showStatus === 'function') {
      showStatus('מסיר רקע...');
    }
    
    // יצירת תמונה זמנית
    const tempImg = new Image();
    tempImg.crossOrigin = 'anonymous';
    
    await new Promise((resolve, reject) => {
      tempImg.onload = resolve;
      tempImg.onerror = reject;
      tempImg.src = sticker.dataUrl;
    });
    
    // הסרת הרקע
    const newDataUrl = await removeImageBackground(tempImg, (key, current, total) => {
      if (typeof showStatus === 'function') {
        const percent = Math.round((current / total) * 100);
        showStatus(`מסיר רקע... ${percent}%`);
      }
    });
    
    // עדכון המדבקה
    if (typeof pushHistory === 'function') {
      pushHistory();
    }
    
    sticker.dataUrl = newDataUrl;
    sticker.fileName = sticker.fileName.replace(/\.(jpg|jpeg|png)$/i, '_no_bg.png');
    
    // רענון התצוגה
    if (typeof renderStickers === 'function') {
      renderStickers();
    }
    
    if (typeof showStatus === 'function') {
      showStatus('הרקע הוסר בהצלחה!');
    }
    
  } catch (error) {
    console.error('Failed to remove background from sticker:', error);
    if (typeof showStatus === 'function') {
      showStatus('שגיאה בהסרת הרקע: ' + error.message, true);
    } else {
      alert('שגיאה בהסרת הרקע: ' + error.message);
    }
  }
}

// פונקציה להסרת רקע מתמונה במדבקה
async function removeImageBackgroundInSticker(stickerIndex, imageId) {
  if (!stickers[stickerIndex]) {
    alert('מדבקה לא נמצאה');
    return;
  }
  
  const sticker = stickers[stickerIndex];
  const image = sticker.images.find(img => img.id === imageId);
  
  if (!image) {
    alert('תמונה לא נמצאה');
    return;
  }
  
  try {
    // הצגת הודעת טעינה
    if (typeof showStatus === 'function') {
      showStatus('מסיר רקע מהתמונה...');
    }
    
    // יצירת תמונה זמנית
    const tempImg = new Image();
    tempImg.crossOrigin = 'anonymous';
    
    await new Promise((resolve, reject) => {
      tempImg.onload = resolve;
      tempImg.onerror = reject;
      tempImg.src = image.dataUrl;
    });
    
    // הסרת הרקע
    const newDataUrl = await removeImageBackground(tempImg, (key, current, total) => {
      if (typeof showStatus === 'function') {
        const percent = Math.round((current / total) * 100);
        showStatus(`מסיר רקע מהתמונה... ${percent}%`);
      }
    });
    
    // עדכון התמונה
    if (typeof pushHistory === 'function') {
      pushHistory();
    }
    
    image.dataUrl = newDataUrl;
    
    // רענון התצוגה
    if (typeof renderStickers === 'function') {
      renderStickers();
    }
    
    if (typeof showStatus === 'function') {
      showStatus('הרקע הוסר מהתמונה בהצלחה!');
    }
    
  } catch (error) {
    console.error('Failed to remove background from image:', error);
    if (typeof showStatus === 'function') {
      showStatus('שגיאה בהסרת הרקע: ' + error.message, true);
    } else {
      alert('שגיאה בהסרת הרקע: ' + error.message);
    }
  }
}

// בדיקה אם הספרייה זמינה
function isBackgroundRemovalAvailable() {
  return typeof imglyRemoveBackground !== 'undefined' || typeof removeBackground !== 'undefined';
}

// הוספת כפתורי הסרת רקע לממשק (נקרא מ-makor-core.js)
function addBackgroundRemovalButtons() {
  // הכפתורים יתווספו ב-makor-core.js בזמן יצירת המדבקות והתמונות
  console.log('Background removal buttons will be added by makor-core.js');
}

// אתחול כשהדף נטען
document.addEventListener('DOMContentLoaded', function() {
  console.log('Background removal module initialized');
  
  // בדיקה אם הספרייה זמינה אחרי טעינה
  setTimeout(() => {
    if (isBackgroundRemovalAvailable()) {
      console.log('Background removal is available');
    } else {
      console.warn('Background removal is not available');
    }
  }, 2000);
});