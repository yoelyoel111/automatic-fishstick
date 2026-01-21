/**
 * Google Drive Integration Module
 * מודול לניהול התחברות עם Google ושמירה/טעינה מ-Google Drive
 */

const GoogleDriveManager = (function() {
  // Configuration
  const CLIENT_ID = '329763357606-hlk2bk3e57vboorsom0688fjdrfhq4b6.apps.googleusercontent.com';
  const SCOPES = 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email';
  const DISCOVERY_DOC = 'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest';
  const APP_FOLDER_NAME = 'הגרלומט-פרויקטים';
  
  // State
  let tokenClient = null;
  let gapiInited = false;
  let gisInited = false;
  let currentUser = null;
  let appFolderId = null;
  let autoSaveTimer = null;
  let hasUnsavedChanges = false;
  let lastSavedData = null;
  
  // Performance Cache
  const projectsCache = {
    data: null,
    timestamp: 0,
    ttl: 5 * 60 * 1000, // 5 דקות
    maxProjects: 25 // הגבלה על מספר פרויקטים
  };
  
  // Callbacks
  let onSignInChange = null;
  
  /**
   * Initialize the Google API client
   */
  async function initializeGapiClient() {
    try {
      await gapi.client.init({
        discoveryDocs: [DISCOVERY_DOC],
      });
      gapiInited = true;
      maybeEnableButtons();
      
      // ניסיון לטעון token שמור - רק אחרי ש-GIS גם מוכן
      maybeRestoreToken();
    } catch (error) {
      console.error('Error initializing GAPI client:', error);
      showStatus(t('googleApiInitError'), true);
    }
  }
  
  /**
   * Try to restore saved token from localStorage
   */
  function tryRestoreSavedToken() {
    try {
      console.log('🔍 Checking for saved token in localStorage...');
      const savedTokenStr = localStorage.getItem('google_drive_token');
      
      if (!savedTokenStr) {
        console.log('❌ No saved token found in localStorage');
        return;
      }
      
      console.log('✅ Found saved token, parsing...');
      const tokenData = JSON.parse(savedTokenStr);
      console.log('Token data:', {
        hasAccessToken: !!tokenData.access_token,
        expiresAt: tokenData.expires_at ? new Date(tokenData.expires_at).toLocaleString('he-IL') : 'N/A',
        scope: tokenData.scope
      });
      
      // בדיקה אם ה-token עדיין תקף
      const now = Date.now();
      const isValid = tokenData.expires_at && now < tokenData.expires_at;
      
      if (isValid) {
        const timeLeft = Math.round((tokenData.expires_at - now) / 1000 / 60);
        console.log(`✅ Token is still valid! ${timeLeft} minutes left`);
        console.log('🔄 Restoring token to gapi.client...');
        
        // שחזור ה-token ל-gapi
        gapi.client.setToken({
          access_token: tokenData.access_token,
          scope: tokenData.scope
        });
        
        console.log('✅ Token restored, fetching user info...');
        
        // קריאה לפונקציה שמביאה את פרטי המשתמש
        setTimeout(() => {
          fetchUserInfo();
        }, 100);
      } else {
        console.log('❌ Saved token expired, removing from localStorage');
        console.log('Expired at:', new Date(tokenData.expires_at).toLocaleString('he-IL'));
        console.log('Current time:', new Date(now).toLocaleString('he-IL'));
        localStorage.removeItem('google_drive_token');
      }
    } catch (e) {
      console.error('❌ Failed to restore token:', e);
      localStorage.removeItem('google_drive_token');
    }
  }
  
  /**
   * Load the GAPI client library
   */
  function gapiLoaded() {
    gapi.load('client', initializeGapiClient);
  }
  
  /**
   * Initialize Google Identity Services
   */
  function gisLoaded() {
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES,
      callback: handleTokenResponse,
    });
    gisInited = true;
    maybeEnableButtons();
    
    // נסה לשחזר token שמור אחרי ש-GIS מוכן
    maybeRestoreToken();
  }

  /**
   * Handle token response from Google
   */
  function handleTokenResponse(response) {
    if (response.error !== undefined) {
      console.error('Token error:', response);
      showStatus(t('googleAuthError'), true);
      return;
    }
    
    // Token is now available in response.access_token
    console.log('✅ Token received successfully from Google');
    
    // שמירת ה-token ב-localStorage להתחברות אוטומטית
    try {
      const expiresAt = Date.now() + (response.expires_in * 1000);
      const tokenData = {
        access_token: response.access_token,
        expires_at: expiresAt,
        scope: response.scope
      };
      
      console.log('💾 Saving token to localStorage...');
      console.log('Token will expire at:', new Date(expiresAt).toLocaleString('he-IL'));
      
      localStorage.setItem('google_drive_token', JSON.stringify(tokenData));
      
      // בדיקה שהשמירה עבדה
      const saved = localStorage.getItem('google_drive_token');
      if (saved) {
        console.log('✅ Token saved successfully to localStorage!');
      } else {
        console.error('❌ Failed to save token to localStorage!');
      }
    } catch (e) {
      console.error('❌ Error saving token to localStorage:', e);
    }
    
    // Small delay to ensure token is set
    setTimeout(() => {
      fetchUserInfo();
    }, 100);
  }
  
  /**
   * Fetch user information
   */
  async function fetchUserInfo() {
    try {
      const token = gapi.client.getToken();
      console.log('Token check:', token);
      
      if (!token || !token.access_token) {
        console.error('No valid token available');
        showStatus(t('googleAuthRetry'), true);
        return;
      }
      
      const response = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: {
          'Authorization': `Bearer ${token.access_token}`
        }
      });
      
      console.log('User info response status:', response.status);
      
      if (response.status === 401) {
        console.error('Token expired or invalid');
        showStatus('פג תוקף ההתחברות - נסה להתחבר שוב', true);
        gapi.client.setToken('');
        updateUIForSignedOut();
        return;
      }
      
      if (response.ok) {
        currentUser = await response.json();
        console.log('User info:', currentUser);
        updateUIForSignedIn();
        if (onSignInChange) onSignInChange(true, currentUser);
        showStatus(t('googleAuthSuccess', { name: currentUser.name || currentUser.email }));
        
        // Get or create app folder
        await getOrCreateAppFolder();
        
        // Handle connection based on content state
        await handleConnectionFlow();
      }
    } catch (error) {
      console.error('Error fetching user info:', error);
      showStatus('שגיאה בקבלת פרטי משתמש', true);
    }
  }
  
  /**
   * Enable buttons when both GAPI and GIS are initialized
   */
  function maybeEnableButtons() {
    if (gapiInited && gisInited) {
      const signInBtn = document.getElementById('googleSignInBtn');
      if (signInBtn) {
        signInBtn.disabled = false;
        signInBtn.classList.remove('opacity-50');
      }
    }
  }
  
  /**
   * Try to restore token only when both GAPI and GIS are ready
   */
  function maybeRestoreToken() {
    if (gapiInited && gisInited) {
      console.log('🚀 Both GAPI and GIS are ready, trying to restore token...');
      tryRestoreSavedToken();
    } else {
      console.log('⏳ Waiting for both GAPI and GIS to be ready...');
    }
  }
  
  /**
   * Sign in with Google
   */
  function signIn() {
    if (!tokenClient) {
      showStatus('Google API לא מאותחל', true);
      return;
    }
    
    if (gapi.client.getToken() === null) {
      // First time - prompt for consent
      tokenClient.requestAccessToken({ prompt: 'consent' });
    } else {
      // Already have token - just request without prompt
      tokenClient.requestAccessToken({ prompt: '' });
    }
  }
  
  /**
   * Sign out from Google
   */
  function signOut() {
    const token = gapi.client.getToken();
    if (token !== null) {
      google.accounts.oauth2.revoke(token.access_token);
      gapi.client.setToken('');
      
      // מחיקת ה-token השמור
      try {
        localStorage.removeItem('google_drive_token');
        console.log('Saved token removed');
      } catch (e) {
        console.error('Failed to remove saved token:', e);
      }
      
      currentUser = null;
      appFolderId = null;
      updateUIForSignedOut();
      if (onSignInChange) onSignInChange(false, null);
      showStatus('התנתקת מ-Google');
    }
  }
  
  /**
   * Check if user is signed in
   */
  function isSignedIn() {
    return gapi.client && gapi.client.getToken() !== null;
  }
  
  /**
   * Update UI for signed in state
   */
  function updateUIForSignedIn() {
    const signInBtn = document.getElementById('googleSignInBtn');
    const signOutBtn = document.getElementById('googleSignOutBtn');
    const userInfo = document.getElementById('googleUserInfo');
    const saveToDriveBtn = document.getElementById('saveToDriveBtn');
    const saveToDriveNamesBtn = document.getElementById('saveToDriveNamesBtn');
    
    if (signInBtn) signInBtn.classList.add('hidden');
    if (signOutBtn) signOutBtn.classList.remove('hidden');
    
    // Show save to drive buttons based on current tab
    const tabWords = document.getElementById('tabWords');
    const tabNamesLottery = document.getElementById('tabNamesLottery');
    
    if (tabWords && tabWords.classList.contains('active')) {
      if (saveToDriveBtn) saveToDriveBtn.classList.remove('hidden');
    } else if (tabNamesLottery && tabNamesLottery.classList.contains('active')) {
      if (saveToDriveNamesBtn) saveToDriveNamesBtn.classList.remove('hidden');
    }
    
    if (userInfo && currentUser) {
      userInfo.classList.remove('hidden');
      const userPic = document.getElementById('googleUserPic');
      const userName = document.getElementById('googleUserName');
      
      if (userPic && currentUser.picture) {
        userPic.src = currentUser.picture;
        userPic.classList.remove('hidden');
      }
      if (userName) {
        userName.textContent = currentUser.name || currentUser.email;
      }
    }
  }
  
  /**
   * Update UI for signed out state
   */
  function updateUIForSignedOut() {
    const signInBtn = document.getElementById('googleSignInBtn');
    const signOutBtn = document.getElementById('googleSignOutBtn');
    const userInfo = document.getElementById('googleUserInfo');
    const saveToDriveBtn = document.getElementById('saveToDriveBtn');
    const saveToDriveNamesBtn = document.getElementById('saveToDriveNamesBtn');
    
    if (signInBtn) signInBtn.classList.remove('hidden');
    if (signOutBtn) signOutBtn.classList.add('hidden');
    if (userInfo) userInfo.classList.add('hidden');
    if (saveToDriveBtn) saveToDriveBtn.classList.add('hidden');
    if (saveToDriveNamesBtn) saveToDriveNamesBtn.classList.add('hidden');
    
    // Hide projects banner
    hideProjectsBanner();
  }

  /**
   * Get cached projects with improved error handling
   */
  async function getCachedProjects() {
    const now = Date.now();
    
    // Return cached data if still valid
    if (projectsCache.data && (now - projectsCache.timestamp) < projectsCache.ttl) {
      console.log('Using cached projects data');
      return projectsCache.data;
    }
    
    console.log('Fetching fresh projects list (names only)');
    showStatus('טוען רשימת פרויקטים...');
    
    // Get only the file list from Drive (much faster!)
    const files = await listProjectsFromDrive();
    const limitedFiles = files.slice(0, projectsCache.maxProjects);
    
    if (limitedFiles.length === 0) {
      projectsCache.data = [];
      projectsCache.timestamp = now;
      showStatus('לא נמצאו פרויקטים');
      return [];
    }
    
    // Create project entries - identify type by filename suffix
    // Files containing _הגרלש are names lottery projects
    const filesWithTypes = limitedFiles.map(file => {
      const isNamesProject = file.name.includes('_הגרלש');
      
      return {
        ...file,
        projectType: isNamesProject ? 'names' : 'stickers',
        thumbnail: null,
        hasData: false,
        loadError: false
      };
    });
    
    // Cache the results
    projectsCache.data = filesWithTypes;
    projectsCache.timestamp = now;
    
    showStatus(`${filesWithTypes.length} פרויקטים נמצאו ✓`);
    
    // Clear status after a moment
    setTimeout(() => {
      const statusEl = document.getElementById('statusMessage');
      if (statusEl && statusEl.textContent.includes('פרויקטים נמצאו')) {
        statusEl.classList.add('hidden');
      }
    }, 2000);
    
    return filesWithTypes;
  }
  
  /**
   * Load project types in background to categorize them correctly
   */
  async function loadProjectTypesInBackground(projects) {
    let needsRefresh = false;
    
    for (const project of projects) {
      // Skip if already loaded
      if (project.hasData) continue;
      
      try {
        const projectData = await loadProjectFromDrive(project.id);
        if (projectData) {
          const actualType = projectData.projectType || 'stickers';
          
          // Check if type changed from our guess
          if (project.projectType !== actualType) {
            needsRefresh = true;
          }
          
          project.projectType = actualType;
          project.hasData = true;
          project.thumbnail = projectData.thumbnail || null;
          
          // Update the card in the UI
          const card = document.querySelector(`[data-file-id="${project.id}"]`);
          if (card) {
            card.dataset.projectType = project.projectType;
          }
        }
      } catch (e) {
        console.log(`Could not load type for ${project.name}:`, e);
        project.projectType = 'stickers'; // Default to stickers
        project.hasData = true;
      }
      
      // Small delay between loads to not overwhelm the API
      await new Promise(resolve => setTimeout(resolve, 200));
    }
    
    // After all types are loaded, refresh the banner if any type changed
    if (needsRefresh) {
      console.log('Refreshing banner to show correct categories');
      await showProjectsBanner();
    }
  }
  
  /**
   * Load thumbnails in background and update UI progressively
   */
  async function loadThumbnailsInBackground(projects) {
    // Filter projects that need thumbnails (stickers without thumbnail)
    const projectsNeedingThumbnails = projects.filter(p => 
      (p.projectType === 'stickers' || p.projectType === 'unknown') && 
      !p.thumbnail && 
      !p.loadError
    );
    
    if (projectsNeedingThumbnails.length === 0) return;
    
    console.log(`Loading thumbnails for ${projectsNeedingThumbnails.length} projects`);
    let loadedCount = 0;
    
    // Load thumbnails in parallel batches for better performance
    const batchSize = 3; // טוען 3 תמונות בו זמנית
    
    for (let i = 0; i < projectsNeedingThumbnails.length; i += batchSize) {
      const batch = projectsNeedingThumbnails.slice(i, i + batchSize);
      
      // Load batch in parallel
      const promises = batch.map(async (project) => {
        try {
          // Skip if already loaded
          if (project.thumbnail) return;
          
          const projectData = await loadProjectFromDrive(project.id);
          if (projectData) {
            // Update project type if it was unknown
            if (project.projectType === 'unknown') {
              project.projectType = projectData.projectType || 'stickers';
            }
            
            project.hasData = true;
            
            if (projectData.thumbnail) {
              project.thumbnail = projectData.thumbnail;
              
              // Update the specific card in the UI immediately
              updateProjectCardThumbnail(project.id, projectData.thumbnail);
              loadedCount++;
            }
          }
        } catch (e) {
          console.log(`Could not load thumbnail for ${project.name}`);
          project.loadError = true;
        }
      });
      
      // Wait for current batch to complete before starting next batch
      await Promise.all(promises);
      
      // Small delay between batches to avoid overwhelming the API
      if (i + batchSize < projectsNeedingThumbnails.length) {
        await new Promise(resolve => setTimeout(resolve, 300));
      }
    }
    
    if (loadedCount > 0) {
      console.log(`Loaded ${loadedCount} thumbnails`);
    }
  }
  
  /**
   * Attach event listeners to a project card
   */
  function attachCardEventListeners(card) {
    // Remove existing listeners
    const newCard = card.cloneNode(true);
    card.parentNode.replaceChild(newCard, card);
    
    // Click to load project
    newCard.addEventListener('click', async (e) => {
      if (e.target.closest('.delete-project-btn')) return;
      if (e.target.closest('.rename-project-btn')) return;
      
      const fileId = newCard.dataset.fileId;
      const fileName = newCard.dataset.fileName;
      const projectType = newCard.dataset.projectType;
      
      // Show loading with better animation
      newCard.innerHTML = '<div class="text-white text-sm flex flex-col items-center gap-2"><div class="w-6 h-6 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>טוען פרויקט...</div>';
      
      try {
        const projectData = await loadProjectFromDrive(fileId);
        if (projectData) {
          // Determine actual type from loaded data
          const actualType = projectData.projectType || 'stickers';
          
          // Update the project type in cache if it was unknown
          if (projectType === 'unknown' && actualType !== 'unknown') {
            if (projectsCache.data) {
              const cachedProject = projectsCache.data.find(p => p.id === fileId);
              if (cachedProject) {
                cachedProject.projectType = actualType;
                cachedProject.hasData = true;
              }
            }
            console.log(`Project ${fileName} is actually type: ${actualType}`);
          }
          
          // Show content below banner (in case it was hidden)
          showContentBelowBanner();
          
          // Switch to correct tab first
          switchToTab(actualType);
          
          if (actualType === 'names' && typeof loadNamesProjectData === 'function') {
            loadNamesProjectData(projectData, fileName);
          } else if (typeof loadProjectData === 'function') {
            loadProjectData(projectData, fileName);
          }
          
          // Set current project file name for auto-save
          window.currentProjectFileName = fileName;
          
          // Restore card by recreating it with proper event listeners
          const restoredCard = createProjectCard({
            id: fileId,
            name: fileName,
            modifiedTime: new Date().toISOString()
          }, actualType);
          newCard.replaceWith(restoredCard);
          attachCardEventListeners(restoredCard);
          
          // Highlight the current project
          highlightCurrentProject();
        }
        
        // Don't refresh banner after loading project - it causes issues
      } catch (error) {
        console.error('Error loading project:', error);
        showStatus('שגיאה בטעינת הפרויקט', true);
        
        // Restore card appearance on error
        const restoredCard = createProjectCard({
          id: fileId,
          name: fileName,
          modifiedTime: new Date().toISOString()
        }, projectType);
        newCard.replaceWith(restoredCard);
        attachCardEventListeners(restoredCard);
      }
    });
    
    // Delete button
    const deleteBtn = newCard.querySelector('.delete-project-btn');
    if (deleteBtn) {
      deleteBtn.addEventListener('click', async (e) => {
        e.stopPropagation();
        const fileName = newCard.dataset.fileName;
        const fileId = newCard.dataset.fileId;
        
        if (confirm(`האם אתה בטוח שברצונך למחוק את "${fileName.replace('.json', '')}"?`)) {
          await deleteProjectFromDrive(fileId);
          // No need to refresh banner - deleteProjectFromDrive handles the UI update
        }
      });
    }
    
    // Rename button
    const renameBtn = newCard.querySelector('.rename-project-btn');
    if (renameBtn) {
      renameBtn.addEventListener('click', async (e) => {
        e.stopPropagation();
        const fileName = newCard.dataset.fileName;
        const fileId = newCard.dataset.fileId;
        
        const newName = prompt('שם חדש לפרויקט:', fileName.replace('.json', ''));
        if (newName && newName.trim() && newName !== fileName.replace('.json', '')) {
          const success = await renameProjectInDrive(fileId, newName.trim());
          if (success) {
            clearProjectsCache();
            await showProjectsBanner();
          }
        }
      });
    }
  }

  /**
   * Update a specific project card with its thumbnail
   */
  function updateProjectCardThumbnail(fileId, thumbnail) {
    const card = document.querySelector(`[data-file-id="${fileId}"]`);
    if (!card || !thumbnail) return;
    
    const fileName = card.dataset.fileName;
    const file = { name: fileName, modifiedTime: new Date().toISOString() }; // Approximate
    
    // Replace entire card content with thumbnail version
    card.innerHTML = `
      <div class="absolute inset-0 bg-cover bg-center opacity-80" style="background-image: url('${thumbnail}')"></div>
      <div class="absolute inset-0 bg-gradient-to-t from-black/70 via-transparent to-black/30"></div>
      <span class="project-name absolute bottom-6 left-0 right-0 text-white text-xs font-bold text-center px-1 truncate drop-shadow-lg" title="${fileName}">${fileName.replace('.json', '')}</span>
      <span class="absolute bottom-1 left-0 right-0 text-white/80 text-[10px] text-center drop-shadow">${formatDate(file.modifiedTime)}</span>
      <button class="rename-project-btn absolute top-1 right-1 w-5 h-5 bg-blue-500/90 hover:bg-blue-600 rounded-full text-white text-xs opacity-0 group-hover:opacity-100 transition-opacity shadow" title="שנה שם">✏️</button>
      <button class="delete-project-btn absolute top-1 left-1 w-5 h-5 bg-red-500/90 hover:bg-red-600 rounded-full text-white text-xs opacity-0 group-hover:opacity-100 transition-opacity shadow" title="מחק">×</button>
    `;
    
    // Re-attach event listeners for the new buttons
    attachCardEventListeners(card);
  }
  
  /**
   * Clear projects cache (call when projects change)
   */
  function clearProjectsCache() {
    projectsCache.data = null;
    projectsCache.timestamp = 0;
  }
  
  /**
   * Highlight the current active project in the banner
   */
  function highlightCurrentProject() {
    const container = document.getElementById('projectsContainer');
    if (!container) return;
    
    const currentFileName = window.currentProjectFileName;
    
    // Remove highlight from all cards
    const allCards = container.querySelectorAll('[data-file-name]');
    allCards.forEach(card => {
      card.classList.remove('ring-4', 'ring-yellow-400', 'ring-offset-2', 'ring-offset-transparent', 'scale-110', 'z-10');
      card.style.transform = '';
    });
    
    // Add highlight to current project
    if (currentFileName) {
      const currentCard = container.querySelector(`[data-file-name="${currentFileName}"]`);
      if (currentCard) {
        currentCard.classList.add('ring-4', 'ring-yellow-400', 'ring-offset-2', 'ring-offset-transparent', 'scale-110', 'z-10');
        currentCard.style.transform = 'scale(1.1)';
        
        // Scroll the card into view
        currentCard.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
      }
    }
  }
  
  /**
   * Show projects banner with cached data
   */
  async function showProjectsBanner() {
    const banner = document.getElementById('driveProjectsBanner');
    const container = document.getElementById('projectsContainer');
    
    if (!banner || !container) return;
    
    // Show banner
    banner.classList.remove('hidden');
    
    // Clear container and show loading placeholders immediately
    container.innerHTML = '';
    
    // Add new project button
    const newProjectBtn = document.createElement('button');
    newProjectBtn.id = 'newProjectBtn';
    newProjectBtn.className = 'flex-shrink-0 w-28 h-28 bg-white/20 hover:bg-white/30 backdrop-blur rounded-xl border-2 border-dashed border-white/50 hover:border-white transition-all flex flex-col items-center justify-center gap-2 group';
    newProjectBtn.innerHTML = `
      <span class="text-3xl group-hover:scale-110 transition-transform">➕</span>
      <span class="text-white text-sm font-medium">פרויקט חדש</span>
    `;
    newProjectBtn.addEventListener('click', showNewProjectDialog);
    container.appendChild(newProjectBtn);
    
    // Show loading placeholders while fetching data
    showLoadingPlaceholders(container);
    
    // Get cached projects data
    const filesWithTypes = await getCachedProjects();
    
    // Clear loading placeholders
    clearLoadingPlaceholders(container);
    
    // Separate projects by type
    const stickerProjects = filesWithTypes.filter(f => f.projectType !== 'names');
    const namesProjects = filesWithTypes.filter(f => f.projectType === 'names');
    
    // Show stickers projects section
    if (stickerProjects.length > 0) {
      const separator = document.createElement('div');
      separator.className = 'flex-shrink-0 flex items-center justify-center px-2';
      separator.innerHTML = '<div class="w-px h-20 bg-white/30"></div>';
      container.appendChild(separator);
      
      // Stickers header
      const stickersHeader = document.createElement('div');
      stickersHeader.className = 'flex-shrink-0 flex flex-col items-center justify-center gap-0 px-4 py-1';
      stickersHeader.innerHTML = `
        <div class="text-3xl opacity-60">🏷️</div>
        <div class="text-white/90 text-sm font-bold border-b-2 border-purple-400 pb-1">מדבקות</div>
        <div class="text-purple-300 text-xs mt-1">${stickerProjects.length} פרויקטים</div>
      `;
      container.appendChild(stickersHeader);
      
      // Add sticker projects
      stickerProjects.forEach(file => {
        const card = createProjectCard(file, 'stickers');
        container.appendChild(card);
      });
    }
    
    // Show names projects section
    if (namesProjects.length > 0) {
      const separator = document.createElement('div');
      separator.className = 'flex-shrink-0 flex items-center justify-center px-2';
      separator.innerHTML = '<div class="w-px h-20 bg-white/30"></div>';
      container.appendChild(separator);
      
      // Names header
      const namesHeader = document.createElement('div');
      namesHeader.className = 'flex-shrink-0 flex flex-col items-center justify-center gap-0 px-4 py-1';
      namesHeader.innerHTML = `
        <div class="text-3xl opacity-60">👥</div>
        <div class="text-white/90 text-sm font-bold border-b-2 border-emerald-400 pb-1">הגרלת שמות</div>
        <div class="text-emerald-300 text-xs mt-1">${namesProjects.length} פרויקטים</div>
      `;
      container.appendChild(namesHeader);
      
      // Add names projects
      namesProjects.forEach(file => {
        const card = createProjectCard(file, 'names');
        container.appendChild(card);
      });
    }
    
    // If no projects at all, show empty message
    if (filesWithTypes.length === 0) {
      const emptyMsg = document.createElement('div');
      emptyMsg.className = 'text-white/70 text-sm py-8 px-4';
      emptyMsg.innerHTML = '💡 אין פרויקטים שמורים עדיין. צור פרויקט חדש!';
      container.appendChild(emptyMsg);
    }
    
    // Highlight current project after banner is shown
    highlightCurrentProject();
    
    // Load thumbnails in background for sticker projects
    if (stickerProjects.length > 0) {
      setTimeout(() => {
        loadThumbnailsInBackground(stickerProjects);
      }, 500);
    }
  }
  
  /**
   * Show loading placeholders while projects are being fetched
   */
  function showLoadingPlaceholders(container) {
    // Add separator
    const separator = document.createElement('div');
    separator.className = 'loading-placeholder flex-shrink-0 flex items-center justify-center px-2';
    separator.innerHTML = '<div class="w-px h-20 bg-white/30"></div>';
    container.appendChild(separator);
    
    // Add loading header
    const loadingHeader = document.createElement('div');
    loadingHeader.className = 'loading-placeholder flex-shrink-0 flex flex-col items-center justify-center gap-0 px-4 py-1';
    loadingHeader.innerHTML = `
      <div class="text-3xl opacity-60">⏳</div>
      <div class="text-white/90 text-sm font-bold border-b-2 border-blue-400 pb-1">טוען פרויקטים</div>
      <div class="text-blue-300 text-xs mt-1">אנא המתן...</div>
    `;
    container.appendChild(loadingHeader);
    
    // Add 6 loading card placeholders
    for (let i = 0; i < 6; i++) {
      const placeholder = document.createElement('div');
      placeholder.className = 'loading-placeholder flex-shrink-0 w-28 h-28 bg-white/10 backdrop-blur rounded-xl border border-white/20 flex flex-col items-center justify-center gap-2 animate-pulse';
      placeholder.innerHTML = `
        <div class="w-12 h-12 bg-white/20 rounded-full flex items-center justify-center">
          <div class="w-6 h-6 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>
        </div>
        <div class="w-16 h-2 bg-white/20 rounded animate-pulse"></div>
        <div class="w-12 h-1.5 bg-white/15 rounded animate-pulse"></div>
      `;
      container.appendChild(placeholder);
    }
  }
  
  /**
   * Add a single new project to the banner without refreshing everything
   */
  async function addProjectToBanner(projectFile, projectType) {
    const container = document.getElementById('projectsContainer');
    if (!container) return;
    
    // Check if project already exists in banner - don't add duplicates
    const existingCard = container.querySelector(`[data-file-id="${projectFile.id}"]`);
    if (existingCard) {
      console.log('Project already exists in banner, skipping add');
      return;
    }
    
    // Also check by name to avoid duplicates
    const existingByName = container.querySelector(`[data-file-name="${projectFile.name}"]`);
    if (existingByName) {
      console.log('Project with same name already exists in banner, skipping add');
      return;
    }
    
    // Add to cache (only if not already there)
    if (projectsCache.data) {
      const existsInCache = projectsCache.data.some(p => p.id === projectFile.id || p.name === projectFile.name);
      if (!existsInCache) {
        projectsCache.data.unshift({
          ...projectFile,
          projectType,
          thumbnail: null,
          hasData: true,
          loadError: false
        });
      }
    }
    
    // Find the projects counter
    const counterElement = container.querySelector('.text-purple-300');
    
    // If section doesn't exist, create it
    if (!counterElement) {
      const newProjectBtn = container.querySelector('#newProjectBtn');
      
      // Add separator
      const separator = document.createElement('div');
      separator.className = 'flex-shrink-0 flex items-center justify-center px-2';
      separator.innerHTML = '<div class="w-px h-20 bg-white/30"></div>';
      container.insertBefore(separator, newProjectBtn.nextSibling);
      
      // Add section header
      const headerDiv = document.createElement('div');
      headerDiv.className = 'flex-shrink-0 flex flex-col items-center justify-center gap-0 px-4 py-1';
      headerDiv.innerHTML = `
        <div class="text-3xl opacity-60">🏷️</div>
        <div class="text-white/90 text-sm font-bold border-b-2 border-purple-400 pb-1">מדבקות</div>
        <div class="text-purple-300 text-xs mt-1">1 פרויקטים</div>
      `;
      
      container.insertBefore(headerDiv, separator.nextSibling);
    } else {
      // Update existing section count
      const currentCount = parseInt(counterElement.textContent.match(/\d+/)?.[0] || '0');
      counterElement.textContent = `${currentCount + 1} פרויקטים`;
    }
    
    // Create and add the new project card
    const displayType = projectType === 'names' ? 'names' : 'stickers';
    const newCard = createProjectCard({
      id: projectFile.id,
      name: projectFile.name || `פרויקט-${Date.now()}.json`,
      modifiedTime: projectFile.modifiedTime || new Date().toISOString()
    }, displayType);
    
    // Add at the end of the projects
    container.appendChild(newCard);
    
    // Add entrance animation
    newCard.style.opacity = '0';
    newCard.style.transform = 'scale(0.8)';
    newCard.style.transition = 'all 0.3s ease';
    
    setTimeout(() => {
      newCard.style.opacity = '1';
      newCard.style.transform = 'scale(1)';
    }, 50);
    
    // Remove empty message if it exists
    const emptyMsg = container.querySelector('.text-white\\/70');
    if (emptyMsg) {
      emptyMsg.remove();
    }
  }

  /**
   * Clear loading placeholders
   */
  function clearLoadingPlaceholders(container) {
    const placeholders = container.querySelectorAll('.loading-placeholder');
    placeholders.forEach(placeholder => placeholder.remove());
  }
  
  /**
   * Create a project card element
   */
  function createProjectCard(file, projectType) {
    const icon = projectType === 'names' ? '👥' : projectType === 'unknown' ? '📄' : '🏷️';
    
    // Different colors for different project types
    let bgColor, iconBg;
    if (projectType === 'names') {
      bgColor = 'bg-emerald-500/20 hover:bg-emerald-500/35 border-emerald-300/50 hover:border-emerald-300';
      iconBg = 'bg-emerald-400/30';
    } else if (projectType === 'unknown') {
      bgColor = 'bg-blue-500/20 hover:bg-blue-500/35 border-blue-300/50 hover:border-blue-300';
      iconBg = 'bg-blue-400/30';
    } else {
      bgColor = 'bg-purple-500/20 hover:bg-purple-500/35 border-purple-300/50 hover:border-purple-300';
      iconBg = 'bg-purple-400/30';
    }
    
    // Clean display name - remove .json and _הגרלש
    const displayName = getDisplayName(file.name);
    
    const card = document.createElement('button');
    card.className = `project-card flex-shrink-0 w-28 h-28 ${bgColor} backdrop-blur rounded-xl border transition-all flex flex-col items-center justify-center gap-1 group relative overflow-hidden`;
    card.dataset.fileId = file.id;
    card.dataset.fileName = file.name;
    card.dataset.projectType = projectType;
    
    // Check if we have a thumbnail (only for stickers)
    const hasThumbnail = projectType === 'stickers' && file.thumbnail;
    const hasError = file.loadError;
    
    if (hasError) {
      // Card with error state
      card.innerHTML = `
        <div class="w-12 h-12 bg-red-400/30 rounded-full flex items-center justify-center mb-1">
          <span class="text-xl text-red-300">⚠️</span>
        </div>
        <span class="project-name text-white/70 text-xs font-medium text-center px-1 truncate w-full" title="${displayName}">${displayName}</span>
        <span class="text-red-300 text-[10px]">שגיאה בטעינה</span>
        <button class="rename-project-btn absolute top-1 right-1 w-5 h-5 bg-blue-500/80 hover:bg-blue-600 rounded-full text-white text-xs opacity-0 group-hover:opacity-100 transition-opacity" title="שנה שם">✏️</button>
        <button class="delete-project-btn absolute top-1 left-1 w-5 h-5 bg-red-500/80 hover:bg-red-600 rounded-full text-white text-xs opacity-0 group-hover:opacity-100 transition-opacity" title="מחק">×</button>
      `;
    } else if (hasThumbnail) {
      // Card with thumbnail background
      card.innerHTML = `
        <div class="absolute inset-0 bg-cover bg-center opacity-80" style="background-image: url('${file.thumbnail}')"></div>
        <div class="absolute inset-0 bg-gradient-to-t from-black/70 via-transparent to-black/30"></div>
        <span class="project-name absolute bottom-6 left-0 right-0 text-white text-xs font-bold text-center px-1 truncate drop-shadow-lg" title="${displayName}">${displayName}</span>
        <span class="absolute bottom-1 left-0 right-0 text-white/80 text-[10px] text-center drop-shadow">${formatDate(file.modifiedTime)}</span>
        <button class="rename-project-btn absolute top-1 right-1 w-5 h-5 bg-blue-500/90 hover:bg-blue-600 rounded-full text-white text-xs opacity-0 group-hover:opacity-100 transition-opacity shadow" title="שנה שם">✏️</button>
        <button class="delete-project-btn absolute top-1 left-1 w-5 h-5 bg-red-500/90 hover:bg-red-600 rounded-full text-white text-xs opacity-0 group-hover:opacity-100 transition-opacity shadow" title="מחק">×</button>
      `;
    } else {
      // Card with icon (for names or stickers without thumbnail)
      const isStickersWithoutThumbnail = projectType === 'stickers' && !file.thumbnail && file.hasData;
      
      card.innerHTML = `
        <div class="w-12 h-12 ${iconBg} rounded-full flex items-center justify-center mb-1 relative">
          <span class="text-2xl group-hover:scale-110 transition-transform">${icon}</span>
          ${isStickersWithoutThumbnail ? '<div class="absolute inset-0 rounded-full border-2 border-white/30 border-t-white animate-spin"></div>' : ''}
        </div>
        <span class="project-name text-white text-xs font-medium text-center px-1 truncate w-full" title="${displayName}">${displayName}</span>
        <span class="text-white/60 text-[10px]">${formatDate(file.modifiedTime)}</span>
        ${isStickersWithoutThumbnail ? '<span class="text-white/50 text-[9px]">טוען תמונה...</span>' : ''}
        <button class="rename-project-btn absolute top-1 right-1 w-5 h-5 bg-blue-500/80 hover:bg-blue-600 rounded-full text-white text-xs opacity-0 group-hover:opacity-100 transition-opacity" title="שנה שם">✏️</button>
        <button class="delete-project-btn absolute top-1 left-1 w-5 h-5 bg-red-500/80 hover:bg-red-600 rounded-full text-white text-xs opacity-0 group-hover:opacity-100 transition-opacity" title="מחק">×</button>
      `;
    }
    
    // Click to load project
    card.addEventListener('click', async (e) => {
      if (e.target.closest('.delete-project-btn')) return;
      if (e.target.closest('.rename-project-btn')) return;
      
      const fileId = card.dataset.fileId;
      const fileName = card.dataset.fileName;
      
      // Show loading
      card.innerHTML = '<div class="text-white text-sm flex flex-col items-center justify-center h-full"><div class="w-6 h-6 border-2 border-white/30 border-t-white rounded-full animate-spin"></div><span class="mt-2">טוען...</span></div>';
      
      const projectData = await loadProjectFromDrive(fileId);
      if (projectData) {
        // Determine actual type from loaded data
        const actualType = projectData.projectType || projectType;
        
        // Show content below banner (in case it was hidden)
        showContentBelowBanner();
        
        // Switch to correct tab first
        switchToTab(actualType);
        
        if (actualType === 'names' && typeof loadNamesProjectData === 'function') {
          loadNamesProjectData(projectData, fileName);
        } else if (typeof loadProjectData === 'function') {
          loadProjectData(projectData, fileName);
        }
        
        // Set current project file name for auto-save
        window.currentProjectFileName = fileName;
      }
      
      // Recreate card with proper event listeners (whether success or failure)
      const restoredCard = createProjectCard({
        id: fileId,
        name: fileName,
        modifiedTime: file.modifiedTime || new Date().toISOString()
      }, projectData ? (projectData.projectType || projectType) : projectType);
      card.replaceWith(restoredCard);
      
      // Highlight the current project
      highlightCurrentProject();
    });
    
    // Rename button
    card.querySelector('.rename-project-btn').addEventListener('click', async (e) => {
      e.stopPropagation();
      showRenameDialog(file.id, file.name);
    });
    
    // Delete button
    card.querySelector('.delete-project-btn').addEventListener('click', async (e) => {
      e.stopPropagation();
      if (confirm(`למחוק את "${file.name}"?`)) {
        await deleteProjectFromDrive(file.id);
        // No need to refresh banner - deleteProjectFromDrive handles the UI update
      }
    });
    
    return card;
  }
  
  /**
   * Show rename dialog
   */
  function showRenameDialog(fileId, currentName) {
    const nameWithoutExt = currentName.replace('.json', '');
    
    const modal = document.createElement('div');
    modal.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50';
    modal.innerHTML = `
      <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full mx-4">
        <div class="p-4 border-b border-gray-200">
          <h3 class="text-xl font-bold text-gray-800">✏️ שינוי שם פרויקט</h3>
        </div>
        <div class="p-4">
          <label class="block text-sm font-medium text-gray-700 mb-2">שם חדש:</label>
          <input type="text" id="renameInput" value="${nameWithoutExt}" class="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-lg">
        </div>
        <div class="p-4 border-t border-gray-200 flex gap-3 justify-end">
          <button id="cancelRenameBtn" class="px-6 py-2 bg-gray-200 text-gray-700 font-bold rounded-lg hover:bg-gray-300 transition-all">
            ביטול
          </button>
          <button id="confirmRenameBtn" class="px-6 py-2 bg-gradient-to-r from-blue-500 to-indigo-600 text-white font-bold rounded-lg hover:from-blue-600 hover:to-indigo-700 transition-all">
            ✓ שמור
          </button>
        </div>
      </div>
    `;
    
    document.body.appendChild(modal);
    
    const input = document.getElementById('renameInput');
    input.focus();
    input.select();
    
    document.getElementById('cancelRenameBtn').addEventListener('click', () => {
      modal.remove();
    });
    
    document.getElementById('confirmRenameBtn').addEventListener('click', async () => {
      let newName = input.value.trim();
      if (!newName) {
        showStatus('יש להזין שם', true);
        return;
      }
      if (!newName.endsWith('.json')) {
        newName += '.json';
      }
      
      modal.remove();
      await renameProjectInDrive(fileId, newName);
    });
    
    modal.addEventListener('click', (e) => {
      if (e.target === modal) modal.remove();
    });
    
    input.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        document.getElementById('confirmRenameBtn').click();
      }
    });
  }
  
  /**
   * Rename project in Google Drive
   */
  async function renameProjectInDrive(fileId, newName) {
    if (!isSignedIn()) {
      showStatus('יש להתחבר ל-Google תחילה', true);
      return false;
    }
    
    try {
      showStatus('משנה שם...');
      
      await gapi.client.drive.files.update({
        fileId: fileId,
        resource: {
          name: newName
        }
      });
      
      // Update current project name if this is the active project
      if (window.currentProjectFileName && window.currentProjectFileName === fileId) {
        window.currentProjectFileName = newName;
      }
      
      showStatus(`השם שונה ל"${newName.replace('.json', '')}" ✓`);
      clearProjectsCache();
      await showProjectsBanner();
      return true;
    } catch (error) {
      console.error('Error renaming file:', error);
      showStatus('שגיאה בשינוי השם', true);
      return false;
    }
  }
  
  /**
   * Switch to the correct tab
   */
  function switchToTab(projectType) {
    if (projectType === 'names') {
      const tabNamesLottery = document.getElementById('tabNamesLottery');
      if (tabNamesLottery) tabNamesLottery.click();
    } else {
      const tabWords = document.getElementById('tabWords');
      if (tabWords) tabWords.click();
    }
  }
  
  /**
   * Handle connection flow based on content state
   * מצב 1: אין תוכן - מציג הכל כרגיל, מוכן לעבודה
   * מצב 2: יש תוכן - שואל שם לפרויקט
   */
  async function handleConnectionFlow() {
    // Check if there's content in the stickers page
    let hasContent = false;
    
    if (typeof window.getProjectData === 'function') {
      const projectData = window.getProjectData();
      if (projectData && projectData.stickers && projectData.stickers.length > 0) {
        hasContent = true;
      }
    }
    
    // Show projects banner first
    await showProjectsBanner();
    
    if (hasContent) {
      // מצב 2: יש תוכן - שואל שם לפרויקט
      await showProjectNameDialog();
    } else {
      // מצב 1: אין תוכן - מציג הכל כרגיל, מוכן לעבודה
      showContentBelowBanner();
      showStatus('מחובר ל-Google Drive ✓ התחל לעבוד או בחר פרויקט קיים');
    }
  }
  
  /**
   * Hide all content below the Drive banner
   */
  function hideContentBelowBanner() {
    // Hide the header section (download/print buttons)
    const headerSection = document.querySelector('.no-print.mb-8.bg-cyan-50');
    if (headerSection) {
      headerSection.classList.add('hidden');
      headerSection.dataset.hiddenByDrive = 'true';
    }
    
    // Hide the main content area (tools + preview)
    const wordsContent = document.getElementById('wordsContent');
    const numbersContent = document.getElementById('numbersContent');
    const namesContent = document.getElementById('namesContent');
    const previewSection = document.getElementById('previewSection');
    const emptyState = document.getElementById('emptyState');
    
    if (wordsContent) {
      wordsContent.classList.add('hidden');
      wordsContent.dataset.hiddenByDrive = 'true';
    }
    if (numbersContent) {
      numbersContent.classList.add('hidden');
      numbersContent.dataset.hiddenByDrive = 'true';
    }
    if (namesContent) {
      namesContent.classList.add('hidden');
      namesContent.dataset.hiddenByDrive = 'true';
    }
    if (previewSection) {
      previewSection.classList.add('hidden');
      previewSection.dataset.hiddenByDrive = 'true';
    }
    if (emptyState) {
      emptyState.classList.add('hidden');
      emptyState.dataset.hiddenByDrive = 'true';
    }
    
    console.log('Content hidden below banner - waiting for project selection');
  }
  
  /**
   * Show all content below the Drive banner
   */
  function showContentBelowBanner() {
    // Show all elements that were hidden by Drive
    const hiddenElements = document.querySelectorAll('[data-hidden-by-drive="true"]');
    hiddenElements.forEach(el => {
      el.classList.remove('hidden');
      delete el.dataset.hiddenByDrive;
    });
    
    // Always show the header section (download/print buttons)
    const headerSection = document.querySelector('.no-print.mb-8.bg-cyan-50');
    if (headerSection) {
      headerSection.classList.remove('hidden');
    }
    
    // Show content based on active tab
    const tabWords = document.getElementById('tabWords');
    const tabNamesLottery = document.getElementById('tabNamesLottery');
    const tabNumbers = document.getElementById('tabNumbers');
    
    const wordsContent = document.getElementById('wordsContent');
    const namesContent = document.getElementById('namesContent');
    const numbersContent = document.getElementById('numbersContent');
    const previewSection = document.getElementById('previewSection');
    const emptyState = document.getElementById('emptyState');
    
    // Hide all content first
    if (wordsContent) wordsContent.classList.add('hidden');
    if (namesContent) namesContent.classList.add('hidden');
    if (numbersContent) numbersContent.classList.add('hidden');
    
    // Show the correct content based on active tab
    if (tabWords && tabWords.classList.contains('active')) {
      if (wordsContent) wordsContent.classList.remove('hidden');
      // Show preview or empty state for stickers
      if (previewSection) previewSection.classList.remove('hidden');
      if (emptyState) emptyState.classList.remove('hidden');
    } else if (tabNamesLottery && tabNamesLottery.classList.contains('active')) {
      if (namesContent) namesContent.classList.remove('hidden');
    } else if (tabNumbers && tabNumbers.classList.contains('active')) {
      if (numbersContent) numbersContent.classList.remove('hidden');
    } else {
      // Default to words/stickers tab
      if (wordsContent) wordsContent.classList.remove('hidden');
      if (previewSection) previewSection.classList.remove('hidden');
      if (emptyState) emptyState.classList.remove('hidden');
    }
    
    console.log('Content shown below banner');
  }
  
  /**
   * Show dialog to ask for project name when user has content
   */
  async function showProjectNameDialog() {
    return new Promise((resolve) => {
      const modal = document.createElement('div');
      modal.id = 'projectNameModal';
      modal.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50';
      modal.innerHTML = `
        <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full mx-4">
          <div class="p-6 border-b border-gray-200">
            <h3 class="text-xl font-bold text-gray-800 flex items-center gap-2">
              <span class="text-2xl">💾</span>
              ${t('saveProject')}
            </h3>
          </div>
          <div class="p-6 space-y-4">
            <p class="text-gray-600">${t('enterProjectName')}</p>
            <div>
              <label class="block text-sm font-medium text-gray-700 mb-2">${t('projectPlaceholder')}:</label>
              <input type="text" id="projectNameInput" placeholder="${t('projectPlaceholder')}..." class="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 text-lg">
            </div>
            <p class="text-sm text-gray-500">💡 ${t('saveToDrive')}</p>
          </div>
          <div class="p-4 border-t border-gray-200 flex gap-3 justify-end">
            <button id="cancelProjectNameBtn" class="px-6 py-3 bg-gray-200 text-gray-700 font-bold rounded-lg hover:bg-gray-300 transition-all">
              ${t('cancel')}
            </button>
            <button id="confirmProjectNameBtn" class="px-6 py-3 bg-gradient-to-r from-indigo-500 to-purple-600 text-white font-bold rounded-lg hover:from-indigo-600 hover:to-purple-700 transition-all">
              ✓ ${t('save')}
            </button>
          </div>
        </div>
      `;
      
      document.body.appendChild(modal);
      
      const nameInput = document.getElementById('projectNameInput');
      nameInput.focus();
      
      // Handle Enter key
      nameInput.addEventListener('keypress', async (e) => {
        if (e.key === 'Enter') {
          await handleProjectNameConfirm(nameInput.value, modal, resolve);
        }
      });
      
      document.getElementById('cancelProjectNameBtn').addEventListener('click', () => {
        modal.remove();
        // ביטול - מסתיר הכל ומנקה את העבודה
        hideContentBelowBanner();
        
        // Clear current work
        if (typeof clearProject === 'function') {
          clearProject();
        } else {
          if (typeof stickers !== 'undefined') {
            stickers.length = 0;
          }
          if (typeof renderStickers === 'function') {
            renderStickers();
          }
        }
        
        window.currentProjectFileName = null;
        showStatus('העבודה בוטלה. בחר פרויקט קיים או צור חדש');
        resolve(false);
      });
      
      document.getElementById('confirmProjectNameBtn').addEventListener('click', async () => {
        await handleProjectNameConfirm(nameInput.value, modal, resolve);
      });
    });
  }
  
  /**
   * Handle project name confirmation
   */
  async function handleProjectNameConfirm(inputValue, modal, resolve) {
    let projectName = inputValue.trim();
    const currentType = getCurrentProjectType();
    
    if (!projectName) {
      // Generate default name
      const prefix = currentType === 'names' ? 'הגרלת שמות' : 'עיצוב מדבקות';
      projectName = `${prefix}-${new Date().toLocaleDateString('he-IL')}`;
    }
    
    // Add -הגש suffix for names projects (for easy identification)
    if (currentType === 'names' && !projectName.includes('-הגש')) {
      projectName = projectName + '-הגש';
    }
    
    modal.remove();
    
    // Set project name
    window.currentProjectFileName = projectName + '.json';
    
    // Save the project (saveProjectToDrive already adds to banner)
    if (typeof window.getProjectData === 'function') {
      const projectData = window.getProjectData();
      if (projectData) {
        showStatus(`שומר פרויקט "${projectName}"...`);
        await saveProjectToDrive(projectData, window.currentProjectFileName);
      }
    }
    
    // Content stays visible
    showContentBelowBanner();
    resolve(true);
  }
  
  /**
   * Hide projects banner
   */
  function hideProjectsBanner() {
    const banner = document.getElementById('driveProjectsBanner');
    if (banner) {
      banner.classList.add('hidden');
    }
  }
  
  /**
   * Format date for display
   */
  function formatDate(dateString) {
    if (!dateString) return 'עכשיו';
    
    const date = new Date(dateString);
    
    // Check if date is valid
    if (isNaN(date.getTime())) {
      return 'עכשיו';
    }
    
    const now = new Date();
    const diffMs = now - date;
    const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
    
    if (diffDays === 0) {
      return 'היום';
    } else if (diffDays === 1) {
      return 'אתמול';
    } else if (diffDays < 7) {
      return `לפני ${diffDays} ימים`;
    } else {
      return date.toLocaleDateString('he-IL');
    }
  }
  
  /**
   * Get display name for a file (removes .json and _הגרלש)
   */
  function getDisplayName(fileName) {
    if (!fileName) return '';
    return fileName
      .replace('.json', '')
      .replace('_הגרלש', '');
  }
  
  /**
   * Check if a file is a names project based on filename
   */
  function isNamesProjectByFilename(fileName) {
    return fileName && fileName.includes('_הגרלש');
  }
  
  /**
   * Show new project dialog
   */
  function showNewProjectDialog() {
    const currentType = getCurrentProjectType();
    
    const modal = document.createElement('div');
    modal.id = 'newProjectModal';
    modal.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50';
    modal.innerHTML = `
      <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full mx-4">
        <div class="p-6 border-b border-gray-200">
          <h3 class="text-xl font-bold text-gray-800 flex items-center gap-2">
            <span class="text-2xl">✨</span>
            פרויקט חדש
          </h3>
        </div>
        <div class="p-6 space-y-4">
          <div>
            <label class="block text-sm font-medium text-gray-700 mb-2">סוג הפרויקט:</label>
            <div class="flex gap-3">
              <button type="button" class="project-type-btn flex-1 p-4 border-2 rounded-xl transition-all ${currentType === 'stickers' ? 'border-indigo-500 bg-indigo-50' : 'border-gray-200 hover:border-gray-300'}" data-type="stickers">
                <div class="text-2xl mb-1">🏷️</div>
                <div class="font-medium">עיצוב מדבקות</div>
              </button>
              <button type="button" class="project-type-btn flex-1 p-4 border-2 rounded-xl transition-all ${currentType === 'names' ? 'border-indigo-500 bg-indigo-50' : 'border-gray-200 hover:border-gray-300'}" data-type="names">
                <div class="text-2xl mb-1">👥</div>
                <div class="font-medium">הגרלת שמות</div>
              </button>
            </div>
          </div>
          <div>
            <label class="block text-sm font-medium text-gray-700 mb-2">שם הפרויקט:</label>
            <input type="text" id="newProjectName" placeholder="הזן שם לפרויקט..." class="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 text-lg">
          </div>
          <p class="text-sm text-gray-500">💡 הפרויקט יישמר אוטומטית ב-Google Drive שלך</p>
        </div>
        <div class="p-4 border-t border-gray-200 flex gap-3 justify-end">
          <button id="cancelNewProjectBtn" class="px-6 py-3 bg-gray-200 text-gray-700 font-bold rounded-lg hover:bg-gray-300 transition-all">
            ביטול
          </button>
          <button id="createNewProjectBtn" class="px-6 py-3 bg-gradient-to-r from-indigo-500 to-purple-600 text-white font-bold rounded-lg hover:from-indigo-600 hover:to-purple-700 transition-all">
            ✨ צור פרויקט
          </button>
        </div>
      </div>
    `;
    
    document.body.appendChild(modal);
    
    let selectedType = currentType;
    const nameInput = document.getElementById('newProjectName');
    nameInput.focus();
    
    // Type selection
    modal.querySelectorAll('.project-type-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        modal.querySelectorAll('.project-type-btn').forEach(b => {
          b.classList.remove('border-indigo-500', 'bg-indigo-50');
          b.classList.add('border-gray-200');
        });
        btn.classList.remove('border-gray-200');
        btn.classList.add('border-indigo-500', 'bg-indigo-50');
        selectedType = btn.dataset.type;
      });
    });
    
    document.getElementById('cancelNewProjectBtn').addEventListener('click', () => {
      modal.remove();
    });
    
    document.getElementById('createNewProjectBtn').addEventListener('click', async () => {
      let projectName = nameInput.value.trim();
      if (!projectName) {
        const typeLabel = selectedType === 'names' ? 'שמות' : 'מדבקות';
        projectName = `${typeLabel}-${new Date().toLocaleDateString('he-IL')}`;
      }
      
      modal.remove();
      
      // Switch to correct tab
      switchToTab(selectedType);
      
      // Show content below banner (in case it was hidden)
      showContentBelowBanner();
      
      // Clear current project
      if (selectedType === 'names') {
        if (typeof clearNamesProject === 'function') {
          clearNamesProject();
        }
      } else {
        if (typeof clearProject === 'function') {
          clearProject();
        } else {
          if (typeof stickers !== 'undefined') {
            stickers.length = 0;
          }
          if (typeof renderStickers === 'function') {
            renderStickers();
          }
          if (typeof updateFileCount === 'function') {
            updateFileCount();
          }
        }
      }
      
      // Set project name
      window.currentProjectFileName = projectName + '.json';
      
      showStatus(`פרויקט חדש "${projectName}" נוצר! התחל לעבוד`);
      
      // Refresh banner only if there are stickers (actual content)
      if (isSignedIn()) {
        setTimeout(async () => {
          // Check if there's actual content before refreshing
          let hasContent = false;
          
          if (selectedType === 'names') {
            // For names projects, check if there are names
            if (typeof getNamesProjectData === 'function') {
              const namesData = getNamesProjectData();
              hasContent = namesData && namesData.names && namesData.names.length > 0;
            }
          } else {
            // For stickers projects, check if there are stickers
            if (typeof window !== 'undefined' && window.stickers && Array.isArray(window.stickers)) {
              hasContent = window.stickers.length > 0;
            } else if (typeof getProjectData === 'function') {
              const projectData = getProjectData();
              hasContent = projectData && projectData.stickers && projectData.stickers.length > 0;
            }
          }
          
          // Only refresh if there's actual content
          if (hasContent) {
            await refreshProjectsBanner();
          }
        }, 500);
      }
    });
    
    modal.addEventListener('click', (e) => {
      if (e.target === modal) modal.remove();
    });
    
    nameInput.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        document.getElementById('createNewProjectBtn').click();
      }
    });
  }
  
  /**
   * Refresh projects banner
   */
  async function refreshProjectsBanner() {
    if (isSignedIn()) {
      showStatus('מרענן רשימת פרויקטים...');
      
      // Clear cache completely to force fresh data from Drive
      projectsCache.data = null;
      projectsCache.timestamp = 0;
      
      await showProjectsBanner();
      
      showStatus('הרשימה עודכנה ✓');
      setTimeout(() => {
        const statusEl = document.getElementById('statusMessage');
        if (statusEl && statusEl.textContent.includes('הרשימה עודכנה')) {
          statusEl.classList.add('hidden');
        }
      }, 2000);
    }
  }

  // Flag to prevent auto-save during project deletion
  let isDeleting = false;
  
  // Flag to prevent auto-save during background removal
  let isProcessingBackground = false;

  /**
   * Mark that there are unsaved changes - triggers auto-save
   */
  function markUnsavedChanges() {
    if (!isSignedIn()) {
      console.log('Auto-save: Not signed in');
      return;
    }
    
    // Don't auto-save during deletion
    if (isDeleting) {
      console.log('Auto-save: Skipping - deletion in progress');
      return;
    }
    
    // Don't auto-save during background removal
    if (isProcessingBackground) {
      console.log('Auto-save: Skipping - background removal in progress');
      return;
    }
    
    // Check for project file name - if not set, don't auto-save
    // User must explicitly create a project first
    const projectFileName = window.currentProjectFileName;
    if (!projectFileName) {
      console.log('Auto-save: No project yet, skipping (user must create project first)');
      return;
    }
    
    hasUnsavedChanges = true;
    
    // Clear existing timer
    if (autoSaveTimer) {
      clearTimeout(autoSaveTimer);
    }
    
    // Clear existing countdown interval
    if (window.autoSaveCountdownInterval) {
      clearInterval(window.autoSaveCountdownInterval);
    }
    
    // Start countdown from 10 seconds
    let secondsLeft = 10;
    updateAutoSaveIndicator('unsaved', secondsLeft);
    
    window.autoSaveCountdownInterval = setInterval(() => {
      secondsLeft--;
      if (secondsLeft > 0) {
        updateAutoSaveIndicator('unsaved', secondsLeft);
      } else {
        clearInterval(window.autoSaveCountdownInterval);
      }
    }, 1000);
    
    // Set new timer - save after 10 seconds of inactivity
    autoSaveTimer = setTimeout(async () => {
      if (window.autoSaveCountdownInterval) {
        clearInterval(window.autoSaveCountdownInterval);
      }
      await performAutoSave();
    }, 10000);
  }
  
  /**
   * Create a new project on first change (when no project exists yet)
   */
  async function createNewProjectOnFirstChange() {
    const currentType = getCurrentProjectType();
    
    // Get current project data based on type
    let projectData = null;
    
    if (currentType === 'names' && typeof window.getNamesProjectData === 'function') {
      projectData = window.getNamesProjectData();
    } else if (typeof window.getProjectData === 'function') {
      projectData = window.getProjectData();
    }
    
    if (!projectData) {
      projectData = {
        stickers: [],
        projectType: currentType,
        settings: {}
      };
    }
    
    // Ensure projectType is set
    projectData.projectType = currentType;
    
    // Capture thumbnail for stickers projects
    if (currentType === 'stickers') {
      try {
        const thumbnail = await captureProjectThumbnail();
        if (thumbnail) {
          projectData.thumbnail = thumbnail;
        }
      } catch (e) {
        console.log('Could not capture thumbnail:', e);
      }
    }
    
    // Create the project
    await createNewProjectWithNumber(projectData);
  }
  
  /**
   * Perform auto-save
   */
  async function performAutoSave() {
    if (!isSignedIn() || !hasUnsavedChanges) {
      console.log('Auto-save: Skipping - not signed in or no changes');
      return;
    }
    
    const projectFileName = window.currentProjectFileName;
    if (!projectFileName) {
      console.log('Auto-save: No project file name set');
      hasUnsavedChanges = false;
      const indicator = document.getElementById('autoSaveIndicator');
      if (indicator) indicator.style.opacity = '0';
      return;
    }
    
    // Get project data based on current type
    const currentType = getCurrentProjectType();
    let projectData = null;
    
    if (currentType === 'names' && typeof window.getNamesProjectData === 'function') {
      projectData = window.getNamesProjectData();
    } else if (typeof window.getProjectData === 'function') {
      projectData = window.getProjectData();
    }
    
    if (!projectData) {
      console.log('Auto-save: No project data to save (empty project)');
      hasUnsavedChanges = false;
      const indicator = document.getElementById('autoSaveIndicator');
      if (indicator) indicator.style.opacity = '0';
      return;
    }
    
    // Ensure projectType is set
    projectData.projectType = currentType;
    
    // Capture thumbnail for stickers projects
    if (currentType === 'stickers') {
      try {
        const thumbnail = await captureProjectThumbnail();
        if (thumbnail) {
          projectData.thumbnail = thumbnail;
        }
      } catch (e) {
        console.log('Could not capture thumbnail:', e);
      }
    }
    
    // Check if data actually changed (excluding thumbnail for comparison)
    const dataForComparison = { ...projectData };
    delete dataForComparison.thumbnail;
    const currentDataStr = JSON.stringify(dataForComparison);
    if (currentDataStr === lastSavedData) {
      console.log('Auto-save: Data unchanged, skipping');
      hasUnsavedChanges = false;
      const indicator = document.getElementById('autoSaveIndicator');
      if (indicator) indicator.style.opacity = '0';
      return;
    }
    
    updateAutoSaveIndicator('saving');
    console.log('Auto-save: Saving to', projectFileName);
    
    try {
      const result = await saveProjectToDriveQuiet(projectData, projectFileName);
      if (result) {
        hasUnsavedChanges = false;
        lastSavedData = currentDataStr;
        updateAutoSaveIndicator('saved');
        console.log('Auto-save: Completed successfully');
        
        // Check if this project already exists in banner
        const existingCard = document.querySelector(`[data-file-name="${projectFileName}"]`);
        if (!existingCard && result.id) {
          // Add new project to banner
          const projectType = projectData.projectType || getCurrentProjectType();
          await addProjectToBanner({
            id: result.id,
            name: projectFileName,
            modifiedTime: new Date().toISOString()
          }, projectType);
        }
      } else {
        console.log('Auto-save: Save returned null');
        updateAutoSaveIndicator('error');
      }
    } catch (error) {
      console.error('Auto-save error:', error);
      updateAutoSaveIndicator('error');
    }
  }
  
  /**
   * Capture a thumbnail of the current project preview
   */
  async function captureProjectThumbnail() {
    const preview = document.getElementById('printPreviewInner');
    if (!preview || !preview.children.length) return null;
    
    try {
      // Use html2canvas to capture the preview
      const canvas = await html2canvas(preview, {
        scale: 0.3, // Small scale for thumbnail
        useCORS: true,
        allowTaint: true,
        backgroundColor: '#ffffff',
        width: Math.min(preview.scrollWidth, 800),
        height: Math.min(preview.scrollHeight, 600)
      });
      
      // Convert to small JPEG
      const thumbnail = canvas.toDataURL('image/jpeg', 0.5);
      return thumbnail;
    } catch (error) {
      console.error('Error capturing thumbnail:', error);
      return null;
    }
  }
  
  /**
   * Save project to Google Drive quietly (no status messages)
   */
  async function saveProjectToDriveQuiet(projectData, fileName) {
    if (!isSignedIn()) return null;
    
    if (!appFolderId) {
      await getOrCreateAppFolder();
    }
    
    try {
      // Add project type to data
      const dataToSave = {
        ...projectData,
        projectType: projectData.projectType || getCurrentProjectType()
      };
      
      const fileContent = JSON.stringify(dataToSave, null, 2);
      const file = new Blob([fileContent], { type: 'application/json' });
      
      const token = gapi.client.getToken().access_token;
      
      // Check if file already exists
      const existingFile = await findFileByName(fileName);
      let actualFileName = fileName;
      
      if (existingFile) {
        // Try to update existing file
        const form = new FormData();
        form.append('metadata', new Blob([JSON.stringify({ name: fileName, mimeType: 'application/json' })], { type: 'application/json' }));
        form.append('file', file);
        
        const response = await fetch(
          `https://www.googleapis.com/upload/drive/v3/files/${existingFile.id}?uploadType=multipart`,
          {
            method: 'PATCH',
            headers: { 'Authorization': `Bearer ${token}` },
            body: form
          }
        );
        
        if (response.ok) {
          return await response.json();
        }
        
        // If 403, try to delete old file and create new one
        if (response.status === 403) {
          console.log('Auto-save: No permission to update, trying to delete and recreate');
          let deletedOldFile = false;
          try {
            await gapi.client.drive.files.delete({ fileId: existingFile.id });
            deletedOldFile = true;
            console.log('Auto-save: Old file deleted successfully');
          } catch (e) {
            console.log('Auto-save: Could not delete old file, will create with timestamp');
            // Add timestamp to filename to make it unique
            const baseName = fileName.replace('.json', '');
            const timestamp = new Date().toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' }).replace(':', '-');
            actualFileName = `${baseName}_${timestamp}.json`;
          }
        }
      }
      
      // Create new file
      const metadata = {
        name: actualFileName,
        mimeType: 'application/json',
        parents: appFolderId ? [appFolderId] : []
      };
      
      const form = new FormData();
      form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
      form.append('file', file);
      
      const response = await fetch(
        'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
        {
          method: 'POST',
          headers: { 'Authorization': `Bearer ${token}` },
          body: form
        }
      );
      
      if (response.ok) {
        const result = await response.json();
        // Update the project file name if it changed
        if (actualFileName !== fileName) {
          window.currentProjectFileName = actualFileName;
        }
        return result;
      }
      
      console.error('Failed to create file:', response.status);
      return null;
    } catch (error) {
      console.error('Error in quiet save:', error);
      return null;
    }
  }
  
  /**
   * Update auto-save indicator
   */
  function updateAutoSaveIndicator(status, secondsLeft = null) {
    let indicator = document.getElementById('autoSaveIndicator');
    
    if (!indicator) {
      indicator = document.createElement('div');
      indicator.id = 'autoSaveIndicator';
      indicator.className = 'fixed bottom-1 left-24 px-1.5 py-0.5 rounded text-[10px] font-medium z-40 transition-all';
      document.body.appendChild(indicator);
    }
    
    // ביטול טיימר קודם אם קיים
    if (indicator.hideTimer) {
      clearTimeout(indicator.hideTimer);
    }
    
    switch (status) {
      case 'unsaved':
        indicator.className = 'fixed bottom-1 left-24 px-1.5 py-0.5 rounded text-[10px] font-medium z-40 bg-yellow-100 text-yellow-700 border border-yellow-300';
        if (secondsLeft !== null) {
          indicator.textContent = `● לא נשמר (${secondsLeft}s)`;
        } else {
          indicator.textContent = '● לא נשמר';
        }
        indicator.style.opacity = '1';
        break;
      case 'saving':
        indicator.className = 'fixed bottom-1 left-24 px-1.5 py-0.5 rounded text-[10px] font-medium z-40 bg-blue-100 text-blue-700 border border-blue-300';
        indicator.textContent = '⏳ שומר...';
        indicator.style.opacity = '1';
        break;
      case 'saved':
        indicator.className = 'fixed bottom-1 left-24 px-1.5 py-0.5 rounded text-[10px] font-medium z-40 bg-green-100 text-green-700 border border-green-300';
        indicator.textContent = '✓ נשמר';
        indicator.style.opacity = '1';
        // הסתרה אוטומטית אחרי 3 שניות
        indicator.hideTimer = setTimeout(() => {
          indicator.style.opacity = '0';
        }, 3000);
        break;
      case 'error':
        indicator.className = 'fixed bottom-1 left-24 px-1.5 py-0.5 rounded text-[10px] font-medium z-40 bg-red-100 text-red-700 border border-red-300';
        indicator.textContent = '✗ שגיאה';
        indicator.style.opacity = '1';
        // הסתרה אוטומטית אחרי 5 שניות
        indicator.hideTimer = setTimeout(() => {
          indicator.style.opacity = '0';
        }, 5000);
        break;
    }
  }
  
  /**
   * Get current project type based on active tab
   */
  function getCurrentProjectType() {
    const tabWords = document.getElementById('tabWords');
    const tabNamesLottery = document.getElementById('tabNamesLottery');
    
    if (tabWords && tabWords.classList.contains('active')) {
      return 'stickers';
    } else if (tabNamesLottery && tabNamesLottery.classList.contains('active')) {
      return 'names';
    }
    return 'stickers';
  }
  
  /**
   * Auto-save current work when signing in - only if there's content
   */
  async function autoSaveCurrentWork() {
    // Check if there's already a project name set
    if (window.currentProjectFileName) {
      console.log('Auto-save on login: Project already has a name');
      return;
    }
    
    // Check if there's content to save
    let hasContent = false;
    let projectData = null;
    
    if (typeof window.getProjectData === 'function') {
      projectData = window.getProjectData();
      if (projectData && projectData.stickers && projectData.stickers.length > 0) {
        hasContent = true;
      }
    }
    
    // If no content, don't create a project yet - wait for first sticker upload
    if (!hasContent) {
      console.log('Auto-save on login: No content yet, will create project on first change');
      return;
    }
    
    // Generate project name - find next available number
    await createNewProjectWithNumber(projectData);
  }
  
  /**
   * Create a new project with auto-numbered name
   */
  async function createNewProjectWithNumber(projectData) {
    const files = await listProjectsFromDrive();
    const currentType = projectData.projectType || getCurrentProjectType();
    
    // Different naming based on type
    const prefix = currentType === 'names' ? 'הגרלת שמות' : 'עיצוב מדבקות';
    let projectNumber = 1;
    
    // Find the highest number for this type
    files.forEach(file => {
      const match = file.name.match(new RegExp(`${prefix}\\s*(\\d+)`));
      if (match) {
        const num = parseInt(match[1], 10);
        if (num >= projectNumber) {
          projectNumber = num + 1;
        }
      }
    });
    
    const projectName = `${prefix} ${projectNumber}.json`;
    
    console.log('Creating new project:', projectName);
    showStatus(`יוצר פרויקט "${projectName.replace('.json', '')}"...`);
    
    // Save the project
    const result = await saveProjectToDrive(projectData, projectName);
    
    if (result) {
      // Set the project name for future auto-saves
      window.currentProjectFileName = projectName;
      showStatus(`פרויקט "${projectName.replace('.json', '')}" נוצר ✓`);
      
      // Add the new project to banner without refreshing everything
      const projectType = projectData.projectType || getCurrentProjectType();
      await addProjectToBanner(result, projectType);
    }
    
    return result;
  }

  /**
   * Get or create the app folder in Google Drive
   */
  async function getOrCreateAppFolder() {
    try {
      // Search for existing folder
      const response = await gapi.client.drive.files.list({
        q: `name='${APP_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
        fields: 'files(id, name)',
        spaces: 'drive'
      });
      
      if (response.result.files && response.result.files.length > 0) {
        appFolderId = response.result.files[0].id;
        console.log('Found existing app folder:', appFolderId);
      } else {
        // Create new folder
        const createResponse = await gapi.client.drive.files.create({
          resource: {
            name: APP_FOLDER_NAME,
            mimeType: 'application/vnd.google-apps.folder'
          },
          fields: 'id'
        });
        appFolderId = createResponse.result.id;
        console.log('Created new app folder:', appFolderId);
      }
      
      return appFolderId;
    } catch (error) {
      console.error('Error with app folder:', error);
      return null;
    }
  }
  
  /**
   * Generate unique filename by checking existing files
   */
  async function generateUniqueFileName(baseName) {
    if (!baseName.endsWith('.json')) {
      baseName += '.json';
    }
    
    // Get list of existing files
    const existingFiles = await listProjectsFromDrive();
    const existingNames = existingFiles.map(f => f.name.toLowerCase());
    
    // If base name doesn't exist, use it
    if (!existingNames.includes(baseName.toLowerCase())) {
      return baseName;
    }
    
    // Extract name without extension
    const nameWithoutExt = baseName.replace('.json', '');
    
    // Try numbered versions
    let counter = 2;
    let uniqueName;
    
    do {
      uniqueName = `${nameWithoutExt} (${counter}).json`;
      counter++;
    } while (existingNames.includes(uniqueName.toLowerCase()) && counter < 100);
    
    return uniqueName;
  }

  /**
   * Save project to Google Drive
   */
  async function saveProjectToDrive(projectData, fileName, projectType) {
    if (!isSignedIn()) {
      showStatus('יש להתחבר ל-Google תחילה', true);
      return null;
    }
    
    if (!appFolderId) {
      await getOrCreateAppFolder();
    }
    
    try {
      // Add project type to data
      const dataToSave = {
        ...projectData,
        projectType: projectType || getCurrentProjectType()
      };
      
      showStatus('בודק שמות קיימים...');
      
      const fileContent = JSON.stringify(dataToSave, null, 2);
      const file = new Blob([fileContent], { type: 'application/json' });
      let actualFileName = fileName || `פרויקט-${new Date().toLocaleDateString('he-IL')}.json`;
      
      // For names projects, add _הגרלש suffix for identification
      const projectTypeToSave = dataToSave.projectType || getCurrentProjectType();
      if (projectTypeToSave === 'names') {
        // Remove .json if exists, add _הגרלש, then add .json back
        actualFileName = actualFileName.replace('.json', '');
        if (!actualFileName.includes('_הגרלש')) {
          actualFileName = actualFileName + '_הגרלש';
        }
        actualFileName = actualFileName + '.json';
      }
      
      // Generate unique filename to avoid conflicts
      actualFileName = await generateUniqueFileName(actualFileName);
      
      showStatus(t('savingToDrive'));
      
      const token = gapi.client.getToken().access_token;
      
      // Create new file with unique name
      const metadata = {
        name: actualFileName,
        mimeType: 'application/json',
        parents: appFolderId ? [appFolderId] : []
      };
      
      const form = new FormData();
      form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
      form.append('file', file);
      
      const response = await fetch(
        'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
        {
          method: 'POST',
          headers: { 'Authorization': `Bearer ${token}` },
          body: form
        }
      );
      
      if (response.ok) {
        const result = await response.json();
        const displayName = actualFileName.replace('.json', '');
        showStatus(`הפרויקט נשמר כ"${displayName}" ✓`);
        
        // Update current project filename
        window.currentProjectFileName = actualFileName;
        
        // Add the new project to banner without refreshing everything
        // Make sure to include proper modifiedTime
        const projectType = dataToSave.projectType || getCurrentProjectType();
        const projectForBanner = {
          ...result,
          name: actualFileName,
          modifiedTime: result.modifiedTime || new Date().toISOString()
        };
        await addProjectToBanner(projectForBanner, projectType);
        
        return result;
      } else {
        throw new Error('Failed to save file');
      }
    } catch (error) {
      console.error('Error saving to Drive:', error);
      showStatus(t('driveSaveError'), true);
      return null;
    }
  }
  
  /**
   * Find file by name in app folder
   */
  async function findFileByName(fileName) {
    try {
      let query = `name='${fileName}' and trashed=false`;
      if (appFolderId) {
        query += ` and '${appFolderId}' in parents`;
      }
      
      const response = await gapi.client.drive.files.list({
        q: query,
        fields: 'files(id, name, modifiedTime)',
        spaces: 'drive'
      });
      
      if (response.result.files && response.result.files.length > 0) {
        return response.result.files[0];
      }
      return null;
    } catch (error) {
      console.error('Error finding file:', error);
      return null;
    }
  }

  /**
   * List all project files from Google Drive
   */
  async function listProjectsFromDrive() {
    if (!isSignedIn()) {
      showStatus('יש להתחבר ל-Google תחילה', true);
      return [];
    }
    
    if (!appFolderId) {
      await getOrCreateAppFolder();
    }
    
    try {
      let query = `mimeType='application/json' and trashed=false`;
      if (appFolderId) {
        query += ` and '${appFolderId}' in parents`;
      }
      
      const response = await gapi.client.drive.files.list({
        q: query,
        fields: 'files(id, name, modifiedTime, size)',
        orderBy: 'modifiedTime desc',
        spaces: 'drive'
      });
      
      return response.result.files || [];
    } catch (error) {
      console.error('Error listing files:', error);
      showStatus('שגיאה בטעינת רשימת הקבצים', true);
      return [];
    }
  }
  
  /**
   * Load project from Google Drive by file ID
   */
  async function loadProjectFromDrive(fileId) {
    if (!isSignedIn()) {
      showStatus('יש להתחבר ל-Google תחילה', true);
      return null;
    }
    
    try {
      showStatus(t('loadingProject'));
      
      const response = await gapi.client.drive.files.get({
        fileId: fileId,
        alt: 'media'
      });
      
      const projectData = response.result;
      showStatus(t('projectLoadedSuccess', { name: '' })); // No name here
      return projectData;
    } catch (error) {
      console.error('Error loading from Drive:', error);
      showStatus(t('projectLoadError'), true);
      return null;
    }
  }
  
  /**
   * Remove a specific project card from the banner without reloading everything
   */
  function removeProjectCardFromBanner(fileId, fileName) {
    const card = document.querySelector(`[data-file-id="${fileId}"]`);
    if (card) {
      // Check if this is the current project being deleted
      const cardFileName = card.dataset.fileName;
      const isCurrentProject = window.currentProjectFileName === cardFileName;
      
      // Add fade out animation
      card.style.transition = 'all 0.3s ease';
      card.style.opacity = '0';
      card.style.transform = 'scale(0.8)';
      
      // Remove after animation
      setTimeout(() => {
        card.remove();
        
        // Update cache to remove the deleted project
        if (projectsCache.data) {
          projectsCache.data = projectsCache.data.filter(p => p.id !== fileId);
        }
        
        // Check if section is now empty and update headers
        updateSectionHeaders();
        
        // If this was the current project, clear everything and collapse
        if (isCurrentProject) {
          // Set deletion flag to prevent auto-save from creating new project
          isDeleting = true;
          
          // IMPORTANT: Clear project name FIRST to stop auto-save
          window.currentProjectFileName = null;
          
          // Cancel any pending auto-save
          if (autoSaveTimer) {
            clearTimeout(autoSaveTimer);
            autoSaveTimer = null;
          }
          hasUnsavedChanges = false;
          
          // Hide auto-save indicator
          const indicator = document.getElementById('autoSaveIndicator');
          if (indicator) indicator.style.opacity = '0';
          
          // Clear current work
          if (typeof clearProject === 'function') {
            clearProject();
          } else {
            if (typeof window.stickers !== 'undefined') {
              window.stickers.length = 0;
            }
            if (typeof renderStickers === 'function') {
              renderStickers();
            }
          }
          
          // Hide content below banner (collapse to initial state)
          hideContentBelowBanner();
          
          // Reset deletion flag after a delay
          setTimeout(() => {
            isDeleting = false;
          }, 500);
          
          showStatus('הפרויקט נמחק. בחר פרויקט קיים או צור חדש');
        }
      }, 300);
    } else {
      // Card not found, just update cache
      if (projectsCache.data) {
        projectsCache.data = projectsCache.data.filter(p => p.id !== fileId);
      }
    }
  }
  
  /**
   * Update section headers after project removal
   */
  function updateSectionHeaders() {
    const container = document.getElementById('projectsContainer');
    if (!container) return;
    
    // Count remaining projects
    const allCards = container.querySelectorAll('[data-project-type]');
    
    // Update the single header counter
    const counterElement = container.querySelector('.text-purple-300');
    if (counterElement) {
      counterElement.textContent = `${allCards.length} פרויקטים`;
    }
    
    // If no projects left, show empty message
    if (allCards.length === 0) {
      const emptyMsg = document.createElement('div');
      emptyMsg.className = 'text-white/70 text-sm py-8 px-4';
      emptyMsg.innerHTML = '💡 אין פרויקטים שמורים עדיין. צור פרויקט חדש!';
      container.appendChild(emptyMsg);
    }
  }

  /**
   * Delete project from Google Drive
   */
  async function deleteProjectFromDrive(fileId) {
    if (!isSignedIn()) {
      showStatus('יש להתחבר ל-Google תחילה', true);
      return false;
    }
    
    try {
      await gapi.client.drive.files.delete({
        fileId: fileId
      });
      showStatus(t('projectDeletedSuccess'));
      
      // Clear cache to force reload from Drive next time
      clearProjectsCache();
      
      // Remove the card from banner immediately without reloading
      removeProjectCardFromBanner(fileId);
      
      return true;
    } catch (error) {
      console.error('Error deleting file:', error);
      showStatus(t('projectDeleteError'), true);
      return false;
    }
  }
  
  /**
   * Show file picker modal
   */
  async function showDriveFilePicker() {
    const files = await listProjectsFromDrive();
    
    if (files.length === 0) {
      showStatus('לא נמצאו פרויקטים שמורים ב-Google Drive');
      return;
    }
    
    // Create modal
    const modal = document.createElement('div');
    modal.id = 'driveFilePickerModal';
    modal.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50';
    modal.innerHTML = `
      <div class="bg-white rounded-2xl shadow-2xl max-w-lg w-full mx-4 max-h-[80vh] flex flex-col">
        <div class="p-4 border-b border-gray-200 flex justify-between items-center">
          <h3 class="text-xl font-bold text-gray-800">📂 בחר פרויקט מ-Google Drive</h3>
          <button id="closeDrivePickerBtn" class="text-gray-500 hover:text-gray-700 text-2xl">&times;</button>
        </div>
        <div class="p-4 overflow-y-auto flex-1">
          <div id="driveFilesList" class="space-y-2">
            ${files.map(file => `
              <div class="drive-file-item p-3 border-2 border-gray-200 rounded-lg hover:border-blue-400 hover:bg-blue-50 cursor-pointer transition-all" data-file-id="${file.id}" data-file-name="${file.name}">
                <div class="flex justify-between items-center">
                  <div>
                    <div class="font-medium text-gray-800">${file.name}</div>
                    <div class="text-sm text-gray-500">${new Date(file.modifiedTime).toLocaleString('he-IL')}</div>
                  </div>
                  <button class="delete-drive-file text-red-500 hover:text-red-700 p-2" data-file-id="${file.id}" title="מחק">🗑️</button>
                </div>
              </div>
            `).join('')}
          </div>
        </div>
      </div>
    `;
    
    document.body.appendChild(modal);
    
    // Event listeners
    document.getElementById('closeDrivePickerBtn').addEventListener('click', () => {
      modal.remove();
    });
    
    modal.addEventListener('click', (e) => {
      if (e.target === modal) modal.remove();
    });
    
    // File selection
    modal.querySelectorAll('.drive-file-item').forEach(item => {
      item.addEventListener('click', async (e) => {
        if (e.target.closest('.delete-drive-file')) return;
        
        const fileId = item.dataset.fileId;
        const fileName = item.dataset.fileName;
        modal.remove();
        
        const projectData = await loadProjectFromDrive(fileId);
        if (projectData && typeof loadProjectData === 'function') {
          loadProjectData(projectData, fileName);
        }
      });
    });
    
    // Delete buttons
    modal.querySelectorAll('.delete-drive-file').forEach(btn => {
      btn.addEventListener('click', async (e) => {
        e.stopPropagation();
        const fileId = btn.dataset.fileId;
        
        if (confirm('האם למחוק את הקובץ מ-Google Drive?')) {
          const success = await deleteProjectFromDrive(fileId);
          if (success) {
            btn.closest('.drive-file-item').remove();
            
            // Check if list is empty
            if (modal.querySelectorAll('.drive-file-item').length === 0) {
              modal.remove();
              showStatus('אין עוד פרויקטים ב-Google Drive');
            }
          }
        }
      });
    });
  }

  /**
   * Show save dialog
   */
  function showSaveToDriverDialog(projectData) {
    const defaultName = currentProjectFileName || `פרויקט-${new Date().toLocaleDateString('he-IL')}.json`;
    
    const modal = document.createElement('div');
    modal.id = 'driveSaveModal';
    modal.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50';
    modal.innerHTML = `
      <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full mx-4">
        <div class="p-4 border-b border-gray-200">
          <h3 class="text-xl font-bold text-gray-800">💾 שמור ל-Google Drive</h3>
        </div>
        <div class="p-4">
          <label class="block text-sm font-medium text-gray-700 mb-2">שם הקובץ:</label>
          <input type="text" id="driveFileName" value="${defaultName}" class="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500">
          <p class="text-xs text-gray-500 mt-2">💡 אם השם כבר קיים, יתווסף מספר אוטומטית (לדוגמה: "פרויקט 1 (2)")</p>
        </div>
        <div class="p-4 border-t border-gray-200 flex gap-3 justify-end">
          <button id="cancelDriveSaveBtn" class="px-6 py-2 bg-gray-200 text-gray-700 font-bold rounded-lg hover:bg-gray-300 transition-all">
            ביטול
          </button>
          <button id="confirmDriveSaveBtn" class="px-6 py-2 bg-gradient-to-r from-green-500 to-emerald-600 text-white font-bold rounded-lg hover:from-green-600 hover:to-emerald-700 transition-all">
            💾 שמור
          </button>
        </div>
      </div>
    `;
    
    document.body.appendChild(modal);
    
    const fileNameInput = document.getElementById('driveFileName');
    fileNameInput.focus();
    fileNameInput.select();
    
    document.getElementById('cancelDriveSaveBtn').addEventListener('click', () => {
      modal.remove();
    });
    
    document.getElementById('confirmDriveSaveBtn').addEventListener('click', async () => {
      let fileName = fileNameInput.value.trim();
      if (!fileName) {
        fileName = defaultName;
      }
      if (!fileName.endsWith('.json')) {
        fileName += '.json';
      }
      
      modal.remove();
      await saveProjectToDrive(projectData, fileName);
    });
    
    modal.addEventListener('click', (e) => {
      if (e.target === modal) modal.remove();
    });
    
    fileNameInput.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        document.getElementById('confirmDriveSaveBtn').click();
      }
    });
  }
  
  /**
   * Set callback for sign in state changes
   */
  function setOnSignInChange(callback) {
    onSignInChange = callback;
  }
  
  /**
   * Get current user
   */
  function getCurrentUser() {
    return currentUser;
  }
  
  // Public API
  return {
    gapiLoaded,
    gisLoaded,
    signIn,
    signOut,
    isSignedIn,
    saveProjectToDrive,
    loadProjectFromDrive,
    listProjectsFromDrive,
    deleteProjectFromDrive,
    showDriveFilePicker,
    showSaveToDriverDialog,
    setOnSignInChange,
    getCurrentUser,
    showProjectsBanner,
    hideProjectsBanner,
    refreshProjectsBanner,
    showNewProjectDialog,
    markUnsavedChanges,
    getCurrentProjectType,
    showContentBelowBanner,
    hideContentBelowBanner,
    setProcessingBackground: (value) => { isProcessingBackground = value; }
  };
})();

// Global functions for script callbacks
function gapiLoaded() {
  GoogleDriveManager.gapiLoaded();
}

function gisLoaded() {
  GoogleDriveManager.gisLoaded();
}
