    // ========== Names Lottery Functions ==========
    
    let namesData = [];
    let namesNoteTemplate = null;
    let namesNotes = [];
    let namesLotteryResults = [];
    
    function handleNamesNoteImageUpload(e) {
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
          namesNoteTemplate = {
            dataUrl: event.target.result,
            fileName: file.name,
            width: img.width,
            height: img.height
          };

          document.getElementById('namesNoteFileName').textContent = `✓ ${file.name}`;
          showStatus(t('noteImageUploadedSuccess', { name: file.name }));
          
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
        showStatus(t('uploadExcelFirst'), true);
        return;
      }

      if (!namesNoteTemplate) {
        showStatus(t('uploadNoteImageFirst'), true);
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
      showStatus(t('notesCreatedSuccess', { count: namesNotes.length }));
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
        showStatus(t('noNotesToCenter'), true);
        return;
      }

      namesNotes.forEach(note => {
        note.textX = note.width / 2;
        note.textY = note.height / 2; // 50% לתצוגה באתר
      });

      renderNamesNotes();
      showStatus(t('textCenteredSuccess'));
    }

    async function downloadNamesNotesAsPDF() {
      if (namesNotes.length === 0) {
        showStatus(t('noNotesToDownload'), true);
        return;
      }

      const { jsPDF } = window.jspdf;
      showStatus(t('preparingPdf'));

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

        pdf.save(t('notesPdfFileName'));
        showStatus(t('pdfDownloadedSuccess'));
      } catch (error) {
        console.error('PDF generation error:', error);
        showStatus(t('pdfError'), true);
      }
    }

    function printNamesNotes() {
      if (namesNotes.length === 0) {
        showStatus(t('noNotesToPrint'), true);
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
        showStatus(t('noLotteryResultsToDownload'), true);
        return;
      }

      showStatus(t('preparingPdf'));

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
        container.style.direction = currentLanguage === 'he' ? 'rtl' : 'ltr';
        container.style.textAlign = currentLanguage === 'he' ? 'right' : 'left';

        // Add title
        const title = document.createElement('div');
        title.textContent = t('lotteryResultsTitle');
        title.style.fontSize = '24px';
        title.style.fontWeight = 'bold';
        title.style.textAlign = 'center';
        title.style.marginBottom = '10mm';
        title.style.borderBottom = '2px solid #333';
        title.style.paddingBottom = '5mm';
        container.appendChild(title);

        // Add date
        const dateDiv = document.createElement('div');
        const date = new Date().toLocaleDateString(currentLanguage === 'he' ? 'he-IL' : 'en-US');
        dateDiv.textContent = t('dateLabel', { date });
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
            const groupName = item.groupName || t('noGroup');
            if (!groups[groupName]) {
              groups[groupName] = [];
            }
            groups[groupName].push(item);
          });
          
          Object.keys(groups).sort().forEach((groupName, groupIndex) => {
            const groupItems = groups[groupName];
            
            // Group header
            const groupHeader = document.createElement('div');
            groupHeader.textContent = t('groupWinnersHeader', { name: groupName, count: groupItems.length });
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
        footer.textContent = t('totalResults', { count: namesLotteryResults.length });
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
        
        pdf.save(t('lotteryPdfFileName'));
        showStatus(t('pdfDownloadedSuccess'));
        
      } catch (error) {
        console.error('PDF generation error:', error);
        showStatus(t('pdfError'), true);
      }
    }

    function printNamesLottery() {
      if (namesLotteryResults.length === 0) {
        showStatus(t('noLotteryResultsToPrint'), true);
        return;
      }

      const originalContent = document.body.innerHTML;
      const resultsSection = document.getElementById('namesLotteryResultsSection');
      
      if (!resultsSection) {
        showStatus(t('printError'), true);
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
        instructions.innerHTML = t('simpleInstructions');
      } else if (type === 'groups') {
        instructions.innerHTML = t('groupsInstructions');
      } else if (type === 'tickets') {
        instructions.innerHTML = t('ticketsInstructions');
      } else if (type === 'groupsTickets') {
        instructions.innerHTML = t('groupsTicketsInstructions');
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
      const fileName = uploadType === 'simple' ? t('templateSimple') : 
                       uploadType === 'groups' ? t('templateGroups') : 
                       uploadType === 'tickets' ? t('templateTickets') :
                       t('templateGroupsTickets');
      link.download = fileName;
      link.href = URL.createObjectURL(blob);
      link.click();
      URL.revokeObjectURL(link.href);
      
      showStatus(t('pdfDownloadedSuccess')); // Or template downloaded
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
          showStatus(t('csvReadError'), true);
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
            showStatus(t('fileReadError'), true);
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
        showStatus(t('noNamesFoundInFile'), true);
        return;
      }
      
      // Update UI
      document.getElementById('namesFileInfo').textContent = `✓ ${fileName} (${t('namesLoadedCount', { count: namesData.length })})`;
      
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
          previewHtml += `<div class="font-bold text-indigo-800 text-base mb-2">${group} (${names.length} ${t('namesLotteryWinners') === 'Winners' ? 'students' : 'תלמידים'})</div>`;
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
        previewHtml += `<div class="font-bold text-indigo-800 text-base mb-2">${t('typeTickets')}</div>`;
        previewHtml += `<div class="text-gray-700 text-sm leading-relaxed">`;
        for (const [name, count] of Object.entries(students)) {
          previewHtml += `${name} (${count} ${t('namesLotteryWinners') === 'Winners' ? 'tickets' : 'כרטיסים'}), `;
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
            previewHtml += `${name} (${count} ${t('namesLotteryWinners') === 'Winners' ? 'tickets' : 'כרטיסים'}), `;
          }
          previewHtml = previewHtml.slice(0, -2); // הסר את הפסיק האחרון
          previewHtml += `</div></div>`;
        }
      } else {
        previewHtml = `<div class="text-gray-700">${namesData.map(item => item.name).join(', ')}</div>`;
      }
      previewList.innerHTML = previewHtml;
      
      showStatus(t('namesLoadedSuccess', { count: namesData.length }));
      
      // שלב 1: הפעלת רק כפתור בחר מדבקה והטאבים
      enableNamesImageButtonOnly();
    }
    
    function generateNamesLottery(mode = 'all') {
      if (namesData.length === 0) {
        showStatus(t('uploadExcelFirst'), true);
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
        showStatus(t('groupLotteryOnlyWithGroups'), true);
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
      
      showStatus(t('lotterySuccessAll', { count: winnersCount, total: namesData.length }));
    }
    
    function generateGroupNamesLottery(winnersCount, startNum) {
      // Group names by group
      const groups = {};
      namesData.forEach(item => {
        const group = item.group || t('noGroup');
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
      showStatus(t('lotterySuccessGroups', { count: winnersCount, total: totalWinners, groups: groupCount }));
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
    
    // ========== Project Save/Load Functions for Names ==========
    
    /**
     * Get names project data for saving to Drive
     */
    function getNamesProjectData() {
      // Check if there's any data to save
      if (namesData.length === 0 && namesNotes.length === 0 && namesLotteryResults.length === 0) {
        return null;
      }
      
      return {
        projectType: 'names',
        namesData: namesData,
        namesNoteTemplate: namesNoteTemplate,
        namesNotes: namesNotes,
        namesLotteryResults: namesLotteryResults,
        settings: {
          uploadType: document.getElementById('namesUploadType')?.value || 'simple',
          notesPerRow: document.getElementById('namesNotesPerRow')?.value || 4,
          notesSpacing: document.getElementById('namesNotesSpacing')?.value || 5,
          notesColor: document.getElementById('namesNotesColor')?.value || '#000000',
          notesFontSize: document.getElementById('namesNotesFontSize')?.value || 18,
          lotteryWinners: document.getElementById('namesLotteryWinners')?.value || 5,
          lotteryPerRow: document.getElementById('namesLotteryPerRow')?.value || 5,
          lotteryFontSize: document.getElementById('namesLotteryFontSize')?.value || 18
        },
        savedAt: new Date().toISOString()
      };
    }
    
    /**
     * Load names project data from Drive
     */
    function loadNamesProjectData(projectData, fileName) {
      if (!projectData) return;
      
      // Load names data
      if (projectData.namesData) {
        namesData = projectData.namesData;
        
        // Update preview
        const previewSection = document.getElementById('namesPreviewSection');
        const previewList = document.getElementById('namesPreviewList');
        
        if (previewSection && previewList && namesData.length > 0) {
          previewSection.classList.remove('hidden');
          previewList.innerHTML = namesData.map((item, i) => {
            let text = `${i + 1}. ${item.name}`;
            if (item.group) text += ` (${item.group})`;
            if (item.totalTickets > 1) text += ` - ${item.ticketNumber}/${item.totalTickets}`;
            return `<div class="py-1 border-b border-gray-200">${text}</div>`;
          }).join('');
          
          document.getElementById('namesFileInfo').textContent = t('namesLoadedCount', { count: namesData.length });
        }
      }
      
      // Load note template
      if (projectData.namesNoteTemplate) {
        namesNoteTemplate = projectData.namesNoteTemplate;
        const fileNameEl = document.getElementById('namesNoteFileName');
        if (fileNameEl && namesNoteTemplate.fileName) {
          fileNameEl.textContent = `✓ ${namesNoteTemplate.fileName}`;
        }
        enableNamesButtons();
      }
      
      // Load notes
      if (projectData.namesNotes) {
        namesNotes = projectData.namesNotes;
        renderNamesNotes();
      }
      
      // Load lottery results
      if (projectData.namesLotteryResults) {
        namesLotteryResults = projectData.namesLotteryResults;
        if (namesLotteryResults.length > 0) {
          renderNamesLotteryResults();
        }
      }
      
      // Load settings
      if (projectData.settings) {
        const s = projectData.settings;
        if (s.uploadType) {
          const el = document.getElementById('namesUploadType');
          if (el) el.value = s.uploadType;
        }
        if (s.notesPerRow) {
          const el = document.getElementById('namesNotesPerRow');
          if (el) el.value = s.notesPerRow;
        }
        if (s.notesSpacing) {
          const el = document.getElementById('namesNotesSpacing');
          if (el) el.value = s.notesSpacing;
        }
        if (s.notesColor) {
          const el = document.getElementById('namesNotesColor');
          if (el) el.value = s.notesColor;
        }
        if (s.notesFontSize) {
          const el = document.getElementById('namesNotesFontSize');
          if (el) el.value = s.notesFontSize;
        }
        if (s.lotteryWinners) {
          const el = document.getElementById('namesLotteryWinners');
          if (el) el.value = s.lotteryWinners;
        }
        if (s.lotteryPerRow) {
          const el = document.getElementById('namesLotteryPerRow');
          if (el) el.value = s.lotteryPerRow;
        }
        if (s.lotteryFontSize) {
          const el = document.getElementById('namesLotteryFontSize');
          if (el) el.value = s.lotteryFontSize;
        }
      }
      
      // Set project file name
      if (fileName) {
        window.currentProjectFileName = fileName;
      }
      
      showStatus(t('projectLoadedSuccess', { name: fileName?.replace('.json', '') || t('tabNamesLottery') }));
    }
    
    /**
     * Clear names project
     */
    function clearNamesProject() {
      namesData = [];
      namesNoteTemplate = null;
      namesNotes = [];
      namesLotteryResults = [];
      
      // Clear UI
      const previewSection = document.getElementById('namesPreviewSection');
      const previewList = document.getElementById('namesPreviewList');
      const fileInfo = document.getElementById('namesFileInfo');
      const noteFileName = document.getElementById('namesNoteFileName');
      
      if (previewSection) previewSection.classList.add('hidden');
      if (previewList) previewList.innerHTML = '';
      if (fileInfo) fileInfo.textContent = '';
      if (noteFileName) noteFileName.textContent = '';
      
      // Clear notes preview
      renderNamesNotes();
      
      // Clear lottery results
      const lotteryResults = document.getElementById('namesLotteryResults');
      const lotterySection = document.getElementById('namesLotteryResultsSection');
      const lotteryEmpty = document.getElementById('namesLotteryEmptyState');
      
      if (lotteryResults) lotteryResults.innerHTML = '';
      if (lotterySection) lotterySection.classList.add('hidden');
      if (lotteryEmpty) lotteryEmpty.classList.remove('hidden');
      
      // Disable buttons
      disableNamesButtons();
    }
    
    // Expose functions globally
    window.getNamesProjectData = getNamesProjectData;
    window.loadNamesProjectData = loadNamesProjectData;
    window.clearNamesProject = clearNamesProject;
    
    // ========== End Names Lottery Functions ==========
