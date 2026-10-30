/* ══════════════════════════════════════════════════════════════
   STATE
══════════════════════════════════════════════════════════════ */
const sessionId = crypto.randomUUID();
let isBusy = false;     // prevents double-sends while streaming
let isTokenLimitReached = false;  // tracks if token limit has been hit
// Get base path from configuration (or empty string if not set)
const base = window.APP_BASE_PATH || '';

/* ══════════════════════════════════════════════════════════════
   TOKEN LIMITING - USER FINGERPRINTING
══════════════════════════════════════════════════════════════ */
/**
 * Generates a unique machine fingerprint for user identification.
 * Uses only machine-level characteristics (not browser-specific) so the same
 * fingerprint is generated across different browsers on the same machine.
 */
window.fingerPrint = async function () {
    try {
        var canvas = document.body.appendChild(document.createElement('canvas'));
        canvas.width = 600;
        canvas.height = 300;
        canvas.style.display = "none";
        const ctx = canvas.getContext("2d");
        const size = 24;
        const diamondSize = 28;
        const gap = 4;
        const startX = 30;
        const startY = 30;
        const blue = "#1A3276";
        const orange = "#F28C00";
        const colorMap = [
            ["blue", "blue", "diamond"],
            ["blue", "orange", "blue"],
            ["blue", "blue", "blue"]
        ];
        function drawSquare(x, y, color) {
            ctx.fillStyle = color;
            ctx.fillRect(x, y, size, size);
        }
        function drawDiamond(centerX, centerY, size, color) {
            ctx.fillStyle = color;
            ctx.beginPath();
            ctx.moveTo(centerX, centerY - size / 2);
            ctx.lineTo(centerX + size / 2, centerY);
            ctx.lineTo(centerX, centerY + size / 2);
            ctx.lineTo(centerX - size / 2, centerY);
            ctx.closePath();
            ctx.fill();
        }
        for (let row = 0; row < 3; row++) {
            for (let col = 0; col < 3; col++) {
                const type = colorMap[row][col];
                const x = startX + col * (size + gap);
                const y = startY + row * (size + gap);
                if (type === "blue") drawSquare(x, y, blue);
                else if (type === "orange") drawSquare(x, y, orange);
                else if (type === "diamond") drawDiamond(x + size / 2, y + size / 2, diamondSize, orange);
            }
        }
        ctx.font = "20px Arial";
        ctx.fillStyle = blue;
        ctx.textBaseline = "middle";
        ctx.fillText("Syncfusion", startX + 3 * (size + gap) + 20, startY + size + gap);
        ctx.globalCompositeOperation = "multiply";
        ctx.fillStyle = "rgb(255,0,255)";
        ctx.beginPath(); ctx.arc(50, 200, 50, 0, Math.PI * 2); ctx.fill();
        ctx.fillStyle = "rgb(0,255,255)";
        ctx.beginPath(); ctx.arc(100, 200, 50, 0, Math.PI * 2); ctx.fill();
        ctx.fillStyle = "rgb(255,255,0)";
        ctx.beginPath(); ctx.arc(75, 250, 50, 0, Math.PI * 2); ctx.fill();
        ctx.fillStyle = "rgb(255,0,255)";
        ctx.beginPath();
        ctx.arc(200, 200, 75, 0, Math.PI * 2, true);
        ctx.arc(200, 200, 25, 0, Math.PI * 2, true);
        ctx.fill("evenodd");
        const sha256 = async function (str) {
            const encoder = new TextEncoder();
            const data = encoder.encode(str);
            const hashBuffer = await crypto.subtle.digest('SHA-256', data);
            const hashArray = Array.from(new Uint8Array(hashBuffer));
            return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
        };

        const visitorID = sha256(canvas.toDataURL());
        return visitorID;
    }
    catch (error) {
        console.error(error);
        return null;
    }
}

/**
 * Simple hash function for generating fingerprint from canvas data.
 */
function hashCode(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        const char = str.charCodeAt(i);
        hash = ((hash << 5) - hash) + char;
        hash = hash & hash;
    }
    return Math.abs(hash);
}

/**
 * Displays a banner notification when token limit is reached.
 */
window.showBanner = function(message) {
    // Remove existing banner if any
    const existing = document.getElementById('token-limit-banner');
    if (existing) {
        existing.remove();
    }
    
    // Create banner element
    const banner = document.createElement('div');
    banner.id = 'token-limit-banner';
    banner.innerHTML = message;
    banner.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        background: #f8d7da;
        color: #721c24;;
        padding: 15px 50px 15px 15px;
        text-align: center;
        z-index: 10000;
        font-size: 14px;
        line-height: 1.5;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    `;
    
    // Add close button
    const closeBtn = document.createElement('button');
    closeBtn.innerHTML = '×';
    closeBtn.style.cssText = `
        position: absolute;
        right: 15px;
        top: 50%;
        transform: translateY(-50%);
        background: none;
        border: none;
        color: #721c24;
        font-size: 24px;
        cursor: pointer;
        font-weight: normal;
        line-height: 1;
        padding: 5px 10px;
        opacity: 0.9;
    `;
    closeBtn.onmouseover = () => closeBtn.style.opacity = '1';
    closeBtn.onmouseout = () => closeBtn.style.opacity = '0.9';
    closeBtn.onclick = () => banner.remove();
    
    banner.appendChild(closeBtn);
    document.body.insertBefore(banner, document.body.firstChild);
};

/**
 * Checks token limit status on page load and disables UI if limit reached.
 */
async function checkTokenLimitOnLoad() {
    const userCode = window.fingerPrint ? await window.fingerPrint() : null;
    if (!userCode) return;

    try {
        const res = await fetch(`${base}/api/chat/tokens/${userCode}`);
        if (!res.ok) return;

        const data = await res.json();
        
        // If remaining tokens is 0 or negative, disable the UI
        if (data.remainingTokens <= 0) {
            isTokenLimitReached = true;
            
            // Show banner with message from server
            if (data.message) {
                window.showBanner(data.message);
            }
            
            // Disable send button
            const sendBtn = document.getElementById('sendBtn');
            if (sendBtn) {
                sendBtn.disabled = true;
                sendBtn.title = 'Token limit reached. Please refresh after reset time.';
                sendBtn.style.cursor = 'not-allowed';
                sendBtn.style.opacity = '0.5';
            }
            
            // Disable textarea
            const textarea = document.getElementById('promptInput');
            if (textarea) {
                textarea.disabled = true;
                textarea.placeholder = 'Token limit reached. Refresh page after reset time.';
                textarea.style.cursor = 'not-allowed';
            }
        }
    } catch (err) {
        console.error('Failed to check token limit on load:', err);
    }
}

/* ══════════════════════════════════════════════════════════════
   UTILITY
══════════════════════════════════════════════════════════════ */
function escapeHtml(str) {
  const d = document.createElement('div');
  d.textContent = str;
  return d.innerHTML;
}

function scrollFeedToBottom() {
  const feed = document.getElementById('chatFeed');
  feed.scrollTop = feed.scrollHeight;
}

/* ══════════════════════════════════════════════════════════════
   CUSTOM MODAL
══════════════════════════════════════════════════════════════ */
let modalOverlay = null;

function createModalOverlay() {
  if (modalOverlay) return modalOverlay;
  
  modalOverlay = document.createElement('div');
  modalOverlay.className = 'modal-overlay';
  modalOverlay.innerHTML = `
    <div class="modal-box">
      <div class="modal-icon"></div>
      <div class="modal-title"></div>
      <div class="modal-message"></div>
      <div class="modal-actions"></div>
    </div>
  `;
  document.body.appendChild(modalOverlay);
  
  // Close modal when clicking overlay
  modalOverlay.addEventListener('click', (e) => {
    if (e.target === modalOverlay) {
      closeModal();
    }
  });
  
  return modalOverlay;
}

function closeModal() {
  if (modalOverlay) {
    modalOverlay.classList.remove('active');
  }
}

function showModal(message, buttons, options = {}) {
  return new Promise((resolve) => {
    const overlay = createModalOverlay();
    const iconEl = overlay.querySelector('.modal-icon');
    const titleEl = overlay.querySelector('.modal-title');
    const messageEl = overlay.querySelector('.modal-message');
    const actionsEl = overlay.querySelector('.modal-actions');
    
    // Set icon if provided
    if (options.icon) {
      iconEl.innerHTML = options.icon;
      iconEl.style.display = 'flex';
    } else {
      iconEl.innerHTML = '';
      iconEl.style.display = 'none';
    }
    
    // Set title if provided
    if (options.title) {
      titleEl.textContent = options.title;
      titleEl.style.display = 'block';
    } else {
      titleEl.textContent = '';
      titleEl.style.display = 'none';
    }
    
    messageEl.textContent = message;
    actionsEl.innerHTML = '';
    
    let primaryButton = null;
    
    buttons.forEach((btn, index) => {
      const button = document.createElement('button');
      button.className = `modal-btn modal-btn-${btn.type}`;
      button.textContent = btn.label;
      button.addEventListener('click', () => {
        removeKeyboardHandler();
        closeModal();
        resolve(btn.value);
      });
      actionsEl.appendChild(button);
      
      // Remember the OK button (primary action)
      if (btn.type === 'ok') {
        primaryButton = button;
      }
    });
    
    overlay.classList.add('active');
    
    // Auto-focus the primary button (OK) after a short delay to ensure modal is visible
    setTimeout(() => {
      if (primaryButton) {
        primaryButton.focus();
      }
    }, 100);
    
    // Add keyboard event handler
    const keyHandler = (e) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        removeKeyboardHandler();
        closeModal();
        // Resolve with false for Escape (same as Cancel)
        resolve(false);
      } else if (e.key === 'Enter') {
        e.preventDefault();
        // Trigger click on the focused button, or primary button if nothing focused
        const focusedBtn = document.activeElement;
        if (focusedBtn && focusedBtn.classList.contains('modal-btn')) {
          focusedBtn.click();
        } else if (primaryButton) {
          primaryButton.click();
        }
      } else if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
        e.preventDefault();
        // Navigate between buttons using arrow keys
        const allButtons = Array.from(actionsEl.querySelectorAll('.modal-btn'));
        if (allButtons.length > 1) {
          const currentIndex = allButtons.indexOf(document.activeElement);
          let nextIndex;
          
          if (e.key === 'ArrowLeft') {
            // Move to previous button (or wrap to last)
            nextIndex = currentIndex <= 0 ? allButtons.length - 1 : currentIndex - 1;
          } else {
            // Move to next button (or wrap to first)
            nextIndex = currentIndex >= allButtons.length - 1 ? 0 : currentIndex + 1;
          }
          
          allButtons[nextIndex].focus();
        }
      }
    };
    
    const removeKeyboardHandler = () => {
      document.removeEventListener('keydown', keyHandler);
    };
    
    document.addEventListener('keydown', keyHandler);
  });
}

function showAlert(message) {
  return showModal(message, [
    { label: 'OK', type: 'ok', value: true }
  ]);
}

function showConfirm(message) {
  return showModal(message, [
    { label: 'Cancel', type: 'cancel', value: false },
    { label: 'OK', type: 'ok', value: true }
  ]);
}

/* ══════════════════════════════════════════════════════════════
   TABS
══════════════════════════════════════════════════════════════ */
function initTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab;

      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      document.querySelectorAll('.tab-pane').forEach(pane => {
        pane.classList.toggle('active', pane.id === `${tab}-tab`);
      });

      // Refresh file list when switching tabs
      if (tab === 'documents') {
        loadFiles('Input', 'documentsList');
      } else if (tab === 'exports') {
        loadFiles('Output', 'exportsList');
      }
    });
  });
}

/* ══════════════════════════════════════════════════════════════
   FILE HELPERS
══════════════════════════════════════════════════════════════ */
function getFileTypeIcon(fileName) {
  const ext = fileName.split('.').pop().toLowerCase();
  const iconStyle = 'width: 24px; height: 24px;';
  
  switch (ext) {
    // Word documents
    case 'docx':
    case 'doc':
    case 'rtf':
      return `<img src="${base}/img/word.png" alt="Word" style="${iconStyle}">`;
    
    // HTML
    case 'html':
      return `<img src="${base}/img/html.png" alt="HTML" style="${iconStyle}">`;
    
    // PowerPoint
    case 'pptx':
      return `<img src="${base}/img/powerpoint.png" alt="PowerPoint" style="${iconStyle}">`;
    
    // Excel
    case 'xlsx':
    case 'xls':
    case 'xlsm':
    case 'csv':
      return `<img src="${base}/img/excel.png" alt="Excel" style="${iconStyle}">`;
    
    // PDF
    case 'pdf':
      return `<img src="${base}/img/pdf.png" alt="PDF" style="${iconStyle}">`;
    
    // JSON
    case 'json':
      return `<img src="${base}/img/json.png" alt="JSON" style="${iconStyle}">`;
    
    // Markdown
    case 'md':
      return `<img src="${base}/img/markdown.png" alt="Markdown" style="${iconStyle}">`;
    
    // Text
    case 'txt':
      return `<img src="${base}/img/txt.png" alt="Text" style="${iconStyle}">`;
    
    // Images
    case 'png':
    case 'jpg':
    case 'jpeg':
      return `<img src="${base}/img/image.png" alt="Image" style="${iconStyle}">`;
    
    // Default
    default:
      return `<img src="${base}/img/others.png" alt="File" style="${iconStyle}">`;
  }
}

const FILE_ICON_SVG = `
  <svg viewBox="0 0 24 24" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
    <path d="M14,2H6A2,2 0 0,0 4,4V20A2,2 0 0,0 6,22H18A2,2 0 0,0 20,20V8L14,2M18,20H6V4H13V9H18V20Z"/>
  </svg>`;


const DOWNLOAD_ICON_SVG = `
  <svg viewBox="0 0 24 24" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
    <polyline points="7 10 12 15 17 10"/>
    <line x1="12" y1="15" x2="12" y2="3"/>
  </svg>`;

const DELETE_ICON_SVG = `
  <svg viewBox="0 0 24 24" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <polyline points="3 6 5 6 21 6"/>
    <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
  </svg>`;

async function loadFiles(folder, listId) {
  const list = document.getElementById(listId);
  list.innerHTML = '<li class="empty-msg">Loading…</li>';

  try {
    const res = await fetch(`${base}/api/files/folders/${folder}`);
    if (!res.ok) throw new Error(await res.text());
    const { files } = await res.json();

    list.innerHTML = '';

    if (!files.length) {
      list.innerHTML = '<li class="empty-msg">No files found</li>';
      return;
    }

    files.forEach(f => list.appendChild(buildFileItem(f, folder)));
  } catch (err) {
    list.innerHTML = `<li class="empty-msg" style="color:#e53e3e">Error: ${escapeHtml(err.message)}</li>`;
  }
}

function buildFileItem(file, folder) {
  const isOutput = folder === 'Output';
  const li = document.createElement('li');
  li.className = 'file-item';
  const iconSvg = getFileTypeIcon(file.name);

  li.innerHTML = `
    <div class="file-icon">${iconSvg}</div>
    <div class="file-info">
      <div class="file-name" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</div>
      <div class="file-size">${escapeHtml(file.size)}</div>
    </div>
    <div class="file-actions">
      ${`<button class="file-action-btn download" title="Download">${DOWNLOAD_ICON_SVG}</button>`}
      <button class="file-action-btn delete" title="Delete">${DELETE_ICON_SVG}</button>
    </div>`;

  
    li.querySelector('.download').addEventListener('click', e => {
      e.stopPropagation();
      triggerDownload(folder, file.name);
    });
  

  li.querySelector('.delete').addEventListener('click', async e => {
    e.stopPropagation();
    const confirmed = await showConfirm(`Are you sure you want to delete ${file.name}?`);
    if (confirmed) {
        await deleteFile(folder, file.name, li);
        showAlert(`File Deleted Successfully ✅`);
    }
  });

  return li;
}

function triggerDownload(folder, fileName) {
  const a = document.createElement('a');
  a.href = `${base}/api/files/download/${encodeURIComponent(folder)}/${encodeURIComponent(fileName)}`;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

async function deleteFile(folder, fileName, li) {
  try {
    const res = await fetch(
      `${base}/api/files/delete/${encodeURIComponent(folder)}/${encodeURIComponent(fileName)}`,
      { method: 'DELETE' }
    );
    if (!res.ok) throw new Error(await res.text());
    li.remove();
    const parent = li.parentElement;
    if (parent && !parent.querySelector('.file-item')) {
      parent.innerHTML = '<li class="empty-msg">No files found</li>';
    }
  } catch (err) {
    await showAlert(`Failed to delete "${fileName}":\n${err.message}`);
  }
}

/* ══════════════════════════════════════════════════════════════
   FILE UPLOAD
══════════════════════════════════════════════════════════════ */
function initUpload() {
  bindUpload('uploadDocumentsBtn', 'uploadDocumentsInput', 'Input',  'documentsList');
  bindUpload('uploadExportsBtn',   'uploadExportsInput',   'Output', 'exportsList');
}

function bindUpload(btnId, inputId, folder, listId) {
  const btn   = document.getElementById(btnId);
  const input = document.getElementById(inputId);

  btn.addEventListener('click', () => input.click());

  input.addEventListener('change', async () => {
    if (!input.files.length) return;
    
    // Validation constants
    const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50 MB in bytes
    const MAX_FILE_COUNT = 10;
    const MAX_FILE_SIZE_MB = (MAX_FILE_SIZE / (1024 * 1024)).toFixed(1);
    
    // Warning icon SVG (orange triangle with exclamation)
    const warningIcon = `
      <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 30 30">
        <path d="M14.9993493,2 C16.3733967,2 17.6477191,2.71745392 18.3640958,3.89849762 L29.4732112,22.4446209 C30.1716648,23.6541832 30.175835,25.1435285 29.4841658,26.3569832 C28.7924967,27.5704379 27.5088761,28.3257424 26.0978207,28.3411725 L3.8864766,28.3411725 C2.48982245,28.3257424 1.20620185,27.5704379 0.514532682,26.3569832 C-0.177136481,25.1435285 -0.172966311,23.6541832 0.536131236,22.4265244 L11.6346027,3.89849762 C12.3509794,2.71745392 13.6253018,2 14.9993493,2 Z M14.9993493,4.62065442 C14.5423377,4.62065442 14.1184222,4.85875817 13.8805847,5.2488356 L2.79494922,23.7551064 C2.56213134,24.1582938 2.56074129,24.6547423 2.79129768,25.0592271 C3.02185406,25.463712 3.4497276,25.7154802 3.90087779,25.7205972 L26.0834195,25.7205972 C26.5489709,25.7154802 26.9768445,25.463712 27.2074008,25.0592271 C27.4379572,24.6547423 27.4365672,24.1582938 27.2143932,23.7732029 L16.119679,5.25140925 C15.8821392,4.85980573 15.4573651,4.62065442 14.9993493,4.62065442 Z M14.072808,19.552668 C14.5845226,19.0409533 15.4141759,19.0409533 15.9258905,19.552668 C16.4376051,20.0643826 16.4376051,20.8940359 15.9258905,21.4057505 C15.4141759,21.9174651 14.5845226,21.9174651 14.072808,21.4057505 C13.5610934,20.8940359 13.5610934,20.0643826 14.072808,19.552668 Z M14.9993493,9.99659153 C15.723023,9.99659153 16.3096765,10.583245 16.3096765,11.3069187 L16.3096765,11.3069187 L16.3096765,16.5482276 C16.3096765,17.2719013 15.723023,17.8585548 14.9993493,17.8585548 C14.2756755,17.8585548 13.6890221,17.2719013 13.6890221,16.5482276 L13.6890221,16.5482276 L13.6890221,11.3069187 C13.6890221,10.583245 14.2756755,9.99659153 14.9993493,9.99659153 Z" fill="#FF9522" fill-rule="nonzero"/>
      </svg>
    `;
    
    // Validate file count
    if (input.files.length > MAX_FILE_COUNT) {
      await showAlert(`You can only upload up to ${MAX_FILE_COUNT} files at once.\nYou selected ${input.files.length} files.`);
      input.value = '';
      return;
    }
    
    // Validate file sizes
    const oversizedFiles = [];
    for (const file of input.files) {
      if (file.size > MAX_FILE_SIZE) {
        oversizedFiles.push({
          name: file.name,
          size: (file.size / (1024 * 1024)).toFixed(1) // Convert to MB
        });
      }
    }
    
    if (oversizedFiles.length > 0) {
      // Format message for single or multiple files
      let message;
        if (oversizedFiles.length === 1) {
            message = `This file can't be uploaded because it (${oversizedFiles[0].size} MB) exceeds the maximum file-size ${MAX_FILE_SIZE_MB} MB for this operation.`;
      } else {
            const fileList = oversizedFiles.map(f => `${f.name} (${f.size} MB)`).join(', ');
            message = `These files can't be uploaded because they exceed the maximum file-size ${MAX_FILE_SIZE_MB} MB:\n\n${fileList}`;
      }
      
      await showModal(message, [{ label: 'OK', type: 'ok', value: true }], {
        icon: warningIcon,
        title: 'Sorry!'
      });
      input.value = '';
      return;
    }
    
    // All validations passed, proceed with upload
    const fd = new FormData();
    for (const f of input.files) fd.append('files', f);
    input.value = '';

    try {
      const res = await fetch(`${base}/api/files/upload`, { method: 'POST', body: fd });
      if (!res.ok) throw new Error(await res.text());
      const data = await res.json();
      
      // Show success message based on file count
      const fileCount = data.files.length;
      let message;
      if (fileCount === 1) {
          // Single file uploaded
          message = `${data.files[0].split('/').pop()} uploaded successfully`;
      } else {
        // Multiple files uploaded
        message = 'Files uploaded successfully';
      }
      
      await showAlert(message);
      await loadFiles(folder, listId);
    } catch (err) {
      await showAlert(`Upload failed:\n${err.message}`);
    }
  });
}



/* ══════════════════════════════════════════════════════════════
   CHAT — message rendering
══════════════════════════════════════════════════════════════ */
function appendMessage(role, text) {
  const feed   = document.getElementById('chatFeed');
  const banner = document.getElementById('welcomeBanner');

  // Hide welcome banner once conversation starts
  if (banner) banner.style.display = 'none';

  const row = document.createElement('div');
  row.className = `msg-row ${role}`;

  const bubble = document.createElement('div');
  bubble.className = 'msg-bubble';
  bubble.textContent = text;

  row.appendChild(bubble);
  feed.appendChild(row);
  scrollFeedToBottom();
  return bubble;   // return bubble so streaming can update it
}

function appendTypingIndicator() {
  const feed = document.getElementById('chatFeed');
  const banner = document.getElementById('welcomeBanner');
  if (banner) banner.style.display = 'none';

  const row = document.createElement('div');
  row.className = 'msg-row bot typing-row';
  row.id = 'typingIndicator';
  row.innerHTML = `
    <div class="msg-bubble-dot">
      <div class="typing-dots">
        <span></span><span></span><span></span>
      </div>
    </div>`;
  feed.appendChild(row);
  scrollFeedToBottom();
}

function removeTypingIndicator() {
  const el = document.getElementById('typingIndicator');
  if (el) el.remove();
}

/* ══════════════════════════════════════════════════════════════
   CHAT — send prompt
══════════════════════════════════════════════════════════════ */
async function sendPrompt(text) {
  text = text.trim();
  if (!text || isBusy || isTokenLimitReached) return;

  isBusy = true;
  setSendBtnState(false);

  appendMessage('user', text);
  appendTypingIndicator();

  try {
    // Get user fingerprint for token limiting
    const userCode = window.fingerPrint ? await window.fingerPrint() : null;

    const res = await fetch(`${base}/api/chat/send`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ 
        sessionId, 
        message: text,
        userCode: userCode
      })
    });

    removeTypingIndicator();

    // Handle token limit error (429)
    if (res.status === 429) {
      const errorData = await res.json();
      
      // Always show banner with either server message or default message
      const bannerMessage = errorData.message || 'You have reached your token limit. Download our Syncfusion Smart AI Samples from GitHub to explore this sample locally with your own API key.';
      console.log('Token limit reached. Showing banner:', bannerMessage);
      window.showBanner(bannerMessage);
      
      // Disable sending permanently until page refresh
      isTokenLimitReached = true;
      setSendBtnState(false);
      
      // Update button to show it's disabled due to limit
      const sendBtn = document.getElementById('sendBtn');
      sendBtn.title = 'Token limit reached. Please refresh after reset time.';
      sendBtn.style.cursor = 'not-allowed';
      sendBtn.style.opacity = '0.5';
      
      // Disable textarea as well
      const textarea = document.getElementById('promptInput');
      textarea.disabled = true;
      textarea.placeholder = 'Token limit reached. Refresh page after reset time.';
      textarea.style.cursor = 'not-allowed';
      
      return;
    }

    if (!res.ok) {
      const errText = await res.text();
      appendMessage('bot', `❌ Error: ${errText}`);
      return;
    }

    // Stream response
    const reader  = res.body.getReader();
    const decoder = new TextDecoder();
    let fullText  = '';
    const bubble  = appendMessage('bot', '');

      while (true) {
          const { done, value } = await reader.read();

          if (done) break;

          fullText += decoder.decode(value, { stream: true });
          bubble.textContent = fullText;
          scrollFeedToBottom();

          // Only re-enable if token limit hasn't been reached
          if (!isTokenLimitReached) {
              setSendBtnState(true);
          }
      }


    if (!fullText.trim()) bubble.textContent = '(No response)';

    // Refresh exports list after AI response completes
    await loadFiles('Output', 'exportsList');

  } catch (err) {
    removeTypingIndicator();
    appendMessage('bot', `❌ ${err.message}`);
  } finally {
    isBusy = false;
    setSendBtnState(true);
  }
}

function setSendBtnState(enabled) {
  document.getElementById('sendBtn').disabled = !enabled;
}

/* ══════════════════════════════════════════════════════════════
   CHAT — input bar wiring
══════════════════════════════════════════════════════════════ */
function initChatInput() {
  const textarea = document.getElementById('promptInput');
  const sendBtn  = document.getElementById('sendBtn');
// Only send if token limit hasn't been reached
      if (!isTokenLimitReached) {
        dispatchSend();
      }
  // Auto-grow textarea
  textarea.addEventListener('input', () => {

      const maxHeight = 125; // same as your current limit

      if (textarea.scrollHeight <= maxHeight) {
          textarea.style.height = textarea.scrollHeight + 'px';
          textarea.style.overflowY = 'hidden'; // ✅ no scrollbar yet
      } else {
          textarea.style.height = maxHeight + 'px';
          textarea.style.overflowY = 'auto';   // ✅ show scrollbar
      }

  });

  // Enter sends, Shift+Enter inserts newline
  textarea.addEventListener('keydown', e => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      dispatchSend();
    // Don't send if token limit reached
    if (isTokenLimitReached) return;
    
    }
  });

  sendBtn.addEventListener('click', dispatchSend);

  function dispatchSend() {
    const text = textarea.value;
    textarea.value = '';
    textarea.style.height = 'auto';
    sendPrompt(text);
  }
}

/* ══════════════════════════════════════════════════════════════
   PROMPT SUGGESTIONS
══════════════════════════════════════════════════════════════ */
function initSuggestions() {
  document.querySelectorAll('.suggestion-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      const promptText = chip.dataset.prompt;
      if (promptText) sendPrompt(promptText);
    });
  });
}

/* ══════════════════════════════════════════════════════════════
   MOBILE SIDEBAR TOGGLE
══════════════════════════════════════════════════════════════ */
function initMobileSidebar() {
  const menuBtn = document.getElementById('mobileMenuBtn');
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('sidebarOverlay');

  console.log('initMobileSidebar called');
  console.log('menuBtn:', menuBtn);
  console.log('sidebar:', sidebar);
  console.log('overlay:', overlay);

  if (!menuBtn || !sidebar || !overlay) {
    console.error('Missing elements!', { menuBtn, sidebar, overlay });
    return;
  }

  // Toggle sidebar
  menuBtn.addEventListener('click', (e) => {
    console.log('Menu button clicked!');
    e.preventDefault();
    e.stopPropagation();
    sidebar.classList.toggle('active');
    overlay.classList.toggle('active');
    console.log('Sidebar active:', sidebar.classList.contains('active'));
  });

  // Close sidebar when clicking overlay
  overlay.addEventListener('click', () => {
    console.log('Overlay clicked!');
    sidebar.classList.remove('active');
    overlay.classList.remove('active');
  });

  // Close sidebar when selecting a file or action
  sidebar.addEventListener('click', (e) => {
    // Close only on file actions (download/delete) or upload
    if (e.target.closest('.file-action-btn') || e.target.closest('.upload-btn')) {
      setTimeout(() => {
        sidebar.classList.remove('active');
        overlay.classList.remove('active');
      }, 200);
    }
  });
}

/* ══════════════════════════════════════════════════════════════
   BOOTSTRAP
══════════════════════════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', async () => {
  // Set refresh icon path using base
  const refreshIcon = document.getElementById('refreshIcon');
  if (refreshIcon) {
    refreshIcon.src = `${base}/img/refresh.svg`;
  }

  initTabs();
  initUpload();
  initChatInput();
  initSuggestions();
  initMobileSidebar();

  // Check token limit status on page load
  await checkTokenLimitOnLoad();

  await Promise.all([
    loadFiles('Input',  'documentsList'),
    loadFiles('Output', 'exportsList')
  ]);
});
