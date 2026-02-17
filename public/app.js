/* ===================================================================
   MD-to-PDF — Frontend Application
   Real-time Markdown editor with Mermaid support & multi-format export
   =================================================================== */

(function () {
  'use strict';

  // ───── Theme Initialization (must run before any rendering) ─────
  // Apply saved theme immediately to prevent flash of wrong theme
  const STORAGE_KEY_THEME = 'md-to-pdf-theme';
  const savedTheme = localStorage.getItem(STORAGE_KEY_THEME) || 'dark';
  document.documentElement.setAttribute('data-theme', savedTheme);

  // Set initial highlight.js theme based on saved theme
  const hljsTheme = document.getElementById('hljs-theme');
  if (hljsTheme) {
    hljsTheme.href = savedTheme === 'dark'
      ? 'https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.11.1/styles/github-dark.min.css'
      : 'https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.11.1/styles/github.min.css';
  }

  // ───── DOM References ─────
  const editor = document.getElementById('editor');
  const preview = document.getElementById('preview');
  const docTitle = document.getElementById('doc-title');
  const charCount = document.getElementById('char-count');
  const wordCount = document.getElementById('word-count');
  const loadingOverlay = document.getElementById('loading-overlay');
  const loadingText = document.getElementById('loading-text');
  const toastContainer = document.getElementById('toast-container');
  const resizeHandle = document.getElementById('resize-handle');
  const hljsThemeLink = document.getElementById('hljs-theme');

  // ───── Mermaid Setup ─────
  mermaid.initialize({
    startOnLoad: false,
    theme: savedTheme === 'dark' ? 'dark' : 'default',
    securityLevel: 'loose',
    flowchart: { useMaxWidth: true, htmlLabels: true, curve: 'basis' },
    sequence: { useMaxWidth: true, wrap: true },
    gantt: { useMaxWidth: true },
  });

  // ───── Marked Setup ─────
  const { Marked } = marked;
  const markedInstance = new Marked();

  markedInstance.use({
    gfm: true,
    breaks: false,
    renderer: {
      code({ text, lang }) {
        if (lang === 'mermaid') {
          return `<pre class="mermaid">${escapeHtml(text)}</pre>`;
        }
        let highlighted;
        if (lang && hljs.getLanguage(lang)) {
          highlighted = hljs.highlight(text, { language: lang }).value;
        } else {
          highlighted = hljs.highlightAuto(text).value;
        }
        return `<pre><code class="hljs language-${lang || 'plaintext'}">${highlighted}</code></pre>`;
      },
    },
  });

  function escapeHtml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ───── State ─────
  let mermaidCounter = 0;
  let isRendering = false;

  // ───── Session Persistence ─────
  const STORAGE_KEY_CONTENT = 'md-to-pdf-content';
  const STORAGE_KEY_TITLE = 'md-to-pdf-title';
  const SAVE_INTERVAL = 1000; // Auto-save every 1s after changes

  function saveSession() {
    try {
      localStorage.setItem(STORAGE_KEY_CONTENT, editor.value);
      localStorage.setItem(STORAGE_KEY_TITLE, docTitle.value);
    } catch (e) {
      // localStorage full or unavailable — silently ignore
    }
  }

  function loadSession() {
    const savedContent = localStorage.getItem(STORAGE_KEY_CONTENT);
    const savedTitle = localStorage.getItem(STORAGE_KEY_TITLE);
    if (savedTitle) docTitle.value = savedTitle;
    return savedContent;
  }

  // Debounced auto-save on every keystroke
  let saveTimer = null;
  function scheduleSave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(saveSession, SAVE_INTERVAL);
  }

  // Save on editor input
  editor.addEventListener('input', scheduleSave);
  // Save on title change
  docTitle.addEventListener('input', scheduleSave);
  // Save immediately before the page unloads
  window.addEventListener('beforeunload', saveSession);

  // ───── Default Content ─────
  const defaultContent = `# Welcome to MD-to-PDF

A powerful Markdown editor with **real-time preview**, Mermaid diagram support, and multi-format export.

---

## Features

- **Live Preview** — See your changes as you type
- **Mermaid Diagrams** — Flowcharts, sequence diagrams, Gantt charts, and more
- **Export Options** — PDF, HTML, Markdown, and Word (.docx)
- **Dark/Light Theme** — Toggle with the moon/sun icon
- **Syntax Highlighting** — Code blocks with automatic language detection

## Code Example

\`\`\`javascript
function fibonacci(n) {
  if (n <= 1) return n;
  return fibonacci(n - 1) + fibonacci(n - 2);
}

console.log(fibonacci(10)); // 55
\`\`\`

## Flowchart

\`\`\`mermaid
graph TD
    A[Start Writing Markdown] --> B{Has Diagrams?}
    B -->|Yes| C[Render Mermaid]
    B -->|No| D[Render Markdown]
    C --> E[Live Preview]
    D --> E
    E --> F{Export?}
    F -->|PDF| G[High Quality PDF]
    F -->|HTML| H[Self-Contained HTML]
    F -->|MD| I[Raw Markdown File]
\`\`\`

## Sequence Diagram

\`\`\`mermaid
sequenceDiagram
    participant User
    participant Editor
    participant Server
    participant Puppeteer

    User->>Editor: Types Markdown
    Editor->>Editor: Parse & Render Preview
    User->>Editor: Clicks Export PDF
    Editor->>Server: POST /export/pdf
    Server->>Puppeteer: Render HTML + Mermaid
    Puppeteer-->>Server: PDF Buffer
    Server-->>Editor: PDF File
    Editor-->>User: Download PDF
\`\`\`

## Table Example

| Feature | Markdown | HTML | PDF |
|---------|:--------:|:----:|:---:|
| Text formatting | Yes | Yes | Yes |
| Code highlighting | Yes | Yes | Yes |
| Mermaid diagrams | Source | Rendered | Rendered |
| Tables | Yes | Yes | Yes |
| Images | Yes | Yes | Yes |

## Gantt Chart

\`\`\`mermaid
gantt
    title Project Timeline
    dateFormat  YYYY-MM-DD
    section Planning
    Research           :a1, 2024-01-01, 7d
    Design             :a2, after a1, 5d
    section Development
    Frontend           :b1, after a2, 14d
    Backend            :b2, after a2, 10d
    Integration        :b3, after b1, 5d
    section Testing
    QA Testing         :c1, after b3, 7d
    Bug Fixes          :c2, after c1, 5d
\`\`\`

## Blockquote

> **Note:** This editor supports all standard Markdown syntax plus GitHub Flavored Markdown extensions including tables, task lists, and strikethrough.

## Task List

- [x] Real-time markdown preview
- [x] Mermaid diagram rendering
- [x] PDF export with Puppeteer
- [x] HTML export (self-contained)
- [x] Markdown file export
- [x] Dark/Light theme toggle
- [x] Syntax highlighting

---

*Start editing to see the magic happen!*
`;

  // ───── Initialize ─────
  const savedContent = loadSession();
  editor.value = savedContent !== null ? savedContent : defaultContent;
  updatePreview();

  // ───── Live Preview with Debounce ─────
  let debounceTimer = null;

  editor.addEventListener('input', () => {
    updateStats();
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(updatePreview, 300);
  });

  function updateStats() {
    const text = editor.value;
    charCount.textContent = text.length.toLocaleString() + ' chars';
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    wordCount.textContent = words.toLocaleString() + ' words';
  }

  async function updatePreview() {
    if (isRendering) return;
    isRendering = true;

    try {
      const md = editor.value;
      const html = markedInstance.parse(md);
      preview.innerHTML = html;
      updateStats();
      await renderMermaidDiagrams();
    } catch (err) {
      console.error('Preview error:', err);
    } finally {
      isRendering = false;
    }
  }

  async function renderMermaidDiagrams() {
    const elements = preview.querySelectorAll('.mermaid');
    if (elements.length === 0) return;

    for (let i = 0; i < elements.length; i++) {
      const el = elements[i];
      const code = el.textContent.trim();
      if (!code) continue;

      const id = 'mermaid-preview-' + mermaidCounter++;

      try {
        const { svg } = await mermaid.render(id, code);
        el.innerHTML = svg;
        el.classList.remove('mermaid-has-error');
      } catch (err) {
        el.innerHTML =
          '<div class="mermaid-error-msg">Diagram Error: ' +
          escapeHtml(err.message || 'Invalid syntax') +
          '</div>';
        el.classList.add('mermaid-has-error');

        // Clean up any orphaned error SVGs mermaid may have appended
        const errSvg = document.getElementById('d' + id);
        if (errSvg) errSvg.remove();
      }
    }
  }

  // ───── Theme Toggle ─────
  const btnTheme = document.getElementById('btn-theme');

  btnTheme.addEventListener('click', () => {
    const current = document.documentElement.getAttribute('data-theme');
    const next = current === 'dark' ? 'light' : 'dark';
    applyTheme(next);
    localStorage.setItem(STORAGE_KEY_THEME, next);
  });

  function applyTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);

    // Switch highlight.js theme
    if (theme === 'dark') {
      hljsThemeLink.href =
        'https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.11.1/styles/github-dark.min.css';
    } else {
      hljsThemeLink.href =
        'https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.11.1/styles/github.min.css';
    }

    // Update mermaid theme
    mermaid.initialize({
      startOnLoad: false,
      theme: theme === 'dark' ? 'dark' : 'default',
      securityLevel: 'loose',
    });

    // Re-render mermaid diagrams with new theme
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(updatePreview, 100);
  }

  // ───── Export Dropdown ─────
  const exportDropdownBtn = document.getElementById('export-dropdown-btn');
  const exportDropdown = document.querySelector('.export-dropdown');

  if (exportDropdownBtn && exportDropdown) {
    // Toggle dropdown on button click
    exportDropdownBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      exportDropdown.classList.toggle('open');
    });

    // Close dropdown when clicking outside
    document.addEventListener('click', (e) => {
      if (!exportDropdown.contains(e.target)) {
        exportDropdown.classList.remove('open');
      }
    });

    // Close dropdown when pressing Escape
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        exportDropdown.classList.remove('open');
      }
    });

    // Close dropdown after selecting an option
    const dropdownItems = exportDropdown.querySelectorAll('.dropdown-item');
    dropdownItems.forEach(item => {
      item.addEventListener('click', () => {
        exportDropdown.classList.remove('open');
      });
    });
  }

  // ───── Export Functions ─────

  // Export as Markdown
  document.getElementById('btn-export-md').addEventListener('click', async () => {
    const markdown = editor.value;
    if (!markdown.trim()) {
      showToast('Nothing to export — editor is empty', 'warn');
      return;
    }

    try {
      const res = await fetch('/export/md', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ markdown, title: docTitle.value }),
      });
      if (!res.ok) throw new Error('Export failed');
      const blob = await res.blob();
      downloadBlob(blob, (docTitle.value || 'document') + '.md');
      showToast('Markdown exported successfully', 'success');
    } catch (err) {
      showToast('Failed to export Markdown: ' + err.message, 'error');
    }
  });

  // Export as HTML
  // Export as HTML
  document.getElementById('btn-export-html').addEventListener('click', async () => {
    const markdown = editor.value;
    if (!markdown.trim()) {
      showToast('Nothing to export — editor is empty', 'warn');
      return;
    }

    try {
      const res = await fetch('/export/html', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ markdown, title: docTitle.value }),
      });
      if (!res.ok) throw new Error('Export failed');
      const blob = await res.blob();
      downloadBlob(blob, (docTitle.value || 'document') + '.html');
      showToast('HTML exported successfully', 'success');
    } catch (err) {
      showToast('Failed to export HTML: ' + err.message, 'error');
    }
  });

  // Export as Word
  document.getElementById('btn-export-docx').addEventListener('click', async () => {
    const markdown = editor.value;
    if (!markdown.trim()) {
      showToast('Nothing to export — editor is empty', 'warn');
      return;
    }

    showLoading('Generating Word document...');

    try {
      const res = await fetch('/export/docx', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ markdown, title: docTitle.value }),
      });
      if (!res.ok) throw new Error('Export failed');
      const blob = await res.blob();
      downloadBlob(blob, (docTitle.value || 'document') + '.docx');
      showToast('Word document exported successfully', 'success');
    } catch (err) {
      showToast('Failed to export Word: ' + err.message, 'error');
    } finally {
      hideLoading();
    }
  });

  // Export as PDF
  document.getElementById('btn-export-pdf').addEventListener('click', async () => {
    const markdown = editor.value;
    if (!markdown.trim()) {
      showToast('Nothing to export — editor is empty', 'warn');
      return;
    }

    showLoading('Generating PDF...');

    try {
      // Get the rendered HTML from the preview for PDF generation
      // We re-parse on the server side with the template, so send the raw rendered HTML
      const html = markedInstance.parse(markdown);

      const res = await fetch('/export/pdf', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ html, title: docTitle.value }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: 'Unknown error' }));
        throw new Error(err.error || 'PDF generation failed');
      }

      const blob = await res.blob();
      downloadBlob(blob, (docTitle.value || 'document') + '.pdf');
      showToast('PDF exported successfully', 'success');
    } catch (err) {
      showToast('Failed to export PDF: ' + err.message, 'error');
    } finally {
      hideLoading();
    }
  });

  // ───── Download Helper ─────
  function sanitizeFilename(filename) {
    // Only remove characters that are actually problematic for filenames
    // Preserve Unicode characters (Japanese, etc.)
    return filename.replace(/[<>:"/\\|?*\x00-\x1f]/g, '_');
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = sanitizeFilename(filename);
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ───── Loading Overlay ─────
  function showLoading(text) {
    loadingText.textContent = text || 'Processing...';
    loadingOverlay.classList.add('visible');
  }

  function hideLoading() {
    loadingOverlay.classList.remove('visible');
  }

  // ───── Toast Notifications ─────
  function showToast(message, type) {
    const toast = document.createElement('div');
    toast.className = 'toast toast-' + (type || 'info');

    const icons = {
      success: '<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="7" stroke="currentColor" stroke-width="1.5"/><path d="M5 8l2 2 4-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>',
      error: '<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="7" stroke="currentColor" stroke-width="1.5"/><path d="M5.5 5.5l5 5M10.5 5.5l-5 5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>',
      warn: '<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M8 1.5l7 13H1z" stroke="currentColor" stroke-width="1.5" fill="none"/><path d="M8 6v3M8 11.5v.5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>',
      info: '<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="7" stroke="currentColor" stroke-width="1.5"/><path d="M8 7v4M8 5v.5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>',
    };

    toast.innerHTML = (icons[type] || icons.info) + '<span>' + escapeHtml(message) + '</span>';
    toastContainer.appendChild(toast);

    // Trigger animation
    requestAnimationFrame(() => toast.classList.add('visible'));

    setTimeout(() => {
      toast.classList.remove('visible');
      setTimeout(() => toast.remove(), 300);
    }, 3500);
  }

  // ───── Resizable Panes ─────
  let isResizing = false;

  resizeHandle.addEventListener('mousedown', (e) => {
    isResizing = true;
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
    e.preventDefault();
  });

  document.addEventListener('mousemove', (e) => {
    if (!isResizing) return;
    const container = document.querySelector('.editor-container');
    const rect = container.getBoundingClientRect();
    const offset = e.clientX - rect.left;
    const percent = (offset / rect.width) * 100;
    const clamped = Math.min(Math.max(percent, 20), 80);

    container.style.gridTemplateColumns = `${clamped}% 6px 1fr`;
  });

  document.addEventListener('mouseup', () => {
    if (isResizing) {
      isResizing = false;
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    }
  });

  // ───── Tab Key Support in Editor ─────
  editor.addEventListener('keydown', (e) => {
    if (e.key === 'Tab') {
      e.preventDefault();
      const start = editor.selectionStart;
      const end = editor.selectionEnd;
      editor.value = editor.value.substring(0, start) + '  ' + editor.value.substring(end);
      editor.selectionStart = editor.selectionEnd = start + 2;
      editor.dispatchEvent(new Event('input'));
    }
  });

  // ───── Synchronized Scrolling (Preview → Editor only) ─────
  // When scrolling the preview, the editor scrolls to track position
  // But scrolling the editor does NOT scroll the preview
  let isPreviewScrolling = false;

  preview.addEventListener('scroll', () => {
    if (isPreviewScrolling) return;
    
    // Calculate scroll percentage in preview
    const previewScrollHeight = preview.scrollHeight - preview.clientHeight;
    const editorScrollHeight = editor.scrollHeight - editor.clientHeight;
    
    if (previewScrollHeight <= 0 || editorScrollHeight <= 0) return;
    
    const scrollPercentage = preview.scrollTop / previewScrollHeight;
    const targetEditorScroll = scrollPercentage * editorScrollHeight;
    
    // Sync editor scroll position
    isPreviewScrolling = true;
    editor.scrollTop = targetEditorScroll;
    
    // Reset flag after a short delay
    setTimeout(() => {
      isPreviewScrolling = false;
    }, 50);
  });

  // ───── Drag & Drop File Support ─────
  let dragCounter = 0;

  editor.addEventListener('dragenter', (e) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounter++;
    editor.classList.add('drag-over');
  });

  editor.addEventListener('dragleave', (e) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounter--;
    if (dragCounter === 0) {
      editor.classList.remove('drag-over');
    }
  });

  editor.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.stopPropagation();
  });

  editor.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounter = 0;
    editor.classList.remove('drag-over');

    const files = e.dataTransfer.files;
    if (files.length === 0) return;

    const file = files[0];
    
    // Check if it's a markdown file
    const validExtensions = ['.md', '.markdown', '.mdown', '.mkd', '.txt'];
    const fileName = file.name.toLowerCase();
    const isValidFile = validExtensions.some(ext => fileName.endsWith(ext));

    if (!isValidFile) {
      showToast('Please drop a Markdown file (.md, .markdown, .mdown, .mkd, .txt)', 'warn');
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      const content = event.target.result;
      editor.value = content;
      
      // Set the document title from the filename (without extension)
      const titleWithoutExt = file.name.replace(/\.[^/.]+$/, '');
      docTitle.value = titleWithoutExt;
      
      // Update preview and save
      updatePreview();
      saveSession();
      updateStats();
      
      showToast(`Loaded: ${file.name}`, 'success');
    };

    reader.onerror = () => {
      showToast('Failed to read file', 'error');
    };

    reader.readAsText(file);
  });

  // ───── Keyboard Shortcuts ─────
  document.addEventListener('keydown', (e) => {
    // Ctrl/Cmd + S — Export PDF
    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
      e.preventDefault();
      document.getElementById('btn-export-pdf').click();
    }
    // Ctrl/Cmd + Shift + S — Export HTML
    if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'S') {
      e.preventDefault();
      document.getElementById('btn-export-html').click();
    }
  });

  // ───── Support Popup ─────
  const supportPopup = document.getElementById('support-popup');
  const supportPopupClose = document.getElementById('support-popup-close');
  
  if (supportPopup) {
    // Show popup with a slight delay
    setTimeout(() => {
      supportPopup.style.display = 'flex';
    }, 3000);
  }
  
  if (supportPopupClose) {
    supportPopupClose.addEventListener('click', () => {
      supportPopup.style.display = 'none';
    });
  }
})();
