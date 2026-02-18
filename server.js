const express = require('express');
const path = require('path');
const puppeteer = require('puppeteer');
const fs = require('fs');
const { Marked } = require('marked');
const hljs = require('highlight.js');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, ImageRun } = require('docx');
const helmet = require('helmet');
const compression = require('compression');
const rateLimit = require('express-rate-limit');

const app = express();
const PORT = process.env.PORT || 3000;

// Helper function to create Content-Disposition header with Unicode support
function getContentDisposition(filename) {
  // RFC 5987 encoding for Unicode filenames
  // Use ASCII fallback + UTF-8 encoded filename
  const encodedFilename = encodeURIComponent(filename);
  return `attachment; filename="${filename.replace(/[^\x00-\x7F]/g, '_')}"; filename*=UTF-8''${encodedFilename}`;
}

// Security middleware
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      scriptSrc: ["'self'", "'unsafe-inline'", "cdnjs.cloudflare.com", "cdn.jsdelivr.net"],
      styleSrc: ["'self'", "'unsafe-inline'", "cdnjs.cloudflare.com", "fonts.googleapis.com"],
      fontSrc: ["'self'", "fonts.gstatic.com", "cdnjs.cloudflare.com"],
      imgSrc: ["'self'", "data:", "blob:"],
      connectSrc: ["'self'"],
    },
  },
  crossOriginEmbedderPolicy: false,
}));

// Gzip/Brotli compression for all responses (Core Web Vitals)
app.use(compression());

// Rate limiting - 100 requests per 15 minutes per IP
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 100,
  message: { error: 'Too many requests, please try again later.' },
  standardHeaders: true,
  legacyHeaders: false,
});
app.use('/export/', limiter);

// Stricter rate limit for export endpoints - 20 exports per 15 minutes
const exportLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 20,
  message: { error: 'Too many export requests, please try again later.' },
  standardHeaders: true,
  legacyHeaders: false,
});
app.use('/export/pdf', exportLimiter);
app.use('/export/docx', exportLimiter);

app.use(express.json({ limit: '50mb' }));

// SEO-friendly static file serving with caching
app.use(express.static(path.join(__dirname, 'public'), {
  maxAge: '1d', // Cache static assets for 1 day
  etag: true,
  lastModified: true,
}));

// SEO headers for all routes
app.use((req, res, next) => {
  // Remove X-Powered-By header for security
  res.removeHeader('X-Powered-By');
  
  // Add SEO-friendly headers
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'DENY');
  res.setHeader('X-XSS-Protection', '1; mode=block');
  
  // Cache control for HTML pages
  if (req.path === '/' || req.path.endsWith('.html')) {
    res.setHeader('Cache-Control', 'public, max-age=3600'); // 1 hour for HTML
  }
  
  next();
});

// Dynamic sitemap for better SEO
app.get('/sitemap.xml', (req, res) => {
  const baseUrl = 'https://markdown-editor-viewer.onrender.com';
  const today = new Date().toISOString().split('T')[0];
  
  const sitemap = `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
  <url>
    <loc>${baseUrl}/</loc>
    <lastmod>${today}</lastmod>
    <changefreq>weekly</changefreq>
    <priority>1.0</priority>
  </url>
</urlset>`;
  
  res.setHeader('Content-Type', 'application/xml; charset=utf-8');
  res.setHeader('Cache-Control', 'public, max-age=86400'); // 1 day
  res.send(sitemap);
});

// Generate PNG from SVG for OG image (social platforms don't support SVG)
let ogImageCache = null;
app.get('/og-image.png', async (req, res) => {
  try {
    if (ogImageCache) {
      res.setHeader('Content-Type', 'image/png');
      res.setHeader('Cache-Control', 'public, max-age=604800'); // 7 days
      return res.end(ogImageCache);
    }
    const svgPath = path.join(__dirname, 'public', 'og-image.svg');
    const svgContent = fs.readFileSync(svgPath, 'utf-8');
    const b = await getBrowser();
    const page = await b.newPage();
    await page.setViewport({ width: 1200, height: 630 });
    await page.setContent(`<!DOCTYPE html><html><body style="margin:0;padding:0;">${svgContent}</body></html>`, { waitUntil: 'networkidle0' });
    const pngBuffer = await page.screenshot({ type: 'png', clip: { x: 0, y: 0, width: 1200, height: 630 } });
    await page.close();
    ogImageCache = Buffer.from(pngBuffer);
    res.setHeader('Content-Type', 'image/png');
    res.setHeader('Cache-Control', 'public, max-age=604800');
    res.end(ogImageCache);
  } catch (err) {
    console.error('OG image generation error:', err);
    // Fall back to SVG
    res.redirect('/og-image.svg');
  }
});

// Health check endpoint for monitoring
app.get('/health', (req, res) => {
  res.status(200).json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Persistent Puppeteer browser instance
let browser = null;

async function getBrowser() {
  if (!browser || !browser.connected) {
    const launchOptions = {
      headless: 'new',
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-gpu',
        '--font-render-hinting=none',
      ],
    };
    // Use system Chromium in Docker/deployment environments
    if (process.env.PUPPETEER_EXECUTABLE_PATH) {
      launchOptions.executablePath = process.env.PUPPETEER_EXECUTABLE_PATH;
    }
    browser = await puppeteer.launch(launchOptions);
  }
  return browser;
}

// Read the PDF template once at startup
const pdfTemplate = fs.readFileSync(
  path.join(__dirname, 'templates', 'pdf-template.html'),
  'utf-8'
);

// Create a configured Marked instance for server-side rendering
function createMarkedInstance() {
  const marked = new Marked();
  marked.use({
    renderer: {
      code({ text, lang }) {
        if (lang === 'mermaid') {
          // Wrap mermaid in a div for better page break control
          return `<div class="mermaid-wrapper"><pre class="mermaid">${text}</pre></div>`;
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
  return marked;
}

const marked = createMarkedInstance();

// Render Mermaid diagrams to PNG images using Puppeteer
async function renderMermaidToImage(mermaidCode, index) {
  const browser = await getBrowser();
  let page = null;
  
  try {
    page = await browser.newPage();
    
    // Set viewport for high-quality rendering
    await page.setViewport({ width: 1200, height: 800, deviceScaleFactor: 2 });
    
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <script src="https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js"></script>
        <style>
          body { margin: 0; padding: 20px; background: white; }
          .mermaid { display: flex; justify-content: center; }
        </style>
      </head>
      <body>
        <pre class="mermaid">${mermaidCode.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</pre>
        <script>
          mermaid.initialize({ 
            startOnLoad: true, 
            theme: 'default',
            securityLevel: 'loose',
            flowchart: { useMaxWidth: false, htmlLabels: true, curve: 'basis' },
            sequence: { useMaxWidth: false, wrap: true },
            gantt: { useMaxWidth: false },
          });
        </script>
      </body>
      </html>
    `;
    
    await page.setContent(html, { waitUntil: 'networkidle0', timeout: 15000 });
    
    // Wait for mermaid to render
    await page.waitForSelector('.mermaid svg', { timeout: 10000 });
    await new Promise(r => setTimeout(r, 500)); // Extra wait for rendering
    
    // Get the SVG element
    const svgElement = await page.$('.mermaid svg');
    if (!svgElement) {
      throw new Error('SVG element not found');
    }
    
    // Get bounding box
    const boundingBox = await svgElement.boundingBox();
    if (!boundingBox) {
      throw new Error('Could not get bounding box');
    }
    
    // Take screenshot of just the SVG
    const imageBuffer = await svgElement.screenshot({
      type: 'png',
      omitBackground: true,
    });
    
    return {
      buffer: imageBuffer,
      width: boundingBox.width,
      height: boundingBox.height,
    };
  } catch (error) {
    console.error(`Error rendering mermaid diagram ${index}:`, error.message);
    return null;
  } finally {
    if (page) await page.close().catch(() => {});
  }
}

// Helper function to convert markdown to docx elements
// Matches PDF styling exactly - colors, borders, layout
async function parseMarkdownToDocx(markdown) {
  const elements = [];
  const lines = markdown.split('\n');
  let inCodeBlock = false;
  let codeContent = [];
  let codeLanguage = '';
  let inTable = false;
  let tableRows = [];
  let tableColumnCount = 0;
  let currentList = [];
  let inOrderedList = false;
  let mermaidDiagrams = [];

  // Preview-matching colors (light blue theme)
  const COLORS = {
    text: '1A1A2E',           // Main text color
    border: 'D0D7DE',         // Border color (light gray)
    accent: '0969DA',         // Blue accent (links, blockquote border)
    codeBg: 'F6F8FA',         // Code background
    codeInlineBg: 'EFF1F3',   // Inline code background
    codeInlineText: 'C7254E', // Inline code text color
    blockquoteBg: 'F0F7FF',   // Blockquote background (light blue)
    tableHeaderBg: 'F6F8FA',  // Table header background
    tableStripeBg: 'F8F9FA',  // Table stripe background
    mutedText: '656D76',      // Muted/secondary text
  };

  // Process inline markdown formatting (bold, italic, code, links)
  const processInlineFormatting = (text) => {
    const runs = [];
    let remaining = text;

    while (remaining.length > 0) {
      // Find all potential matches - more flexible patterns
      const boldMatch = remaining.match(/^(\*\*.+?\*\*)/);
      const italicMatch = remaining.match(/^(\*.+?\*)/);
      const codeMatch = remaining.match(/^(`.+?`)/);
      const linkMatch = remaining.match(/^(\[.+?\]\(.+?\))/);

      let match = null;
      let type = null;

      // Check what matches at the start of remaining text
      if (linkMatch) {
        match = linkMatch;
        type = 'link';
      } else if (boldMatch) {
        match = boldMatch;
        type = 'bold';
      } else if (italicMatch) {
        match = italicMatch;
        type = 'italic';
      } else if (codeMatch) {
        match = codeMatch;
        type = 'code';
      }

      if (match) {
        // Extract the content (without the markdown markers)
        let content = match[1];
        
        if (type === 'link') {
          // Links: [text](url) -> extract text
          const linkTextMatch = content.match(/\[(.+?)\]/);
          if (linkTextMatch) {
            content = linkTextMatch[1];
          }
        } else if (type === 'bold') {
          // Bold: **text** -> text
          content = content.slice(2, -2);
        } else if (type === 'italic') {
          // Italic: *text* -> text
          content = content.slice(1, -1);
        } else if (type === 'code') {
          // Code: `code` -> code
          content = content.slice(1, -1);
        }
        
        if (type === 'link') {
          // Links in Word - blue color, underlined
          runs.push(new TextRun({ 
            text: content, 
            size: 22, // 11pt
            color: '0969DA',
            underline: {},
            font: 'Calibri',
          }));
        } else if (type === 'bold') {
          runs.push(new TextRun({ 
            text: content, 
            size: 22, // 11pt
            bold: true,
            font: 'Calibri',
          }));
        } else if (type === 'italic') {
          runs.push(new TextRun({ 
            text: content, 
            size: 22, // 11pt
            italics: true,
            font: 'Calibri',
          }));
        } else if (type === 'code') {
          runs.push(new TextRun({ 
            text: content, 
            size: 18, // 9pt for inline code
            font: 'Consolas',
            shading: { fill: 'EFF1F3' },
            color: 'C7254E',
          }));
        }
        remaining = remaining.slice(match[0].length);
      } else {
        // No match, add remaining text as-is
        runs.push(new TextRun({ 
          text: remaining, 
          size: 22, // 11pt
          font: 'Calibri',
        }));
        break;
      }
    }
    return runs;
  };

  const flushList = () => {
    if (currentList.length > 0) {
      currentList.forEach((item, index) => {
        // Process inline formatting for list items
        const runs = processInlineFormatting(item);
        
        // Add bullet/number prefix
        const prefix = inOrderedList ? `${index + 1}. ` : `• `;
        
        elements.push(new Paragraph({
          children: [
            new TextRun({
              text: prefix,
              size: 22,
              font: 'Calibri',
            }),
            ...runs
          ],
          indent: { left: 360 }, // Smaller indent for lists
          spacing: { after: 40, line: 300 },
        }));
      });
      currentList = [];
    }
  };

  const flushTable = () => {
    if (tableRows.length > 0) {
      const docxTable = new Table({
        rows: tableRows.map((row, rowIndex) => {
          while (row.length < tableColumnCount) {
            row.push('');
          }
          
          return new TableRow({
            children: row.map((cell, cellIndex) => {
              const cellContent = cell.trim();
              // Process inline formatting in table cells
              const cellRuns = processInlineFormatting(cellContent);
              if (cellRuns.length === 0) {
                cellRuns.push(new TextRun({ text: '', size: 26 }));
              }
              
              return new TableCell({
                children: [new Paragraph({
                  children: cellRuns.map(run => {
                    // Adjust size for table cells (slightly smaller)
                    if (run.options) {
                      run.options.size = 20; // 10pt
                    }
                    return run;
                  }),
                  spacing: { after: 0 },
                })],
                shading: rowIndex === 0 
                  ? { fill: 'F6F8FA' } 
                  : (rowIndex % 2 === 0 ? { fill: 'F8F9FA' } : undefined),
                margins: { 
                  top: 80, 
                  bottom: 80, 
                  left: 100, 
                  right: 100 
                },
                width: { 
                  size: tableColumnCount > 0 ? Math.floor(100 / tableColumnCount) : 20, 
                  type: WidthType.PERCENTAGE 
                },
              });
            }),
          });
        }),
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 1, color: 'D0D7DE' },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: 'D0D7DE' },
          left: { style: BorderStyle.SINGLE, size: 1, color: 'D0D7DE' },
          right: { style: BorderStyle.SINGLE, size: 1, color: 'D0D7DE' },
          insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'D0D7DE' },
          insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'D0D7DE' },
        },
      });
      elements.push(docxTable);
      elements.push(new Paragraph({ children: [], spacing: { after: 200 } }));
      tableRows = [];
      tableColumnCount = 0;
    }
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Mermaid diagram handling
    if (line.startsWith('```mermaid')) {
      inCodeBlock = true;
      codeContent = [];
      codeLanguage = 'mermaid';
      continue;
    }

    // Code block handling
    if (line.startsWith('```')) {
      if (!inCodeBlock) {
        inCodeBlock = true;
        codeLanguage = line.slice(3).trim();
        codeContent = [];
      } else {
        if (codeLanguage === 'mermaid') {
          mermaidDiagrams.push(codeContent.join('\n'));
          elements.push(new Paragraph({
            children: [new TextRun({
              text: '[Mermaid Diagram]',
              size: 18, // 9pt
              color: '666666',
            })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 160, after: 160 },
          }));
        } else {
          // Code block with background and border (matching PDF)
          const codeLines = codeContent.join('\n').split('\n');
          
          // Add top border paragraph
          elements.push(new Paragraph({
            children: [],
            spacing: { before: 100, after: 0 },
          }));
          
          codeLines.forEach((codeLine, idx) => {
            elements.push(new Paragraph({
              children: [new TextRun({
                text: codeLine || ' ', // Use space for empty lines
                font: 'Consolas',
                size: 18, // 9pt
              })],
              shading: { fill: 'F6F8FA' },
              spacing: { after: 20, line: 280 }, // Proper line spacing
            }));
          });
          
          // Add bottom spacing
          elements.push(new Paragraph({ 
            children: [], 
            spacing: { after: 100 },
              shading: { fill: 'F6F8FA' },
          }));
        }
        inCodeBlock = false;
        codeLanguage = '';
      }
      continue;
    }

    if (inCodeBlock) {
      codeContent.push(line);
      continue;
    }

    // Table handling
    if (line.startsWith('|')) {
      if (!inTable) {
        inTable = true;
        tableRows = [];
        tableColumnCount = 0;
      }
      
      const rawCells = line.split('|');
      const cells = rawCells.slice(1, -1).map(c => c.trim());
      
      if (cells.some(c => /^[-:]+$/.test(c))) {
        if (cells.length > tableColumnCount) {
          tableColumnCount = cells.length;
        }
        continue;
      }
      
      if (cells.length > tableColumnCount) {
        tableColumnCount = cells.length;
      }
      
      tableRows.push(cells);
      continue;
    } else if (inTable) {
      flushTable();
      inTable = false;
    }

    // Headers - Word-style sizes (H1=18pt, H2=14pt, H3=12pt, H4=11pt)
    if (line.startsWith('# ')) {
      flushList();
      const isFirstElement = elements.length === 0;
      elements.push(new Paragraph({
        children: [new TextRun({ 
          text: line.slice(2), 
          bold: true, 
          size: 36, // 18pt
          font: 'Calibri',
          color: '1A1A2E',
        })],
        heading: HeadingLevel.HEADING_1,
        spacing: { before: isFirstElement ? 0 : 240, after: 120 },
        border: {
          bottom: { style: BorderStyle.SINGLE, size: 8, color: 'D0D7DE' }
        },
      }));
      continue;
    }
    if (line.startsWith('## ')) {
      flushList();
      elements.push(new Paragraph({
        children: [new TextRun({ 
          text: line.slice(3), 
          bold: true, 
          size: 28, // 14pt
          font: 'Calibri',
          color: '1A1A2E',
        })],
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 100 },
        border: { 
          bottom: { style: BorderStyle.SINGLE, size: 4, color: 'D0D7DE' } 
        },
      }));
      continue;
    }
    if (line.startsWith('### ')) {
      flushList();
      elements.push(new Paragraph({
        children: [new TextRun({ 
          text: line.slice(4), 
          bold: true, 
          size: 24, // 12pt
          font: 'Calibri',
          color: '1A1A2E',
        })],
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 160, after: 80 },
      }));
      continue;
    }
    if (line.startsWith('#### ')) {
      flushList();
      elements.push(new Paragraph({
        children: [new TextRun({ 
          text: line.slice(5), 
          bold: true, 
          size: 22, // 11pt
          font: 'Calibri',
          color: '1A1A2E',
        })],
        heading: HeadingLevel.HEADING_4,
        spacing: { before: 140, after: 60 },
      }));
      continue;
    }

    // List handling
    if (line.startsWith('- ') || line.startsWith('* ')) {
      flushList();
      inOrderedList = false;
      currentList.push(line.slice(2));
      continue;
    }

    if (/^\d+\.\s/.test(line)) {
      flushList();
      inOrderedList = true;
      currentList.push(line.replace(/^\d+\.\s/, ''));
      continue;
    }

    // Horizontal rule - Word-style
    if (line.match(/^(---|\*\*\*|___)$/)) {
      flushList();
      elements.push(new Paragraph({
        children: [],
        border: { 
          bottom: { style: BorderStyle.SINGLE, size: 8, color: 'D0D7DE' } 
        },
        spacing: { before: 200, after: 200 },
      }));
      continue;
    }

    // Blockquote - Word-style with blue left border and padding
    if (line.startsWith('> ')) {
      flushList();
      let quoteText = line.slice(2);
      // Check if there's already a leading space
      const hasLeadingSpace = quoteText.startsWith(' ');
      // Remove leading space for processing
      quoteText = quoteText.trimLeft();
      // Process inline formatting within blockquote
      const runs = processInlineFormatting(quoteText);
      // Add leading space back at the start if needed
      if (hasLeadingSpace || runs.length > 0) {
        runs.unshift(new TextRun({ 
          text: ' ', 
          size: 22,
          font: 'Calibri',
        }));
      }
      
      elements.push(new Paragraph({
        children: runs,
        border: { 
          left: { style: BorderStyle.SINGLE, size: 24, color: '0969DA' } 
        },
        spacing: { before: 80, after: 80, line: 300 },
        margin: { left: 240 }, // Add left margin for space from border
        shading: { fill: 'F0F7FF' },
      }));
      continue;
    }

    // Empty line
    if (line.trim() === '') {
      flushList();
      continue;
    }

    // Regular paragraph - Word-style
    flushList();

    const runs = processInlineFormatting(line);
    
    if (runs.length > 0) {
      elements.push(new Paragraph({
        children: runs,
        spacing: { after: 120, line: 300 }, // Compact line spacing
      }));
    }
  }

  flushList();
  flushTable();

  return { elements, mermaidDiagrams };
}


// Input validation helper
function validateInput(req, res, next) {
  const { html, markdown, title } = req.body;
  
  // Validate title if provided
  if (title && typeof title !== 'string') {
    return res.status(400).json({ error: 'Title must be a string' });
  }
  if (title && title.length > 500) {
    return res.status(400).json({ error: 'Title too long (max 500 characters)' });
  }
  
  // Validate html if provided
  if (html !== undefined) {
    if (typeof html !== 'string') {
      return res.status(400).json({ error: 'HTML content must be a string' });
    }
    if (html.length > 50 * 1024 * 1024) { // 50MB limit
      return res.status(400).json({ error: 'HTML content too large (max 50MB)' });
    }
  }
  
  // Validate markdown if provided
  if (markdown !== undefined) {
    if (typeof markdown !== 'string') {
      return res.status(400).json({ error: 'Markdown content must be a string' });
    }
    if (markdown.length > 50 * 1024 * 1024) { // 50MB limit
      return res.status(400).json({ error: 'Markdown content too large (max 50MB)' });
    }
  }
  
  next();
}

// Apply validation to export routes
app.use('/export/', validateInput);

// POST /export/pdf — Generate PDF from rendered HTML
app.post('/export/pdf', async (req, res) => {
  const { html, title } = req.body;

  if (!html) {
    return res.status(400).json({ error: 'Missing html content' });
  }

  let page = null;
  try {
    const b = await getBrowser();
    page = await b.newPage();

    const finalHtml = pdfTemplate.replace('{{CONTENT}}', html);

    // Use domcontentloaded instead of networkidle0 for faster, more reliable loading
    // networkidle0 can timeout on slow CDN resources or complex diagrams
    await page.setContent(finalHtml, { waitUntil: 'domcontentloaded', timeout: 60000 });

    // Wait for Mermaid diagrams to render (the template sets this flag)
    // Increased timeout for complex diagrams
    await page.waitForFunction(() => window.__MERMAID_RENDERED__ === true, {
      timeout: 60000,
    }).catch(() => {
      // If no mermaid diagrams or timeout, proceed anyway
    });

    // Extra wait for any remaining resources (images, fonts, SVG rendering)
    await new Promise((r) => setTimeout(r, 2000));

    const pdfUint8 = await page.pdf({
      format: 'A4',
      printBackground: true,
      preferCSSPageSize: false,
      margin: {
        top: '20mm',
        bottom: '20mm',
        left: '15mm',
        right: '15mm',
      },
      displayHeaderFooter: true,
      headerTemplate: '<span></span>',
      footerTemplate: `
        <div style="width:100%;text-align:center;font-size:9px;color:#888;padding:5px 0;">
          <span class="pageNumber"></span> / <span class="totalPages"></span>
        </div>`,
    });

    // Convert Uint8Array to Buffer so Express sends it as binary
    const pdfBuffer = Buffer.from(pdfUint8);
    // Allow Unicode characters (Japanese, etc.) in filename, only remove problematic chars
    const filename = (title || 'document').replace(/[<>:"/\\|?*\x00-\x1f]/g, '_') + '.pdf';

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': getContentDisposition(filename),
      'Content-Length': pdfBuffer.length,
    });
    res.end(pdfBuffer);
  } catch (err) {
    console.error('PDF generation error:', err);
    res.status(500).json({ error: 'Failed to generate PDF: ' + err.message });
  } finally {
    if (page) await page.close().catch(() => {});
  }
});

// POST /export/html — Generate self-contained HTML file
app.post('/export/html', async (req, res) => {
  const { markdown, title } = req.body;

  if (!markdown) {
    return res.status(400).json({ error: 'Missing markdown content' });
  }

  const renderedHtml = marked.parse(markdown);
  // Allow Unicode characters (Japanese, etc.) in filename, only remove problematic chars
  const filename = (title || 'document').replace(/[<>:"/\\|?*\x00-\x1f]/g, '_') + '.html';

  const fullHtml = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${title || 'Markdown Document'}</title>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.11.1/styles/github.min.css">
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
      max-width: 900px;
      margin: 0 auto;
      padding: 40px 20px;
      line-height: 1.7;
      color: #1a1a2e;
      background: #fff;
    }
    h1, h2, h3, h4, h5, h6 { margin-top: 1.5em; margin-bottom: 0.5em; font-weight: 600; }
    h1 { font-size: 2em; border-bottom: 2px solid #e1e4e8; padding-bottom: 0.3em; }
    h2 { font-size: 1.5em; border-bottom: 1px solid #e1e4e8; padding-bottom: 0.3em; }
    pre { background: #f6f8fa; border-radius: 6px; padding: 16px; overflow-x: auto; }
    code { font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace; font-size: 0.9em; }
    :not(pre) > code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; }
    blockquote { border-left: 4px solid #0969da; margin: 1em 0; padding: 0.5em 1em; background: #f0f7ff; }
    table { border-collapse: collapse; width: 100%; margin: 1em 0; }
    th, td { border: 1px solid #d0d7de; padding: 8px 12px; text-align: left; }
    th { background: #f6f8fa; font-weight: 600; }
    tr:nth-child(even) { background: #f8f9fa; }
    img { max-width: 100%; }
    a { color: #0969da; }
    hr { border: none; border-top: 2px solid #e1e4e8; margin: 2em 0; }
    .mermaid { text-align: center; margin: 1.5em 0; }
    .mermaid svg { max-width: 100%; height: auto; }
  </style>
</head>
<body>
  ${renderedHtml}
  <script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.11.1/highlight.min.js"><\/script>
  <script src="https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js"><\/script>
  <script>
    mermaid.initialize({ startOnLoad: true, theme: 'default', securityLevel: 'loose' });
  <\/script>
</body>
</html>`;

  res.set({
    'Content-Type': 'text/html; charset=utf-8',
    'Content-Disposition': getContentDisposition(filename),
  });
  res.send(fullHtml);
});

// POST /export/md — Return raw markdown as a downloadable file
app.post('/export/md', (req, res) => {
  const { markdown, title } = req.body;

  if (!markdown) {
    return res.status(400).json({ error: 'Missing markdown content' });
  }

  // Allow Unicode characters (Japanese, etc.) in filename, only remove problematic chars
  const filename = (title || 'document').replace(/[<>:"/\\|?*\x00-\x1f]/g, '_') + '.md';

  res.set({
    'Content-Type': 'text/markdown; charset=utf-8',
    'Content-Disposition': getContentDisposition(filename),
  });
  res.send(markdown);
});

// POST /export/docx — Generate Microsoft Word document
app.post('/export/docx', async (req, res) => {
  const { markdown, title } = req.body;

  if (!markdown) {
    return res.status(400).json({ error: 'Missing markdown content' });
  }

  try {
    const { elements: docxElements, mermaidDiagrams } = await parseMarkdownToDocx(markdown);

    // Render Mermaid diagrams to images
    const mermaidImages = [];
    for (let i = 0; i < mermaidDiagrams.length; i++) {
      const imageData = await renderMermaidToImage(mermaidDiagrams[i], i);
      mermaidImages.push(imageData);
    }

    // Replace placeholder paragraphs with actual images
    const finalElements = [];
    let mermaidIndex = 0;
    
    for (const element of docxElements) {
      // Check if this is a mermaid placeholder paragraph
      if (element.constructor.name === 'Paragraph') {
        const children = element.children || [];
        if (children.length === 1 && children[0].constructor.name === 'TextRun') {
          const text = children[0].text || '';
          if (text === '[Mermaid Diagram]' && mermaidImages[mermaidIndex]) {
            const imageData = mermaidImages[mermaidIndex];
            // Convert pixel dimensions to EMUs (English Metric Units)
            // 1 inch = 914400 EMUs, 1 pixel = 9525 EMUs at 96 DPI
            const maxWidthInches = 6; // Max width for Word document
            const maxWidthEmus = maxWidthInches * 914400;
            
            let widthEmus = Math.min(imageData.width * 9525, maxWidthEmus);
            let heightEmus = (imageData.height * 9525 * widthEmus) / (imageData.width * 9525);
            
            // Create paragraph with centered image
            finalElements.push(new Paragraph({
              children: [
                new ImageRun({
                  data: imageData.buffer,
                  transformation: {
                    width: widthEmus / 9525,
                    height: heightEmus / 9525,
                  },
                  type: 'png',
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { before: 200, after: 200 },
            }));
            mermaidIndex++;
            continue;
          }
        }
      }
      finalElements.push(element);
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: {
              top: 1134,    // 20mm in twips (20 * 56.7)
              right: 850,   // 15mm in twips
              bottom: 1134, // 20mm
              left: 850,    // 15mm
            },
          },
        },
        children: [
          ...finalElements,
        ],
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    // Allow Unicode characters (Japanese, etc.) in filename, only remove problematic chars
    const filename = (title || 'document').replace(/[<>:"/\\|?*\x00-\x1f]/g, '_') + '.docx';

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': getContentDisposition(filename),
      'Content-Length': buffer.length,
    });
    res.send(buffer);
  } catch (err) {
    console.error('DOCX generation error:', err);
    res.status(500).json({ error: 'Failed to generate Word document: ' + err.message });
  }
});

// Cleanup on exit
process.on('SIGINT', async () => {
  if (browser) await browser.close();
  process.exit();
});

process.on('SIGTERM', async () => {
  if (browser) await browser.close();
  process.exit();
});

app.listen(PORT, '0.0.0.0', () => {
  const nets = require('os').networkInterfaces();
  const ips = Object.values(nets).flat().filter(n => n.family === 'IPv4' && !n.internal).map(n => n.address);
  console.log(`\n  MD-to-PDF server running on:`);
  console.log(`    Local:   http://localhost:${PORT}`);
  ips.forEach(ip => console.log(`    Network: http://${ip}:${PORT}`));
  console.log();
});
