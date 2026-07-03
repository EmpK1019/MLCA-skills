/**
 * export_slides.js
 *
 * Exports PPTX slides to PNG images.
 * Strategy: PowerPoint COM (Windows) → LibreOffice headless → fail
 *
 * Usage (from engine):
 *   node export_slides.js --input path/to/deck.pptx --output dir/for/images [--width 1920] [--height 1080]
 *
 * Output: JSON to stdout
 *   { success: true, slideCount: N, images: ["slide_001.png", ...], method: "powerpoint" | "libreoffice" }
 *   { success: false, error: "..." }
 */

const { execFile, exec } = require('child_process');
const fs = require('fs');
const path = require('path');

// Module-level config (shared by CLI and programmatic API)
let inputFile = '';
let outputDir = '';
let width = 1920;
let height = 1080;

// ── Strategy 1: PowerPoint export (platform-specific) ──────────────────
function tryPowerPoint() {
  const platform = process.platform;

  if (platform === 'win32') {
    return tryPowerPointWindows();
  } else if (platform === 'darwin') {
    return tryPowerPointMac();
  }
  return Promise.reject(new Error(`PowerPoint COM not available on ${platform}`));
}

// ── Windows: PowerShell + PowerPoint COM ────────────────────────────────
function tryPowerPointWindows() {
  return new Promise((resolve, reject) => {
    const ps1Path = path.join(__dirname, 'export_slides_to_images.ps1');
    if (!fs.existsSync(ps1Path)) {
      reject(new Error('PowerShell script not found'));
      return;
    }

    const cmd = `powershell -ExecutionPolicy Bypass -File "${ps1Path}" -InputFile "${inputFile}" -OutputDir "${outputDir}" -Width ${width} -Height ${height}`;
    exec(cmd, { timeout: 120000 }, (err, stdout, stderr) => {
      if (err) {
        reject(new Error(`PowerPoint COM failed: ${err.message}`));
        return;
      }
      try {
        const result = JSON.parse(stdout.trim());
        if (result.success) {
          resolve({ ...result, method: 'powerpoint' });
        } else {
          reject(new Error(result.error || 'PowerPoint export returned failure'));
        }
      } catch (e) {
        reject(new Error(`Failed to parse PowerShell output: ${e.message}`));
      }
    });
  });
}

// ── macOS: AppleScript + PowerPoint ─────────────────────────────────────
function tryPowerPointMac() {
  return new Promise((resolve, reject) => {
    const script = `
tell application "Microsoft PowerPoint"
  set thePresentation to open POSIX file "${inputFile}"
  set slideCount to count of slides of thePresentation
  set fileList to {}
  repeat with i from 1 to slideCount
    set theSlide to slide i of thePresentation
    set thePath to POSIX path of (outputDir & "/slide_" & text -3 thru -1 of ("000" & i) & ".png")
    export theSlide in thePath as PNG
    set end of fileList to "slide_" & text -3 thru -1 of ("000" & i) & ".png"
  end repeat
  close thePresentation saving no
  return fileList
end tell`.replace(/\$\{inputFile\}/g, inputFile.replace(/"/g, '\\"'))
             .replace(/\$\{outputDir\}/g, outputDir.replace(/"/g, '\\"'));

    const cmd = `osascript -e '${script.replace(/'/g, "'\\''")}'`;
    exec(cmd, { timeout: 120000 }, (err, stdout, stderr) => {
      if (err) {
        reject(new Error(`PowerPoint AppleScript failed: ${err.message}`));
        return;
      }
      // Check output directory for PNG files
      try {
        const files = fs.readdirSync(outputDir).filter(f => f.endsWith('.png')).sort();
        if (files.length === 0) {
          reject(new Error('PowerPoint exported no PNG files'));
          return;
        }
        resolve({
          success: true,
          slideCount: files.length,
          files,
          method: 'powerpoint',
        });
      } catch (e) {
        reject(new Error(`Failed to read output directory: ${e.message}`));
      }
    });
  });
}

// ── Strategy 2: LibreOffice headless ───────────────────────────────────
function findSoffice() {
  // Environment variable override
  if (process.env.LIBREOFFICE_PATH && fs.existsSync(process.env.LIBREOFFICE_PATH)) {
    return process.env.LIBREOFFICE_PATH;
  }

  // Check bundled LibreOffice
  const bundled = process.platform === 'win32'
    ? path.join(__dirname, '..', 'libreoffice', 'program', 'soffice.exe')
    : path.join(__dirname, '..', 'libreoffice', 'program', 'soffice');
  if (fs.existsSync(bundled)) return bundled;

  // Check platform-specific install paths
  const candidates = process.platform === 'win32' ? [
    // Portable (common locations)
    'D:\\tmp\\LibreOfficePortable\\App\\libreoffice\\program\\soffice.exe',
    'D:\\LibreOfficePortable\\App\\libreoffice\\program\\soffice.exe',
    // Standard install
    'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, 'Programs', 'LibreOffice', 'program', 'soffice.exe'),
    process.env.PROGRAMFILES && path.join(process.env.PROGRAMFILES, 'LibreOffice', 'program', 'soffice.exe'),
  ].filter(Boolean) : process.platform === 'darwin' ? [
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    '/usr/local/bin/soffice',
  ] : [
    '/usr/bin/soffice',
    '/usr/local/bin/soffice',
  ];

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) return candidate;
  }

  // Try PATH
  return 'soffice';
}

async function tryLibreOffice() {
  const soffice = findSoffice();
  const tmpDir = path.join(outputDir, '__lo_tmp');

  return new Promise((resolve, reject) => {
    // LibreOffice can convert to PDF, then we need pdf2image
    // Step 1: Convert PPTX → PDF
    const cmd = `"${soffice}" --headless --convert-to pdf --outdir "${tmpDir}" "${inputFile}"`;
    exec(cmd, { timeout: 120000 }, (err, stdout, stderr) => {
      if (err) {
        reject(new Error(`LibreOffice conversion failed: ${err.message}`));
        return;
      }

      // Find the generated PDF
      const pdfName = path.basename(inputFile, path.extname(inputFile)) + '.pdf';
      const pdfPath = path.join(tmpDir, pdfName);

      if (!fs.existsSync(pdfPath)) {
        // Try finding any PDF in tmpDir
        const pdfs = fs.readdirSync(tmpDir).filter(f => f.endsWith('.pdf'));
        if (pdfs.length === 0) {
          reject(new Error('LibreOffice did not produce a PDF file'));
          return;
        }
      }

      resolve({ soffice, pdfPath: pdfPath, tmpDir, method: 'libreoffice' });
    });
  }).then(async (loResult) => {
    // Step 2: Convert PDF pages to PNG using pdfjs-dist
    const images = await pdfToImages(loResult.pdfPath, outputDir, width);
    // Cleanup temp dir
    try { fs.rmSync(loResult.tmpDir, { recursive: true, force: true }); } catch {}
    return { success: true, slideCount: images.length, images, method: 'libreoffice' };
  });
}

// ── PDF → Images using pdfjs-dist (Node.js) ────────────────────────────
async function pdfToImages(pdfPath, outDir, targetWidth) {
  // Try to load pdfjs-dist from desktop node_modules
  let pdfjs;
  const pdfjsPaths = [
    path.resolve(__dirname, '..', '..', '..', '..', 'desktop', 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.mjs'),
    path.resolve(__dirname, '..', '..', '..', '..', 'desktop', 'node_modules', 'pdfjs-dist'),
  ];

  // Try canvas for Node.js rendering
  let canvas;
  try { canvas = require('canvas'); } catch {}

  if (!canvas) {
    throw new Error('node-canvas not available for PDF rendering');
  }

  // Load pdfjs-dist
  for (const p of pdfjsPaths) {
    try { pdfjs = require(p); break; } catch {}
  }
  if (!pdfjs) {
    try { pdfjs = require('pdfjs-dist'); } catch {}
  }
  if (!pdfjs) {
    throw new Error('pdfjs-dist not available');
  }

  const data = new Uint8Array(fs.readFileSync(pdfPath));
  const doc = await pdfjs.getDocument({ data }).promise;
  const images = [];

  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const viewport = page.getViewport({ scale: targetWidth / page.getViewport({ scale: 1.0 }).width });
    const canvasInstance = canvas.createCanvas(viewport.width, viewport.height);
    const ctx = canvasInstance.getContext('2d');
    await page.render({ canvasContext: ctx, viewport }).promise;
    const filename = `slide_${String(i).padStart(3, '0')}.png`;
    const outPath = path.join(outDir, filename);
    fs.writeFileSync(outPath, canvasInstance.toBuffer('image/png'));
    images.push(filename);
  }

  return images;
}

// ── Programmatic API ───────────────────────────────────────────────────
/**
 * Export PPTX slides to PNG images (for require() use).
 * Temporarily swaps module-level vars so strategy functions work unchanged.
 * @param {object} opts
 * @param {string} opts.inputFile - Absolute path to the PPTX file.
 * @param {string} opts.outputDir - Directory to write PNG images into.
 * @param {number} [opts.width=1920] - Output image width.
 * @param {number} [opts.height=1080] - Output image height.
 * @returns {Promise<{success: boolean, slideCount: number, images: string[], method: string, error?: string}>}
 */
async function exportSlides(opts) {
  const savedInput = inputFile;
  const savedOutput = outputDir;
  const savedWidth = width;
  const savedHeight = height;

  inputFile = path.resolve(opts.inputFile);
  outputDir = path.resolve(opts.outputDir);
  width = opts.width || 1920;
  height = opts.height || 1080;

  if (!fs.existsSync(inputFile)) {
    inputFile = savedInput; outputDir = savedOutput; width = savedWidth; height = savedHeight;
    return { success: false, slideCount: 0, images: [], method: '', error: `Input file not found: ${opts.inputFile}` };
  }
  fs.mkdirSync(outputDir, { recursive: true });

  try {
    const result = await tryPowerPoint();
    return { ...result, images: result.files || result.images || [] };
  } catch (ppErr) {
    // PowerPoint not available
  }

  try {
    const result = await tryLibreOffice();
    return { ...result, images: result.images || [] };
  } catch (loErr) {
    // LibreOffice not available
  }

  inputFile = savedInput; outputDir = savedOutput; width = savedWidth; height = savedHeight;
  return { success: false, slideCount: 0, images: [], method: '', error: 'Neither PowerPoint nor LibreOffice is available for slide export' };
}

// ── Main (CLI) ─────────────────────────────────────────────────────────
async function main() {
  const args = process.argv.slice(2);
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--input' && args[i + 1]) inputFile = path.resolve(args[++i]);
    if (args[i] === '--output' && args[i + 1]) outputDir = path.resolve(args[++i]);
    if (args[i] === '--width' && args[i + 1]) width = parseInt(args[++i], 10);
    if (args[i] === '--height' && args[i + 1]) height = parseInt(args[++i], 10);
  }

  if (!inputFile || !outputDir) {
    console.log(JSON.stringify({ success: false, error: '--input and --output are required' }));
    process.exit(1);
  }

  if (!fs.existsSync(inputFile)) {
    console.log(JSON.stringify({ success: false, error: `Input file not found: ${inputFile}` }));
    process.exit(1);
  }

  fs.mkdirSync(outputDir, { recursive: true });

  // Try PowerPoint COM first (best quality, fastest)
  try {
    const result = await tryPowerPoint();
    console.log(JSON.stringify(result));
    return;
  } catch (ppErr) {
    // PowerPoint not available or failed
  }

  // Try LibreOffice headless
  try {
    const result = await tryLibreOffice();
    console.log(JSON.stringify(result));
    return;
  } catch (loErr) {
    // LibreOffice not available either
  }

  console.log(JSON.stringify({
    success: false,
    error: 'Neither PowerPoint nor LibreOffice is available for slide export',
  }));
  process.exit(1);
}

if (require.main === module) {
  main().catch((err) => {
    console.log(JSON.stringify({ success: false, error: err.message }));
    process.exit(1);
  });
}

module.exports = { exportSlides };