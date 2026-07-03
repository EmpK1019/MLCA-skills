#!/usr/bin/env node
'use strict';

/**
 * merge_html_decks.js
 *
 * Merge multiple batch HTML deck files into a single self-contained deck.html.
 *
 * Usage:
 *   node merge_html_decks.js --output deck.html deck-batch1.html deck-batch2.html ...
 *   node merge_html_decks.js --output deck.html --dir . --prefix deck-batch
 *
 * Each batch HTML must contain <div class="slide"> elements inside <body>.
 * The first batch provides the <head> (CSS, JS, theme) as the skeleton.
 * All slide divs from subsequent batches are appended.
 * Navigation and progress bar are rebuilt to cover all merged slides.
 */

const fs = require('fs');
const path = require('path');

function parseArgs(argv) {
  const args = argv.slice(2);
  let output = 'deck.html';
  let dir = '';
  let prefix = 'deck-batch';
  const files = [];

  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--output' && args[i + 1]) {
      output = args[++i];
    } else if (args[i] === '--dir' && args[i + 1]) {
      dir = args[++i];
    } else if (args[i] === '--prefix' && args[i + 1]) {
      prefix = args[++i];
    } else if (!args[i].startsWith('--')) {
      files.push(args[i]);
    }
  }

  if (files.length === 0 && dir) {
    const allFiles = fs.readdirSync(dir).sort();
    for (const f of allFiles) {
      if (f.startsWith(prefix) && f.endsWith('.html')) {
        files.push(path.join(dir, f));
      }
    }
  }

  return { output, files };
}

function extractSlides(html) {
  const slides = [];
  const regex = /<div\s+class="slide[^"]*"[^>]*>[\s\S]*?<\/div>\s*(?=<div\s+class="slide|$)/g;
  let match;
  while ((match = regex.exec(html)) !== null) {
    slides.push(match[0]);
  }

  if (slides.length === 0) {
    const fallbackRegex = /<div\s+class="slide[^"]*"[^>]*>[\s\S]*?<\/div>/g;
    while ((match = fallbackRegex.exec(html)) !== null) {
      slides.push(match[0]);
    }
  }

  return slides;
}

function extractHead(html) {
  const match = html.match(/<head[\s\S]*?<\/head>/i);
  return match ? match[0] : '';
}

function extractBodyContent(html) {
  const match = html.match(/<body[^>]*>([\s\S]*)<\/body>/i);
  return match ? match[1].trim() : '';
}

function buildMergedHtml(headHtml, allSlides) {
  const slideCount = allSlides.length;

  const navJs = `
  <script>
  (function() {
    let current = 0;
    const slides = document.querySelectorAll('.slide');
    const total = slides.length;
    const progressBar = document.getElementById('deck-progress');
    const counterEl = document.getElementById('deck-counter');

    function showSlide(n) {
      if (n < 0) n = 0;
      if (n >= total) n = total - 1;
      current = n;
      slides.forEach((s, i) => {
        s.style.display = i === current ? 'flex' : 'none';
        s.classList.toggle('active', i === current);
      });
      if (progressBar) progressBar.style.width = ((current + 1) / total * 100) + '%';
      if (counterEl) counterEl.textContent = (current + 1) + ' / ' + total;
    }

    document.addEventListener('keydown', function(e) {
      if (e.key === 'ArrowRight' || e.key === 'ArrowDown' || e.key === ' ') {
        e.preventDefault();
        showSlide(current + 1);
      } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
        e.preventDefault();
        showSlide(current - 1);
      } else if (e.key === 'Home') {
        e.preventDefault();
        showSlide(0);
      } else if (e.key === 'End') {
        e.preventDefault();
        showSlide(total - 1);
      }
    });

    document.addEventListener('click', function(e) {
      if (e.target.closest('a, button, input, select, textarea')) return;
      const w = window.innerWidth;
      if (e.clientX > w * 0.65) showSlide(current + 1);
      else if (e.clientX < w * 0.35) showSlide(current - 1);
    });

    showSlide(0);
  })();
  </script>`;

  return `<!DOCTYPE html>
<html lang="zh-CN">
${headHtml}
<body>
  <div id="deck-container" style="position:relative;width:100%;height:100vh;overflow:hidden;background:#1a1a2e;">
    <div id="deck-progress" style="position:fixed;top:0;left:0;height:3px;background:linear-gradient(90deg,#4facfe,#00f2fe);z-index:1000;transition:width 0.3s ease;"></div>
    <div id="deck-counter" style="position:fixed;bottom:16px;right:24px;font-size:13px;color:rgba(255,255,255,0.5);z-index:1000;font-family:system-ui,sans-serif;"></div>
${allSlides.map(s => '    ' + s).join('\n')}
  </div>
${navJs}
</body>
</html>`;
}

function main() {
  const { output, files } = parseArgs(process.argv);

  if (files.length === 0) {
    console.error('No batch files found to merge.');
    process.exit(1);
  }

  console.log(`Merging ${files.length} batch file(s) into ${output}...`);

  let headHtml = '';
  const allSlides = [];

  for (const filePath of files) {
    if (!fs.existsSync(filePath)) {
      console.warn(`Warning: ${filePath} not found, skipping.`);
      continue;
    }
    const html = fs.readFileSync(filePath, 'utf-8');

    if (!headHtml) {
      headHtml = extractHead(html);
    }

    const slides = extractSlides(html);
    console.log(`  ${path.basename(filePath)}: ${slides.length} slide(s)`);
    allSlides.push(...slides);
  }

  if (allSlides.length === 0) {
    console.error('No slides found in any batch file.');
    process.exit(1);
  }

  const merged = buildMergedHtml(headHtml, allSlides);
  fs.writeFileSync(output, merged, 'utf-8');
  console.log(`Done: ${output} with ${allSlides.length} slide(s).`);
}

main();
