/**
 * audit_deck_structural.js
 *
 * Structural audit for PPTX decks — checks layout quality without rendering images.
 * Reads the audit JSON produced by render_deck.js and outputs a list of issues.
 *
 * Usage:
 *   node audit_deck_structural.js <audit_report.json>
 *
 * Output (stdout, JSON):
 *   {
 *     slideCount: N,
 *     issues: [ { slideIndex, severity, rule, message, suggestion } ],
 *     summary: { error: N, warning: N, info: N },
 *     passed: boolean
 *   }
 */

const fs = require('fs');
const path = require('path');

const SLIDE_W = 13.33; // WIDE layout
const SLIDE_H = 7.5;
const SLIDE_AREA = SLIDE_W * SLIDE_H;

function loadAudit(filePath) {
  const raw = fs.readFileSync(filePath, 'utf8');
  return JSON.parse(raw);
}

function elementArea(el) {
  return Math.max(0, (el.w || 0)) * Math.max(0, (el.h || 0));
}

function isTextElement(el) {
  return el.textLength > 0;
}

function isCardOrRect(el) {
  return (el.type || '').includes('rect') || (el.type || '').includes('rounded') || (el.type || '').includes('ROUNDED');
}

function auditSlide(slideAudit, slideIndex) {
  const issues = [];
  const els = slideAudit.elements || [];
  const kind = slideAudit.kind || 'unknown';
  const title = slideAudit.title || '';

  // ── 1. Page fill rate ──────────────────────────────────────────
  const totalElArea = els.reduce((sum, el) => sum + elementArea(el), 0);
  const fillRate = totalElArea / SLIDE_AREA;

  // Background rects are decorative, discount them
  const bgRects = els.filter(el => isCardOrRect(el) && !isTextElement(el));
  const bgArea = bgRects.reduce((sum, el) => sum + elementArea(el), 0);
  const contentArea = totalElArea - bgArea;
  const contentFillRate = contentArea / SLIDE_AREA;

  if (contentFillRate < 0.18) {
    issues.push({
      slideIndex,
      severity: 'error',
      rule: 'fill-rate',
      message: `Content fills only ${(contentFillRate * 100).toFixed(1)}% of slide (minimum 18%)`,
      suggestion: 'Increase element sizes, add more content, or reduce margins',
    });
  } else if (contentFillRate < 0.28) {
    issues.push({
      slideIndex,
      severity: 'warning',
      rule: 'fill-rate',
      message: `Content fills ${(contentFillRate * 100).toFixed(1)}% of slide — may look sparse`,
      suggestion: 'Consider expanding element sizes or reducing whitespace',
    });
  }

  // ── 2. Font size consistency ────────────────────────────────────
  const textEls = els.filter(isTextElement);
  const fontSizes = textEls.filter(el => el.fontSize).map(el => el.fontSize);
  const uniqueSizes = [...new Set(fontSizes)].sort((a, b) => b - a);

  // Group by role: titles (>14), body (10-14), small (<10)
  const titleSize = fontSizes.filter(s => s >= 14);
  const bodySize = fontSizes.filter(s => s >= 10 && s < 14);
  const smallSize = fontSizes.filter(s => s < 10);

  // Check title font size variance
  if (titleSize.length > 1) {
    const maxTitle = Math.max(...titleSize);
    const minTitle = Math.min(...titleSize);
    if (maxTitle - minTitle > 4) {
      issues.push({
        slideIndex,
        severity: 'warning',
        rule: 'title-font-consistency',
        message: `Title font sizes vary from ${minTitle}pt to ${maxTitle}pt (spread > 4pt)`,
        suggestion: 'Unify title font sizes across this slide',
      });
    }
  }

  // Check if any text is too small
  const tinyText = textEls.filter(el => el.fontSize && el.fontSize < 7);
  if (tinyText.length > 0) {
    issues.push({
      slideIndex,
      severity: 'warning',
      rule: 'tiny-text',
      message: `${tinyText.length} text element(s) with font size < 7pt may be unreadable`,
      suggestion: 'Increase font size to at least 7pt, or reduce content',
    });
  }

  // ── 3. Alignment checks ─────────────────────────────────────────
  // Check elements that should be horizontally aligned
  const xPositions = textEls.map(el => Math.round((el.x || 0) * 10) / 10);
  const xCounts = {};
  xPositions.forEach(x => { xCounts[x] = (xCounts[x] || 0) + 1; });
  const dominantX = Object.entries(xCounts).sort((a, b) => b[1] - a[1])[0];

  // Check elements that should share the same y position (same row)
  const yPositions = textEls.map(el => Math.round((el.y || 0) * 10) / 10);
  const yCounts = {};
  yPositions.forEach(y => { yCounts[y] = (yCounts[y] || 0) + 1; });

  // ── 4. Text overflow risk ────────────────────────────────────────
  for (const el of textEls) {
    if (!el.textLength || !el.w || !el.h) continue;
    const boxArea = el.w * el.h;
    // Rough estimate: at ~14pt, each char takes ~0.01 square inches
    // Chinese chars are wider, so use a higher estimate
    const avgCharArea = el.fontSize >= 14 ? 0.03 : el.fontSize >= 10 ? 0.02 : 0.015;
    const estimatedTextArea = el.textLength * avgCharArea;
    const ratio = estimatedTextArea / boxArea;

    if (ratio > 1.5) {
      issues.push({
        slideIndex,
        severity: 'error',
        rule: 'text-overflow-risk',
        message: `Text "${(el.textPreview || '').substring(0, 40)}" likely overflows its box (density ${ratio.toFixed(2)}x)`,
        suggestion: 'Reduce text length, increase box size, or decrease font size',
      });
    } else if (ratio > 1.0) {
      issues.push({
        slideIndex,
        severity: 'info',
        rule: 'text-overflow-risk',
        message: `Text "${(el.textPreview || '').substring(0, 40)}" may be tight (density ${ratio.toFixed(2)}x)`,
        suggestion: 'Consider reducing text or increasing box size',
      });
    }
  }

  // ── 5. Empty elements ─────────────────────────────────────────────
  const emptyText = textEls.filter(el => el.textLength === 0 || (el.textPreview || '').trim() === '');
  if (emptyText.length > 2) {
    issues.push({
      slideIndex,
      severity: 'info',
      rule: 'empty-elements',
      message: `${emptyText.length} empty text elements on this slide`,
      suggestion: 'Remove placeholder empty elements or fill them with content',
    });
  }

  // ── 6. Cover/closing specific checks ──────────────────────────────
  if (kind === 'cover') {
    const hasTitle = textEls.some(el => el.fontSize >= 20 && el.textLength > 0);
    if (!hasTitle) {
      issues.push({
        slideIndex,
        severity: 'warning',
        rule: 'cover-title',
        message: 'Cover slide appears to lack a large title',
        suggestion: 'Ensure the cover has a visible, large-font title',
      });
    }
  }

  if (kind === 'closing') {
    const hasMessage = textEls.some(el => el.textLength > 20 && el.fontSize >= 10);
    if (!hasMessage) {
      issues.push({
        slideIndex,
        severity: 'warning',
        rule: 'closing-content',
        message: 'Closing slide appears to lack substantive content',
        suggestion: 'Add a closing message or key takeaways',
      });
    }
  }

  // ── 7. Element overlap detection ──────────────────────────────────
  const rects = els.filter(el => el.w > 0 && el.h > 0).map(el => ({
    x: el.x || 0, y: el.y || 0, w: el.w, h: el.h, preview: el.textPreview || ''
  }));
  for (let i = 0; i < rects.length; i++) {
    for (let j = i + 1; j < rects.length; j++) {
      const a = rects[i], b = rects[j];
      const overlapX = Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x));
      const overlapY = Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y));
      const overlapArea = overlapX * overlapY;
      const smallerArea = Math.min(a.w * a.h, b.w * b.h);
      if (overlapArea > smallerArea * 0.5 && smallerArea > 0.01) {
        // Significant overlap — but skip background rects overlapping intentionally
        const isBgA = a.preview === '' && a.w > 5;
        const isBgB = b.preview === '' && b.w > 5;
        if (!isBgA && !isBgB) {
          issues.push({
            slideIndex,
            severity: 'warning',
            rule: 'element-overlap',
            message: `Elements "${a.preview.substring(0, 20)}" and "${b.preview.substring(0, 20)}" overlap significantly`,
            suggestion: 'Adjust positions to avoid text overlap',
          });
          break; // One overlap warning per element is enough
        }
      }
    }
    // Only report first overlap per slide to avoid noise
    if (issues.some(i => i.rule === 'element-overlap' && i.slideIndex === slideIndex)) break;
  }

  return issues;
}

function runAudit(auditData) {
  const allIssues = [];
  for (let i = 0; i < auditData.slides.length; i++) {
    const slideIssues = auditSlide(auditData.slides[i], i);
    allIssues.push(...slideIssues);
  }

  const summary = { error: 0, warning: 0, info: 0 };
  allIssues.forEach(i => { summary[i.severity] = (summary[i.severity] || 0) + 1; });

  return {
    slideCount: auditData.slides.length,
    issues: allIssues,
    summary,
    passed: summary.error === 0 && summary.warning <= 2,
  };
}

// ── Main ────────────────────────────────────────────────────────────
function main() {
  const filePath = process.argv[2];
  if (!filePath) {
    console.error('Usage: node audit_deck_structural.js <audit_report.json>');
    process.exit(1);
  }

  const auditData = loadAudit(path.resolve(filePath));
  const result = runAudit(auditData);
  console.log(JSON.stringify(result, null, 2));
}

if (require.main === module) {
  main();
}

module.exports = { runAudit, auditSlide };