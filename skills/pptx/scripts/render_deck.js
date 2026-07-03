const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const PptxGenJS = require('pptxgenjs');
const fallbackExampleSpecPath = path.resolve(__dirname, '..', 'deck_spec_example.json');
const REMOTE_IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg']);
const RENDER_WARNING_PREFIX = 'MLCA_RENDER_WARNING:';

const embeddedDeckSpec = {
  outputFile: 'replace-me.pptx',
  title: 'Replace With Your Deck Title',
  subtitle: 'Use this file as a renderer template. Replace the embedded spec or pass a JSON file path via argv[2].',
  author: 'MLCA',
  deck: {
    showPageNumbers: true,
    appendSourcesSlide: true,
    writeSpecCopy: true,
    specOutputFile: 'deck-spec.json',
  },
  theme: {
    dominant: '0F172A',
    support: '2563EB',
    accent: 'F59E0B',
    background: 'F8FAFC',
    surface: 'FFFFFF',
    text: '111827',
    muted: '64748B',
    line: 'CBD5E1',
    headerFont: 'Microsoft YaHei UI',
    bodyFont: 'Microsoft YaHei',
  },
  slides: [
    {
      kind: 'cover',
      title: 'Replace With Your Deck Title',
      subtitle: 'A polished opening slide',
      eyebrow: 'PRESENTATION',
    },
    {
      kind: 'agenda',
      title: 'Agenda',
      items: ['Context', 'Key insights', 'Implications', 'Next steps'],
    },
    {
      kind: 'closing',
      title: 'Key Takeaways',
      bullets: ['Replace this example spec', 'Add your own slide kinds', 'Write the final PPTX to the requested filename'],
    },
  ],
};

const THEME_PRESETS = {
  mckinsey: {
    dominant: '1E293B',
    dominantMid: '334155',
    dominantLight: '475569',
    support: '0F766E',
    accent: '0D9488',
    accentSecondary: 'D97706',
    background: 'FFFFFF',
    backgroundAlt: 'F8FAFC',
    surface: 'FFFFFF',
    surfaceAlt: 'F1F5F9',
    surfaceMuted: 'F8FAFC',
    text: '1E293B',
    textSecondary: '475569',
    muted: '64748B',
    mutedLight: '94A3B8',
    line: 'E2E8F0',
    lineDark: '1E293B',
    success: '059669',
    warning: 'D97706',
    danger: 'DC2626',
    headerFont: 'Microsoft YaHei UI',
    bodyFont: 'Microsoft YaHei',
    motif: 'header-bar',
    accentColors: ['0F766E', '0D9488', 'D97706', '7C3AED', 'DC2626'],
    stepGradient: ['1E293B', '334155', '0F766E', '0D9488', 'D97706'],
    borderAccent: '0D9488',
    borderAccentWidth: 4,
  },
  'dark-tech': {
    dominant: '18181B',
    dominantSecondary: '27272A',
    support: '3F3F46',
    accent: '8B5CF6',
    accentBlue: '3B82F6',
    accentAmber: 'F59E0B',
    accentPurple: '8B5CF6',
    accentRed: 'EF4444',
    background: 'FAFAFA',
    surface: 'FFFFFF',
    surfaceMuted: 'F4F4F5',
    text: '18181B',
    textSecondary: '3F3F46',
    textMuted: '71717A',
    textHighlight: '7C3AED',
    textAmber: 'B45309',
    textGreen: '047857',
    line: 'E4E4E7',
    muted: '71717A',
    success: '059669',
    warning: 'D97706',
    danger: 'DC2626',
    headerFont: 'Microsoft YaHei UI',
    bodyFont: 'Microsoft YaHei',
    motif: 'header-bar',
    accentColors: ['8B5CF6', '3B82F6', 'F59E0B', '059669', 'EF4444'],
    stepGradient: ['8B5CF6', '3B82F6', 'F59E0B', '059669', 'EF4444'],
    borderAccent: '8B5CF6',
    borderAccentWidth: 4,
  },
  genspark: {
    dominant: '0F172A',
    support: '1D4ED8',
    accent: '059669',
    accentIndigo: '4F46E5',
    accentAmber: 'D97706',
    background: 'FFFFFF',
    surface: 'FFFFFF',
    surfaceBlue: 'EFF6FF',
    surfaceGreen: 'ECFDF5',
    surfaceIndigo: 'EEF2FF',
    surfaceAmber: 'FFFBEB',
    text: '0F172A',
    textSecondary: '374151',
    muted: '6B7280',
    line: 'E5E7EB',
    success: '059669',
    warning: 'D97706',
    danger: 'DC2626',
    headerFont: 'Microsoft YaHei UI',
    bodyFont: 'Microsoft YaHei',
    motif: 'gradient-cover',
    accentColors: ['1D4ED8', '059669', '4F46E5', 'D97706', 'DC2626'],
    stepGradient: ['1D4ED8', '059669', '4F46E5', 'D97706', 'DC2626'],
    borderAccent: '1D4ED8',
    borderAccentWidth: 4,
  },
};

function asArray(value) {
  if (Array.isArray(value)) return value.filter(Boolean);
  if (value === null || value === undefined || value === '') return [];
  return [value];
}

function hasMeaningfulValue(value) {
  if (value === null || value === undefined) return false;
  if (typeof value === 'string') return value.trim() !== '';
  return true;
}

function pickFirstDefined(...values) {
  for (const value of values) {
    if (hasMeaningfulValue(value)) {
      return value;
    }
  }
  return undefined;
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function toNumber(value, fallback = 0) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : fallback;
}

function toBoolean(value, fallback = false) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (['true', '1', 'yes', 'y'].includes(normalized)) return true;
    if (['false', '0', 'no', 'n'].includes(normalized)) return false;
  }
  return fallback;
}

// ─── Dynamic Layout Engine ───────────────────────────────────────────────────
// Supports all PptxGenJS layouts:
//   LAYOUT_16x9: 10" × 5.625"
//   LAYOUT_16x10: 10" × 6.25"
//   LAYOUT_4x3: 10" × 7.5"
//   LAYOUT_WIDE: 13.33" × 7.5"
const LAYOUT_DIMENSIONS = {
  'LAYOUT_16x9': { w: 10, h: 5.625 },
  'LAYOUT_16x10': { w: 10, h: 6.25 },
  'LAYOUT_4x3': { w: 10, h: 7.5 },
  'LAYOUT_WIDE': { w: 13.33, h: 7.5 },
};

function getCanvas(spec) {
  const layout = (spec && spec.layout) || 'LAYOUT_WIDE';
  return LAYOUT_DIMENSIONS[layout] || LAYOUT_DIMENSIONS['LAYOUT_WIDE'];
}

function getMargin(canvas) {
  return { left: canvas.w * 0.049, right: canvas.w * 0.049, bottom: canvas.h * 0.047 };
}

function getHeaderHeight(canvas) {
  return canvas.h * 0.193;
}

function getContentRect(spec, deck) {
  const canvas = getCanvas(spec);
  const margin = getMargin(canvas);
  const headerH = getHeaderHeight(canvas);
  let topOffset = headerH;
  if (deck && (deck.intro || deck.message)) topOffset += canvas.h * 0.06;
  return {
    x: margin.left,
    y: topOffset,
    w: canvas.w - margin.left - margin.right,
    h: canvas.h - topOffset - margin.bottom,
  };
}

function calcEqualColumns(count, rect, gap) {
  if (!gap && gap !== 0) gap = rect.w * 0.025;
  const n = Math.max(1, count);
  const totalGap = gap * (n - 1);
  const colW = (rect.w - totalGap) / n;
  const cols = [];
  for (let i = 0; i < n; i++) {
    cols.push({
      x: rect.x + i * (colW + gap),
      y: rect.y,
      w: colW,
      h: rect.h,
    });
  }
  return cols;
}

function calcEqualGrid(count, rect, gapX, gapY) {
  if (!gapX && gapX !== 0) gapX = rect.w * 0.029;
  if (!gapY && gapY !== 0) gapY = rect.h * 0.06;
  const n = Math.max(1, count);
  let cols, rows;
  if (n <= 2) { cols = n; rows = 1; }
  else if (n <= 4) { cols = 2; rows = Math.ceil(n / 2); }
  else if (n <= 6) { cols = 3; rows = Math.ceil(n / 3); }
  else if (n <= 9) { cols = 3; rows = Math.ceil(n / 3); }
  else { cols = 4; rows = Math.ceil(n / 4); }
  const cellW = (rect.w - gapX * (cols - 1)) / cols;
  const cellH = (rect.h - gapY * (rows - 1)) / rows;
  const cells = [];
  for (let i = 0; i < n; i++) {
    const col = i % cols;
    const row = Math.floor(i / cols);
    cells.push({
      x: rect.x + col * (cellW + gapX),
      y: rect.y + row * (cellH + gapY),
      w: cellW,
      h: cellH,
    });
  }
  return cells;
}

function estimateTextLines(text, widthInches, fontSize) {
  const str = String(text || '');
  if (!str) return 0;
  const charsPerLine = Math.max(1, Math.floor((widthInches * 72) / (fontSize * 1.05)));
  let lines = 0;
  const paragraphs = str.split('\n');
  for (const para of paragraphs) {
    lines += Math.max(1, Math.ceil(para.length / charsPerLine));
  }
  return lines;
}

function estimateTextHeight(text, widthInches, fontSize) {
  const lines = estimateTextLines(text, widthInches, fontSize);
  return lines * fontSize * 1.45 / 72;
}

function adaptiveFontSize(text, widthInches, maxHeight, baseFontSize, minFontSize) {
  if (!minFontSize) minFontSize = 7;
  let fs = baseFontSize;
  while (fs > minFontSize) {
    const h = estimateTextHeight(text, widthInches, fs);
    if (h <= maxHeight) return fs;
    fs -= 0.5;
  }
  return minFontSize;
}

function calcTwoPanelSplit(leftText, rightText, rect, gap) {
  if (!gap) gap = rect.w * 0.033;
  const leftLen = String(leftText || '').length || 1;
  const rightLen = String(rightText || '').length || 1;
  const ratio = clamp(leftLen / (leftLen + rightLen), 0.35, 0.65);
  const leftW = (rect.w - gap) * ratio;
  const rightW = rect.w - gap - leftW;
  return {
    left: { x: rect.x, y: rect.y, w: leftW, h: rect.h },
    right: { x: rect.x + leftW + gap, y: rect.y, w: rightW, h: rect.h },
  };
}
// ─── End Dynamic Layout Engine ───────────────────────────────────────────────

function normalizeHexColor(value, fallback) {
  const cleaned = String(value || '')
    .trim()
    .replace(/^#/, '');
  if (!cleaned) return fallback;
  if (/^[0-9a-fA-F]{3}$/.test(cleaned)) {
    return cleaned
      .split('')
      .map((char) => char + char)
      .join('')
      .toUpperCase();
  }
  if (/^[0-9a-fA-F]{6}$/.test(cleaned) || /^[0-9a-fA-F]{8}$/.test(cleaned)) {
    return cleaned.slice(cleaned.length - 6).toUpperCase();
  }
  return fallback;
}

function resolveTransparencyOption(transparencyValue, opacityValue, fallback = 0) {
  if (hasMeaningfulValue(transparencyValue)) {
    const numeric = toNumber(transparencyValue, fallback);
    return numeric <= 1 ? clamp(numeric * 100, 0, 100) : clamp(numeric, 0, 100);
  }
  if (hasMeaningfulValue(opacityValue)) {
    const opacity = toNumber(opacityValue, 1);
    const normalizedOpacity = opacity <= 1 ? opacity : opacity / 100;
    return clamp((1 - normalizedOpacity) * 100, 0, 100);
  }
  return fallback;
}

function normalizeFillOption(rawFill, fallbackColor = 'FFFFFF', fallbackTransparency = 0) {
  if (rawFill === false) {
    return {
      color: fallbackColor,
      transparency: 100,
    };
  }
  const fill = rawFill && typeof rawFill === 'object' ? rawFill : {};
  const flatColor = rawFill && typeof rawFill === 'string' ? rawFill : undefined;
  return {
    color: normalizeHexColor(
      pickFirstDefined(fill.color, fill.backgroundColor, flatColor),
      fallbackColor
    ),
    transparency: resolveTransparencyOption(fill.transparency, fill.opacity, fallbackTransparency),
  };
}

function normalizeLineOption(rawLine, fallbackColor = '000000', fallbackWidth = 1, fallbackTransparency = 0) {
  if (rawLine === false) {
    return {
      color: fallbackColor,
      transparency: 100,
      width: 0,
    };
  }
  const line = rawLine && typeof rawLine === 'object' ? rawLine : {};
  const flatColor = rawLine && typeof rawLine === 'string' ? rawLine : undefined;
  const normalized = {
    color: normalizeHexColor(
      pickFirstDefined(line.color, line.borderColor, flatColor),
      fallbackColor
    ),
    transparency: resolveTransparencyOption(line.transparency, line.opacity, fallbackTransparency),
    width: toNumber(pickFirstDefined(line.width, line.borderWidth), fallbackWidth),
  };
  if (hasMeaningfulValue(line.dash)) normalized.dash = String(line.dash);
  if (hasMeaningfulValue(line.beginArrowType)) normalized.beginArrowType = String(line.beginArrowType);
  if (hasMeaningfulValue(line.endArrowType)) normalized.endArrowType = String(line.endArrowType);
  if (hasMeaningfulValue(line.beginArrowhead)) normalized.beginArrowType = String(line.beginArrowhead);
  if (hasMeaningfulValue(line.endArrowhead)) normalized.endArrowType = String(line.endArrowhead);
  return normalized;
}

function resolveElementBox(element, defaults = {}) {
  const box = element && typeof element.box === 'object' ? element.box : {};
  return {
    x: toNumber(pickFirstDefined(element?.x, box.x), defaults.x ?? 0),
    y: toNumber(pickFirstDefined(element?.y, box.y), defaults.y ?? 0),
    w: toNumber(pickFirstDefined(element?.w, box.w), defaults.w ?? 1),
    h: toNumber(pickFirstDefined(element?.h, box.h), defaults.h ?? 1),
  };
}

function normalizeTheme(raw, themePreset) {
  const presetName = String(themePreset || '').trim().toLowerCase();
  const preset = THEME_PRESETS[presetName] || {};
  const theme = raw || {};
  const merged = { ...preset };
  for (const key of Object.keys(theme)) {
    if (hasMeaningfulValue(theme[key])) {
      merged[key] = theme[key];
    }
  }
  return {
    dominant: merged.dominant || '0F172A',
    support: merged.support || '2563EB',
    accent: merged.accent || 'F59E0B',
    background: merged.background || 'F8FAFC',
    surface: merged.surface || 'FFFFFF',
    text: merged.text || '111827',
    muted: merged.muted || '64748B',
    line: merged.line || 'CBD5E1',
    headerFont: merged.headerFont || 'Microsoft YaHei UI',
    bodyFont: merged.bodyFont || 'Microsoft YaHei',
    accentSecondary: merged.accentSecondary || merged.accent || 'F59E0B',
    success: merged.success || merged.accentSecondary || '4CAF50',
    warning: merged.warning || 'F59E0B',
    danger: merged.danger || 'E53E3E',
    motif: merged.motif || '',
    accentColors: merged.accentColors || [],
    stepGradient: merged.stepGradient || [],
    borderAccent: merged.borderAccent || merged.accent || 'F59E0B',
    borderAccentWidth: merged.borderAccentWidth || 4,
    cardTransparency: merged.cardTransparency || 0,
  };
}

function normalizeDeckSpec(raw) {
  const spec = raw || {};
  return {
    outputFile: spec.outputFile || spec.fileName || 'deck-output.pptx',
    title: spec.title || 'Untitled Deck',
    subtitle: spec.subtitle || '',
    author: spec.author || 'MLCA',
    layout: spec.layout || 'LAYOUT_WIDE',
    themePreset: spec.themePreset || '',
    _auditOutput: spec._auditOutput || '',
    deck: {
      showPageNumbers: spec.deck?.showPageNumbers !== false,
      appendSourcesSlide: spec.deck?.appendSourcesSlide !== false,
      writeSpecCopy: spec.deck?.writeSpecCopy !== false,
      specOutputFile: spec.deck?.specOutputFile || 'deck-spec.json',
    },
    theme: normalizeTheme(spec.theme, spec.themePreset),
    assets: asArray(spec.assets),
    slides: asArray(spec.slides),
  };
}

function pushRenderWarning(spec, message) {
  const text = String(message || '').trim();
  if (!text) {
    return;
  }
  if (spec && typeof spec === 'object') {
    if (!Array.isArray(spec._warnings)) {
      spec._warnings = [];
    }
    if (!spec._warnings.includes(text)) {
      spec._warnings.push(text);
    }
  }
  console.warn(`${RENDER_WARNING_PREFIX}${text}`);
}

function trimForLog(value, maxLength = 96) {
  const text = String(value || '').trim();
  if (!text) {
    return '';
  }
  return text.length > maxLength ? `${text.slice(0, maxLength - 1)}…` : text;
}

function describeImageRef(imageSpec) {
  if (hasMeaningfulValue(imageSpec.assetId)) {
    return `image asset "${trimForLog(imageSpec.assetId, 48)}"`;
  }
  if (hasMeaningfulValue(imageSpec.filename)) {
    return `image "${trimForLog(imageSpec.filename, 48)}"`;
  }
  if (hasMeaningfulValue(imageSpec.path)) {
    return `image path "${trimForLog(imageSpec.path)}"`;
  }
  if (hasMeaningfulValue(imageSpec.url)) {
    return `image url "${trimForLog(imageSpec.url)}"`;
  }
  if (hasMeaningfulValue(imageSpec.src)) {
    return `image source "${trimForLog(imageSpec.src)}"`;
  }
  return 'image resource';
}

function normalizeImageFallbackMode(imageSpec, layer = 'foreground') {
  const rawValue = pickFirstDefined(imageSpec?.fallback, imageSpec?.onError);
  if (typeof rawValue === 'string') {
    const normalized = rawValue.trim().toLowerCase();
    if (['skip', 'omit', 'none'].includes(normalized)) {
      return 'skip';
    }
    if (['placeholder', 'box', 'frame'].includes(normalized)) {
      return 'placeholder';
    }
  }
  if (typeof rawValue === 'boolean') {
    return rawValue ? 'placeholder' : 'skip';
  }
  return layer === 'background' ? 'skip' : 'placeholder';
}

function buildSerializableSpec(spec, slides) {
  return {
    outputFile: spec.outputFile,
    title: spec.title,
    subtitle: spec.subtitle,
    author: spec.author,
    layout: spec.layout,
    themePreset: spec.themePreset || '',
    deck: {
      ...spec.deck,
    },
    theme: {
      ...spec.theme,
    },
    assets: asArray(spec.assets).map((asset) => (asset && typeof asset === 'object' ? { ...asset } : asset)),
    slides: slides.map((slide) => ({ ...slide })),
  };
}

function firstNonEmptyString(...values) {
  for (const value of values) {
    const text = String(value ?? '').trim();
    if (text) {
      return text;
    }
  }
  return '';
}

function normalizeTextList(values) {
  return asArray(values)
    .map((value) => String(value ?? '').trim())
    .filter(Boolean);
}

function buildBulletParagraph(values) {
  return normalizeTextList(values)
    .map((value) => (/^[-*•]/.test(value) ? value : `- ${value}`))
    .join('\n');
}

function splitStatementText(value, fallbackTitle) {
  const cleaned = String(value ?? '').trim();
  if (!cleaned) {
    return {
      title: fallbackTitle,
      body: '',
    };
  }

  const matched = cleaned.match(/^([^:：]{2,24})[:：]\s*(.+)$/);
  if (matched) {
    return {
      title: matched[1].trim(),
      body: matched[2].trim(),
    };
  }

  return {
    title: fallbackTitle,
    body: cleaned,
  };
}

function normalizeGridBody(entry) {
  const subtitle = firstNonEmptyString(entry.subtitle, entry.note, entry.desc);
  const bulletBody = buildBulletParagraph(entry.bullets || entry.items || entry.features);
  const directBody = firstNonEmptyString(entry.body, entry.text, entry.message);

  if (subtitle && bulletBody) {
    return `${subtitle}\n${bulletBody}`;
  }
  return bulletBody || directBody || subtitle;
}

function normalizeSlideForRenderer(rawSlide) {
  const slide = rawSlide && typeof rawSlide === 'object' ? { ...rawSlide } : {};
  const kind = String(slide.kind || '').trim().toLowerCase() || 'bullets';
  slide.kind = kind;
  slide.images = asArray(slide.images).filter((image) => image && typeof image === 'object');
  slide.elements = asArray(slide.elements).filter((element) => element && typeof element === 'object');

  if (kind === 'agenda' && !asArray(slide.items).length) {
    slide.items = normalizeTextList(slide.bullets);
  }

  if (kind === 'stats' && !asArray(slide.items).length) {
    const source = asArray(slide.stats).length ? slide.stats : asArray(slide.columns).length ? slide.columns : [];
    if (source.length) {
      slide.items = asArray(source).map((entry) => ({
        ...entry,
        value: firstNonEmptyString(entry.value, entry.metric, '--'),
        label: firstNonEmptyString(entry.label, entry.title),
        note: firstNonEmptyString(entry.note, entry.body, entry.trend, entry.subtitle),
      }));
    }
  }

  if (kind === 'grid' && !asArray(slide.items).length && asArray(slide.columns).length) {
    slide.items = asArray(slide.columns).map((entry, index) => ({
      ...entry,
      title: firstNonEmptyString(entry.title, `Block ${index + 1}`),
      body: normalizeGridBody(entry),
    }));
  }

  if (kind === 'feature-grid' && !asArray(slide.items).length && asArray(slide.columns).length) {
    slide.items = asArray(slide.columns).map((entry, index) => ({
      ...entry,
      title: firstNonEmptyString(entry.title, `Module ${index + 1}`),
      subtitle: firstNonEmptyString(entry.subtitle, entry.desc),
      features: normalizeTextList(entry.features || entry.bullets || entry.items),
      valueTag: firstNonEmptyString(entry.valueTag, entry.tag, entry.metric),
    }));
  }

  if (kind === 'process-flow') {
    if (!asArray(slide.items).length && asArray(slide.steps).length) {
      slide.items = slide.steps;
    }
    slide.items = asArray(slide.items).map((entry, index) => {
      if (!entry || typeof entry !== 'object') {
        return {
          title: `Step ${index + 1}`,
          body: String(entry ?? '').trim(),
        };
      }
      return {
        ...entry,
        title: firstNonEmptyString(entry.title, entry.label, entry.name, `Step ${index + 1}`),
        body: firstNonEmptyString(entry.body, entry.description, entry.desc, entry.message),
        step: firstNonEmptyString(entry.step, entry.label ? `STEP ${String(entry.label).padStart(2, '0')}` : undefined),
        outputTitle: firstNonEmptyString(entry.outputTitle, entry.outputLabel),
        outputs: normalizeTextList(entry.outputs || entry.results || entry.bullets),
      };
    });
  }

  if (kind === 'architecture-stack' || kind === 'layered-architecture') {
    slide.layers = asArray(slide.layers).map((entry, index) => ({
      ...entry,
      title: firstNonEmptyString(entry.title, entry.label, `Layer ${index + 1}`),
      subtitle: firstNonEmptyString(entry.subtitle, entry.desc),
      modules: normalizeTextList(entry.modules || entry.items),
    }));
    slide.sidebarTitle = firstNonEmptyString(slide.sidebarTitle, slide.rightTitle);
    if (!asArray(slide.services).length) {
      slide.services = normalizeTextList(slide.services || slide.rightBullets);
    }
  }

  if (kind === 'kpi-dashboard' && !asArray(slide.items).length && asArray(slide.columns).length) {
    slide.items = asArray(slide.columns).map((entry) => ({
      ...entry,
      value: firstNonEmptyString(entry.value, entry.metric, '--'),
      label: firstNonEmptyString(entry.label, entry.title),
      note: firstNonEmptyString(entry.note, entry.trend, entry.subtitle),
    }));
  }

  if (kind === 'timeline-dots' && !asArray(slide.items).length && asArray(slide.columns).length) {
    slide.items = asArray(slide.columns).map((entry, index) => ({
      ...entry,
      title: firstNonEmptyString(entry.title, entry.label, `Phase ${index + 1}`),
      body: firstNonEmptyString(entry.body, entry.desc),
    }));
  }

  if (kind === 'comparison') {
    slide.columns = asArray(slide.columns).map((entry, index) => ({
      ...entry,
      title: firstNonEmptyString(entry.title, `Column ${index + 1}`),
      items: normalizeTextList(entry.items || entry.bullets),
      body: firstNonEmptyString(entry.body, entry.text),
    }));
  }

  if (kind === 'timeline' && !asArray(slide.items).length && asArray(slide.steps).length) {
    slide.items = slide.steps;
  }
  if (kind === 'timeline') {
    slide.items = asArray(slide.items).map((entry, index) => ({
      ...entry,
      title: firstNonEmptyString(entry.title, entry.label, `Stage ${index + 1}`),
      body: firstNonEmptyString(entry.body, entry.description, entry.note),
    }));
  }

  if (kind === 'capability-radar' && !asArray(slide.items).length && asArray(slide.columns).length) {
    slide.items = asArray(slide.columns).map((entry, index) => ({
      ...entry,
      title: firstNonEmptyString(entry.title, `Capability ${index + 1}`),
      subtitle: firstNonEmptyString(entry.subtitle, entry.desc),
      features: normalizeTextList(entry.features || entry.bullets || entry.items),
    }));
  }

  if (kind === 'value-closing' && !asArray(slide.items).length) {
    slide.items = normalizeTextList(slide.bullets).slice(0, 4).map((entry, index) => {
      const mapped = splitStatementText(entry, `Value ${index + 1}`);
      return {
        title: mapped.title,
        body: mapped.body,
      };
    });
  }

  return slide;
}

function normalizeSourceEntry(entry) {
  if (!entry) return null;
  if (typeof entry === 'string') {
    return { label: entry.trim(), url: '' };
  }
  const label = String(entry.label || entry.name || entry.title || '').trim();
  const url = String(entry.url || entry.href || '').trim();
  if (!label && !url) return null;
  return { label: label || url, url };
}

function sanitizeAssetFilename(value, fallback = 'asset') {
  const cleaned = String(value || '')
    .trim()
    .replace(/[\\\/]+/g, '/')
    .split('/')
    .filter(Boolean)
    .at(-1) || fallback;
  return cleaned.replace(/[^A-Za-z0-9._-]+/g, '_');
}

function assetExtensionFromUrl(url) {
  try {
    const pathname = new URL(url).pathname || '';
    const extension = path.extname(pathname).toLowerCase();
    return REMOTE_IMAGE_EXTENSIONS.has(extension) ? extension : '';
  } catch {
    return '';
  }
}

function assetExtensionFromContentType(contentType) {
  const lowered = String(contentType || '').toLowerCase();
  if (lowered.includes('image/png')) return '.png';
  if (lowered.includes('image/jpeg')) return '.jpg';
  if (lowered.includes('image/gif')) return '.gif';
  if (lowered.includes('image/webp')) return '.webp';
  if (lowered.includes('image/svg')) return '.svg';
  return '';
}

function isRemoteAssetDownloadEnabled() {
  return /^(1|true|yes)$/i.test(String(process.env.MLCA_PPTX_ALLOW_REMOTE_ASSETS || '').trim());
}

async function downloadRemoteAsset(asset) {
  if (typeof fetch !== 'function') {
    throw new Error('Remote asset download requires a fetch-capable Node runtime.');
  }
  const url = String(asset.url || '').trim();
  if (!url) {
    throw new Error(`Asset "${asset.id}" is missing url.`);
  }
  const cacheDir = path.resolve(process.env.MLCA_PPTX_ASSET_CACHE_DIR || '.mlca-pptx-assets');
  fs.mkdirSync(cacheDir, { recursive: true });

  const hash = crypto.createHash('sha1').update(url).digest('hex');
  const explicitName = sanitizeAssetFilename(asset.filename || '', '');
  const hintedExt = path.extname(explicitName).toLowerCase() || assetExtensionFromUrl(url);
  const provisionalExt = REMOTE_IMAGE_EXTENSIONS.has(hintedExt) ? hintedExt : '.img';
  const provisionalName = explicitName || `${asset.id || 'asset'}-${hash.slice(0, 10)}${provisionalExt}`;
  const provisionalPath = path.resolve(cacheDir, provisionalName);
  if (fs.existsSync(provisionalPath)) {
    return provisionalPath;
  }

  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to download asset "${asset.id}" (${response.status} ${response.statusText}).`);
  }
  const contentType = response.headers.get('content-type') || '';
  const contentExt = assetExtensionFromContentType(contentType);
  const finalExt = contentExt || hintedExt || '.png';
  if (!REMOTE_IMAGE_EXTENSIONS.has(finalExt)) {
    throw new Error(`Asset "${asset.id}" is not a supported image type: ${contentType || finalExt}`);
  }

  const finalName = explicitName
    ? `${path.parse(explicitName).name}${finalExt}`
    : `${asset.id || 'asset'}-${hash.slice(0, 10)}${finalExt}`;
  const finalPath = path.resolve(cacheDir, sanitizeAssetFilename(finalName, `${asset.id || 'asset'}${finalExt}`));
  if (fs.existsSync(finalPath)) {
    return finalPath;
  }

  const arrayBuffer = await response.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer);
  fs.writeFileSync(finalPath, buffer);
  return finalPath;
}

async function resolveDeckAssets(spec) {
  const assetMap = new Map();
  for (const [index, rawAsset] of asArray(spec.assets).entries()) {
    const candidate = typeof rawAsset === 'string' ? { id: `asset-${index + 1}`, path: rawAsset } : rawAsset;
    if (!candidate || typeof candidate !== 'object') {
      continue;
    }
    const assetId = String(candidate.id || `asset-${index + 1}`).trim();
    if (!assetId) {
      continue;
    }

    const pathValue = String(candidate.path || '').trim();
    const urlValue = String(candidate.url || '').trim();
    if (!pathValue && !urlValue) {
      continue;
    }

    let resolvedPath = '';
    if (pathValue) {
      resolvedPath = path.resolve(pathValue);
      if (!fs.existsSync(resolvedPath)) {
        pushRenderWarning(spec, `Asset "${assetId}" path does not exist: ${trimForLog(pathValue)}`);
        assetMap.set(assetId, {
          ...candidate,
          id: assetId,
          resolvedPath: '',
          error: `Asset "${assetId}" path does not exist: ${pathValue}`,
        });
        continue;
      }
    } else {
      if (!isRemoteAssetDownloadEnabled()) {
        pushRenderWarning(spec, `Asset "${assetId}" requires remote download, but remote asset fetching is disabled.`);
        assetMap.set(assetId, {
          ...candidate,
          id: assetId,
          resolvedPath: '',
          error: `Asset "${assetId}" requires remote download, but allow_remote_assets is disabled.`,
        });
        continue;
      }
      try {
        resolvedPath = await downloadRemoteAsset({
          id: assetId,
          url: urlValue,
          filename: candidate.filename || '',
        });
      } catch (error) {
        pushRenderWarning(spec, `Asset "${assetId}" download failed: ${trimForLog(error && error.message ? error.message : error)}`);
        assetMap.set(assetId, {
          ...candidate,
          id: assetId,
          resolvedPath: '',
          error: error && error.message ? error.message : String(error),
        });
        continue;
      }
    }

    assetMap.set(assetId, {
      ...candidate,
      id: assetId,
      resolvedPath,
      error: '',
    });
  }
  return assetMap;
}

function resolveImageSourceFields(imageSpec) {
  const rawSrc = String(imageSpec.src || '').trim();
  const rawPath = String(imageSpec.path || '').trim();
  const rawUrl = String(imageSpec.url || '').trim();
  if (rawPath || rawUrl) {
    return {
      path: rawPath,
      url: rawUrl,
    };
  }
  if (/^https?:\/\//i.test(rawSrc)) {
    return {
      path: '',
      url: rawSrc,
    };
  }
  return {
    path: rawSrc,
    url: '',
  };
}

async function resolveImagePath(assetMap, imageSpec) {
  const assetId = String(imageSpec.assetId || '').trim();
  if (assetId) {
    const asset = assetMap.get(assetId);
    if (!asset) {
      return {
        path: '',
        error: `Slide image asset "${assetId}" is not available.`,
      };
    }
    if (!asset.resolvedPath) {
      return {
        path: '',
        error: asset.error || `Slide image asset "${assetId}" is not available.`,
      };
    }
    return {
      path: asset.resolvedPath,
      error: '',
    };
  }

  const source = resolveImageSourceFields(imageSpec);
  if (source.path) {
    const resolvedPath = path.resolve(source.path);
    if (!fs.existsSync(resolvedPath)) {
      return {
        path: '',
        error: `Slide image path does not exist: ${source.path}`,
      };
    }
    return {
      path: resolvedPath,
      error: '',
    };
  }

  if (source.url) {
    if (!isRemoteAssetDownloadEnabled()) {
      return {
        path: '',
        error: 'Slide image url requires remote download, but allow_remote_assets is disabled.',
      };
    }
    const hash = crypto.createHash('sha1').update(source.url).digest('hex');
    try {
      return {
        path: await downloadRemoteAsset({
          id: assetId || String(imageSpec.id || `inline-image-${hash.slice(0, 10)}`),
          url: source.url,
          filename: imageSpec.filename || '',
        }),
        error: '',
      };
    } catch (error) {
      return {
        path: '',
        error: error && error.message ? error.message : String(error),
      };
    }
  }

  return {
    path: '',
    error: '',
  };
}

function renderImagePlaceholder(slide, imageSpec, theme, message = 'Image unavailable') {
  const box = resolveElementBox(imageSpec, { x: 0, y: 0, w: 1, h: 1 });
  if (box.w <= 0 || box.h <= 0) {
    return;
  }
  const inset = Math.min(0.12, Math.max(Math.min(box.w, box.h) * 0.08, 0.04));
  slide.addShape(slide._pptx.shapes.RECTANGLE, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    fill: { color: theme.background, transparency: 12 },
    line: { color: theme.line, width: 1.1, transparency: 20, dash: 'dash' },
  });
  slide.addShape(slide._pptx.shapes.LINE, {
    x: box.x + inset,
    y: box.y + inset,
    w: Math.max(box.w - inset * 2, 0),
    h: Math.max(box.h - inset * 2, 0),
    line: { color: theme.muted, width: 1.1, transparency: 30 },
  });
  slide.addShape(slide._pptx.shapes.LINE, {
    x: box.x + inset,
    y: box.y + box.h - inset,
    w: Math.max(box.w - inset * 2, 0),
    h: -Math.max(box.h - inset * 2, 0),
    line: { color: theme.muted, width: 1.1, transparency: 30 },
  });
  if (box.w >= 1.45 && box.h >= 0.48) {
    slide.addText(trimForLog(message, 48), {
      x: box.x + 0.1,
      y: box.y + Math.max(box.h - 0.25, 0) / 2,
      w: Math.max(box.w - 0.2, 0.2),
      h: 0.25,
      fontFace: theme.bodyFont,
      fontSize: Math.min(11, Math.max(8, box.h * 8)),
      color: theme.muted,
      margin: 0,
      align: 'center',
      valign: 'mid',
      fit: 'shrink',
    });
  }
}

async function renderSlideImages(slide, rawSlide, assetMap, theme, spec, layer = 'foreground') {
  const images = asArray(rawSlide.images).filter((image) => image && typeof image === 'object');
  for (const image of images) {
    const imageLayer = String(image.layer || 'foreground').trim().toLowerCase();
    if (imageLayer !== layer) {
      continue;
    }
    const resolved = await resolveImagePath(assetMap, image);
    if (resolved.error) {
      pushRenderWarning(spec, `${describeImageRef(image)}: ${trimForLog(resolved.error)}`);
      if (normalizeImageFallbackMode(image, layer) === 'placeholder') {
        renderImagePlaceholder(slide, image, theme);
      }
      continue;
    }
    if (!resolved.path) {
      continue;
    }
    const box = resolveElementBox(image, { x: 0, y: 0, w: 1, h: 1 });
    try {
      slide.addImage({
        path: resolved.path,
        x: box.x,
        y: box.y,
        w: box.w,
        h: box.h,
      });
    } catch (error) {
      pushRenderWarning(spec, `${describeImageRef(image)}: ${trimForLog(error && error.message ? error.message : error)}`);
      if (normalizeImageFallbackMode(image, layer) === 'placeholder') {
        renderImagePlaceholder(slide, image, theme);
      }
    }
  }
}

function resolveElementShapeType(slide, rawShape) {
  const shape = String(rawShape || 'rect').trim().toLowerCase();
  if (shape === 'line') {
    return slide._pptx.shapes.LINE;
  }
  if (['roundrect', 'rounded-rect', 'roundedrect', 'rounded_rectangle', 'round-rect'].includes(shape)) {
    return slide._pptx.shapes.ROUNDED_RECTANGLE;
  }
  if (['oval', 'ellipse', 'circle'].includes(shape)) {
    return slide._pptx.shapes.OVAL;
  }
  return slide._pptx.shapes.RECTANGLE;
}

function resolveElementType(element) {
  const explicitType = String(element.type || '').trim().toLowerCase();
  if (explicitType) {
    return explicitType;
  }
  if (hasMeaningfulValue(element.shape) && String(element.shape).trim().toLowerCase() === 'line') {
    return 'line';
  }
  if (hasMeaningfulValue(element.assetId) || hasMeaningfulValue(element.path) || hasMeaningfulValue(element.src) || hasMeaningfulValue(element.url)) {
    return 'image';
  }
  if (hasMeaningfulValue(element.x2) || hasMeaningfulValue(element.y2) || element.from || element.to) {
    return 'line';
  }
  if (hasMeaningfulValue(element.text) || Array.isArray(element.text) || hasMeaningfulValue(element.title) || hasMeaningfulValue(element.body) || hasMeaningfulValue(element.label)) {
    return 'text';
  }
  return 'shape';
}

function buildTextContent(element) {
  const rawText = pickFirstDefined(element.text, element.value, element.label, element.title, element.body, '');
  if (Array.isArray(rawText)) {
    if (rawText.every((item) => item && typeof item === 'object' && hasMeaningfulValue(item.text))) {
      return rawText.map((item) => ({
        text: String(item.text),
        options: item.options && typeof item.options === 'object' ? { ...item.options } : item.options,
      }));
    }
    return rawText.map((item) => String(item ?? '')).join('\n');
  }
  return String(rawText ?? '');
}

function renderTextElement(slide, element, theme) {
  const style = element.style && typeof element.style === 'object' ? element.style : {};
  const box = resolveElementBox(element, { x: 0, y: 0, w: 2, h: 0.4 });
  const content = buildTextContent(element);
  if (!content || (Array.isArray(content) && !content.length)) {
    return;
  }

  const options = {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    fontFace: firstNonEmptyString(style.fontFace, element.fontFace, theme.bodyFont),
    fontSize: toNumber(pickFirstDefined(style.fontSize, element.fontSize), 12),
    color: normalizeHexColor(pickFirstDefined(style.color, element.color), theme.text),
    bold: toBoolean(pickFirstDefined(style.bold, element.bold), false),
    italic: toBoolean(pickFirstDefined(style.italic, element.italic), false),
    margin: pickFirstDefined(style.margin, element.margin, 0),
  };
  if (hasMeaningfulValue(style.align) || hasMeaningfulValue(element.align)) {
    options.align = String(pickFirstDefined(style.align, element.align)).trim().toLowerCase();
  }
  if (hasMeaningfulValue(style.valign) || hasMeaningfulValue(element.valign)) {
    options.valign = String(pickFirstDefined(style.valign, element.valign)).trim().toLowerCase();
  }
  if (hasMeaningfulValue(style.breakLine) || hasMeaningfulValue(element.breakLine)) {
    options.breakLine = toBoolean(pickFirstDefined(style.breakLine, element.breakLine), false);
  }
  if (hasMeaningfulValue(style.fit) || hasMeaningfulValue(element.fit)) {
    options.fit = String(pickFirstDefined(style.fit, element.fit)).trim();
  }
  if (hasMeaningfulValue(style.rotate) || hasMeaningfulValue(element.rotate)) {
    options.rotate = toNumber(pickFirstDefined(style.rotate, element.rotate), 0);
  }
  if (hasMeaningfulValue(style.charSpacing) || hasMeaningfulValue(element.charSpacing)) {
    options.charSpacing = toNumber(pickFirstDefined(style.charSpacing, element.charSpacing), 0);
  }
  if (hasMeaningfulValue(style.underline) || hasMeaningfulValue(element.underline)) {
    options.underline = pickFirstDefined(style.underline, element.underline);
  }
  if (hasMeaningfulValue(style.fill) || hasMeaningfulValue(element.fill)) {
    options.fill = normalizeFillOption(pickFirstDefined(style.fill, element.fill), 'FFFFFF', 100);
  }
  if (hasMeaningfulValue(style.line) || hasMeaningfulValue(element.line)) {
    options.line = normalizeLineOption(pickFirstDefined(style.line, element.line), theme.line, 1, 100);
  }
  slide.addText(content, options);
}

function renderShapeElement(slide, element, theme) {
  const style = element.style && typeof element.style === 'object' ? element.style : {};
  const box = resolveElementBox(element, { x: 0, y: 0, w: 1, h: 1 });
  const shapeType = resolveElementShapeType(slide, pickFirstDefined(style.shape, element.shape, element.shapeType, 'rect'));
  const shapeOptions = {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    fill: normalizeFillOption(
      pickFirstDefined(style.fill, element.fill),
      normalizeHexColor(pickFirstDefined(style.fillColor, element.fillColor), theme.surface),
      resolveTransparencyOption(
        pickFirstDefined(style.fillTransparency, element.fillTransparency),
        pickFirstDefined(style.fillOpacity, element.fillOpacity),
        0
      )
    ),
    line: normalizeLineOption(
      pickFirstDefined(style.line, element.line),
      normalizeHexColor(pickFirstDefined(style.lineColor, element.lineColor), theme.line),
      toNumber(pickFirstDefined(style.lineWidth, element.lineWidth), 1),
      resolveTransparencyOption(
        pickFirstDefined(style.lineTransparency, element.lineTransparency),
        pickFirstDefined(style.lineOpacity, element.lineOpacity),
        0
      )
    ),
  };
  if (shapeType === slide._pptx.shapes.ROUNDED_RECTANGLE) {
    shapeOptions.rectRadius = toNumber(pickFirstDefined(style.rectRadius, element.rectRadius, element.radius), 0.08);
  }
  if (hasMeaningfulValue(style.rotate) || hasMeaningfulValue(element.rotate)) {
    shapeOptions.rotate = toNumber(pickFirstDefined(style.rotate, element.rotate), 0);
  }
  if (style.shadow && typeof style.shadow === 'object') {
    shapeOptions.shadow = { ...style.shadow };
  } else if (element.shadow && typeof element.shadow === 'object') {
    shapeOptions.shadow = { ...element.shadow };
  }
  slide.addShape(shapeType, shapeOptions);

  if (hasMeaningfulValue(element.text) || Array.isArray(element.text)) {
    renderTextElement(
      slide,
      {
        ...element,
        type: 'text',
        margin: pickFirstDefined(element.textMargin, element.margin, 0),
        fill: false,
        line: false,
        style: {
          ...(style.textStyle && typeof style.textStyle === 'object' ? style.textStyle : {}),
          ...(element.textStyle && typeof element.textStyle === 'object' ? element.textStyle : {}),
        },
      },
      theme
    );
  }
}

async function renderImageElement(slide, element, assetMap, theme, spec, layer = 'foreground') {
  const resolved = await resolveImagePath(assetMap, element);
  if (resolved.error) {
    pushRenderWarning(spec, `${describeImageRef(element)}: ${trimForLog(resolved.error)}`);
    if (normalizeImageFallbackMode(element, layer) === 'placeholder') {
      renderImagePlaceholder(slide, element, theme);
    }
    return;
  }
  if (!resolved.path) {
    return;
  }
  const box = resolveElementBox(element, { x: 0, y: 0, w: 1, h: 1 });
  try {
    slide.addImage({
      path: resolved.path,
      x: box.x,
      y: box.y,
      w: box.w,
      h: box.h,
    });
  } catch (error) {
    pushRenderWarning(spec, `${describeImageRef(element)}: ${trimForLog(error && error.message ? error.message : error)}`);
    if (normalizeImageFallbackMode(element, layer) === 'placeholder') {
      renderImagePlaceholder(slide, element, theme);
    }
  }
}

function renderLineElement(slide, element, theme) {
  const from = element.from && typeof element.from === 'object' ? element.from : {};
  const to = element.to && typeof element.to === 'object' ? element.to : {};
  const x = toNumber(pickFirstDefined(element.x, element.x1, from.x), 0);
  const y = toNumber(pickFirstDefined(element.y, element.y1, from.y), 0);
  const derivedW = hasMeaningfulValue(element.x2) || hasMeaningfulValue(to.x) ? toNumber(pickFirstDefined(element.x2, to.x), x) - x : 1;
  const derivedH = hasMeaningfulValue(element.y2) || hasMeaningfulValue(to.y) ? toNumber(pickFirstDefined(element.y2, to.y), y) - y : 0;
  slide.addShape(slide._pptx.shapes.LINE, {
    x,
    y,
    w: toNumber(pickFirstDefined(element.w), derivedW),
    h: toNumber(pickFirstDefined(element.h), derivedH),
    line: normalizeLineOption(
      pickFirstDefined(element.line, element.style?.line),
      normalizeHexColor(pickFirstDefined(element.lineColor, element.color), theme.line),
      toNumber(pickFirstDefined(element.lineWidth, element.width), 1.2),
      resolveTransparencyOption(
        pickFirstDefined(element.lineTransparency),
        pickFirstDefined(element.lineOpacity, element.opacity),
        0
      )
    ),
  });
}

async function renderSlideElements(slide, rawSlide, assetMap, theme, spec, layer = 'foreground') {
  const elements = asArray(rawSlide.elements).filter((element) => element && typeof element === 'object');
  for (const element of elements) {
    const elementLayer = String(element.layer || 'foreground').trim().toLowerCase();
    if (elementLayer !== layer) {
      continue;
    }
    const elementType = resolveElementType(element);
    if (elementType === 'text') {
      renderTextElement(slide, element, theme);
    } else if (elementType === 'shape') {
      renderShapeElement(slide, element, theme);
    } else if (elementType === 'image') {
      await renderImageElement(slide, element, assetMap, theme, spec, layer);
    } else if (elementType === 'line') {
      renderLineElement(slide, element, theme);
    }
  }
}

function collectUniqueSources(slides) {
  const map = new Map();
  for (const slide of slides) {
    for (const entry of asArray(slide.sources).map(normalizeSourceEntry).filter(Boolean)) {
      const key = `${entry.label}||${entry.url}`;
      if (!map.has(key)) map.set(key, entry);
    }
  }
  return [...map.values()];
}

function footerText(slide) {
  const entries = asArray(slide.sources).map(normalizeSourceEntry).filter(Boolean);
  if (!entries.length) return '';
  return `Sources: ${entries.map((entry) => entry.label).join('; ')}`;
}

function addPageChrome(slide, spec, theme, index, total, dark = false) {
  const lineColor = dark ? 'FFFFFF' : theme.line;
  const textColor = dark ? 'D1D5DB' : theme.muted;
  slide.addShape(slide._pptx.shapes.RECTANGLE, {
    x: 0.55,
    y: 6.95,
    w: 12.2,
    h: 0.015,
    fill: { color: lineColor, transparency: dark ? 70 : 15 },
    line: { color: lineColor, transparency: 100 },
  });

  const footer = footerText(slide._mlca || {});
  if (footer) {
    slide.addText(footer, {
      x: 0.65,
      y: 7.0,
      w: 8.9,
      h: 0.22,
      fontFace: theme.bodyFont,
      fontSize: 8.5,
      color: textColor,
      margin: 0,
    });
  }

  if (spec.deck.showPageNumbers) {
    slide.addText(`${index + 1}/${total}`, {
      x: 11.55,
      y: 6.98,
      w: 0.75,
      h: 0.2,
      fontFace: theme.bodyFont,
      fontSize: 8.5,
      color: textColor,
      margin: 0,
      align: 'right',
    });
  }
}

function normalizePageChrome(rawValue, kind) {
  if (!hasMeaningfulValue(rawValue)) {
    return {
      enabled: kind !== 'free-layout',
      dark: false,
    };
  }
  if (typeof rawValue === 'boolean') {
    return {
      enabled: rawValue,
      dark: false,
    };
  }
  const options = rawValue && typeof rawValue === 'object' ? rawValue : {};
  return {
    enabled: options.enabled !== false,
    dark: toBoolean(options.dark, false),
  };
}

function getAccentForIndex(theme, index, fallback) {
  const colors = theme.accentColors || [];
  if (colors.length > 0) {
    return colors[index % colors.length];
  }
  return fallback || (index % 2 === 0 ? theme.support : theme.accent);
}

function getStepColor(theme, index) {
  const gradient = theme.stepGradient || [];
  if (gradient.length > 0) {
    return gradient[index % gradient.length];
  }
  const colors = theme.accentColors || [];
  if (colors.length > 0) {
    return colors[index % colors.length];
  }
  return [theme.dominant, theme.support, theme.accent][index % 3] || theme.support;
}

function isDarkTheme(theme) {
  const bg = normalizeHexColor(theme.background, 'FFFFFF');
  const r = parseInt(bg.slice(0, 2), 16);
  const g = parseInt(bg.slice(2, 4), 16);
  const b = parseInt(bg.slice(4, 6), 16);
  return (r * 299 + g * 587 + b * 114) / 1000 < 128;
}

function addMotifDecorations(slide, theme, region) {
  const motif = theme.motif || '';
  const dark = isDarkTheme(theme);
  if (motif === 'left-accent-rail') {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 0.16,
      h: 7.5,
      fill: { color: theme.borderAccent || theme.accent },
      line: { color: theme.borderAccent || theme.accent, transparency: 100 },
    });
  } else if (motif === 'header-bar') {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 13.33,
      h: 0.06,
      fill: { color: theme.borderAccent || theme.accent },
      line: { color: theme.borderAccent || theme.accent, transparency: 100 },
    });
  } else if (motif === 'gradient-cover' && region === 'cover') {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 13.33,
      h: 7.5,
      fill: { color: theme.dominant, transparency: 90 },
      line: { color: theme.dominant, transparency: 100 },
    });
  }
}

function renderFreeLayout(slide, deck, theme) {
  slide.background = {
    color: normalizeHexColor(deck.backgroundColor, theme.background),
  };
}

function addLightHeader(slide, title, eyebrow, theme) {
  slide.background = { color: theme.background };
  const motif = theme.motif || '';
  if (motif === 'header-bar') {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 13.33,
      h: 0.06,
      fill: { color: theme.borderAccent || theme.accent },
      line: { color: theme.borderAccent || theme.accent, transparency: 100 },
    });
  } else if (motif === 'left-accent-rail') {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 0.16,
      h: 7.5,
      fill: { color: theme.borderAccent || theme.accent },
      line: { color: theme.borderAccent || theme.accent, transparency: 100 },
    });
  } else {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 0.16,
      h: 7.5,
      fill: { color: theme.support },
      line: { color: theme.support, transparency: 100 },
    });
  }
  if (eyebrow) {
    slide.addText(eyebrow.toUpperCase(), {
      x: 0.65,
      y: 0.42,
      w: 3.3,
      h: 0.22,
      fontFace: theme.bodyFont,
      fontSize: 10,
      color: theme.support,
      bold: true,
      charSpacing: 2.4,
      margin: 0,
    });
  }
  slide.addText(title, {
    x: 0.65,
    y: eyebrow ? 0.7 : 0.52,
    w: 7.9,
    h: 0.55,
    fontFace: theme.headerFont,
    fontSize: 24,
    color: theme.text,
    bold: true,
    margin: 0,
  });
}

function addCard(slide, x, y, w, h, theme, accent) {
  const cardBg = theme.cardBackground || theme.surface;
  const cardTrans = theme.cardTransparency || 0;
  const borderAccentColor = theme.borderAccent || '';
  const borderAccentW = theme.borderAccentWidth || 0;
  slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fill: { color: cardBg, transparency: cardTrans },
    line: { color: theme.line, transparency: 55, width: 1 },
    shadow: { type: 'outer', color: '000000', blur: 2, offset: 1, angle: 45, opacity: 0.08 },
  });
  if (borderAccentColor && borderAccentW > 0) {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x,
      y,
      w: borderAccentW / 72,
      h,
      fill: { color: borderAccentColor },
      line: { color: borderAccentColor, transparency: 100 },
    });
  }
  slide.addShape(slide._pptx.shapes.RECTANGLE, {
    x,
    y,
    w,
    h: 0.06,
    fill: { color: accent || theme.support },
    line: { color: accent || theme.support, transparency: 100 },
  });
}

function addPill(slide, x, y, w, h, text, fillColor, textColor, theme) {
  slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fill: { color: fillColor },
    line: { color: fillColor, transparency: 100 },
  });
  slide.addText(String(text || ''), {
    x,
    y: y + 0.02,
    w,
    h,
    fontFace: theme.bodyFont,
    fontSize: 8.5,
    color: textColor,
    bold: true,
    margin: 0,
    align: 'center',
    valign: 'mid',
  });
}

function renderTagRow(slide, tags, x, y, maxWidth, theme, fillColor = 'F8FAFC', textColor = '') {
  const items = asArray(tags).map((item) => String(item || '').trim()).filter(Boolean);
  let currentX = x;
  let currentY = y;
  items.forEach((item) => {
    const width = Math.min(1.2 + item.length * 0.08, 1.9);
    if (currentX + width > x + maxWidth) {
      currentX = x;
      currentY += 0.26;
    }
    addPill(slide, currentX, currentY, width, 0.22, item, fillColor, textColor || theme.muted, theme);
    currentX += width + 0.08;
  });
}

function renderCover(slide, deck, theme) {
  slide.background = { color: theme.dominant };
  const motif = theme.motif || '';
  if (motif === 'left-accent-rail') {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 0.35,
      h: 7.5,
      fill: { color: theme.borderAccent || theme.accent },
      line: { color: theme.borderAccent || theme.accent, transparency: 100 },
    });
  } else if (motif === 'header-bar') {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 13.33,
      h: 0.08,
      fill: { color: theme.borderAccent || theme.accent },
      line: { color: theme.borderAccent || theme.accent, transparency: 100 },
    });
    slide.addShape(slide._pptx.shapes.OVAL, {
      x: 8.8,
      y: -1.0,
      w: 4.0,
      h: 4.0,
      fill: { color: theme.accent, transparency: 68 },
      line: { color: theme.accent, transparency: 100 },
    });
  } else {
    slide.addShape(slide._pptx.shapes.OVAL, {
      x: 8.8,
      y: -1.0,
      w: 4.0,
      h: 4.0,
      fill: { color: theme.support, transparency: 68 },
      line: { color: theme.support, transparency: 100 },
    });
    slide.addShape(slide._pptx.shapes.OVAL, {
      x: 10.3,
      y: -0.2,
      w: 3.0,
      h: 3.0,
      fill: { color: theme.accent, transparency: 78 },
      line: { color: theme.accent, transparency: 100 },
    });
  }
  if (deck.eyebrow) {
    slide.addText(deck.eyebrow.toUpperCase(), {
      x: 0.85,
      y: 1.1,
      w: 4.5,
      h: 0.25,
      fontFace: theme.bodyFont,
      fontSize: 11,
      color: theme.accent,
      bold: true,
      charSpacing: 3,
      margin: 0,
    });
  }
  slide.addText(deck.title || 'Untitled Deck', {
    x: 0.85,
    y: 1.55,
    w: 7.4,
    h: 1.0,
    fontFace: theme.headerFont,
    fontSize: 27,
    color: 'FFFFFF',
    bold: true,
    margin: 0,
  });
  if (deck.subtitle) {
    slide.addText(deck.subtitle, {
      x: 0.85,
      y: 2.7,
      w: 6.9,
      h: 0.55,
      fontFace: theme.bodyFont,
      fontSize: 13,
      color: 'D6E4FF',
      margin: 0,
    });
  }
  slide.addShape(slide._pptx.shapes.RECTANGLE, {
    x: 0.85,
    y: 3.55,
    w: 1.2,
    h: 0.05,
    fill: { color: theme.accent },
    line: { color: theme.accent, transparency: 100 },
  });
}

function renderAgenda(slide, deck, theme) {
  addLightHeader(slide, deck.title || 'Agenda', deck.eyebrow || '', theme);
  const items = asArray(deck.items);
  const cols = items.length > 4 ? 3 : 2;
  const colWidth = cols === 3 ? 3.85 : 5.85;
  items.forEach((item, index) => {
    const col = index % cols;
    const row = Math.floor(index / cols);
    const x = 0.65 + col * (colWidth + 0.25);
    const y = 1.55 + row * 1.45;
    addCard(slide, x, y, colWidth, 1.08, theme, col % 2 === 0 ? theme.support : theme.accent);
    const isObj = item && typeof item === 'object';
    const num = isObj ? (item.num || String(index + 1).padStart(2, '0')) : String(index + 1).padStart(2, '0');
    const title = isObj ? (item.title || '') : String(item);
    const body = isObj ? (item.body || item.subtitle || '') : '';
    slide.addText(num, {
      x: x + 0.22,
      y: y + 0.18,
      w: 0.55,
      h: 0.25,
      fontFace: theme.bodyFont,
      fontSize: 14,
      color: theme.support,
      bold: true,
      margin: 0,
    });
    slide.addText(title, {
      x: x + 0.22,
      y: y + (body ? 0.32 : 0.46),
      w: colWidth - 0.44,
      h: 0.32,
      fontFace: theme.headerFont,
      fontSize: 15,
      color: theme.text,
      bold: true,
      margin: 0,
      fit: 'shrink',
    });
    if (body) {
      slide.addText(body, {
        x: x + 0.22,
        y: y + 0.68,
        w: colWidth - 0.44,
        h: 0.3,
        fontFace: theme.bodyFont,
        fontSize: 10,
        color: theme.muted,
        margin: 0,
        fit: 'shrink',
      });
    }
  });
}

function renderSection(slide, deck, theme) {
  slide.background = { color: theme.support };
  const motif = theme.motif || '';
  if (motif === 'left-accent-rail') {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 0.35,
      h: 7.5,
      fill: { color: theme.borderAccent || theme.accent },
      line: { color: theme.borderAccent || theme.accent, transparency: 100 },
    });
  } else if (motif === 'header-bar') {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 13.33,
      h: 0.08,
      fill: { color: theme.borderAccent || theme.accent },
      line: { color: theme.borderAccent || theme.accent, transparency: 100 },
    });
  } else {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 0.9,
      y: 1.45,
      w: 4.2,
      h: 0.08,
      fill: { color: 'FFFFFF', transparency: 20 },
      line: { color: 'FFFFFF', transparency: 100 },
    });
  }
  slide.addText(deck.title || 'Section', {
    x: 0.9,
    y: 1.9,
    w: 8.2,
    h: 0.9,
    fontFace: theme.headerFont,
    fontSize: 30,
    color: 'FFFFFF',
    bold: true,
    margin: 0,
  });
  if (deck.subtitle) {
    slide.addText(deck.subtitle, {
      x: 0.92,
      y: 2.95,
      w: 7.5,
      h: 0.42,
      fontFace: theme.bodyFont,
      fontSize: 13,
      color: 'E5EEF9',
      margin: 0,
    });
  }
}

function renderBullets(slide, deck, theme, spec) {
  addLightHeader(slide, deck.title || 'Key Points', deck.eyebrow || '', theme);
  const rect = getContentRect(spec, deck);
  const hasAside = !!(deck.asideTitle || deck.asideText || deck.message);
  const gap = rect.w * 0.025;
  const leftRatio = hasAside ? 0.62 : 1.0;
  const leftW = hasAside ? (rect.w - gap) * leftRatio : rect.w;
  const rightW = hasAside ? rect.w - gap - leftW : 0;

  addCard(slide, rect.x, rect.y, leftW, rect.h, theme, theme.support);
  const bulletPad = leftW * 0.04;
  const runs = [];
  const bulletItems = asArray(deck.bullets).length ? asArray(deck.bullets) : asArray(deck.items);
  const bulletFontSize = adaptiveFontSize(bulletItems.join('\n'), leftW - bulletPad * 2, rect.h * 0.8, 14, 10);
  bulletItems.forEach((item, index, arr) => {
    runs.push({ text: String(item), options: { bullet: true, breakLine: index < arr.length - 1, color: theme.text, fontSize: bulletFontSize } });
  });
  slide.addText(runs.length ? runs : [{ text: 'Add bullet items here', options: { bullet: true, color: theme.text, fontSize: 14 } }], {
    x: rect.x + bulletPad,
    y: rect.y + rect.h * 0.09,
    w: leftW - bulletPad * 2,
    h: rect.h * 0.85,
    fontFace: theme.bodyFont,
    margin: 0,
    valign: 'top',
    fit: 'shrink',
  });

  if (hasAside) {
    const rightX = rect.x + leftW + gap;
    addCard(slide, rightX, rect.y, rightW, rect.h, theme, theme.accent);
    const asidePad = rightW * 0.06;
    const asideInnerW = rightW - asidePad * 2;
    slide.addText(deck.asideTitle || 'So what', {
      x: rightX + asidePad,
      y: rect.y + rect.h * 0.06,
      w: asideInnerW,
      h: rect.h * 0.1,
      fontFace: theme.headerFont,
      fontSize: Math.min(15, rightW * 3.5),
      color: theme.text,
      bold: true,
      margin: 0,
      fit: 'shrink',
    });
    const asideText = deck.asideText || deck.message || '';
    slide.addText(asideText, {
      x: rightX + asidePad,
      y: rect.y + rect.h * 0.2,
      w: asideInnerW,
      h: rect.h * 0.72,
      fontFace: theme.bodyFont,
      fontSize: adaptiveFontSize(asideText, asideInnerW, rect.h * 0.72, 12.5, 9),
      color: theme.muted,
      margin: 0,
      valign: 'top',
      fit: 'shrink',
    });
  }
}

function renderTwoColumn(slide, deck, theme, spec) {
  addLightHeader(slide, deck.title || 'Overview', deck.eyebrow || '', theme);
  const rect = getContentRect(spec, deck);
  const leftContent = String(deck.leftBody || asArray(deck.leftBullets).join('\n') || '');
  const rightContent = String(deck.rightBody || asArray(deck.rightBullets).join('\n') || '');
  const panels = calcTwoPanelSplit(leftContent, rightContent, rect);

  addCard(slide, panels.left.x, panels.left.y, panels.left.w, panels.left.h, theme, theme.support);
  addCard(slide, panels.right.x, panels.right.y, panels.right.w, panels.right.h, theme, theme.accent);

  const padL = panels.left.w * 0.04;
  const padR = panels.right.w * 0.04;
  const innerLW = panels.left.w - padL * 2;
  const innerRW = panels.right.w - padR * 2;
  const titleH = panels.left.h * 0.08;
  const bodyH = panels.left.h * 0.8;

  slide.addText(deck.leftTitle || 'Left column', {
    x: panels.left.x + padL,
    y: panels.left.y + panels.left.h * 0.05,
    w: innerLW,
    h: titleH,
    fontFace: theme.headerFont,
    fontSize: Math.min(15, panels.left.w * 2.5),
    color: theme.text,
    bold: true,
    margin: 0,
    fit: 'shrink',
  });
  const leftBullets = asArray(deck.leftBullets);
  if (leftBullets.length) {
    const runs = [];
    const lbFontSize = adaptiveFontSize(leftBullets.join('\n'), innerLW, bodyH, 13, 9);
    leftBullets.forEach((item, index) => {
      runs.push({ text: String(item), options: { bullet: true, breakLine: index < leftBullets.length - 1, color: theme.text, fontSize: lbFontSize } });
    });
    slide.addText(runs, {
      x: panels.left.x + padL,
      y: panels.left.y + panels.left.h * 0.16,
      w: innerLW,
      h: bodyH,
      fontFace: theme.bodyFont,
      margin: 0,
      valign: 'top',
      fit: 'shrink',
    });
  } else {
    const leftBodyText = String(deck.leftBody || '');
    slide.addText(leftBodyText, {
      x: panels.left.x + padL,
      y: panels.left.y + panels.left.h * 0.16,
      w: innerLW,
      h: bodyH,
      fontFace: theme.bodyFont,
      fontSize: adaptiveFontSize(leftBodyText, innerLW, bodyH, 11.5, 8.5),
      color: theme.muted,
      margin: 0,
      valign: 'top',
      fit: 'shrink',
    });
  }

  slide.addText(deck.rightTitle || 'Right column', {
    x: panels.right.x + padR,
    y: panels.right.y + panels.right.h * 0.05,
    w: innerRW,
    h: titleH,
    fontFace: theme.headerFont,
    fontSize: Math.min(15, panels.right.w * 2.5),
    color: theme.text,
    bold: true,
    margin: 0,
    fit: 'shrink',
  });
  const rightBullets = asArray(deck.rightBullets);
  if (rightBullets.length) {
    const runs = [];
    const rbFontSize = adaptiveFontSize(rightBullets.join('\n'), innerRW, bodyH, 13, 9);
    rightBullets.forEach((item, index) => {
      runs.push({ text: String(item), options: { bullet: true, breakLine: index < rightBullets.length - 1, color: theme.text, fontSize: rbFontSize } });
    });
    slide.addText(runs, {
      x: panels.right.x + padR,
      y: panels.right.y + panels.right.h * 0.16,
      w: innerRW,
      h: bodyH,
      fontFace: theme.bodyFont,
      margin: 0,
      valign: 'top',
      fit: 'shrink',
    });
  } else {
    const rightBodyText = String(deck.rightBody || '');
    slide.addText(rightBodyText, {
      x: panels.right.x + padR,
      y: panels.right.y + panels.right.h * 0.16,
      w: innerRW,
      h: bodyH,
      fontFace: theme.bodyFont,
      fontSize: adaptiveFontSize(rightBodyText, innerRW, bodyH, 11.5, 8.5),
      color: theme.muted,
      margin: 0,
      valign: 'top',
      fit: 'shrink',
    });
  }
}

function renderStats(slide, deck, theme, spec) {
  addLightHeader(slide, deck.title || 'Key Metrics', deck.eyebrow || '', theme);
  const items = asArray(deck.items);
  const rect = getContentRect(spec, deck);
  const cells = calcEqualGrid(items.length, rect);
  items.forEach((item, index) => {
    const cell = cells[index];
    if (!cell) return;
    const pad = cell.w * 0.05;
    const innerW = cell.w - pad * 2;
    addCard(slide, cell.x, cell.y, cell.w, cell.h, theme, getAccentForIndex(theme, index));
    const valueFontSize = Math.min(24, cell.h * 3.2);
    slide.addText(String(item.value || '--'), {
      x: cell.x + pad,
      y: cell.y + cell.h * 0.1,
      w: innerW,
      h: cell.h * 0.3,
      fontFace: theme.headerFont,
      fontSize: valueFontSize,
      color: getAccentForIndex(theme, index),
      bold: true,
      margin: 0,
      align: 'center',
      fit: 'shrink',
    });
    slide.addText(String(item.label || ''), {
      x: cell.x + pad,
      y: cell.y + cell.h * 0.44,
      w: innerW,
      h: cell.h * 0.22,
      fontFace: theme.headerFont,
      fontSize: Math.min(13, cell.h * 1.7),
      color: theme.text,
      bold: true,
      margin: 0,
      align: 'center',
      fit: 'shrink',
    });
    slide.addText(String(item.body || item.note || ''), {
      x: cell.x + pad,
      y: cell.y + cell.h * 0.68,
      w: innerW,
      h: cell.h * 0.25,
      fontFace: theme.bodyFont,
      fontSize: Math.min(9.5, cell.h * 1.3),
      color: theme.muted,
      margin: 0,
      align: 'center',
      fit: 'shrink',
    });
  });
}

function renderGrid(slide, deck, theme, spec) {
  addLightHeader(slide, deck.title || 'Framework', deck.eyebrow || '', theme);
  const items = asArray(deck.items);
  const rect = getContentRect(spec, deck);
  const cells = calcEqualGrid(items.length, rect);
  items.forEach((item, index) => {
    const cell = cells[index];
    if (!cell) return;
    const pad = cell.w * 0.04;
    const innerW = cell.w - pad * 2;
    addCard(slide, cell.x, cell.y, cell.w, cell.h, theme, item.accent || getAccentForIndex(theme, index));
    const titleH = cell.h * 0.18;
    const bodyH = cell.h * 0.7;
    const bodyText = String(item.body || item.text || '');
    const bodyFontSize = adaptiveFontSize(bodyText, innerW, bodyH, 11.5, 8);
    slide.addText(String(item.title || `Block ${index + 1}`), {
      x: cell.x + pad,
      y: cell.y + pad,
      w: innerW,
      h: titleH,
      fontFace: theme.headerFont,
      fontSize: Math.min(15, cell.h * 1.9),
      color: theme.text,
      bold: true,
      margin: 0,
      fit: 'shrink',
    });
    slide.addText(bodyText, {
      x: cell.x + pad,
      y: cell.y + pad + titleH + cell.h * 0.05,
      w: innerW,
      h: bodyH,
      fontFace: theme.bodyFont,
      fontSize: bodyFontSize,
      color: theme.muted,
      margin: 0,
      valign: 'top',
      fit: 'shrink',
    });
  });
}

function renderCareerJourney(slide, deck, theme) {
  addLightHeader(slide, deck.title || '个人背景与转型动机', deck.eyebrow || '', theme);
  addCard(slide, 0.65, 1.58, 5.35, 4.75, theme, theme.support);
  slide.addText(String(deck.intro || deck.message || ''), {
    x: 0.95,
    y: 1.9,
    w: 4.75,
    h: 0.42,
    fontFace: theme.bodyFont,
    fontSize: 11.5,
    color: theme.muted,
    margin: 0,
  });

  asArray(deck.sections).slice(0, 3).forEach((section, index) => {
    const y = 2.45 + index * 1.18;
    slide.addText(String(section.title || `Section ${index + 1}`), {
      x: 0.95,
      y,
      w: 4.65,
      h: 0.24,
      fontFace: theme.headerFont,
      fontSize: 13,
      color: theme.support,
      bold: true,
      margin: 0,
    });
    const bullets = asArray(section.bullets);
    const runs = [];
    bullets.forEach((item, bulletIndex) => {
      runs.push({
        text: String(item),
        options: {
          bullet: true,
          breakLine: bulletIndex < bullets.length - 1,
          color: theme.text,
          fontSize: 10,
        },
      });
    });
    slide.addText(runs, {
      x: 1.0,
      y: y + 0.28,
      w: 4.45,
      h: 0.72,
      fontFace: theme.bodyFont,
      margin: 0,
      valign: 'top',
    });
  });

  slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
    x: 6.25,
    y: 1.58,
    w: 5.95,
    h: 4.75,
    rectRadius: 0.08,
    fill: { color: 'F0F4F8' },
    line: { color: theme.line, transparency: 55, width: 1 },
  });
  slide.addText(String(deck.rightEyebrow || 'CAREER PATH'), {
    x: 10.4,
    y: 1.82,
    w: 1.4,
    h: 0.18,
    fontFace: theme.bodyFont,
    fontSize: 8,
    color: theme.muted,
    bold: true,
    charSpacing: 1.8,
    margin: 0,
    align: 'right',
  });

  asArray(deck.phases).slice(0, 3).forEach((phase, index, list) => {
    const y = 2.1 + index * 1.32;
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 6.68,
      y,
      w: 0.05,
      h: 1.02,
      fill: { color: getStepColor(theme, index) },
      line: { color: theme.line, transparency: 100 },
    });
    slide.addShape(slide._pptx.shapes.OVAL, {
      x: 6.56,
      y: y + 0.15,
      w: 0.28,
      h: 0.28,
      fill: { color: 'FFFFFF' },
      line: { color: getStepColor(theme, index), width: 2 },
    });
    addCard(slide, 6.95, y - 0.04, 4.82, 1.06, theme, getStepColor(theme, index));
    slide.addText(String(phase.title || `Phase ${index + 1}`), {
      x: 7.18,
      y: y + 0.12,
      w: 3.2,
      h: 0.24,
      fontFace: theme.headerFont,
      fontSize: 12.5,
      color: theme.text,
      bold: true,
      margin: 0,
    });
    addPill(slide, 10.6, y + 0.08, 0.95, 0.22, String(phase.stageLabel || phase.stage || ''), 'F1F5F9', theme.muted, theme);
    slide.addText(String(phase.focus || phase.body || ''), {
      x: 7.18,
      y: y + 0.42,
      w: 4.1,
      h: 0.2,
      fontFace: theme.bodyFont,
      fontSize: 9.4,
      color: theme.muted,
      margin: 0,
    });
    renderTagRow(slide, phase.chips || [], 7.18, y + 0.68, 3.9, theme, 'FFFFFF', theme.muted);
    if (index < list.length - 1) {
      slide.addText('↓', {
        x: 6.48,
        y: y + 1.03,
        w: 0.4,
        h: 0.24,
        fontFace: theme.headerFont,
        fontSize: 13,
        color: 'BFD5EA',
        bold: true,
        margin: 0,
        align: 'center',
      });
    }
  });
}

function renderProcessFlow(slide, deck, theme, spec) {
  addLightHeader(slide, deck.title || '我的产品方法论', deck.eyebrow || '', theme);
  const canvas = getCanvas(spec);
  const margin = getMargin(canvas);
  const headerH = getHeaderHeight(canvas);

  const hasIntro = !!(deck.intro || deck.message);
  const introH = hasIntro ? canvas.h * 0.05 : 0;
  if (hasIntro) {
    slide.addText(String(deck.intro || deck.message || ''), {
      x: margin.left,
      y: headerH,
      w: canvas.w * 0.62,
      h: introH,
      fontFace: theme.bodyFont,
      fontSize: 11.5,
      color: theme.muted,
      margin: 0,
      fit: 'shrink',
    });
    if (deck.loopLabel) {
      const pillW = canvas.w * 0.14;
      addPill(slide, canvas.w - margin.right - pillW, headerH, pillW, introH * 0.75, String(deck.loopLabel), 'FFFFFF', theme.text, theme);
    }
  }

  const contentTop = headerH + introH + canvas.h * 0.02;
  const contentH = canvas.h - contentTop - margin.bottom;
  const contentW = canvas.w - margin.left - margin.right;
  const contentRect = { x: margin.left, y: contentTop, w: contentW, h: contentH };

  const items = asArray(deck.items);
  const cols = calcEqualColumns(items.length, contentRect);

  items.forEach((item, index) => {
    const col = cols[index];
    if (!col) return;
    const pad = col.w * 0.07;
    const innerW = col.w - pad * 2;
    const accent = item.accent || getStepColor(theme, index);
    addCard(slide, col.x, col.y, col.w, col.h, theme, accent);

    const stepH = col.h * 0.05;
    const titleH = col.h * 0.08;
    const bodyH = col.h * 0.35;
    const outputBoxTop = col.y + col.h * 0.55;
    const outputBoxH = col.h * 0.4;

    slide.addText(String(item.step || `STEP ${String(index + 1).padStart(2, '0')}`), {
      x: col.x + pad,
      y: col.y + col.h * 0.04,
      w: innerW,
      h: stepH,
      fontFace: theme.bodyFont,
      fontSize: Math.min(8.2, col.w * 4),
      color: theme.muted,
      bold: true,
      margin: 0,
      align: 'center',
    });
    slide.addText(String(item.title || `步骤 ${index + 1}`), {
      x: col.x + pad,
      y: col.y + col.h * 0.11,
      w: innerW,
      h: titleH,
      fontFace: theme.headerFont,
      fontSize: adaptiveFontSize(item.title || '', innerW, titleH, 11.5, 8.5),
      color: theme.text,
      bold: true,
      margin: 0,
      align: 'center',
      fit: 'shrink',
    });

    const bodyText = String(item.body || item.message || '');
    slide.addText(bodyText, {
      x: col.x + pad,
      y: col.y + col.h * 0.22,
      w: innerW,
      h: bodyH,
      fontFace: theme.bodyFont,
      fontSize: adaptiveFontSize(bodyText, innerW, bodyH, 8.8, 7),
      color: theme.muted,
      margin: 0,
      align: 'center',
      valign: 'mid',
      fit: 'shrink',
    });

    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x: col.x + pad * 0.8,
      y: outputBoxTop,
      w: innerW + pad * 0.4,
      h: outputBoxH,
      rectRadius: 0.05,
      fill: { color: 'F7FAFC' },
      line: { color: theme.line, transparency: 35, width: 1 },
    });
    slide.addText(String(item.outputTitle || '关键产出'), {
      x: col.x + pad,
      y: outputBoxTop + outputBoxH * 0.06,
      w: innerW,
      h: outputBoxH * 0.15,
      fontFace: theme.bodyFont,
      fontSize: Math.min(7.8, col.w * 3.8),
      color: theme.muted,
      bold: true,
      margin: 0,
      align: 'center',
    });
    const outputsText = asArray(item.outputs).map((output) => `• ${output}`).join('\n');
    slide.addText(outputsText, {
      x: col.x + pad,
      y: outputBoxTop + outputBoxH * 0.25,
      w: innerW,
      h: outputBoxH * 0.7,
      fontFace: theme.bodyFont,
      fontSize: adaptiveFontSize(outputsText, innerW, outputBoxH * 0.7, 8.2, 6.5),
      color: theme.text,
      margin: 0,
      valign: 'top',
      align: 'left',
      fit: 'shrink',
    });

    if (index < items.length - 1) {
      const arrowX = col.x + col.w;
      const arrowY = col.y + col.h * 0.5;
      const gapW = cols[index + 1] ? cols[index + 1].x - (col.x + col.w) : 0;
      if (gapW > 0) {
        slide.addText('›', {
          x: arrowX,
          y: arrowY - 0.1,
          w: gapW,
          h: 0.2,
          fontFace: theme.headerFont,
          fontSize: 12,
          color: 'CBD5E0',
          bold: true,
          margin: 0,
          align: 'center',
        });
      }
    }
  });
}

function renderCapabilityMatrix(slide, deck, theme) {
  addLightHeader(slide, deck.title || '产品经理核心能力结构', deck.eyebrow || '', theme);
  addCard(slide, 0.65, 1.58, 3.45, 4.75, theme, theme.dominant);
  slide.addText(String(deck.summaryTitle || '能力模型评估'), {
    x: 0.9,
    y: 1.9,
    w: 2.7,
    h: 0.24,
    fontFace: theme.headerFont,
    fontSize: 13,
    color: theme.support,
    bold: true,
    margin: 0,
  });
  asArray(deck.items).slice(0, 5).forEach((item, index) => {
    const y = 2.35 + index * 0.52;
    const score = Math.max(0, Math.min(Number(item.score || 0), 5));
    slide.addText(String(item.title || `能力 ${index + 1}`), {
      x: 0.95,
      y,
      w: 1.35,
      h: 0.16,
      fontFace: theme.bodyFont,
      fontSize: 8.8,
      color: theme.text,
      margin: 0,
    });
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x: 2.05,
      y: y + 0.02,
      w: 1.62,
      h: 0.1,
      rectRadius: 0.03,
      fill: { color: 'E2E8F0' },
      line: { color: 'E2E8F0', transparency: 100 },
    });
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x: 2.05,
      y: y + 0.02,
      w: 1.62 * (score / 5),
      h: 0.1,
      rectRadius: 0.03,
      fill: { color: getAccentForIndex(theme, index, theme.dominant) },
      line: { color: getAccentForIndex(theme, index, theme.dominant), transparency: 100 },
    });
    slide.addText(String(item.scoreLabel || score.toFixed(1)), {
      x: 3.74,
      y: y - 0.02,
      w: 0.2,
      h: 0.16,
      fontFace: theme.bodyFont,
      fontSize: 8.8,
      color: theme.muted,
      bold: true,
      margin: 0,
      align: 'right',
    });
  });
  slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.95,
    y: 4.95,
    w: 2.85,
    h: 1.05,
    rectRadius: 0.05,
    fill: { color: 'F8FAFC' },
    line: { color: theme.line, transparency: 45, width: 1 },
  });
  slide.addText(String(deck.summaryHeadline || 'T型人才画像'), {
    x: 1.1,
    y: 5.18,
    w: 2.45,
    h: 0.18,
    fontFace: theme.headerFont,
    fontSize: 10.5,
    color: theme.text,
    bold: true,
    margin: 0,
    align: 'center',
  });
  slide.addText(String(deck.summaryBody || ''), {
    x: 1.12,
    y: 5.48,
    w: 2.4,
    h: 0.34,
    fontFace: theme.bodyFont,
    fontSize: 8.4,
    color: theme.muted,
    margin: 0,
    align: 'center',
    valign: 'mid',
  });

  const items = asArray(deck.items).slice(0, 5);
  items.forEach((item, index) => {
    const y = 1.58 + index * 0.94;
    addCard(slide, 4.35, y, 7.85, 0.78, theme, item.accent || getStepColor(theme, index));
    slide.addText(String(item.title || `能力 ${index + 1}`), {
      x: 4.58,
      y: y + 0.16,
      w: 1.2,
      h: 0.18,
      fontFace: theme.headerFont,
      fontSize: 10.5,
      color: theme.text,
      bold: true,
      margin: 0,
    });
    slide.addText(String(item.subtitle || ''), {
      x: 4.58,
      y: y + 0.37,
      w: 1.2,
      h: 0.14,
      fontFace: theme.bodyFont,
      fontSize: 7.4,
      color: theme.muted,
      margin: 0,
    });
    slide.addText(asArray(item.tools).map(String).join(' · '), {
      x: 5.95,
      y: y + 0.18,
      w: 2.2,
      h: 0.34,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: theme.muted,
      margin: 0,
      valign: 'mid',
    });
    slide.addText(String(item.output || item.outputs || ''), {
      x: 8.3,
      y: y + 0.16,
      w: 3.45,
      h: 0.34,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: theme.text,
      margin: 0,
      valign: 'mid',
    });
  });
}

function renderMarketDriverStack(slide, deck, theme) {
  addLightHeader(slide, deck.title || '案例背景：碳足迹/LCA数字化平台', deck.eyebrow || '', theme);
  asArray(deck.sections).slice(0, 4).forEach((section, index) => {
    const y = 1.7 + index * 1.02;
    slide.addText(String(section.title || `Section ${index + 1}`), {
      x: 0.65,
      y,
      w: 5.45,
      h: 0.22,
      fontFace: theme.headerFont,
      fontSize: 12.5,
      color: theme.support,
      bold: true,
      margin: 0,
    });
    slide.addText(asArray(section.bullets).map((item) => `• ${item}`).join('\n'), {
      x: 0.72,
      y: y + 0.28,
      w: 5.1,
      h: 0.62,
      fontFace: theme.bodyFont,
      fontSize: 9.2,
      color: theme.text,
      margin: 0,
      valign: 'top',
    });
  });

  slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
    x: 6.55,
    y: 1.58,
    w: 5.65,
    h: 4.78,
    rectRadius: 0.08,
    fill: { color: 'F0F4F8' },
    line: { color: theme.line, transparency: 55, width: 1 },
  });
  slide.addText(String(deck.rightEyebrow || 'MARKET DRIVERS & SOLUTION'), {
    x: 9.6,
    y: 1.84,
    w: 2.1,
    h: 0.18,
    fontFace: theme.bodyFont,
    fontSize: 7.8,
    color: theme.muted,
    bold: true,
    charSpacing: 1.8,
    margin: 0,
    align: 'right',
  });

  asArray(deck.drivers).slice(0, 3).forEach((driver, index, list) => {
    const y = 2.15 + index * 1.08;
    addCard(slide, 6.9, y, 4.85, 0.82, theme, driver.accent || getAccentForIndex(theme, index));
    addPill(
      slide,
      7.12,
      y + 0.1,
      1.45,
      0.18,
      String(driver.badge || ''),
      driver.badgeColor || (driver.accent || theme.support),
      'FFFFFF',
      theme
    );
    slide.addText(String(driver.title || `Driver ${index + 1}`), {
      x: 7.12,
      y: y + 0.34,
      w: 3.7,
      h: 0.18,
      fontFace: theme.headerFont,
      fontSize: 11,
      color: theme.text,
      bold: true,
      margin: 0,
    });
    slide.addText(String(driver.body || ''), {
      x: 7.12,
      y: y + 0.55,
      w: 4.05,
      h: 0.14,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: theme.muted,
      margin: 0,
    });
    if (index < list.length - 1) {
      slide.addText('↓', {
        x: 9.18,
        y: y + 0.84,
        w: 0.3,
        h: 0.18,
        fontFace: theme.headerFont,
        fontSize: 12,
        color: 'CBD5E0',
        bold: true,
        margin: 0,
        align: 'center',
      });
    }
  });

  slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
    x: 7.18,
    y: 5.52,
    w: 4.28,
    h: 0.52,
    rectRadius: 0.05,
    fill: { color: 'E6F6FC' },
    line: { color: theme.accent, width: 1.2, dash: 'dash' },
  });
  addPill(slide, 8.82, 5.39, 0.92, 0.18, String(deck.solutionLabel || 'SOLUTION'), 'E6F6FC', theme.support, theme);
  slide.addText(String(deck.solutionTitle || 'LCA 数字化平台'), {
    x: 7.48,
    y: 5.66,
    w: 3.7,
    h: 0.16,
    fontFace: theme.headerFont,
    fontSize: 11.5,
    color: theme.dominant,
    bold: true,
    margin: 0,
    align: 'center',
  });
  slide.addText(String(deck.solutionBody || '标准化 · 自动化 · 协同化'), {
    x: 7.48,
    y: 5.86,
    w: 3.7,
    h: 0.12,
    fontFace: theme.bodyFont,
    fontSize: 7.8,
    color: theme.muted,
    margin: 0,
    align: 'center',
  });
}

function renderPersonaCards(slide, deck, theme) {
  addLightHeader(slide, deck.title || '用户与场景分析', deck.eyebrow || '', theme);
  const personas = asArray(deck.items).slice(0, 3);
  personas.forEach((persona, index) => {
    const x = 0.65 + index * 4.0;
    const accent = persona.accent || getAccentForIndex(theme, index);
    addCard(slide, x, 1.6, 3.72, 4.72, theme, accent);
    slide.addShape(slide._pptx.shapes.OVAL, {
      x: x + 1.48,
      y: 1.92,
      w: 0.72,
      h: 0.72,
      fill: { color: accent, transparency: 88 },
      line: { color: accent, transparency: 100 },
    });
    slide.addText(String(persona.role || persona.title || `Persona ${index + 1}`), {
      x: x + 0.25,
      y: 2.78,
      w: 3.2,
      h: 0.24,
      fontFace: theme.headerFont,
      fontSize: 12,
      color: theme.text,
      bold: true,
      margin: 0,
      align: 'center',
    });
    slide.addText(String(persona.subtitle || ''), {
      x: x + 0.25,
      y: 3.05,
      w: 3.2,
      h: 0.16,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: theme.muted,
      margin: 0,
      align: 'center',
    });
    slide.addText('核心职责', {
      x: x + 0.22,
      y: 3.36,
      w: 0.9,
      h: 0.16,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: accent,
      bold: true,
      margin: 0,
    });
    slide.addText(asArray(persona.responsibilities).map((item) => `• ${item}`).join('\n'), {
      x: x + 0.28,
      y: 3.56,
      w: 3.05,
      h: 0.72,
      fontFace: theme.bodyFont,
      fontSize: 8.4,
      color: theme.text,
      margin: 0,
      valign: 'top',
    });
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.22,
      y: 4.42,
      w: 3.18,
      h: 0.82,
      rectRadius: 0.04,
      fill: { color: 'FFF5F5' },
      line: { color: 'FEB2B2', transparency: 35, width: 1 },
    });
    slide.addText('核心痛点', {
      x: x + 0.34,
      y: 4.53,
      w: 0.8,
      h: 0.14,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: 'C53030',
      bold: true,
      margin: 0,
    });
    slide.addText(asArray(persona.pains).map((item) => `• ${item}`).join('\n'), {
      x: x + 0.34,
      y: 4.73,
      w: 2.86,
      h: 0.38,
      fontFace: theme.bodyFont,
      fontSize: 7.6,
      color: 'C53030',
      margin: 0,
      valign: 'top',
    });
    slide.addText('关键任务', {
      x: x + 0.22,
      y: 5.42,
      w: 0.8,
      h: 0.14,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: theme.muted,
      bold: true,
      margin: 0,
    });
    renderTagRow(slide, persona.jobs || [], x + 0.24, 5.64, 3.0, theme, 'F8FAFC', theme.muted);
  });
}

function renderProblemCards(slide, deck, theme) {
  addLightHeader(slide, deck.title || '问题定义：数据、计算、协作的断裂点', deck.eyebrow || '', theme);
  const columns = asArray(deck.columns).slice(0, 3);
  columns.forEach((column, index) => {
    const x = 0.65 + index * 4.0;
    const accent = column.accent || getAccentForIndex(theme, index);
    addCard(slide, x, 1.6, 3.72, 4.72, theme, accent);
    addPill(slide, x + 2.72, 1.78, 0.74, 0.2, String(column.priority || ''), 'FFF5F5', 'C53030', theme);
    slide.addText(String(column.title || `Problem ${index + 1}`), {
      x: x + 0.24,
      y: 2.05,
      w: 2.1,
      h: 0.22,
      fontFace: theme.headerFont,
      fontSize: 12.2,
      color: theme.text,
      bold: true,
      margin: 0,
    });
    slide.addText(String(column.subtitle || ''), {
      x: x + 0.24,
      y: 2.3,
      w: 2.4,
      h: 0.14,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: theme.muted,
      margin: 0,
    });
    slide.addText('核心问题', {
      x: x + 0.24,
      y: 2.68,
      w: 0.8,
      h: 0.14,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: accent,
      bold: true,
      margin: 0,
    });
    slide.addText(String(column.problem || column.body || ''), {
      x: x + 0.24,
      y: 2.88,
      w: 3.05,
      h: 0.9,
      fontFace: theme.bodyFont,
      fontSize: 9,
      color: theme.text,
      margin: 0,
      valign: 'top',
    });
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.24,
      y: 4.12,
      w: 3.05,
      h: 0.9,
      rectRadius: 0.04,
      fill: { color: 'FFF5F5' },
      line: { color: 'FEB2B2', transparency: 35, width: 1 },
    });
    slide.addText('业务影响', {
      x: x + 0.35,
      y: 4.25,
      w: 0.72,
      h: 0.14,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: 'C53030',
      bold: true,
      margin: 0,
    });
    slide.addText(asArray(column.impacts).map((item) => `• ${item}`).join('\n'), {
      x: x + 0.34,
      y: 4.46,
      w: 2.82,
      h: 0.42,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: 'C53030',
      margin: 0,
      valign: 'top',
    });
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.24,
      y: 5.26,
      w: 3.05,
      h: 0.74,
      rectRadius: 0.04,
      fill: { color: 'F7FAFC' },
      line: { color: theme.line, transparency: 35, width: 1, dash: 'dash' },
    });
    slide.addText('典型场景', {
      x: x + 0.35,
      y: 5.38,
      w: 0.72,
      h: 0.14,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: theme.muted,
      bold: true,
      margin: 0,
    });
    slide.addText(String(column.example || ''), {
      x: x + 0.34,
      y: 5.58,
      w: 2.82,
      h: 0.22,
      fontFace: theme.bodyFont,
      fontSize: 7.7,
      color: theme.muted,
      italic: true,
      margin: 0,
      valign: 'mid',
    });
  });
}

function renderArchitectureStack(slide, deck, theme) {
  addLightHeader(slide, deck.title || '产品方案设计：平台架构', deck.eyebrow || '', theme);
  const layers = asArray(deck.layers).slice(0, 3);
  layers.forEach((layer, index) => {
    const y = 1.64 + index * 1.48;
    const accent = layer.accent || getAccentForIndex(theme, index);
    addCard(slide, 0.65, y, 8.35, 1.15, theme, accent);
    slide.addText(String(layer.title || `Layer ${index + 1}`), {
      x: 0.9,
      y: y + 0.18,
      w: 3.0,
      h: 0.18,
      fontFace: theme.headerFont,
      fontSize: 11.5,
      color: theme.text,
      bold: true,
      margin: 0,
    });
    slide.addText(String(layer.subtitle || ''), {
      x: 6.3,
      y: y + 0.18,
      w: 2.2,
      h: 0.14,
      fontFace: theme.bodyFont,
      fontSize: 7.5,
      color: theme.muted,
      margin: 0,
      align: 'right',
    });
    asArray(layer.modules).slice(0, 5).forEach((moduleText, moduleIndex, moduleList) => {
      const moduleGap = 7.6 / Math.max(moduleList.length, 1);
      const moduleW = moduleGap - 0.16;
      const moduleX = 0.9 + moduleIndex * moduleGap;
      slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
        x: moduleX,
        y: y + 0.46,
        w: moduleW,
        h: 0.42,
        rectRadius: 0.04,
        fill: { color: 'F8FAFC' },
        line: { color: theme.line, transparency: 25, width: 1 },
      });
      slide.addText(String(moduleText), {
        x: moduleX + 0.06,
        y: y + 0.55,
        w: moduleW - 0.12,
        h: 0.16,
        fontFace: theme.bodyFont,
        fontSize: Math.min(7.8, moduleW * 6),
        color: theme.text,
        margin: 0,
        align: 'center',
        valign: 'mid',
        fit: 'shrink',
      });
    });
    if (index < layers.length - 1) {
      slide.addText('↕', {
        x: 4.2,
        y: y + 1.14,
        w: 0.3,
        h: 0.16,
        fontFace: theme.headerFont,
        fontSize: 11,
        color: 'BFD5EA',
        bold: true,
        margin: 0,
        align: 'center',
      });
    }
  });
  slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
    x: 9.25,
    y: 1.64,
    w: 2.95,
    h: 4.98,
    rectRadius: 0.08,
    fill: { color: 'F1F5F9' },
    line: { color: theme.line, transparency: 35, width: 1, dash: 'dash' },
  });
  slide.addText(String(deck.sidebarTitle || '关键服务支持'), {
    x: 9.55,
    y: 1.9,
    w: 2.2,
    h: 0.18,
    fontFace: theme.headerFont,
    fontSize: 11.2,
    color: theme.dominant,
    bold: true,
    margin: 0,
  });
  asArray(deck.services).slice(0, 5).forEach((service, index) => {
    const y = 2.3 + index * 0.54;
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x: 9.55,
      y,
      w: 2.3,
      h: 0.32,
      rectRadius: 0.03,
      fill: { color: 'FFFFFF' },
      line: { color: theme.line, transparency: 35, width: 1 },
    });
    slide.addText(String(service), {
      x: 9.7,
      y: y + 0.07,
      w: 2.0,
      h: 0.12,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: theme.text,
      margin: 0,
      valign: 'mid',
    });
  });
  if (deck.deployment) {
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x: 9.55,
      y: 5.55,
      w: 2.3,
      h: 0.48,
      rectRadius: 0.04,
      fill: { color: 'E6F6FC' },
      line: { color: theme.accent, transparency: 35, width: 1 },
    });
    slide.addText('DEPLOYMENT', {
      x: 9.7,
      y: 5.66,
      w: 2.0,
      h: 0.12,
      fontFace: theme.bodyFont,
      fontSize: 7.2,
      color: theme.support,
      bold: true,
      margin: 0,
      align: 'center',
    });
    slide.addText(String(deck.deployment), {
      x: 9.7,
      y: 5.84,
      w: 2.0,
      h: 0.1,
      fontFace: theme.bodyFont,
      fontSize: 7.4,
      color: theme.muted,
      margin: 0,
      align: 'center',
    });
  }
}

function renderFeatureGrid(slide, deck, theme, spec) {
  addLightHeader(slide, deck.title || '核心功能模块', deck.eyebrow || '', theme);
  const items = asArray(deck.items);
  const rect = getContentRect(spec, deck);
  const cells = calcEqualGrid(items.length, rect);
  items.forEach((item, index) => {
    const cell = cells[index];
    if (!cell) return;
    const pad = cell.w * 0.05;
    const innerW = cell.w - pad * 2;
    const accent = item.accent || getAccentForIndex(theme, index);
    addCard(slide, cell.x, cell.y, cell.w, cell.h, theme, accent);
    const titleH = cell.h * 0.14;
    const subtitleH = cell.h * 0.08;
    const featuresH = cell.h * 0.55;
    slide.addText(String(item.title || `Module ${index + 1}`), {
      x: cell.x + pad,
      y: cell.y + cell.h * 0.06,
      w: innerW * 0.6,
      h: titleH,
      fontFace: theme.headerFont,
      fontSize: Math.min(12, cell.w * 2.1),
      color: theme.text,
      bold: true,
      margin: 0,
      fit: 'shrink',
    });
    slide.addText(String(item.subtitle || ''), {
      x: cell.x + pad,
      y: cell.y + cell.h * 0.06 + titleH,
      w: innerW * 0.5,
      h: subtitleH,
      fontFace: theme.bodyFont,
      fontSize: Math.min(7.8, cell.w * 1.4),
      color: theme.muted,
      margin: 0,
    });
    const featuresText = asArray(item.features).map((feature) => `• ${feature}`).join('\n');
    slide.addText(featuresText, {
      x: cell.x + pad,
      y: cell.y + cell.h * 0.32,
      w: innerW,
      h: featuresH,
      fontFace: theme.bodyFont,
      fontSize: adaptiveFontSize(featuresText, innerW, featuresH, 8.3, 6.5),
      color: theme.text,
      margin: 0,
      valign: 'top',
      fit: 'shrink',
    });
    if (item.valueTag) {
      const pillW = Math.min(cell.w * 0.25, 1.3);
      addPill(slide, cell.x + cell.w - pad - pillW, cell.y + cell.h * 0.82, pillW, cell.h * 0.1, String(item.valueTag), 'F8FAFC', theme.muted, theme);
    }
  });
}

function renderSwimlane(slide, deck, theme) {
  addLightHeader(slide, deck.title || '产品原型/功能流程：一次碳足迹计算', deck.eyebrow || '', theme);
  const phases = asArray(deck.phases).slice(0, 4);
  const actors = asArray(deck.actors).slice(0, 3);
  const phaseStartX = 2.32;
  const phaseWidth = 2.38;
  const laneStartY = 2.05;
  const laneHeight = 1.26;

  phases.forEach((phase, index) => {
    const x = phaseStartX + index * phaseWidth;
    slide.addText(String(phase.step || index + 1), {
      x,
      y: 1.56,
      w: 0.18,
      h: 0.18,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: 'FFFFFF',
      bold: true,
      margin: 0,
      align: 'center',
      valign: 'mid',
      shape: {
        type: slide._pptx.shapes.OVAL,
        fill: { color: theme.dominant },
        line: { color: theme.dominant, transparency: 100 },
      },
    });
    slide.addText(String(phase.title || ''), {
      x: x + 0.28,
      y: 1.55,
      w: 1.8,
      h: 0.16,
      fontFace: theme.bodyFont,
      fontSize: 8.3,
      color: theme.muted,
      bold: true,
      margin: 0,
    });
  });

  actors.forEach((actor, actorIndex) => {
    const y = laneStartY + actorIndex * laneHeight;
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.65,
      y,
      w: 1.35,
      h: 0.94,
      rectRadius: 0.05,
      fill: { color: 'F8FAFC' },
      line: { color: theme.line, transparency: 35, width: 1 },
    });
    slide.addText(String(actor.name || `Actor ${actorIndex + 1}`), {
      x: 0.78,
      y: y + 0.24,
      w: 1.05,
      h: 0.16,
      fontFace: theme.headerFont,
      fontSize: 9.8,
      color: theme.text,
      bold: true,
      margin: 0,
      align: 'center',
    });
    slide.addText(String(actor.subtitle || ''), {
      x: 0.78,
      y: y + 0.46,
      w: 1.05,
      h: 0.12,
      fontFace: theme.bodyFont,
      fontSize: 7.1,
      color: theme.muted,
      margin: 0,
      align: 'center',
    });
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x: 2.08,
      y,
      w: 9.95,
      h: 0.94,
      rectRadius: 0.05,
      fill: { color: actorIndex === 1 ? 'F0F9FF' : 'FFFFFF' },
      line: { color: theme.line, transparency: 55, width: 1 },
    });
  });

  asArray(deck.nodes).forEach((node) => {
    const actorIndex = Math.max(0, Math.min(Number(node.actorIndex || 0), 2));
    const phaseIndex = Math.max(0, Math.min(Number(node.phaseIndex || 0), 3));
    const x = phaseStartX + phaseIndex * phaseWidth;
    const y = laneStartY + actorIndex * laneHeight + 0.08;
    const accent =
      node.type === 'user' ? theme.dominant : node.type === 'external' ? theme.muted : theme.accent;
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x,
      y,
      w: 2.0,
      h: 0.78,
      rectRadius: 0.04,
      fill: { color: node.type === 'system' ? 'F0F9FF' : 'FFFFFF' },
      line: { color: accent, width: 1.4 },
    });
    slide.addText(String(node.title || ''), {
      x: x + 0.12,
      y: y + 0.12,
      w: 1.52,
      h: 0.14,
      fontFace: theme.headerFont,
      fontSize: 8.8,
      color: theme.text,
      bold: true,
      margin: 0,
    });
    slide.addText(String(node.body || ''), {
      x: x + 0.12,
      y: y + 0.34,
      w: 1.72,
      h: 0.24,
      fontFace: theme.bodyFont,
      fontSize: 7.2,
      color: theme.muted,
      margin: 0,
      valign: 'mid',
    });
    if (node.status) {
      addPill(slide, x + 1.55, y - 0.08, 0.38, 0.18, String(node.status), 'E53E3E', 'FFFFFF', theme);
    }
  });

  slide.addText(String(deck.legend || ''), {
    x: 0.65,
    y: 6.45,
    w: 11.55,
    h: 0.16,
    fontFace: theme.bodyFont,
    fontSize: 7.6,
    color: theme.muted,
    margin: 0,
  });
}

function renderHubSpoke(slide, deck, theme) {
  addLightHeader(slide, deck.title || '平台生态', deck.eyebrow || '', theme);
  const spokes = asArray(deck.spokes || deck.items).slice(0, 5);

  slide.addShape(slide._pptx.shapes.OVAL, {
    x: 4.78,
    y: 2.15,
    w: 2.1,
    h: 2.1,
    fill: { color: theme.dominant },
    line: { color: theme.dominant, transparency: 100 },
    shadow: { type: 'outer', color: '000000', blur: 3, offset: 1, angle: 45, opacity: 0.1 },
  });
  if (deck.centerTag) {
    addPill(slide, 5.1, 1.86, 1.45, 0.2, String(deck.centerTag), 'E6F6FC', theme.support, theme);
  }
  slide.addText(String(deck.hubTitle || deck.centerTitle || '中心平台'), {
    x: 5.04,
    y: 2.62,
    w: 1.58,
    h: 0.42,
    fontFace: theme.headerFont,
    fontSize: 13,
    color: 'FFFFFF',
    bold: true,
    margin: 0,
    align: 'center',
    valign: 'mid',
  });
  if (deck.centerBody) {
    slide.addText(String(deck.centerBody), {
      x: 5.0,
      y: 3.08,
      w: 1.65,
      h: 0.58,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: 'D6E4FF',
      margin: 0,
      align: 'center',
      valign: 'mid',
    });
  }

  const positions = [
    { x: 0.82, y: 1.68, w: 2.4, h: 1.18, anchorX: 3.22, anchorY: 2.58, lineToX: 4.72, lineToY: 2.78 },
    { x: 9.95, y: 1.68, w: 2.4, h: 1.18, anchorX: 9.95, anchorY: 2.58, lineToX: 6.88, lineToY: 2.78 },
    { x: 0.82, y: 4.18, w: 2.4, h: 1.18, anchorX: 3.22, anchorY: 4.44, lineToX: 4.72, lineToY: 3.82 },
    { x: 9.95, y: 4.18, w: 2.4, h: 1.18, anchorX: 9.95, anchorY: 4.44, lineToX: 6.88, lineToY: 3.82 },
    { x: 4.47, y: 5.2, w: 2.72, h: 0.92, anchorX: 5.83, anchorY: 5.2, lineToX: 5.83, lineToY: 4.26 },
  ];

  spokes.forEach((spoke, index) => {
    const pos = positions[index];
    if (!pos) return;
    const accent = spoke.accent || getAccentForIndex(theme, index);
    slide.addShape(slide._pptx.shapes.LINE, {
      x: pos.lineToX,
      y: pos.lineToY,
      w: pos.anchorX - pos.lineToX,
      h: pos.anchorY - pos.lineToY,
      line: { color: theme.line, width: 1.2 },
    });
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x: pos.x,
      y: pos.y,
      w: pos.w,
      h: pos.h,
      rectRadius: 0.06,
      fill: { color: 'FFFFFF' },
      line: { color: theme.line, transparency: 35, width: 1 },
      shadow: { type: 'outer', color: '000000', blur: 2, offset: 1, angle: 45, opacity: 0.08 },
    });
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: pos.x,
      y: pos.y,
      w: pos.w,
      h: 0.06,
      fill: { color: accent },
      line: { color: accent, transparency: 100 },
    });
    slide.addText(String(spoke.title || `Role ${index + 1}`), {
      x: pos.x + 0.16,
      y: pos.y + 0.16,
      w: pos.w - 0.32,
      h: 0.2,
      fontFace: theme.headerFont,
      fontSize: 10.5,
      color: theme.text,
      bold: true,
      margin: 0,
      align: 'center',
    });
    slide.addText(String(spoke.body || spoke.text || ''), {
      x: pos.x + 0.16,
      y: pos.y + 0.46,
      w: pos.w - 0.32,
      h: index === 4 ? 0.22 : 0.48,
      fontFace: theme.bodyFont,
      fontSize: index === 4 ? 7.4 : 7.8,
      color: theme.muted,
      margin: 0,
      align: 'center',
      valign: 'mid',
    });
  });
}

function renderMetricsDashboard(slide, deck, theme, spec) {
  addLightHeader(slide, deck.title || '数据与指标 (Success Metrics)', deck.eyebrow || '', theme);
  const canvas = getCanvas(spec);
  const margin = getMargin(canvas);
  const headerH = getHeaderHeight(canvas);
  const contentW = canvas.w - margin.left - margin.right;
  const contentH = canvas.h - headerH - margin.bottom;
  const contentTop = headerH;

  const metrics = asArray(deck.metrics || deck.items);
  const hasChart = !!(asArray(deck.series).length || deck.chartTitle);
  const metricsW = hasChart ? contentW * 0.48 : contentW;
  const metricsRect = { x: margin.left, y: contentTop, w: metricsW, h: contentH };
  const metricCells = calcEqualGrid(metrics.length, metricsRect);

  metrics.forEach((metric, index) => {
    const cell = metricCells[index];
    if (!cell) return;
    const pad = cell.w * 0.07;
    const innerW = cell.w - pad * 2;
    const accent = metric.accent || getAccentForIndex(theme, index);
    addCard(slide, cell.x, cell.y, cell.w, cell.h, theme, accent);
    slide.addText(String(metric.title || metric.label || `Metric ${index + 1}`), {
      x: cell.x + pad,
      y: cell.y + cell.h * 0.08,
      w: innerW,
      h: cell.h * 0.14,
      fontFace: theme.headerFont,
      fontSize: Math.min(10.5, cell.w * 1.8),
      color: theme.text,
      bold: true,
      margin: 0,
      fit: 'shrink',
    });
    const rows = asArray(metric.rows);
    if (rows.length) {
      const rowH = (cell.h * 0.7) / Math.max(rows.length, 1);
      rows.forEach((rowItem, rowIndex) => {
        const rowY = cell.y + cell.h * 0.28 + rowIndex * rowH;
        slide.addText(String(rowItem.label || ''), {
          x: cell.x + pad,
          y: rowY,
          w: innerW * 0.6,
          h: rowH * 0.5,
          fontFace: theme.bodyFont,
          fontSize: Math.min(7.2, rowH * 5),
          color: theme.muted,
          margin: 0,
        });
        slide.addText(String(rowItem.value || ''), {
          x: cell.x + pad + innerW * 0.55,
          y: rowY,
          w: innerW * 0.4,
          h: rowH * 0.5,
          fontFace: theme.headerFont,
          fontSize: Math.min(10, rowH * 7),
          color: theme.text,
          bold: true,
          margin: 0,
          align: 'right',
        });
        if (rowItem.note) {
          slide.addText(String(rowItem.note), {
            x: cell.x + pad + innerW * 0.6,
            y: rowY + rowH * 0.5,
            w: innerW * 0.35,
            h: rowH * 0.4,
            fontFace: theme.bodyFont,
            fontSize: Math.min(6.8, rowH * 4.5),
            color: theme.muted,
            margin: 0,
            align: 'right',
          });
        }
      });
    } else if (metric.value) {
      slide.addText(String(metric.value || '--'), {
        x: cell.x + pad,
        y: cell.y + cell.h * 0.35,
        w: innerW,
        h: cell.h * 0.25,
        fontFace: theme.headerFont,
        fontSize: Math.min(20, cell.h * 2.5),
        color: accent,
        bold: true,
        margin: 0,
        align: 'center',
        fit: 'shrink',
      });
      slide.addText(String(metric.note || ''), {
        x: cell.x + pad,
        y: cell.y + cell.h * 0.65,
        w: innerW,
        h: cell.h * 0.2,
        fontFace: theme.bodyFont,
        fontSize: Math.min(8, cell.h * 1.2),
        color: theme.muted,
        margin: 0,
        align: 'center',
        fit: 'shrink',
      });
    }
  });

  if (hasChart) {
    const chartX = margin.left + metricsW + contentW * 0.02;
    const chartW = contentW - metricsW - contentW * 0.02;
    const chartH = contentH * 0.55;
    addCard(slide, chartX, contentTop, chartW, chartH, theme, theme.support);
    slide.addText(String(deck.chartTitle || '核心业务增长趋势'), {
      x: chartX + chartW * 0.04,
      y: contentTop + chartH * 0.08,
      w: chartW * 0.5,
      h: chartH * 0.1,
      fontFace: theme.headerFont,
      fontSize: 10.5,
      color: theme.text,
      bold: true,
      margin: 0,
    });
    const chartSeries = asArray(deck.series).slice(0, 2);
    const quarterLabels = asArray(deck.quarters).slice(0, 4);
    const barAreaTop = contentTop + chartH * 0.25;
    const barAreaH = chartH * 0.6;
    const barColW = (chartW * 0.8) / Math.max(quarterLabels.length, 1);
    quarterLabels.forEach((label, index) => {
      slide.addText(String(label), {
        x: chartX + chartW * 0.15 + index * barColW,
        y: contentTop + chartH * 0.85,
        w: barColW * 0.7,
        h: chartH * 0.08,
        fontFace: theme.bodyFont,
        fontSize: 7,
        color: theme.muted,
        margin: 0,
        align: 'center',
      });
    });
    chartSeries.forEach((series, seriesIndex) => {
      const seriesY = barAreaTop + seriesIndex * (barAreaH / Math.max(chartSeries.length, 1));
      slide.addText(String(series.name || `Series ${seriesIndex + 1}`), {
        x: chartX + chartW * 0.04,
        y: seriesY,
        w: chartW * 0.12,
        h: barAreaH * 0.1,
        fontFace: theme.bodyFont,
        fontSize: 7.4,
        color: series.color || (seriesIndex === 0 ? theme.accent : theme.dominant),
        bold: true,
        margin: 0,
      });
      const values = asArray(series.values).slice(0, 4).map((v) => Number(v || 0));
      const maxValue = Math.max(...values, 1);
      values.forEach((value, index) => {
        const baseX = chartX + chartW * 0.15 + index * barColW;
        const barMaxW = barColW * 0.65;
        const barW = barMaxW * (value / maxValue);
        slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
          x: baseX,
          y: seriesY + barAreaH * 0.12,
          w: barMaxW,
          h: barAreaH * 0.06,
          rectRadius: 0.02,
          fill: { color: 'E2E8F0' },
          line: { color: 'E2E8F0', transparency: 100 },
        });
        slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
          x: baseX,
          y: seriesY + barAreaH * 0.12,
          w: Math.max(barW, 0.08),
          h: barAreaH * 0.06,
          rectRadius: 0.02,
          fill: { color: series.color || (seriesIndex === 0 ? theme.accent : theme.dominant) },
          line: { color: series.color || (seriesIndex === 0 ? theme.accent : theme.dominant), transparency: 100 },
        });
        slide.addText(String(value), {
          x: baseX,
          y: seriesY + barAreaH * 0.2,
          w: barMaxW,
          h: barAreaH * 0.08,
          fontFace: theme.bodyFont,
          fontSize: 6.8,
          color: theme.muted,
          margin: 0,
          align: 'center',
        });
      });
    });

    if (deck.northStarLabel || deck.northStarValue) {
      const nsY = contentTop + chartH + contentH * 0.03;
      const nsH = contentH - chartH - contentH * 0.03;
      slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
        x: chartX,
        y: nsY,
        w: chartW,
        h: nsH,
        rectRadius: 0.08,
        fill: { color: theme.dominant },
        line: { color: theme.dominant, transparency: 100 },
      });
      slide.addText(String(deck.northStarLabel || 'North Star Metric'), {
        x: chartX + chartW * 0.05,
        y: nsY + nsH * 0.12,
        w: chartW * 0.35,
        h: nsH * 0.15,
        fontFace: theme.bodyFont,
        fontSize: 8,
        color: '90CDF4',
        bold: true,
        margin: 0,
      });
      slide.addText(String(deck.northStarValue || ''), {
        x: chartX + chartW * 0.05,
        y: nsY + nsH * 0.35,
        w: chartW * 0.4,
        h: nsH * 0.35,
        fontFace: theme.headerFont,
        fontSize: Math.min(24, nsH * 3),
        color: 'FFFFFF',
        bold: true,
        margin: 0,
        fit: 'shrink',
      });
      slide.addText(String(deck.northStarBody || ''), {
        x: chartX + chartW * 0.45,
        y: nsY + nsH * 0.2,
        w: chartW * 0.48,
        h: nsH * 0.6,
        fontFace: theme.bodyFont,
        fontSize: 8.2,
        color: 'D6E4FF',
        margin: 0,
        valign: 'mid',
        fit: 'shrink',
      });
    }
  }
}

function renderRoadmapPhases(slide, deck, theme) {
  addLightHeader(slide, deck.title || '产品迭代路线图：MVP → 增强版 → 生态平台', deck.eyebrow || '', theme);
  const phases = asArray(deck.items).slice(0, 3);
  phases.forEach((phase, index) => {
    const x = 0.65 + index * 4.0;
    const accent = phase.accent || getStepColor(theme, index);
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x,
      y: 1.6,
      w: 3.72,
      h: 3.95,
      rectRadius: 0.08,
      fill: { color: 'FFFFFF' },
      line: { color: theme.line, transparency: 35, width: 1 },
    });
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x,
      y: 1.6,
      w: 3.72,
      h: 0.82,
      rectRadius: 0.08,
      fill: { color: accent },
      line: { color: accent, transparency: 100 },
    });
    addPill(slide, x + 0.18, 1.76, 0.9, 0.2, String(phase.time || ''), 'FFFFFF', accent, theme);
    slide.addText(String(phase.title || `Phase ${index + 1}`), {
      x: x + 0.18,
      y: 2.04,
      w: 2.85,
      h: 0.2,
      fontFace: theme.headerFont,
      fontSize: 11.8,
      color: 'FFFFFF',
      bold: true,
      margin: 0,
    });
    slide.addText(String(phase.subtitle || ''), {
      x: x + 0.18,
      y: 2.24,
      w: 2.9,
      h: 0.12,
      fontFace: theme.bodyFont,
      fontSize: 7.4,
      color: 'E6F6FC',
      margin: 0,
    });
    slide.addText(asArray(phase.features).map((feature) => `• ${feature}`).join('\n'), {
      x: x + 0.22,
      y: 2.68,
      w: 3.05,
      h: 1.75,
      fontFace: theme.bodyFont,
      fontSize: 8.2,
      color: theme.text,
      margin: 0,
      valign: 'top',
    });
    slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.18,
      y: 4.72,
      w: 3.1,
      h: 0.56,
      rectRadius: 0.04,
      fill: { color: 'F8FAFC' },
      line: { color: theme.line, transparency: 35, width: 1 },
    });
    slide.addText('Key Milestone', {
      x: x + 0.32,
      y: 4.84,
      w: 0.8,
      h: 0.12,
      fontFace: theme.bodyFont,
      fontSize: 7.2,
      color: theme.muted,
      bold: true,
      margin: 0,
    });
    slide.addText(String(phase.milestone || ''), {
      x: x + 0.32,
      y: 5.02,
      w: 2.7,
      h: 0.12,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: theme.text,
      margin: 0,
    });
  });

  slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.65,
    y: 5.8,
    w: 11.55,
    h: 0.62,
    rectRadius: 0.05,
    fill: { color: 'FFFFFF' },
    line: { color: theme.line, transparency: 35, width: 1 },
  });
  slide.addText(String(deck.risksTitle || 'Risks'), {
    x: 0.95,
    y: 5.96,
    w: 0.6,
    h: 0.12,
    fontFace: theme.headerFont,
    fontSize: 8.4,
    color: 'C53030',
    bold: true,
    margin: 0,
  });
  slide.addText(asArray(deck.risks).map((item) => `• ${item}`).join('  '), {
    x: 1.52,
    y: 5.96,
    w: 4.4,
    h: 0.12,
    fontFace: theme.bodyFont,
    fontSize: 7.2,
    color: theme.text,
    margin: 0,
  });
  slide.addText(String(deck.dependenciesTitle || 'Dependencies'), {
    x: 6.4,
    y: 5.96,
    w: 0.9,
    h: 0.12,
    fontFace: theme.headerFont,
    fontSize: 8.4,
    color: theme.muted,
    bold: true,
    margin: 0,
  });
  slide.addText(asArray(deck.dependencies).map((item) => `• ${item}`).join('  '), {
    x: 7.3,
    y: 5.96,
    w: 4.2,
    h: 0.12,
    fontFace: theme.bodyFont,
    fontSize: 7.2,
    color: theme.text,
    margin: 0,
  });
}

function renderStrengthsSidebar(slide, deck, theme) {
  addLightHeader(slide, deck.title || '我的优势', deck.eyebrow || '', theme);
  asArray(deck.items).slice(0, 4).forEach((item, index) => {
    const col = index % 2;
    const row = Math.floor(index / 2);
    const x = 0.65 + col * 3.85;
    const y = 1.62 + row * 2.2;
    const accent = item.accent || getAccentForIndex(theme, index);
    addCard(slide, x, y, 3.45, 1.92, theme, accent);
    slide.addText(String(item.title || `Strength ${index + 1}`), {
      x: x + 0.2,
      y: y + 0.2,
      w: 1.75,
      h: 0.18,
      fontFace: theme.headerFont,
      fontSize: 10.8,
      color: theme.text,
      bold: true,
      margin: 0,
    });
    slide.addText(String(item.subtitle || ''), {
      x: x + 0.2,
      y: y + 0.42,
      w: 1.75,
      h: 0.12,
      fontFace: theme.bodyFont,
      fontSize: 7.2,
      color: theme.muted,
      margin: 0,
    });
    slide.addText(asArray(item.bullets).map((bullet) => `• ${bullet}`).join('\n'), {
      x: x + 0.22,
      y: y + 0.68,
      w: 2.92,
      h: 0.76,
      fontFace: theme.bodyFont,
      fontSize: 7.8,
      color: theme.text,
      margin: 0,
      valign: 'top',
    });
    renderTagRow(slide, item.tags || [], x + 0.22, y + 1.54, 2.8, theme, 'F8FAFC', theme.muted);
  });

  slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
    x: 8.45,
    y: 1.62,
    w: 3.75,
    h: 4.78,
    rectRadius: 0.08,
    fill: { color: theme.dominant },
    line: { color: theme.dominant, transparency: 100 },
  });
  slide.addText(String(deck.sidebarTitle || 'Certifications & Tool Stack'), {
    x: 8.8,
    y: 1.92,
    w: 2.7,
    h: 0.18,
    fontFace: theme.headerFont,
    fontSize: 11,
    color: '90CDF4',
    bold: true,
    margin: 0,
  });
  slide.addText(asArray(deck.certifications).map((item) => `• ${item}`).join('\n'), {
    x: 8.8,
    y: 2.3,
    w: 2.8,
    h: 1.0,
    fontFace: theme.bodyFont,
    fontSize: 8.2,
    color: 'FFFFFF',
    margin: 0,
    valign: 'top',
  });
  slide.addText('Tool Stack', {
    x: 8.8,
    y: 3.58,
    w: 1.1,
    h: 0.16,
    fontFace: theme.headerFont,
    fontSize: 9.6,
    color: '90CDF4',
    bold: true,
    margin: 0,
  });
  renderTagRow(slide, deck.tools || [], 8.82, 3.88, 2.8, theme, '1E3A5F', 'E2E8F0');
  slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
    x: 8.82,
    y: 5.22,
    w: 2.95,
    h: 0.78,
    rectRadius: 0.04,
    fill: { color: '0F2F57' },
    line: { color: '0F2F57', transparency: 100 },
  });
  slide.addText(String(deck.quote || ''), {
    x: 9.0,
    y: 5.38,
    w: 2.58,
    h: 0.42,
    fontFace: theme.bodyFont,
    fontSize: 8,
    color: 'E2E8F0',
    italic: true,
    margin: 0,
    valign: 'mid',
  });
}

function renderValueClosing(slide, deck, theme) {
  slide.background = { color: theme.background };
  slide.addShape(slide._pptx.shapes.OVAL, {
    x: -0.45,
    y: 4.9,
    w: 2.6,
    h: 2.6,
    fill: { color: theme.support, transparency: 92 },
    line: { color: theme.support, transparency: 100 },
  });
  slide.addShape(slide._pptx.shapes.OVAL, {
    x: 10.35,
    y: 0.2,
    w: 2.2,
    h: 2.2,
    fill: { color: theme.accent, transparency: 94 },
    line: { color: theme.accent, transparency: 100 },
  });
  slide.addShape(slide._pptx.shapes.RECTANGLE, {
    x: 0,
    y: 0.78,
    w: 13.33,
    h: 0.02,
    fill: { color: theme.line },
    line: { color: theme.line, transparency: 100 },
  });

  addPill(
    slide,
    4.95,
    1.28,
    1.9,
    0.26,
    String(deck.valueLabel || deck.eyebrow || 'VALUE PROPOSITION'),
    'E6F6FC',
    theme.support,
    theme
  );
  slide.addText(String(deck.title || '我能为团队带来的价值'), {
    x: 2.0,
    y: 1.78,
    w: 9.3,
    h: 0.5,
    fontFace: theme.headerFont,
    fontSize: 24,
    color: theme.text,
    bold: true,
    margin: 0,
    align: 'center',
  });
  slide.addText(String(deck.subtitle || ''), {
    x: 2.25,
    y: 2.42,
    w: 8.8,
    h: 0.36,
    fontFace: theme.bodyFont,
    fontSize: 11.5,
    color: theme.muted,
    margin: 0,
    align: 'center',
    valign: 'mid',
  });
  if (deck.statement) {
    slide.addText(String(deck.statement), {
      x: 1.95,
      y: 2.95,
      w: 9.45,
      h: 0.32,
      fontFace: theme.bodyFont,
      fontSize: 10.2,
      color: theme.text,
      italic: true,
      margin: 0,
      align: 'center',
      valign: 'mid',
    });
  }

  const items = asArray(deck.items).slice(0, 3);
  items.forEach((item, index) => {
    const x = 1.1 + index * 3.98;
    const accent = item.accent || getAccentForIndex(theme, index);
    addCard(slide, x, 3.55, 3.5, 1.7, theme, accent);
    slide.addText(String(item.title || `Value ${index + 1}`), {
      x: x + 0.2,
      y: 3.84,
      w: 3.0,
      h: 0.2,
      fontFace: theme.headerFont,
      fontSize: 11.5,
      color: theme.text,
      bold: true,
      margin: 0,
      align: 'center',
    });
    slide.addText(String(item.body || item.message || ''), {
      x: x + 0.26,
      y: 4.2,
      w: 2.9,
      h: 0.52,
      fontFace: theme.bodyFont,
      fontSize: 8.8,
      color: theme.muted,
      margin: 0,
      align: 'center',
      valign: 'mid',
    });
    if (item.tag) {
      addPill(slide, x + 1.1, 4.82, 1.3, 0.2, String(item.tag), 'F8FAFC', theme.muted, theme);
    }
  });

  slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
    x: 3.35,
    y: 5.72,
    w: 6.6,
    h: 0.48,
    rectRadius: 0.05,
    fill: { color: 'FFFFFF' },
    line: { color: theme.line, transparency: 30, width: 1 },
  });
  slide.addText(String(deck.footerNote || '期待把行业理解、产品抽象与数据驱动执行真正带进团队。'), {
    x: 3.58,
    y: 5.87,
    w: 6.15,
    h: 0.14,
    fontFace: theme.bodyFont,
    fontSize: 8.2,
    color: theme.muted,
    margin: 0,
    align: 'center',
  });
}

function renderComparison(slide, deck, theme, spec) {
  addLightHeader(slide, deck.title || 'Comparison', deck.eyebrow || '', theme);
  const cols = asArray(deck.columns);
  const rect = getContentRect(spec, deck);
  const columns = calcEqualColumns(cols.length, rect);
  cols.forEach((col, index) => {
    const cell = columns[index];
    if (!cell) return;
    const pad = cell.w * 0.04;
    const innerW = cell.w - pad * 2;
    const accentColor = index === 0 ? theme.support : index === 1 ? theme.accent : theme.dominant;
    addCard(slide, cell.x, cell.y, cell.w, cell.h, theme, accentColor);
    const titleH = cell.h * 0.1;
    const bodyH = cell.h * 0.8;
    slide.addText(String(col.title || `Column ${index + 1}`), {
      x: cell.x + pad,
      y: cell.y + cell.h * 0.06,
      w: innerW,
      h: titleH,
      fontFace: theme.headerFont,
      fontSize: Math.min(15, cell.w * 2.5),
      color: theme.text,
      bold: true,
      margin: 0,
      align: 'center',
      fit: 'shrink',
    });
    const lines = asArray(col.items || col.bullets).length ? asArray(col.items || col.bullets).map(String).join('\n') : String(col.body || '');
    slide.addText(lines, {
      x: cell.x + pad,
      y: cell.y + cell.h * 0.2,
      w: innerW,
      h: bodyH,
      fontFace: theme.bodyFont,
      fontSize: adaptiveFontSize(lines, innerW, bodyH, 11.5, 8),
      color: theme.muted,
      margin: 0,
      valign: 'top',
      align: 'left',
      fit: 'shrink',
    });
  });
}

function renderTimeline(slide, deck, theme, spec) {
  addLightHeader(slide, deck.title || 'Timeline', deck.eyebrow || '', theme);
  const items = asArray(deck.items);
  const canvas = getCanvas(spec);
  const margin = getMargin(canvas);
  const headerH = getHeaderHeight(canvas);
  const contentW = canvas.w - margin.left - margin.right;
  const contentTop = headerH;

  const laneY = contentTop + (canvas.h - contentTop - margin.bottom) * 0.42;
  const dotSize = Math.min(0.34, contentW / items.length * 0.14);
  const cardAreaTop = contentTop + (canvas.h - contentTop - margin.bottom) * 0.02;
  const cardAreaH = (laneY - cardAreaTop) * 0.9;

  const cols = calcEqualColumns(items.length, { x: margin.left, y: cardAreaTop, w: contentW, h: cardAreaH });

  items.forEach((item, index) => {
    const col = cols[index];
    if (!col) return;
    const centerX = col.x + col.w / 2;

    slide.addShape(slide._pptx.shapes.OVAL, {
      x: centerX - dotSize / 2,
      y: laneY,
      w: dotSize,
      h: dotSize,
      fill: { color: getStepColor(theme, index) },
      line: { color: getStepColor(theme, index), transparency: 100 },
    });

    if (index < items.length - 1) {
      const nextCol = cols[index + 1];
      const nextCenterX = nextCol.x + nextCol.w / 2;
      slide.addShape(slide._pptx.shapes.RECTANGLE, {
        x: centerX + dotSize / 2,
        y: laneY + dotSize * 0.4,
        w: nextCenterX - centerX - dotSize,
        h: dotSize * 0.15,
        fill: { color: theme.line },
        line: { color: theme.line, transparency: 100 },
      });
    }

    const pad = col.w * 0.04;
    const innerW = col.w - pad * 2;
    addCard(slide, col.x, col.y, col.w, col.h, theme, getAccentForIndex(theme, index));
    const titleH = col.h * 0.25;
    const bodyH = col.h * 0.6;
    slide.addText(String(item.title || `Stage ${index + 1}`), {
      x: col.x + pad,
      y: col.y + col.h * 0.08,
      w: innerW,
      h: titleH,
      fontFace: theme.headerFont,
      fontSize: adaptiveFontSize(item.title || '', innerW, titleH, 13, 9),
      color: theme.text,
      bold: true,
      margin: 0,
      align: 'center',
      fit: 'shrink',
    });
    const bodyText = String(item.body || item.note || '');
    slide.addText(bodyText, {
      x: col.x + pad,
      y: col.y + col.h * 0.35,
      w: innerW,
      h: bodyH,
      fontFace: theme.bodyFont,
      fontSize: adaptiveFontSize(bodyText, innerW, bodyH, 9.8, 7),
      color: theme.muted,
      margin: 0,
      align: 'center',
      valign: 'mid',
      fit: 'shrink',
    });
  });
}

function renderQuote(slide, deck, theme) {
  slide.background = { color: theme.background };
  slide.addShape(slide._pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.95,
    y: 1.15,
    w: 11.0,
    h: 4.75,
    rectRadius: 0.08,
    fill: { color: theme.surface },
    line: { color: theme.line, transparency: 55, width: 1 },
    shadow: { type: 'outer', color: '000000', blur: 2, offset: 1, angle: 45, opacity: 0.08 },
  });
  slide.addText('“', {
    x: 1.45,
    y: 1.65,
    w: 0.8,
    h: 0.8,
    fontFace: theme.headerFont,
    fontSize: 42,
    color: theme.support,
    bold: true,
    margin: 0,
  });
  slide.addText(String(deck.quote || deck.body || 'Add a strong quote or single key statement here.'), {
    x: 2.2,
    y: 1.92,
    w: 8.8,
    h: 1.75,
    fontFace: theme.headerFont,
    fontSize: 20,
    color: theme.text,
    bold: true,
    margin: 0,
    align: 'center',
    valign: 'mid',
  });
  slide.addText(String(deck.attribution || deck.subtitle || ''), {
    x: 2.2,
    y: 4.18,
    w: 8.8,
    h: 0.28,
    fontFace: theme.bodyFont,
    fontSize: 11,
    color: theme.muted,
    italic: true,
    margin: 0,
    align: 'center',
  });
}

function renderClosing(slide, deck, theme) {
  slide.background = { color: theme.dominant };
  const motif = theme.motif || '';
  if (motif === 'left-accent-rail') {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 0.35,
      h: 7.5,
      fill: { color: theme.borderAccent || theme.accent },
      line: { color: theme.borderAccent || theme.accent, transparency: 100 },
    });
  } else if (motif === 'header-bar') {
    slide.addShape(slide._pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 13.33,
      h: 0.08,
      fill: { color: theme.borderAccent || theme.accent },
      line: { color: theme.borderAccent || theme.accent, transparency: 100 },
    });
  } else {
    slide.addShape(slide._pptx.shapes.OVAL, {
      x: 9.8,
      y: 4.6,
      w: 3.4,
      h: 3.4,
      fill: { color: theme.support, transparency: 76 },
      line: { color: theme.support, transparency: 100 },
    });
  }
  slide.addText(deck.title || 'Closing', {
    x: 0.9,
    y: 1.25,
    w: 8.0,
    h: 0.7,
    fontFace: theme.headerFont,
    fontSize: 25,
    color: 'FFFFFF',
    bold: true,
    margin: 0,
  });
  if (deck.subtitle) {
    slide.addText(deck.subtitle, {
      x: 0.9,
      y: 2.0,
      w: 6.8,
      h: 0.35,
      fontFace: theme.bodyFont,
      fontSize: 12.5,
      color: 'D6E4FF',
      margin: 0,
    });
  }
  if (deck.message) {
    slide.addText(deck.message, {
      x: 0.9,
      y: 2.45,
      w: 8.0,
      h: 0.45,
      fontFace: theme.bodyFont,
      fontSize: 11.5,
      color: 'E2E8F0',
      margin: 0,
      fit: 'shrink',
    });
  }
  const runs = [];
  const bulletItems = asArray(deck.bullets);
  bulletItems.forEach((item, index) => {
    runs.push({ text: String(item), options: { bullet: true, breakLine: index < bulletItems.length - 1, color: 'F8FAFC', fontSize: 14 } });
  });
  const bulletsY = deck.message ? 3.1 : 2.75;
  slide.addText(runs.length ? runs : [{ text: 'Add the final takeaway here', options: { bullet: true, color: 'F8FAFC', fontSize: 14 } }], {
    x: 0.95,
    y: bulletsY,
    w: 6.8,
    h: 2.6,
    fontFace: theme.bodyFont,
    margin: 0,
    valign: 'top',
  });
  if (deck.footer) {
    slide.addText(deck.footer, {
      x: 0.9,
      y: 6.5,
      w: 11.0,
      h: 0.35,
      fontFace: theme.bodyFont,
      fontSize: 9.5,
      color: theme.accent,
      margin: 0,
      align: 'center',
      fit: 'shrink',
    });
  }
}

function renderSourcesSlide(slide, sources, theme) {
  addLightHeader(slide, 'Sources', 'APPENDIX', theme);
  slide.addText(
    sources.map((entry, index) => `${index + 1}. ${entry.label}${entry.url ? ` - ${entry.url}` : ''}`).join('\n'),
    {
      x: 0.85,
      y: 1.6,
      w: 11.2,
      h: 5.5,
      fontFace: theme.bodyFont,
      fontSize: 10.5,
      color: theme.text,
      margin: 0,
      valign: 'top',
    }
  );
}

function renderExplicitSourcesSlide(slide, rawSlide, spec) {
  const entries = asArray(rawSlide.sources)
    .concat(asArray(rawSlide.items))
    .map(normalizeSourceEntry)
    .filter(Boolean);
  renderSourcesSlide(slide, entries.length ? entries : collectUniqueSources(spec.slides), spec.theme);
}

async function renderSlide(pptx, slide, rawSlide, spec, index, total) {
  const normalizedSlide = normalizeSlideForRenderer(rawSlide);
  slide._pptx = pptx;
  slide._mlca = normalizedSlide;
  const theme = spec.theme;
  const kind = String(normalizedSlide.kind || 'bullets').toLowerCase();
  const assetMap = spec._assetMap || new Map();
  const pageChrome = normalizePageChrome(normalizedSlide.pageChrome, kind);

  await renderSlideImages(slide, normalizedSlide, assetMap, theme, spec, 'background');
  await renderSlideElements(slide, normalizedSlide, assetMap, theme, spec, 'background');

  if (kind === 'free-layout') {
    renderFreeLayout(slide, normalizedSlide, theme);
  } else if (kind === 'cover') renderCover(slide, normalizedSlide, theme);
  else if (kind === 'agenda') renderAgenda(slide, normalizedSlide, theme);
  else if (kind === 'section') renderSection(slide, normalizedSlide, theme);
  else if (kind === 'career-journey') renderCareerJourney(slide, normalizedSlide, theme);
  else if (kind === 'process-flow') renderProcessFlow(slide, normalizedSlide, theme, spec);
  else if (kind === 'capability-matrix') renderCapabilityMatrix(slide, normalizedSlide, theme);
  else if (kind === 'market-driver-stack') renderMarketDriverStack(slide, normalizedSlide, theme);
  else if (kind === 'persona-cards') renderPersonaCards(slide, normalizedSlide, theme);
  else if (kind === 'problem-cards') renderProblemCards(slide, normalizedSlide, theme);
  else if (kind === 'architecture-stack') renderArchitectureStack(slide, normalizedSlide, theme);
  else if (kind === 'feature-grid') renderFeatureGrid(slide, normalizedSlide, theme, spec);
  else if (kind === 'swimlane') renderSwimlane(slide, normalizedSlide, theme);
  else if (kind === 'hub-spoke') renderHubSpoke(slide, normalizedSlide, theme);
  else if (kind === 'metrics-dashboard') renderMetricsDashboard(slide, normalizedSlide, theme, spec);
  else if (kind === 'roadmap-phases') renderRoadmapPhases(slide, normalizedSlide, theme);
  else if (kind === 'strengths-sidebar') renderStrengthsSidebar(slide, normalizedSlide, theme);
  else if (kind === 'kpi-dashboard') renderMetricsDashboard(slide, normalizedSlide, theme, spec);
  else if (kind === 'layered-architecture') renderArchitectureStack(slide, normalizedSlide, theme);
  else if (kind === 'timeline-dots') renderTimeline(slide, normalizedSlide, theme, spec);
  else if (kind === 'capability-radar') renderCapabilityMatrix(slide, normalizedSlide, theme);
  else if (kind === 'value-closing') renderValueClosing(slide, normalizedSlide, theme);
  else if (kind === 'two-column') renderTwoColumn(slide, normalizedSlide, theme, spec);
  else if (kind === 'stats') renderStats(slide, normalizedSlide, theme, spec);
  else if (kind === 'grid') renderGrid(slide, normalizedSlide, theme, spec);
  else if (kind === 'comparison') renderComparison(slide, normalizedSlide, theme, spec);
  else if (kind === 'timeline') renderTimeline(slide, normalizedSlide, theme, spec);
  else if (kind === 'quote') renderQuote(slide, normalizedSlide, theme);
  else if (kind === 'closing') renderClosing(slide, normalizedSlide, theme);
  else if (kind === 'sources') renderExplicitSourcesSlide(slide, normalizedSlide, spec);
  else renderBullets(slide, normalizedSlide, theme, spec);

  await renderSlideImages(slide, normalizedSlide, assetMap, theme, spec, 'foreground');
  await renderSlideElements(slide, normalizedSlide, assetMap, theme, spec, 'foreground');

  if (pageChrome.enabled) {
    const darkKinds = new Set(['cover', 'section', 'closing']);
    addPageChrome(slide, spec, theme, index, total, pageChrome.dark || darkKinds.has(kind));
  }
  delete slide._pptx;
  delete slide._mlca;
}

function loadDeckSpec() {
  const externalJsonPath = process.argv[2] ? path.resolve(process.argv[2]) : '';
  if (externalJsonPath && fs.existsSync(externalJsonPath)) {
    return normalizeDeckSpec(JSON.parse(fs.readFileSync(externalJsonPath, 'utf8')));
  }
  if (process.env.MLCA_DECK_SPEC_JSON) {
    return normalizeDeckSpec(JSON.parse(process.env.MLCA_DECK_SPEC_JSON));
  }
  if (fs.existsSync(fallbackExampleSpecPath)) {
    return normalizeDeckSpec(JSON.parse(fs.readFileSync(fallbackExampleSpecPath, 'utf8')));
  }
  return normalizeDeckSpec(embeddedDeckSpec);
}

function collectSlideAudit(pptx, normalizedSlides) {
  const audit = { slides: [] };
  const slidesArr = pptx.slides;
  if (!slidesArr || !slidesArr.length) return audit;
  const layout = pptx.layout;
  const slideW = layout === 'LAYOUT_WIDE' ? 13.33 : 10;
  const slideH = 7.5;

  for (let si = 0; si < slidesArr.length; si++) {
    const slideObj = slidesArr[si];
    const shapes = slideObj._slideObjects || [];
    const slideAudit = {
      slideIndex: si,
      kind: (normalizedSlides[si] && normalizedSlides[si].kind) || 'unknown',
      title: (normalizedSlides[si] && normalizedSlides[si].title) || '',
      width: slideW,
      height: slideH,
      elements: [],
    };

    for (const shape of shapes) {
      const opts = shape.options || shape;
      const el = {
        type: shape.shape || 'unknown',
        x: opts.x ?? 0,
        y: opts.y ?? 0,
        w: opts.w ?? opts.width ?? 0,
        h: opts.h ?? opts.height ?? 0,
      };

      // Extract text content
      let textContent = '';
      if (Array.isArray(shape.text)) {
        textContent = shape.text.map(t => {
          if (typeof t === 'string') return t;
          return t.text || '';
        }).join(' ');
      } else if (typeof shape.text === 'string') {
        textContent = shape.text;
      }
      if (textContent) {
        el.textLength = textContent.length;
        el.textPreview = textContent.substring(0, 120);
      }

      // Extract font info from first text frame
      if (Array.isArray(shape.text) && shape.text[0] && shape.text[0].options) {
        const to = shape.text[0].options;
        if (to.fontSize) el.fontSize = to.fontSize;
        if (to.fontFace) el.fontFace = to.fontFace;
        if (to.color) el.color = typeof to.color === 'string' ? to.color : '';
        if (to.bold) el.bold = true;
        if (to.align) el.align = to.align;
      } else if (opts.fontSize) {
        el.fontSize = opts.fontSize;
        if (opts.fontFace) el.fontFace = opts.fontFace;
        if (opts.color) el.color = typeof opts.color === 'string' ? opts.color : '';
        if (opts.bold) el.bold = true;
        if (opts.align) el.align = opts.align;
      }

      slideAudit.elements.push(el);
    }

    audit.slides.push(slideAudit);
  }
  return audit;
}

async function renderDeck(rawSpec) {
  const spec = normalizeDeckSpec(rawSpec || {});
  spec._assetMap = await resolveDeckAssets(spec);
  const pptx = new PptxGenJS();
  pptx.layout = spec.layout;
  pptx.author = spec.author;
  pptx.company = 'MLCA';
  pptx.subject = spec.title;
  pptx.title = spec.title;
  pptx.lang = 'zh-CN';
  pptx.theme = {
    headFontFace: spec.theme.headerFont,
    bodyFontFace: spec.theme.bodyFont,
    lang: 'zh-CN',
  };

  const slides = spec.slides.length ? spec.slides : embeddedDeckSpec.slides;
  // Always collect audit data (no longer requires _auditOutput in spec)
  const auditEnabled = true;

  for (const [index, rawSlide] of slides.entries()) {
    const normalizedSlide = normalizeSlideForRenderer(rawSlide);
    const slide = pptx.addSlide();
    slide._pptx = pptx;
    slide._mlca = normalizedSlide;
    slide._auditKind = String(normalizedSlide.kind || 'bullets').toLowerCase();
    if (auditEnabled) {
      slide._auditElements = [];
    }
    await renderSlide(pptx, slide, rawSlide, spec, index, slides.length);
    delete slide._pptx;
    delete slide._mlca;
    delete slide._auditKind;
    delete slide._auditElements;
  }

  const uniqueSources = collectUniqueSources(slides);
  if (spec.deck.appendSourcesSlide && uniqueSources.length) {
    const slide = pptx.addSlide();
    renderSourcesSlide(slide, uniqueSources, spec.theme);
    addPageChrome(slide, spec, spec.theme, slides.length, slides.length + 1, false);
  }

  // Collect audit data before writing file
  let auditData = null;
  const normalizedSlides = slides.map(s => normalizeSlideForRenderer(s));
  if (auditEnabled) {
    auditData = collectSlideAudit(pptx, normalizedSlides);
  }

  await pptx.writeFile({ fileName: spec.outputFile });
  if (spec.deck.writeSpecCopy) {
    fs.writeFileSync(
      path.resolve(spec.deck.specOutputFile || 'deck-spec.json'),
      JSON.stringify(buildSerializableSpec(spec, slides), null, 2),
      'utf8'
    );
  }
  console.log(`Wrote ${spec.outputFile}`);

  // Write audit report (always, using outputFile name as base)
  if (auditData) {
    const auditFileName = spec._auditOutput || spec.outputFile.replace('.pptx', '_audit.json');
    const auditPath = path.resolve(auditFileName);
    fs.writeFileSync(auditPath, JSON.stringify(auditData, null, 2), 'utf8');
    console.log(`Wrote audit report: ${auditPath}`);
  }

  delete spec._assetMap;
  return spec.outputFile;
}

async function renderDeckFromFile(specPath) {
  const resolvedPath = path.resolve(specPath);
  const rawSpec = JSON.parse(fs.readFileSync(resolvedPath, 'utf8'));
  return renderDeck(rawSpec);
}

async function main() {
  const spec = loadDeckSpec();
  await renderDeck(spec);
}

if (require.main === module) {
  main().catch((error) => {
    console.error(error && error.stack ? error.stack : String(error));
    process.exit(1);
  });
}

module.exports = {
  buildSerializableSpec,
  loadDeckSpec,
  normalizeDeckSpec,
  renderDeck,
  renderDeckFromFile,
};
