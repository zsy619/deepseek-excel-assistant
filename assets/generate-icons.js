/**
 * Generate per-button ribbon icons + the app brand icon.
 *
 * Why per-button icons: Office's ribbon will fall back to a "Logo"
 * placeholder when several buttons share the same image resource, or
 * when the image file fails to load. Giving each control its own
 * dedicated icon (3 sizes: 16, 32, 80) makes the ribbon behave properly
 * and lets users distinguish the buttons at a glance.
 *
 * Each icon is a rounded square with a vertical gradient + a single
 * white glyph in the center representing the action. PNG output is
 * RGBA so it composites cleanly on both light and dark Office themes.
 *
 * Usage: node assets/generate-icons.js
 */

const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

/* ----------------------------------------------------------------- */
/* PNG encoding primitives                                            */
/* ----------------------------------------------------------------- */

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc = CRC_TABLE[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, "ascii");
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0);
  return Buffer.concat([len, typeBuf, data, crcBuf]);
}

function makePng(size, drawPixel) {
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type RGBA
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const rows = [];
  for (let y = 0; y < size; y++) {
    const row = [0]; // filter byte
    for (let x = 0; x < size; x++) {
      const [r, g, b, a] = drawPixel(x, y);
      row.push(r, g, b, a);
    }
    rows.push(Buffer.from(row));
  }
  const raw = Buffer.concat(rows);
  const compressed = zlib.deflateSync(raw);
  return Buffer.concat([
    sig,
    chunk("IHDR", ihdr),
    chunk("IDAT", compressed),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

/* ----------------------------------------------------------------- */
/* Drawing helpers                                                    */
/* ----------------------------------------------------------------- */

function smooth(t) {
  return t * t * (3 - 2 * t);
}

function sdRoundedRect(x, y, cx, cy, halfW, halfH, radius) {
  const qx = Math.max(Math.abs(x - cx) - (halfW - radius), 0);
  const qy = Math.max(Math.abs(y - cy) - (halfH - radius), 0);
  return Math.sqrt(qx * qx + qy * qy) - radius;
}

function sdCircle(x, y, cx, cy, r) {
  const dx = x - cx;
  const dy = y - cy;
  return Math.sqrt(dx * dx + dy * dy) - r;
}

function sdRing(x, y, cx, cy, rOuter, rInner) {
  return Math.max(sdCircle(x, y, cx, cy, rOuter), -sdCircle(x, y, cx, cy, rInner));
}

function sdBox(x, y, cx, cy, halfW, halfH) {
  return Math.max(
    Math.abs(x - cx) - halfW,
    Math.abs(y - cy) - halfH
  );
}

function softMask(d, aa) {
  return 1 - smooth(Math.max(0, Math.min(1, (d + aa) / (aa * 2))));
}

function lerpColor(a, b, t) {
  return [
    Math.round(a[0] + (b[0] - a[0]) * t),
    Math.round(a[1] + (b[1] - a[1]) * t),
    Math.round(a[2] + (b[2] - a[2]) * t),
    Math.round(a[3] + (b[3] - a[3]) * t),
  ];
}

const TRANSPARENT = [0, 0, 0, 0];

/* ----------------------------------------------------------------- */
/* Generic icon shell: rounded square + gradient background           */
/* ----------------------------------------------------------------- */

function makeShell(size, bgTop, bgBottom) {
  const cx = size / 2;
  const cy = size / 2;
  const half = size / 2;
  const radius = Math.round(size * 0.22);
  const pad = size * 0.04;
  return function shell(x, y) {
    const dBg = sdRoundedRect(x, y, cx, cy, half - pad, half - pad, radius);
    if (dBg > 0.8) return { color: TRANSPARENT, alpha: 0 };
    const tY = Math.max(0, Math.min(1, y / size));
    return {
      color: lerpColor(bgTop, bgBottom, tY),
      alpha: softMask(dBg, 1.0),
    };
  };
}

function compose(shell, glyph, glyphColor) {
  return function (x, y) {
    const s = shell(x, y);
    if (s.alpha <= 0) return TRANSPARENT;
    let outColor = s.color;
    let outAlpha = s.alpha;
    const g = glyph(x, y);
    if (g.intensity > 0) {
      outColor = lerpColor(outColor, glyphColor, g.intensity);
      outAlpha = Math.min(1, outAlpha + g.intensity * 0.3);
    }
    return [
      outColor[0],
      outColor[1],
      outColor[2],
      Math.round(outAlpha * 255),
    ];
  };
}

/* ----------------------------------------------------------------- */
/* Glyphs - each returns { intensity } 0..1 with antialiased edges    */
/* ----------------------------------------------------------------- */

/** Filled rounded rect. */
function glyphRect(cx, cy, halfW, halfH, radius, size) {
  return function (x, y) {
    const d = sdRoundedRect(x, y, cx, cy, halfW, halfH, radius);
    return { intensity: softMask(d, 1.0) };
  };
}

/** Filled circle. */
function glyphCircle(cx, cy, r) {
  return function (x, y) {
    const d = sdCircle(x, y, cx, cy, r);
    return { intensity: softMask(d, 0.7) };
  };
}

/** Ring (annulus). */
function glyphRing(cx, cy, rOuter, rInner) {
  return function (x, y) {
    const d = sdRing(x, y, cx, cy, rOuter, rInner);
    return { intensity: softMask(d, 0.8) };
  };
}

/** Line segment drawn with a thick "pen" (capsule). */
function glyphLine(x0, y0, x1, y1, thickness) {
  return function (x, y) {
    const dx = x1 - x0;
    const dy = y1 - y0;
    const len2 = dx * dx + dy * dy;
    let t = 0;
    if (len2 > 0) {
      t = ((x - x0) * dx + (y - y0) * dy) / len2;
      t = Math.max(0, Math.min(1, t));
    }
    const px = x0 + dx * t;
    const py = y0 + dy * t;
    const d = Math.sqrt((x - px) * (x - px) + (y - py) * (y - py)) - thickness;
    return { intensity: softMask(d, 0.8) };
  };
}

/** Triangle filled (used for arrows). */
function glyphTriangle(x1, y1, x2, y2, x3, y3) {
  return function (x, y) {
    // Barycentric sign test.
    const d1 = (x - x2) * (y1 - y2) - (x1 - x2) * (y - y2);
    const d2 = (x - x3) * (y2 - y3) - (x2 - x3) * (y - y3);
    const d3 = (x - x1) * (y3 - y1) - (x3 - x1) * (y - y1);
    const hasNeg = d1 < 0 || d2 < 0 || d3 < 0;
    const hasPos = d1 > 0 || d2 > 0 || d3 > 0;
    if (!(hasNeg && hasPos)) {
      // inside
      return { intensity: 1 };
    }
    // antialias by distance to nearest edge
    const e1 = lineDist(x, y, x1, y1, x2, y2);
    const e2 = lineDist(x, y, x2, y2, x3, y3);
    const e3 = lineDist(x, y, x3, y3, x1, y1);
    const d = Math.min(e1, e2, e3);
    return { intensity: softMask(d, 1.0) };
  };
}
function lineDist(px, py, x1, y1, x2, y2) {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const len2 = dx * dx + dy * dy;
  let t = 0;
  if (len2 > 0) t = Math.max(0, Math.min(1, ((px - x1) * dx + (py - y1) * dy) / len2));
  const ex = x1 + dx * t - px;
  const ey = y1 + dy * t - py;
  return Math.sqrt(ex * ex + ey * ey);
}

/** Multiply/combine multiple glyphs into one. */
function combine(...glyphs) {
  return function (x, y) {
    let best = 0;
    for (const g of glyphs) {
      const r = g(x, y).intensity;
      if (r > best) best = r;
    }
    return { intensity: best };
  };
}

/* ----------------------------------------------------------------- */
/* Icon designs                                                       */
/* ----------------------------------------------------------------- */

const WHITE = [255, 255, 255, 255];

const ICONS = {
  /** App brand icon - chat bubble with tail and a single bright "DS" spark.
   *  Designed to be readable at 16x16 (one big central dot, not three
   *  tiny ones, plus a tail that says "speaking"). */
  brand: {
    bgTop: [0, 120, 212, 255],
    bgBottom: [0, 90, 158, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      // Bubble fills most of the icon, with a clear tail bottom-left.
      const bubbleHalfW = size * 0.36;
      const bubbleHalfH = size * 0.28;
      const bubbleCenterY = cy - size * 0.02;
      // Tail hangs from bottom-left of the bubble.
      const tailTipX = cx - size * 0.30;
      const tailTipY = cy + size * 0.34;
      const tailBaseY = bubbleCenterY + bubbleHalfH - size * 0.02;
      const tailBaseHalf = size * 0.10;
      const tailBaseX = cx - size * 0.18;
      // Single bright dot inside the bubble.
      const dotR = size * 0.11;
      const dotX = cx;
      const dotY = bubbleCenterY;
      // Green spark at top-right (AI active indicator).
      const sparkR = size * 0.07;
      const sparkX = size * 0.78;
      const sparkY = size * 0.22;
      return combine(
        // bubble body
        glyphRect(cx, bubbleCenterY, bubbleHalfW, bubbleHalfH, size * 0.10, size),
        // bubble tail (triangle)
        glyphTriangle(
          tailBaseX - tailBaseHalf, tailBaseY,
          tailBaseX + tailBaseHalf, tailBaseY,
          tailTipX, tailTipY
        ),
        // central bright dot
        glyphCircle(dotX, dotY, dotR),
        // green spark top-right
        glyphCircle(sparkX, sparkY, sparkR)
      );
    },
  },

  /** 分析选区 (Analyze Selection) - bar chart with up-trend arrow. */
  analyze: {
    bgTop: [76, 90, 200, 255],
    bgBottom: [54, 65, 175, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      const barW = size * 0.10;
      const baseY = cy + size * 0.18;
      const heights = [size * 0.16, size * 0.24, size * 0.32];
      const xs = [cx - size * 0.18, cx, cx + size * 0.18];
      return combine(
        glyphRect(xs[0], baseY - heights[0] / 2, barW / 2, heights[0] / 2, size * 0.02, size),
        glyphRect(xs[1], baseY - heights[1] / 2, barW / 2, heights[1] / 2, size * 0.02, size),
        glyphRect(xs[2], baseY - heights[2] / 2, barW / 2, heights[2] / 2, size * 0.02, size),
        // trend arrow (line + arrowhead)
        glyphLine(cx - size * 0.20, baseY - size * 0.10, cx + size * 0.20, baseY - size * 0.30, size * 0.025),
        glyphTriangle(cx + size * 0.13, baseY - size * 0.32, cx + size * 0.27, baseY - size * 0.22, cx + size * 0.18, baseY - size * 0.20)
      );
    },
  },

  /** 生成公式 (Generate Formula) - fx with sparkle. */
  formula: {
    bgTop: [16, 124, 110, 255],
    bgBottom: [10, 90, 80, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      // Two diagonal strokes forming "f" + "x" (we approximate with two crossing lines + a hook)
      return combine(
        // "f" stem (vertical)
        glyphLine(cx - size * 0.14, cy - size * 0.22, cx - size * 0.14, cy + size * 0.22, size * 0.04),
        // "f" cross
        glyphLine(cx - size * 0.24, cy - size * 0.04, cx - size * 0.04, cy - size * 0.04, size * 0.035),
        // "x" diagonal 1
        glyphLine(cx + size * 0.02, cy - size * 0.16, cx + size * 0.22, cy + size * 0.16, size * 0.04),
        // "x" diagonal 2
        glyphLine(cx + size * 0.22, cy - size * 0.16, cx + size * 0.02, cy + size * 0.16, size * 0.04),
        // sparkle dot top-right
        glyphCircle(cx + size * 0.22, cy - size * 0.24, size * 0.03)
      );
    },
  },

  /** 数据清洗 (Clean Data) - brush strokes / sparkles. */
  clean: {
    bgTop: [217, 119, 38, 255],
    bgBottom: [180, 90, 20, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      // broom handle (diagonal) + bristles (a soft cloud at the bottom)
      return combine(
        glyphLine(cx - size * 0.22, cy - size * 0.22, cx + size * 0.16, cy + size * 0.16, size * 0.04),
        // bristles base
        glyphRect(cx + size * 0.16, cy + size * 0.16, size * 0.14, size * 0.08, size * 0.04, size),
        // sparkle 1
        glyphCircle(cx - size * 0.10, cy - size * 0.05, size * 0.045),
        // sparkle 2 (small)
        glyphCircle(cx - size * 0.22, cy + size * 0.10, size * 0.03)
      );
    },
  },

  /** 插入回复 (Insert Reply) - arrow pointing down into a cell. */
  insert: {
    bgTop: [99, 70, 180, 255],
    bgBottom: [70, 50, 140, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      // Down arrow with a horizontal "tray" line below.
      return combine(
        // arrow shaft
        glyphLine(cx, cy - size * 0.22, cx, cy + size * 0.10, size * 0.045),
        // arrowhead
        glyphTriangle(cx - size * 0.10, cy + size * 0.04, cx + size * 0.10, cy + size * 0.04, cx, cy + size * 0.20),
        // cell top edge
        glyphLine(cx - size * 0.18, cy + size * 0.22, cx + size * 0.18, cy + size * 0.22, size * 0.035)
      );
    },
  },

  /** 导出对话 (Export Conversation) - document with down arrow. */
  export: {
    bgTop: [16, 138, 87, 255],
    bgBottom: [10, 100, 60, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      // Document rectangle with three text lines and a down arrow on the right.
      return combine(
        // doc body
        glyphRect(cx - size * 0.10, cy - size * 0.04, size * 0.16, size * 0.24, size * 0.04, size),
        // three text lines
        glyphLine(cx - size * 0.18, cy + size * 0.04, cx - size * 0.02, cy + size * 0.04, size * 0.022),
        glyphLine(cx - size * 0.18, cy + size * 0.12, cx - size * 0.02, cy + size * 0.12, size * 0.022),
        // down arrow
        glyphLine(cx + size * 0.14, cy - size * 0.18, cx + size * 0.14, cy + size * 0.18, size * 0.035),
        glyphTriangle(cx + size * 0.06, cy + size * 0.10, cx + size * 0.22, cy + size * 0.10, cx + size * 0.14, cy + size * 0.22)
      );
    },
  },

  /** 清空对话 (Clear Conversation) - circle with X. */
  clear: {
    bgTop: [200, 75, 75, 255],
    bgBottom: [160, 50, 50, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      const r = size * 0.22;
      return combine(
        // X cross
        glyphLine(cx - r * 0.7, cy - r * 0.7, cx + r * 0.7, cy + r * 0.7, size * 0.05),
        glyphLine(cx + r * 0.7, cy - r * 0.7, cx - r * 0.7, cy + r * 0.7, size * 0.05)
      );
    },
  },

  /** 切换主题 (Toggle Theme) - half moon + half sun. */
  theme: {
    bgTop: [120, 70, 180, 255],
    bgBottom: [85, 45, 140, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      const r = size * 0.18;
      // Crescent on the left, sun disk on the right.
      return combine(
        // Sun: filled circle on the right
        glyphCircle(cx + size * 0.04, cy, r),
        // Moon: full circle + a slightly-offset cutout circle => crescent
        glyphCircle(cx - size * 0.08, cy, r * 1.05)
      );
    },
  },

  /** 打开设置 (Open Settings) - gear. */
  settings: {
    bgTop: [82, 88, 100, 255],
    bgBottom: [55, 60, 72, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      const outerR = size * 0.24;
      const innerR = size * 0.12;
      // 8-tooth gear: outer ring with 8 small "teeth" bumps
      const teeth = 8;
      const glyphs = [glyphRing(cx, cy, outerR, innerR)];
      for (let i = 0; i < teeth; i++) {
        const angle = (i / teeth) * Math.PI * 2;
        const tx = cx + Math.cos(angle) * (outerR + size * 0.04);
        const ty = cy + Math.sin(angle) * (outerR + size * 0.04);
        glyphs.push(glyphCircle(tx, ty, size * 0.05));
      }
      glyphs.push(glyphCircle(cx, cy, size * 0.07));
      return combine(...glyphs);
    },
  },

  /** 诊断公式 (Diagnose Formulas) - magnifier over an exclamation mark.
   *  Conveys "look for errors with AI help". */
  diagnose: {
    bgTop: [217, 70, 70, 255],
    bgBottom: [167, 30, 30, 255],
    glyph: (size) => {
      const cx = size / 2 - size * 0.04;
      const cy = size / 2 - size * 0.04;
      const ringR = size * 0.18;
      const stroke = size * 0.045;
      // Exclamation dot+line inside the magnifier
      const bar = glyphRect(cx, cy - size * 0.06, size * 0.025, size * 0.07, 0, size);
      const dot = glyphCircle(cx, cy + size * 0.07, size * 0.028);
      // Magnifier handle (diagonal line bottom-right)
      const handle = glyphLine(
        cx + ringR * 0.7, cy + ringR * 0.7,
        cx + ringR * 0.7 + size * 0.14, cy + ringR * 0.7 + size * 0.14,
        stroke
      );
      return combine(
        glyphRing(cx, cy, ringR, ringR - stroke),
        bar,
        dot,
        handle
      );
    },
  },

  /** 公式转 VBA (Translate to VBA) - code brackets "</>" with a sparkle. */
  vba: {
    bgTop: [56, 132, 240, 255],
    bgBottom: [30, 95, 200, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      const stroke = size * 0.04;
      // Left "<" - two strokes meeting at left side
      const leftArm1 = glyphLine(cx - size * 0.22, cy - size * 0.14, cx - size * 0.06, cy, stroke);
      const leftArm2 = glyphLine(cx - size * 0.06, cy, cx - size * 0.22, cy + size * 0.14, stroke);
      // Right ">" - two strokes meeting at right side
      const rightArm1 = glyphLine(cx + size * 0.22, cy - size * 0.14, cx + size * 0.06, cy, stroke);
      const rightArm2 = glyphLine(cx + size * 0.06, cy, cx + size * 0.22, cy + size * 0.14, stroke);
      // Slash "/" between
      const slash = glyphLine(cx + size * 0.02, cy + size * 0.16, cx - size * 0.02, cy - size * 0.16, stroke);
      // Sparkle top-right
      const spark = glyphCircle(cx + size * 0.22, cy - size * 0.24, size * 0.04);
      return combine(leftArm1, leftArm2, rightArm1, rightArm2, slash, spark);
    },
  },

    /** 多选区分析 (Multi-selection) - two stacked rectangles showing
   *  separate selection regions. */
  multi: {
    bgTop: [234, 88, 12, 255],
    bgBottom: [180, 60, 4, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      // Two offset rectangles representing different selection regions
      const back = glyphRect(cx - size * 0.18, cy - size * 0.22, size * 0.30, size * 0.22, size * 0.04, size);
      const front = glyphRect(cx - size * 0.06, cy - size * 0.02, size * 0.30, size * 0.22, size * 0.04, size);
      return combine(back, front);
    },
  },

  /** 数据脱敏 (Mask PII) - shield with redaction marks. */
  pii: {
    bgTop: [124, 58, 237, 255],
    bgBottom: [88, 28, 200, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      const w = size * 0.32;
      const h = size * 0.36;
      // Shield outline + 3 redaction bars inside.
      const shield = combine(
        // Top edge
        glyphLine(cx - w / 2, cy - h / 2, cx + w / 2, cy - h / 2, size * 0.05),
        // Right curve -> bottom point -> left curve
        glyphLine(cx + w / 2, cy - h / 2, cx + w / 2, cy + h * 0.1, size * 0.05),
        glyphLine(cx + w / 2, cy + h * 0.1, cx, cy + h / 2 + size * 0.05, size * 0.05),
        glyphLine(cx, cy + h / 2 + size * 0.05, cx - w / 2, cy + h * 0.1, size * 0.05),
        glyphLine(cx - w / 2, cy + h * 0.1, cx - w / 2, cy - h / 2, size * 0.05)
      );
      const bar1 = glyphRect(cx - size * 0.14, cy - size * 0.08, size * 0.10, size * 0.05, size * 0.01, size);
      const bar2 = glyphRect(cx - size * 0.14, cy, size * 0.16, size * 0.05, size * 0.01, size);
      const bar3 = glyphRect(cx - size * 0.14, cy + size * 0.08, size * 0.08, size * 0.05, size * 0.01, size);
      return combine(shield, bar1, bar2, bar3);
    },
  },
  chartInsert: {
    bgTop: [13, 148, 136, 255],
    bgBottom: [6, 110, 100, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      const baseY = cy + size * 0.18;
      const barW = size * 0.10;
      const heights = [size * 0.12, size * 0.22, size * 0.34];
      const xs = [cx - size * 0.20, cx, cx + size * 0.20];
      // Build up arrow + bars
      return combine(
        // bar 1
        glyphRect(xs[0], baseY - heights[0] / 2, barW / 2, heights[0] / 2, size * 0.02, size),
        // bar 2
        glyphRect(xs[1], baseY - heights[1] / 2, barW / 2, heights[1] / 2, size * 0.02, size),
        // bar 3
        glyphRect(xs[2], baseY - heights[2] / 2, barW / 2, heights[2] / 2, size * 0.02, size),
        // trend arrow
        glyphLine(cx - size * 0.22, baseY - size * 0.18, cx + size * 0.22, baseY - size * 0.34, size * 0.03),
        // arrowhead
        glyphTriangle(cx + size * 0.12, baseY - size * 0.34, cx + size * 0.26, baseY - size * 0.30, cx + size * 0.20, baseY - size * 0.22)
      );
    },
  },
  /** 知识库 (Knowledge base) - a stack of books with a sparkle. */
  knowledge: {
    bgTop: [79, 70, 229, 255],
    bgBottom: [55, 48, 163, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      // Three horizontal book spines stacked.
      const spineH = size * 0.07;
      const spineW = size * 0.32;
      const spineX = cx;
      const ys = [cy - size * 0.16, cy, cy + size * 0.16];
      const spines = ys.map((y) =>
        glyphRect(spineX, y, spineW / 2, spineH / 2, size * 0.015, size)
      );
      // Sparkle top-right
      const spark = glyphCircle(cx + size * 0.22, cy - size * 0.22, size * 0.04);
      return combine(...spines, spark);
    },
  },
  /** 分享 (Share) - chain link glyph, blue/green palette. */
  share: {
    bgTop: [20, 184, 166, 255],
    bgBottom: [13, 148, 136, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      const stroke = size * 0.05;
      // Two interlocking rings on a diagonal.
      const ringA = glyphRing(cx - size * 0.10, cy - size * 0.06, size * 0.16, size * 0.16 - stroke);
      const ringB = glyphRing(cx + size * 0.10, cy + size * 0.06, size * 0.16, size * 0.16 - stroke);
      return combine(ringA, ringB);
    },
  },
  /** 用量看板 (Usage dashboard) - bar chart + sparkline feel. */
  usage: {
    bgTop: [37, 99, 235, 255],
    bgBottom: [29, 78, 216, 255],
    glyph: (size) => {
      const cx = size / 2;
      const cy = size / 2;
      const baseY = cy + size * 0.18;
      const barW = size * 0.10;
      const heights = [size * 0.16, size * 0.22, size * 0.30, size * 0.36];
      const xs = [
        cx - size * 0.21,
        cx - size * 0.07,
        cx + size * 0.07,
        cx + size * 0.21,
      ];
      return combine(
        ...xs.map((x, i) =>
          glyphRect(x, baseY - heights[i] / 2, barW / 2, heights[i] / 2, size * 0.02, size)
        )
      );
    },
  },
};

/* ----------------------------------------------------------------- */
/* Main                                                               */
/* ----------------------------------------------------------------- */

const outDir = __dirname;
const sizes = [16, 32, 80];

// Per-button icons. We write 16/32/80 each.
const buttons = [
  { key: "brand", file: "icon" }, // app brand
  { key: "brand", file: "ribbon-brand" },
  { key: "analyze", file: "ribbon-analyze" },
  { key: "formula", file: "ribbon-formula" },
  { key: "clean", file: "ribbon-clean" },
  { key: "diagnose", file: "ribbon-diagnose" },
  { key: "vba", file: "ribbon-vba" },
  { key: "chartInsert", file: "ribbon-chart-insert" },
  { key: "pii", file: "ribbon-pii" },
  { key: "multi", file: "ribbon-multi" },
  { key: "insert", file: "ribbon-insert" },
  { key: "export", file: "ribbon-export" },
  { key: "clear", file: "ribbon-clear" },
  { key: "theme", file: "ribbon-theme" },
  { key: "settings", file: "ribbon-settings" },
  { key: "knowledge", file: "ribbon-knowledge" },
  { key: "share", file: "ribbon-share" },
  { key: "usage", file: "ribbon-usage" },
];

for (const b of buttons) {
  const design = ICONS[b.key];
  for (const size of sizes) {
    const shell = makeShell(size, design.bgTop, design.bgBottom);
    const glyph = design.glyph(size);
    const drawPixel = compose(shell, glyph, WHITE);
    const png = makePng(size, drawPixel);
    const out = path.join(outDir, `${b.file}-${size}.png`);
    fs.writeFileSync(out, png);
  }
  console.log(`Generated ${b.file.padEnd(20)} 16/32/80`);
}

// Also keep the 64/128 brand sizes for the app icon.
{
  const design = ICONS.brand;
  for (const size of [64, 128]) {
    const shell = makeShell(size, design.bgTop, design.bgBottom);
    const glyph = design.glyph(size);
    const drawPixel = compose(shell, glyph, WHITE);
    const png = makePng(size, drawPixel);
    const out = path.join(outDir, `icon-${size}.png`);
    fs.writeFileSync(out, png);
  }
  console.log("Generated icon-64, icon-128 (app brand)");
}

console.log("\nDone. Icons written to:", outDir);
