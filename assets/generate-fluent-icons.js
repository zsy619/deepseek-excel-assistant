/**
 * Fluent-style ribbon icon generator.
 *
 * Outputs PNG icons that match the Office Fluent UI design tokens:
 *   - Transparent background (no coloured tile).
 *   - 2px stroke at 32px reference (1px @ 16px, 5px @ 80px).
 *   - Single brand colour (#217346 Excel green) on outline,
 *     optional #FFB300 accent (#D13438 alert) on selected glyphs.
 *   - Round line caps and joins.
 *
 * Reuses the SDF primitives from generate-icons.js, then swaps the
 * shell for a transparent one and replaces the per-button glyphs.
 *
 * Usage: node assets/generate-fluent-icons.js
 */

const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

/* ------------------------------------------------------------------ */
/* PNG encoding primitives (same as generate-icons.js)                */
/* ------------------------------------------------------------------ */

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
  ihdr[8] = 8;
  ihdr[9] = 6;
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;
  const rows = [];
  for (let y = 0; y < size; y++) {
    const row = [0];
    for (let x = 0; x < size; x++) {
      const [r, g, b, a] = drawPixel(x, y);
      row.push(r, g, b, a);
    }
    rows.push(Buffer.from(row));
  }
  const compressed = zlib.deflateSync(Buffer.concat(rows));
  return Buffer.concat([sig, chunk("IHDR", ihdr), chunk("IDAT", compressed), chunk("IEND", Buffer.alloc(0))]);
}

/* ------------------------------------------------------------------ */
/* SDF primitives (duplicate from generate-icons.js for portability)  */
/* ------------------------------------------------------------------ */

function smooth(t) { return t * t * (3 - 2 * t); }
function sdCircle(x, y, cx, cy, r) {
  const dx = x - cx, dy = y - cy;
  return Math.sqrt(dx * dx + dy * dy) - r;
}
function sdRing(x, y, cx, cy, rOuter, rInner) {
  return Math.max(sdCircle(x, y, cx, cy, rOuter), -sdCircle(x, y, cx, cy, rInner));
}
function sdRoundedRect(x, y, cx, cy, halfW, halfH, radius) {
  const qx = Math.max(Math.abs(x - cx) - (halfW - radius), 0);
  const qy = Math.max(Math.abs(y - cy) - (halfH - radius), 0);
  return Math.sqrt(qx * qx + qy * qy) - radius;
}
function sdBox(x, y, cx, cy, halfW, halfH) {
  return Math.max(Math.abs(x - cx) - halfW, Math.abs(y - cy) - halfH);
}
function softMask(d, aa) { return 1 - smooth(Math.max(0, Math.min(1, (d + aa) / (aa * 2)))); }
function lerpColor(a, b, t) {
  return [
    Math.round(a[0] + (b[0] - a[0]) * t),
    Math.round(a[1] + (b[1] - a[1]) * t),
    Math.round(a[2] + (b[2] - a[2]) * t),
  ];
}

/* ------------------------------------------------------------------ */
/* Line stroke primitives (round caps)                                */
/* ------------------------------------------------------------------ */

/* Cap-aware line: returns signed distance, accounting for round endcaps. */
function sdCapsule(x, y, x0, y0, x1, y1, r) {
  const dx = x1 - x0, dy = y1 - y0;
  const len2 = dx * dx + dy * dy;
  let t = 0;
  if (len2 > 0) t = Math.max(0, Math.min(1, ((x - x0) * dx + (y - y0) * dy) / len2));
  const px = x0 + dx * t;
  const py = y0 + dy * t;
  return Math.sqrt((x - px) * (x - px) + (y - py) * (y - py)) - r;
}

/* Round-cap line glyph. */
function glyphLineStroke(x0, y0, x1, y1, thickness) {
  return (x, y) => ({ intensity: softMask(sdCapsule(x, y, x0, y0, x1, y1, thickness / 2), 1.0) });
}

/* Filled disc. */
function glyphDisc(cx, cy, r) {
  return (x, y) => ({ intensity: softMask(sdCircle(x, y, cx, cy, r), 0.8) });
}

/* Filled ring. */
function glyphRing(cx, cy, rOuter, rInner) {
  return (x, y) => ({ intensity: softMask(sdRing(x, y, cx, cy, rOuter, rInner), 1.0) });
}

/* Filled rounded rect. */
function glyphRect(cx, cy, halfW, halfH, radius) {
  return (x, y) => ({ intensity: softMask(sdRoundedRect(x, y, cx, cy, halfW, halfH, radius), 1.0) });
}

/* Filled square (no rounding). */
function glyphBox(cx, cy, halfW, halfH) {
  return (x, y) => ({ intensity: softMask(sdBox(x, y, cx, cy, halfW, halfH), 1.0) });
}

/* Filled triangle (3-line union - sign test). */
function glyphTri(x1, y1, x2, y2, x3, y3) {
  function lineDist(px, py, ax, ay, bx, by) {
    const dx = bx - ax, dy = by - ay;
    const len2 = dx * dx + dy * dy;
    let t = 0;
    if (len2 > 0) t = Math.max(0, Math.min(1, ((px - ax) * dx + (py - ay) * dy) / len2));
    const ex = ax + dx * t - px, ey = ay + dy * t - py;
    return Math.sqrt(ex * ex + ey * ey);
  }
  return (x, y) => {
    const d1 = (x - x2) * (y1 - y2) - (x1 - x2) * (y - y2);
    const d2 = (x - x3) * (y2 - y3) - (x2 - x3) * (y - y3);
    const d3 = (x - x1) * (y3 - y1) - (x3 - x1) * (y - y1);
    const hasNeg = d1 < 0 || d2 < 0 || d3 < 0;
    const hasPos = d1 > 0 || d2 > 0 || d3 > 0;
    if (!(hasNeg && hasPos)) return { intensity: 1 };
    const d = Math.min(lineDist(x, y, x1, y1, x2, y2), lineDist(x, y, x2, y2, x3, y3), lineDist(x, y, x3, y3, x1, y1));
    return { intensity: softMask(d, 1.0) };
  };
}

/* Combine (max) of multiple glyphs. */
function union(...glyphs) {
  return (x, y) => {
    let best = 0;
    for (const g of glyphs) {
      const r = g(x, y).intensity;
      if (r > best) best = r;
    }
    return { intensity: best };
  };
}

/* Subtract one glyph from another. */
function subtract(base, cut) {
  return (x, y) => {
    const a = base(x, y).intensity;
    const b = cut(x, y).intensity;
    return { intensity: Math.max(0, a - b) };
  };
}

/* ------------------------------------------------------------------ */
/* Fluent colour tokens                                               */
/* ------------------------------------------------------------------ */

const TRANSPARENT = [0, 0, 0, 0];

// Excel brand green — used on outline strokes.
const BRAND = [33, 115, 70];
// Neutral mid-grey — used for support glyphs (e.g. PII, theme).
const NEUTRAL = [96, 94, 92];
// Accent amber — used for sparkle/highlight on a few glyphs.
const ACCENT = [255, 179, 0];
// Alert red — used only for PII redaction strikes.
const ALERT = [209, 52, 56];

/* ------------------------------------------------------------------ */
/* Render pipeline                                                    */
/*                                                                    */
/* Each button defines { strokes: [...], fills: [...] }. Every layer  */
/* has { glyph, color }. We render onto a transparent canvas by       */
/* stacking: each pixel's alpha is the union of every layer's         */
/* intensity; final RGB is a weighted average of layer colours.       */
/* ------------------------------------------------------------------ */

function composeStrokeOnly(layers) {
  // Single-colour line drawing where colour comes from glyph metadata.
  return function (x, y) {
    let best = -1;
    let bestIdx = -1;
    for (let i = 0; i < layers.length; i++) {
      const v = layers[i].glyph(x, y).intensity;
      if (v > best) { best = v; bestIdx = i; }
    }
    if (best <= 0.01) return TRANSPARENT;
    const c = layers[bestIdx].color;
    return [c[0], c[1], c[2], Math.round(best * 255)];
  };
}

function composeBlend(layers) {
  // Blend any overlapping layer colours using alpha-over compositing.
  return function (x, y) {
    let acc = [0, 0, 0];
    let accA = 0;
    for (const layer of layers) {
      const v = layer.glyph(x, y).intensity;
      if (v <= 0.01) continue;
      const a = v * (layer.alpha || 1.0);
      const c = layer.color;
      acc[0] = acc[0] + (c[0] - acc[0]) * (a * (1 - accA));
      acc[1] = acc[1] + (c[1] - acc[1]) * (a * (1 - accA));
      acc[2] = acc[2] + (c[2] - acc[2]) * (a * (1 - accA));
      accA = accA + a * (1 - accA);
    }
    if (accA <= 0.01) return TRANSPARENT;
    return [Math.round(acc[0]), Math.round(acc[1]), Math.round(acc[2]), Math.round(accA * 255)];
  };
}

/* ------------------------------------------------------------------ */
/* Icon designs — Fluent 2px stroke @32px reference                   */
/*                                                                    */
/* For every button we define:                                        */
/*   strokes: line segments / rings drawn at thickness `s`            */
/*   fills:   rounded rects / discs / triangles at radius 0           */
/* The composer blends all layers, transparent canvas.                */
/* ------------------------------------------------------------------ */

const s = 2; // 32px reference stroke width

function makeIcon(designFn) {
  // designFn(size) -> { strokes: [...], fills: [...] }
  // Convert each spec into a glyph function.
  return function (size) {
    const spec = designFn(size);
    const layers = [];
    if (spec.strokes) {
      for (const stroke of spec.strokes) layers.push({ glyph: glyphLineStroke(stroke.x0, stroke.y0, stroke.x1, stroke.y1, stroke.w || s), color: stroke.color, alpha: 1.0 });
    }
    if (spec.rings) {
      for (const ring of spec.rings) {
        const g = glyphRing(ring.cx, ring.cy, ring.rOuter, ring.rInner);
        layers.push({ glyph: g, color: ring.color, alpha: ring.fill ? 1.0 : 0.85 });
      }
    }
    if (spec.discs) {
      for (const d of spec.discs) {
        const g = glyphDisc(d.cx, d.cy, d.r);
        layers.push({ glyph: g, color: d.color, alpha: d.fill === false ? 0.85 : 1.0 });
      }
    }
    if (spec.rects) {
      for (const r of spec.rects) {
        const g = glyphRect(r.cx, r.cy, r.halfW, r.halfH, r.radius || 0);
        layers.push({ glyph: g, color: r.color, alpha: 1.0 });
      }
    }
    if (spec.boxes) {
      for (const r of spec.boxes) {
        const g = glyphBox(r.cx, r.cy, r.halfW, r.halfH);
        layers.push({ glyph: g, color: r.color, alpha: 1.0 });
      }
    }
    if (spec.tris) {
      for (const t of spec.tris) {
        const g = glyphTri(t.x1, t.y1, t.x2, t.y2, t.x3, t.y3);
        layers.push({ glyph: g, color: t.color, alpha: 1.0 });
      }
    }
    return layers;
  };
}

/* ---- Per-button designs ---- */

// 1. ShowTaskpane — chat bubble with three dots, rounded tail bottom-left
function showTaskpane(size) {
  const cx = size / 2, cy = size / 2;
  const w = size * 0.40, h = size * 0.30, r = size * 0.08;
  const cxOff = size * 0.02, cyOff = -size * 0.02;
  // Bubble rectangle outline
  const bubX0 = cx - w / 2, bubY0 = cy - h / 2 + cyOff;
  const bubX1 = cx + w / 2, bubY1 = cy + h / 2 + cyOff;
  // Tail (down-left)
  const tailP1 = [cx - w / 2 + size * 0.06, bubY1];
  const tailP2 = [cx - w / 2 + size * 0.20, bubY1];
  const tailP3 = [cx - w / 2 + size * 0.02, bubY1 + size * 0.12];
  return {
    strokes: [
      // bubble outline drawn as a U-shape (open at bottom-left for tail)
      { x0: bubX0, y0: bubY0 + r, x1: bubX0, y1: bubY1, w: s, color: BRAND },
      { x0: bubX0, y0: bubY0, x1: bubX1 - r, y1: bubY0, w: s, color: BRAND },
      { x0: bubX1, y0: bubY0, x1: bubX1, y1: bubY1 - r, w: s, color: BRAND },
      // bottom-right edge ending at tail root
      { x0: bubX1, y0: bubY1 - r, x1: cx + w / 2 - size * 0.06, y1: bubY1, w: s, color: BRAND },
      { x0: cx + w / 2 - size * 0.06, y0: bubY1, x1: tailP2[0], y1: tailP2[1], w: s, color: BRAND },
      { x0: tailP1[0], y0: tailP1[1], x1: tailP3[0], y1: tailP3[1], w: s, color: BRAND },
      // Round corners via 4 quarter-arcs simulated by short angled segments
      { x0: bubX0, y0: bubY0 + r, x1: bubX0 + r * 0.7, y1: bubY0 + r * 0.7, w: s, color: BRAND },
      { x0: bubX1 - r * 0.7, y0: bubY0 + r * 0.7, x1: bubX1, y1: bubY0 + r, w: s, color: BRAND },
      { x0: bubX1, y0: bubY1 - r, x1: bubX1 - r * 0.7, y1: bubY1 - r * 0.7, w: s, color: BRAND },
      { x0: bubX0, y0: bubY1, x1: tailP3[0], y1: tailP3[1], w: s, color: BRAND },
    ],
    discs: [
      // Three dots inside the bubble
      { cx: cx - size * 0.10, cy: cy + cyOff, r: size * 0.045, color: BRAND, fill: true },
      { cx: cx,             cy: cy + cyOff, r: size * 0.045, color: BRAND, fill: true },
      { cx: cx + size * 0.10, cy: cy + cyOff, r: size * 0.045, color: BRAND, fill: true },
    ],
  };
}

// 2. AnalyzeSelection — bullseye/crosshair target with four corner brackets
function analyzeSelection(size) {
  const cx = size / 2, cy = size / 2;
  return {
    rings: [
      // outer ring
      { cx, cy, rOuter: size * 0.36, rInner: size * 0.36 - s, color: BRAND },
      // inner dot (filled center)
    ],
    discs: [
      { cx, cy, r: size * 0.08, color: BRAND, fill: true },
      // four corner bracket dots
    ],
    strokes: [
      // crosshair lines extending past the ring (top/bottom/left/right)
      { x0: cx, y0: cy - size * 0.42, x1: cx, y1: cy + size * 0.42, w: s, color: BRAND },
      { x0: cx - size * 0.42, y0: cy, x1: cx + size * 0.42, y1: cy, w: s, color: BRAND },
    ],
  };
}

// 3. GenerateFormula — italic "f" with a sparkle accent (top-right)
function generateFormula(size) {
  const cx = size / 2, cy = size / 2;
  // Italic-like f drawn with two arcs approximated by capsule caps.
  // For simplicity we draw an upright "f" + "x" with a sparkle.
  const fX = cx - size * 0.16;
  const xL = cx + size * 0.04;
  const xR = cx + size * 0.20;
  return {
    strokes: [
      // f: vertical bar (slightly tilted left at top)
      { x0: fX + size * 0.02, y0: cy - size * 0.30, x1: fX - size * 0.02, y1: cy + size * 0.20, w: s, color: BRAND },
      // f: crossbar
      { x0: fX - size * 0.14, y0: cy - size * 0.08, x1: fX + size * 0.10, y1: cy - size * 0.08, w: s, color: BRAND },
      // f: curve (top hook approximated as a quarter-arc line)
      { x0: fX + size * 0.02, y0: cy - size * 0.30, x1: fX + size * 0.10, y1: cy - size * 0.34, w: s, color: BRAND },
      // x: two crossing diagonals
      { x0: xL, y0: cy - size * 0.18, x1: xR, y1: cy + size * 0.20, w: s, color: BRAND },
      { x0: xR, y0: cy - size * 0.18, x1: xL, y1: cy + size * 0.20, w: s, color: BRAND },
    ],
    discs: [
      // sparkle dot top-right
      { cx: cx + size * 0.28, cy: cy - size * 0.30, r: size * 0.06, color: ACCENT, fill: true },
    ],
  };
}

// 4. DiagnoseFormulas — pulse / heartbeat line + check mark
function diagnoseFormulas(size) {
  const cx = size / 2, cy = size / 2;
  const baseY = cy + size * 0.10;
  const points = [
    [cx - size * 0.36, baseY],
    [cx - size * 0.20, baseY],
    [cx - size * 0.10, cy - size * 0.20],
    [cx, cy + size * 0.18],
    [cx + size * 0.10, cy - size * 0.10],
    [cx + size * 0.36, baseY],
  ];
  const strokes = [];
  for (let i = 0; i < points.length - 1; i++) {
    strokes.push({ x0: points[i][0], y0: points[i][1], x1: points[i + 1][0], y1: points[i + 1][1], w: s, color: BRAND });
  }
  return { strokes, discs: [{ cx: cx - size * 0.40, cy: cy - size * 0.20, r: size * 0.04, color: BRAND, fill: true }] };
}

// 5. CleanData — magic sparkles: a large 4-pointed star and 2 small dots
function cleanData(size) {
  const cx = size / 2, cy = size / 2;
  // 4-point star: vertical + horizontal spike built from triangles.
  // Spike centre, slightly above centre for optical balance.
  const bx = cx - size * 0.04;
  const by = cy + size * 0.02;
  const bigL = size * 0.32, bigW = size * 0.10; // thicker so it survives at 16/32px
  // 4 acute isoceles triangles forming a 4-point star (vertical + horizontal spikes).
  return {
    tris: [
      // top spike (apex up)
      { x1: bx,        y1: by - bigL, x2: bx - bigW, y2: by, x3: bx + bigW, y3: by, color: BRAND },
      // bottom spike (apex down)
      { x1: bx,        y1: by + bigL, x2: bx - bigW, y2: by, x3: bx + bigW, y3: by, color: BRAND },
      // left spike (apex left)
      { x1: bx - bigL, y1: by,        x2: bx,        y2: by - bigW, x3: bx, y3: by + bigW, color: BRAND },
      // right spike (apex right)
      { x1: bx + bigL, y1: by,        x2: bx,        y2: by - bigW, x3: bx, y3: by + bigW, color: BRAND },
      // small accent star top-right — vertical
      { x1: cx + size * 0.26, y1: cy - size * 0.28, x2: cx + size * 0.22, y2: cy - size * 0.20, x3: cx + size * 0.30, y3: cy - size * 0.20, color: ACCENT },
      // small accent star top-right — horizontal
      { x1: cx + size * 0.20, y1: cy - size * 0.24, x2: cx + size * 0.32, y2: cy - size * 0.24, x3: cx + size * 0.26, y3: cy - size * 0.28, color: ACCENT },
      { x1: cx + size * 0.20, y1: cy - size * 0.24, x2: cx + size * 0.32, y2: cy - size * 0.24, x3: cx + size * 0.26, y3: cy - size * 0.20, color: ACCENT },
      // tiny bottom-left sparkle (vertical)
      { x1: cx - size * 0.22, y1: cy + size * 0.24, x2: cx - size * 0.25, y2: cy + size * 0.18, x3: cx - size * 0.19, y3: cy + size * 0.18, color: BRAND },
      // tiny bottom-left sparkle (horizontal)
      { x1: cx - size * 0.27, y1: cy + size * 0.21, x2: cx - size * 0.17, y2: cy + size * 0.21, x3: cx - size * 0.22, y3: cy + size * 0.18, color: BRAND },
      { x1: cx - size * 0.27, y1: cy + size * 0.21, x2: cx - size * 0.17, y2: cy + size * 0.21, x3: cx - size * 0.22, y3: cy + size * 0.24, color: BRAND },
    ],
    discs: [
      { cx: bx, cy: by, r: size * 0.04, color: BRAND, fill: true },
    ],
  };
}

// 6. InsertChart — three rising bars + trend arrow
function insertChart(size) {
  const cx = size / 2, cy = size / 2;
  const baseY = cy + size * 0.28;
  const heights = [size * 0.16, size * 0.26, size * 0.36];
  const widths = size * 0.08;
  const xs = [cx - size * 0.22, cx, cx + size * 0.22];
  return {
    rects: [
      { cx: xs[0], cy: baseY - heights[0] / 2, halfW: widths / 2, halfH: heights[0] / 2, radius: 1, color: BRAND },
      { cx: xs[1], cy: baseY - heights[1] / 2, halfW: widths / 2, halfH: heights[1] / 2, radius: 1, color: BRAND },
      { cx: xs[2], cy: baseY - heights[2] / 2, halfW: widths / 2, halfH: heights[2] / 2, radius: 1, color: BRAND },
    ],
    strokes: [
      // trend arrow line passing diagonally above bars
      { x0: cx - size * 0.36, y0: cy - size * 0.10, x1: cx + size * 0.30, y1: cy - size * 0.36, w: s, color: ACCENT },
      // arrowhead
      { x0: cx + size * 0.30, y0: cy - size * 0.36, x1: cx + size * 0.16, y1: cy - size * 0.34, w: s, color: ACCENT },
      { x0: cx + size * 0.30, y0: cy - size * 0.36, x1: cx + size * 0.22, y1: cy - size * 0.22, w: s, color: ACCENT },
    ],
  };
}

// 7. TranslateToCode — code brackets "</>"
function translateToCode(size) {
  const cx = size / 2, cy = size / 2;
  return {
    strokes: [
      // left "<" (top stroke)
      { x0: cx - size * 0.30, y0: cy - size * 0.18, x1: cx - size * 0.10, y1: cy, w: s + 0.2, color: BRAND },
      // left "<" (bottom stroke)
      { x0: cx - size * 0.10, y0: cy, x1: cx - size * 0.30, y1: cy + size * 0.18, w: s + 0.2, color: BRAND },
      // right ">" top
      { x0: cx + size * 0.30, y0: cy - size * 0.18, x1: cx + size * 0.10, y1: cy, w: s + 0.2, color: BRAND },
      // right ">" bottom
      { x0: cx + size * 0.10, y0: cy, x1: cx + size * 0.30, y1: cy + size * 0.18, w: s + 0.2, color: BRAND },
      // slash "/"
      { x0: cx + size * 0.04, y0: cy + size * 0.22, x1: cx - size * 0.04, y1: cy - size * 0.22, w: s + 0.2, color: ACCENT },
    ],
  };
}

// 8. MaskPII — shield outline + crossed-out redaction
function maskPii(size) {
  const cx = size / 2, cy = size / 2;
  const w = size * 0.30, h = size * 0.38;
  // Shield outline vertices
  const topL = [cx - w / 2, cy - h / 2];
  const topR = [cx + w / 2, cy - h / 2];
  const midR = [cx + w / 2, cy + h * 0.05];
  const bottom = [cx, cy + h / 2 + size * 0.02];
  const midL = [cx - w / 2, cy + h * 0.05];
  return {
    strokes: [
      // shield outline
      { x0: topL[0], y0: topL[1], x1: topR[0], y1: topR[1], w: s, color: BRAND },
      { x0: topR[0], y0: topR[1], x1: midR[0], y1: midR[1], w: s, color: BRAND },
      { x0: midR[0], y0: midR[1], x1: bottom[0], y1: bottom[1], w: s, color: BRAND },
      { x0: bottom[0], y0: bottom[1], x1: midL[0], y1: midL[1], w: s, color: BRAND },
      { x0: midL[0], y0: midL[1], x1: topL[0], y1: topL[1], w: s, color: BRAND },
      // redaction bar (warning red slashes)
      { x0: cx - size * 0.16, y0: cy - size * 0.02, x1: cx + size * 0.16, y1: cy + size * 0.02, w: s + 1, color: ALERT },
      { x0: cx - size * 0.16, y0: cy + size * 0.10, x1: cx + size * 0.10, y1: cy + size * 0.10, w: s + 1, color: ALERT },
    ],
  };
}

// 10. CorrelationMatrix — 3×3 grid of cells with intensity gradient (blue→amber→red)
function correlationMatrix(size) {
  const cx = size / 2, cy = size / 2;
  const cell = size * 0.18, gap = size * 0.02;
  const totalW = cell * 3 + gap * 2;
  const x0 = cx - totalW / 2, y0 = cy - totalW / 2;
  // 3×3 grid — fill intensities follow a "diagonal hot" pattern
  // (correlation is 1 on diagonal, lower off-diagonal).
  const intensity = [
    [1.00, 0.70, 0.40],
    [0.70, 1.00, 0.85],
    [0.40, 0.85, 1.00],
  ];
  const rects = [];
  for (let r = 0; r < 3; r++) {
    for (let c = 0; c < 3; c++) {
      const cx2 = x0 + cell / 2 + c * (cell + gap);
      const cy2 = y0 + cell / 2 + r * (cell + gap);
      // Diagonal (correlation=1) → brand green; mid → amber; low → neutral.
      const v = intensity[r][c];
      let color;
      if (v >= 0.85) color = BRAND;
      else if (v >= 0.55) color = ACCENT;
      else color = NEUTRAL;
      rects.push({ cx: cx2, cy: cy2, halfW: cell / 2, halfH: cell / 2, radius: size * 0.018, color });
    }
  }
  return { rects };
}

// 11. DetectOutliers — scatter of dots with one circled outlier
function detectOutliers(size) {
  const cx = size / 2, cy = size / 2;
  // Trend line (slight upward diagonal) — strokes behind dots.
  const linePts = [
    [cx - size * 0.36, cy + size * 0.20],
    [cx - size * 0.10, cy + size * 0.05],
    [cx + size * 0.10, cy - size * 0.05],
    [cx + size * 0.36, cy - size * 0.20],
  ];
  const strokes = [
    { x0: linePts[0][0], y0: linePts[0][1], x1: linePts[3][0], y1: linePts[3][1], w: s * 0.7, color: NEUTRAL },
  ];
  // Sample points — most along trend, one outlier circled in alert red.
  const samples = [
    { cx: cx - size * 0.30, cy: cy + size * 0.18, r: size * 0.045, color: BRAND },
    { cx: cx - size * 0.18, cy: cy + size * 0.10, r: size * 0.045, color: BRAND },
    { cx: cx - size * 0.05, cy: cy + size * 0.04, r: size * 0.045, color: BRAND },
    { cx: cx + size * 0.08, cy: cy - size * 0.05, r: size * 0.045, color: BRAND },
    { cx: cx + size * 0.22, cy: cy - size * 0.14, r: size * 0.045, color: BRAND },
    // outlier (above the trend) — alert red, larger.
    { cx: cx + size * 0.18, cy: cy - size * 0.32, r: size * 0.055, color: ALERT },
  ];
  return {
    rings: [
      // Alert ring around outlier
      { cx: cx + size * 0.18, cy: cy - size * 0.32, rOuter: size * 0.14, rInner: size * 0.14 - s, color: ALERT },
    ],
    discs: samples,
    strokes,
  };
}

// 12. CreatePivot — pivot table schema: rows header + columns header + values cell
function createPivot(size) {
  const cx = size / 2, cy = size / 2;
  const w = size * 0.66, h = size * 0.50;
  const x0 = cx - w / 2, y0 = cy - h / 2;
  return {
    strokes: [
      // outer table border
      { x0, y0, x1: x0 + w, y1: y0, w: s, color: BRAND },
      { x0: x0 + w, y0, x1: x0 + w, y1: y0 + h, w: s, color: BRAND },
      { x0: x0 + w, y0: y0 + h, x1: x0, y1: y0 + h, w: s, color: BRAND },
      { x0, y0: y0 + h, x1: x0, y1: y0, w: s, color: BRAND },
      // column header divider
      { x0, y0: y0 + h * 0.30, x1: x0 + w, y1: y0 + h * 0.30, w: s, color: BRAND },
      // vertical column dividers (3 columns)
      { x0: x0 + w / 3, y0: y0 + h * 0.30, x1: x0 + w / 3, y1: y0 + h, w: s, color: BRAND },
      { x0: x0 + (w / 3) * 2, y0: y0 + h * 0.30, x1: x0 + (w / 3) * 2, y1: y0 + h, w: s, color: BRAND },
    ],
    rects: [
      // Row-header column (left)
      { cx: x0 + w / 6, cy: y0 + h * 0.15, halfW: w / 6, halfH: h * 0.15, radius: 1, color: BRAND },
      // Column header cells (top-right 2)
      { cx: x0 + w / 2, cy: y0 + h * 0.15, halfW: w / 6, halfH: h * 0.10, radius: 1, color: NEUTRAL },
      { cx: x0 + (5 * w) / 6, cy: y0 + h * 0.15, halfW: w / 6, halfH: h * 0.10, radius: 1, color: NEUTRAL },
      // Value cell (highlighted, bottom-right)
      { cx: x0 + (5 * w) / 6, cy: y0 + h * 0.65, halfW: w / 8, halfH: h * 0.10, radius: 1, color: ACCENT },
    ],
  };
}

// 13. QuickReport — document with bullet rows + sparkline
function quickReport(size) {
  const cx = size / 2, cy = size / 2;
  const w = size * 0.60, h = size * 0.72;
  const x0 = cx - w / 2, y0 = cy - h / 2;
  return {
    strokes: [
      // Page outline
      { x0, y0, x1: x0 + w, y1: y0, w: s, color: BRAND },
      { x0: x0 + w, y0, x1: x0 + w, y1: y0 + h, w: s, color: BRAND },
      { x0: x0 + w, y0: y0 + h, x1: x0, y1: y0 + h, w: s, color: BRAND },
      { x0, y0: y0 + h, x1: x0, y1: y0, w: s, color: BRAND },
      // Title underline (thick)
      { x0: x0 + size * 0.06, y0: y0 + h * 0.22, x1: x0 + w * 0.55, y1: y0 + h * 0.22, w: s + 1, color: BRAND },
      // Bullet rows (3 lines)
      { x0: x0 + size * 0.06, y0: y0 + h * 0.42, x1: x0 + w * 0.80, y1: y0 + h * 0.42, w: s * 0.7, color: NEUTRAL },
      { x0: x0 + size * 0.06, y0: y0 + h * 0.58, x1: x0 + w * 0.70, y1: y0 + h * 0.58, w: s * 0.7, color: NEUTRAL },
      { x0: x0 + size * 0.06, y0: y0 + h * 0.74, x1: x0 + w * 0.60, y1: y0 + h * 0.74, w: s * 0.7, color: NEUTRAL },
      // Sparkline at bottom (zig-zag)
      { x0: x0 + size * 0.06, y0: y0 + h * 0.88, x1: x0 + size * 0.18, y1: y0 + h * 0.86, w: s * 0.8, color: ACCENT },
      { x0: x0 + size * 0.18, y0: y0 + h * 0.86, x1: x0 + size * 0.30, y1: y0 + h * 0.92, w: s * 0.8, color: ACCENT },
      { x0: x0 + size * 0.30, y0: y0 + h * 0.92, x1: x0 + size * 0.42, y1: y0 + h * 0.84, w: s * 0.8, color: ACCENT },
      { x0: x0 + size * 0.42, y0: y0 + h * 0.84, x1: x0 + size * 0.54, y1: y0 + h * 0.90, w: s * 0.8, color: ACCENT },
    ],
    discs: [
      // Bullet dots
      { cx: x0 + size * 0.025, cy: y0 + h * 0.42, r: size * 0.022, color: BRAND, fill: true },
      { cx: x0 + size * 0.025, cy: y0 + h * 0.58, r: size * 0.022, color: BRAND, fill: true },
      { cx: x0 + size * 0.025, cy: y0 + h * 0.74, r: size * 0.022, color: BRAND, fill: true },
    ],
  };
}

// 14. InferColumnTypes — column of cells with type tags
function inferColumnTypes(size) {
  const cx = size / 2, cy = size / 2;
  // 3 column blocks side-by-side, each with header + 2 cells + a type pill.
  const colW = size * 0.18, gap = size * 0.04;
  const totalW = colW * 3 + gap * 2;
  const x0 = cx - totalW / 2;
  const colY0 = cy - size * 0.32;
  const colH = size * 0.56;
  const strokes = [];
  for (let i = 0; i < 3; i++) {
    const x = x0 + i * (colW + gap);
    // Column header line (top)
    strokes.push({ x0: x, y0: colY0, x1: x + colW, y1: colY0, w: s, color: BRAND });
    // Column body outline
    strokes.push({ x0: x, y0: colY0 + colH, x1: x + colW, y1: colY0 + colH, w: s, color: BRAND });
    strokes.push({ x0: x, y0: colY0, x1: x, y1: colY0 + colH, w: s, color: BRAND });
    strokes.push({ x0: x + colW, y0: colY0, x1: x + colW, y1: colY0 + colH, w: s, color: BRAND });
    // Type pill bottom
    strokes.push({ x0: x + size * 0.02, y0: colY0 + colH - size * 0.14, x1: x + colW - size * 0.02, y1: colY0 + colH - size * 0.14, w: s * 0.7, color: NEUTRAL });
  }
  return {
    strokes,
    rects: [
      // Type pills (filled accents at bottom of each column)
      { cx: x0 + colW / 2, cy: colY0 + colH - size * 0.07, halfW: colW / 2 - size * 0.02, halfH: size * 0.04, radius: size * 0.02, color: ACCENT },
      { cx: x0 + colW + gap + colW / 2, cy: colY0 + colH - size * 0.07, halfW: colW / 2 - size * 0.02, halfH: size * 0.04, radius: size * 0.02, color: BRAND },
      { cx: x0 + 2 * (colW + gap) + colW / 2, cy: colY0 + colH - size * 0.07, halfW: colW / 2 - size * 0.02, halfH: size * 0.04, radius: size * 0.02, color: NEUTRAL },
    ],
  };
}

// 9. (Tab logo) brand — same bubble as showTaskpane but with one larger spark
function brandTab(size) {
  const cx = size / 2, cy = size / 2;
  const w = size * 0.40, h = size * 0.32, r = size * 0.10;
  const cyOff = -size * 0.02;
  const bubX0 = cx - w / 2, bubY0 = cy - h / 2 + cyOff;
  const bubX1 = cx + w / 2, bubY1 = cy + h / 2 + cyOff;
  const tailP2 = [cx - w / 2 + size * 0.20, bubY1];
  const tailP3 = [cx - w / 2 + size * 0.02, bubY1 + size * 0.14];
  return {
    strokes: [
      { x0: bubX0 + r, y0: bubY0, x1: bubX1 - r, y1: bubY0, w: s, color: BRAND },
      { x0: bubX1, y0: bubY0, x1: bubX1, y1: bubY1 - r, w: s, color: BRAND },
      { x0: bubX1, y0: bubY1 - r, x1: cx + w / 2 - size * 0.06, y1: bubY1, w: s, color: BRAND },
      { x0: cx + w / 2 - size * 0.06, y0: bubY1, x1: tailP2[0], y1: tailP2[1], w: s, color: BRAND },
      { x0: bubX0, y0: bubY0 + r, x1: bubX0, y1: bubY1, w: s, color: BRAND },
      { x0: bubX0, y0: bubY1, x1: tailP3[0], y1: tailP3[1], w: s, color: BRAND },
      // corner suggestions
      { x0: bubX0, y0: bubY0 + r, x1: bubX0 + r * 0.6, y1: bubY0 + r * 0.6, w: s, color: BRAND },
      { x0: bubX1 - r * 0.6, y0: bubY0 + r * 0.6, x1: bubX1, y1: bubY0 + r, w: s, color: BRAND },
      { x0: bubX1, y0: bubY1 - r, x1: bubX1 - r * 0.6, y1: bubY1 - r * 0.6, w: s, color: BRAND },
    ],
    discs: [
      // center dot
      { cx, cy: cy + cyOff, r: size * 0.06, color: ACCENT, fill: true },
      // small accent dot top-right
      { cx: cx + size * 0.24, cy: cy - size * 0.24, r: size * 0.04, color: BRAND, fill: true },
    ],
  };
}

/* ------------------------------------------------------------------ */
/* Main                                                               */
/* ------------------------------------------------------------------ */

const outDir = __dirname;
const sizes = [16, 32, 80];

const buttons = [
  { key: "brand", file: "ribbon-brand" },
  { key: "showTaskpane", file: "ribbon-show-taskpane" },
  { key: "analyzeSelection", file: "ribbon-analyze" },
  { key: "generateFormula", file: "ribbon-formula" },
  { key: "diagnoseFormulas", file: "ribbon-diagnose" },
  { key: "cleanData", file: "ribbon-clean" },
  { key: "insertChart", file: "ribbon-chart-insert" },
  { key: "translateToCode", file: "ribbon-vba" },
  { key: "maskPii", file: "ribbon-pii" },
  { key: "correlationMatrix", file: "ribbon-correlation" },
  { key: "detectOutliers", file: "ribbon-outlier" },
  { key: "createPivot", file: "ribbon-pivot" },
  { key: "quickReport", file: "ribbon-report" },
  { key: "inferColumnTypes", file: "ribbon-coltypes" },
];

for (const b of buttons) {
  const designFn = {
    brand: brandTab,
    showTaskpane,
    analyzeSelection,
    generateFormula,
    diagnoseFormulas,
    cleanData,
    insertChart,
    translateToCode,
    maskPii,
    correlationMatrix,
    detectOutliers,
    createPivot,
    quickReport,
    inferColumnTypes,
  }[b.key];

  const factory = makeIcon(designFn);
  for (const size of sizes) {
    const layers = factory(size);
    const drawPixel = composeBlend(layers);
    const png = makePng(size, drawPixel);
    const out = path.join(outDir, `${b.file}-${size}.png`);
    fs.writeFileSync(out, png);
  }
  console.log(`Generated ${b.file.padEnd(22)} 16/32/80`);
}

// Also produce 64/128 brand icons for the existing app icon URLs.
const brandFactory = makeIcon(brandTab);
for (const size of [64, 128]) {
  const layers = brandFactory(size);
  const drawPixel = composeBlend(layers);
  const png = makePng(size, drawPixel);
  fs.writeFileSync(path.join(outDir, `icon-${size}.png`), png);
}
console.log("Regenerated icon-64, icon-128 in Fluent style");

console.log("\nDone.");
