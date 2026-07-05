#!/usr/bin/env node
/**
 * generate_gradient.js
 * 
 * 渐变底图生成器 — 读取已提炼的 visual_dna.json，
 * 根据其中的色彩规则用 Node.js Canvas 生成高保真、纯净的渐变背景图，
 * 自动保存到编译器 assets/ 目录供所有幻灯片共享，
 * 完全替代"直接把模板原图当底图"的粗糙做法。
 * 
 * 不使用第三方图形库（仅 Node.js 内置），通过 PPM 格式生成原始像素数据，
 * 再转为 JPEG。如果系统安装了 sharp 则优先使用。
 * 
 * Usage:
 *   node scripts/generate_gradient.js <theme_key> [dna_dir] [assets_dir]
 *
 * Example:
 *   node scripts/generate_gradient.js jryw4 /path/to/project/ /path/to/assets/
 */

const fs = require('fs');
const path = require('path');

// ─── CLI Args ──────────────────────────────────────────────────────────────────
const [,, themeKey, dnaDir, assetsDir] = process.argv;

if (!themeKey) {
  console.error(`
Usage: node generate_gradient.js <theme_key> [dna_dir] [assets_dir]

  theme_key  — 主题键名，如 jryw4 (需要已有对应的 <themeKey>_visual_dna.json)
  dna_dir    — [可选] visual_dna.json 所在目录，默认同 theme_key 目录
  assets_dir — [可选] 输出 JPG 底图目录，默认为 scripts/assets/

Example:
  node scripts/generate_gradient.js jryw4 /mnt/c/Downloads/edu_project/
`);
  process.exit(1);
}

const resolvedDnaDir = dnaDir ? path.resolve(dnaDir) : path.dirname(path.resolve('.'));
const defaultAssetsDir = path.join(__dirname, 'assets');
const resolvedAssetsDir = assetsDir ? path.resolve(assetsDir) : defaultAssetsDir;

// ─── Read DNA ──────────────────────────────────────────────────────────────────
const dnaPath = path.join(resolvedDnaDir, `${themeKey}_visual_dna.json`);
if (!fs.existsSync(dnaPath)) {
  console.error(`❌ 找不到视觉 DNA 文件: ${dnaPath}`);
  console.error(`   请先运行: node scripts/extract_visual_dna.js <image_path> ${themeKey} ${resolvedDnaDir}`);
  process.exit(1);
}

const dna = JSON.parse(fs.readFileSync(dnaPath, 'utf8'));
const palette = dna.palette || {};

// ─── Helper: hex → rgb ────────────────────────────────────────────────────────
function hexToRgb(hex) {
  if (!hex) return { r: 200, g: 200, b: 200 };
  const clean = hex.replace('#', '');
  const full = clean.length === 3
    ? clean.split('').map(c => c + c).join('')
    : clean;
  return {
    r: parseInt(full.slice(0, 2), 16),
    g: parseInt(full.slice(2, 4), 16),
    b: parseInt(full.slice(4, 6), 16)
  };
}

function lerp(a, b, t) { return Math.round(a + (b - a) * t); }

// ─── Generate gradient pixel buffer ──────────────────────────────────────────
function generateGradientBuffer(width, height, startHex, endHex, direction, decorations) {
  const startRgb = hexToRgb(startHex);
  const endRgb = hexToRgb(endHex);
  
  // Create pixel buffer: RGB flat array
  const buffer = new Uint8Array(width * height * 3);

  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      let t;
      switch (direction) {
        case 'to-right':    t = x / width; break;
        case 'to-bottom':   t = y / height; break;
        case 'diagonal':    t = (x / width + y / height) / 2; break;
        case 'radial': {
          const cx = width / 2, cy = height / 2;
          const maxDist = Math.sqrt(cx * cx + cy * cy);
          t = Math.sqrt((x - cx) ** 2 + (y - cy) ** 2) / maxDist;
          break;
        }
        default:            t = y / height; break;
      }
      
      // Clamp t
      t = Math.max(0, Math.min(1, t));
      // Ease in-out for softer gradient
      t = t < 0.5 ? 2 * t * t : -1 + (4 - 2 * t) * t;

      const idx = (y * width + x) * 3;
      buffer[idx]     = lerp(startRgb.r, endRgb.r, t);
      buffer[idx + 1] = lerp(startRgb.g, endRgb.g, t);
      buffer[idx + 2] = lerp(startRgb.b, endRgb.b, t);
    }
  }

  // Overlay subtle geometric decorations based on DNA decorativeVocabulary
  if (decorations && decorations.length > 0) {
    overlayDecorations(buffer, width, height, decorations, palette);
  }

  return buffer;
}

// ─── Overlay decorative elements as subtle pixel patterns ────────────────────
function overlayDecorations(buffer, width, height, decorations, palette) {
  for (const dec of decorations) {
    const opacity = dec.opacity || 0.15;
    const col = hexToRgb(dec.color || palette.accentColor || '#FFFFFF');
    
    // Draw circle rings
    if (dec.type && (dec.type.includes('circle') || dec.type.includes('ring') || dec.type.includes('orb'))) {
      let cx, cy;
      switch (dec.placement) {
        case 'top-right':   cx = width * 0.85; cy = height * 0.15; break;
        case 'top-left':    cx = width * 0.15; cy = height * 0.15; break;
        case 'bottom-right':cx = width * 0.85; cy = height * 0.85; break;
        case 'bottom-left': cx = width * 0.15; cy = height * 0.85; break;
        case 'center':      cx = width * 0.5;  cy = height * 0.5;  break;
        default:            cx = width * 0.75; cy = height * 0.25; break;
      }
      
      const sizeMap = { small: 0.12, medium: 0.22, large: 0.35, fullwidth: 0.5 };
      const radius = width * (sizeMap[dec.size] || 0.2);
      
      // Draw anti-aliased circle ring
      for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
          const dist = Math.sqrt((x - cx) ** 2 + (y - cy) ** 2);
          const ringThickness = Math.max(2, radius * 0.04);
          const distFromRing = Math.abs(dist - radius);
          
          if (distFromRing < ringThickness) {
            const alpha = opacity * (1 - distFromRing / ringThickness);
            const idx = (y * width + x) * 3;
            buffer[idx]     = Math.round(buffer[idx]     * (1 - alpha) + col.r * alpha);
            buffer[idx + 1] = Math.round(buffer[idx + 1] * (1 - alpha) + col.g * alpha);
            buffer[idx + 2] = Math.round(buffer[idx + 2] * (1 - alpha) + col.b * alpha);
          }
        }
      }
    }
    
    // Draw subtle hexagon grid dots
    if (dec.type && (dec.type.includes('hex') || dec.type.includes('dots') || dec.type.includes('grid'))) {
      const spacing = 48;
      for (let gy = 0; gy < height; gy += spacing) {
        for (let gx = 0; gx < width; gx += spacing) {
          const offset = (Math.floor(gy / spacing) % 2) * (spacing / 2);
          const px = gx + offset;
          const py = gy;
          
          // Draw small dot at grid point
          const dotRadius = 2;
          for (let dy = -dotRadius; dy <= dotRadius; dy++) {
            for (let dx = -dotRadius; dx <= dotRadius; dx++) {
              if (dx * dx + dy * dy <= dotRadius * dotRadius) {
                const nx = Math.round(px + dx), ny = Math.round(py + dy);
                if (nx >= 0 && nx < width && ny >= 0 && ny < height) {
                  const idx = (ny * width + nx) * 3;
                  const alpha = opacity * 0.4;
                  buffer[idx]     = Math.round(buffer[idx]     * (1 - alpha) + col.r * alpha);
                  buffer[idx + 1] = Math.round(buffer[idx + 1] * (1 - alpha) + col.g * alpha);
                  buffer[idx + 2] = Math.round(buffer[idx + 2] * (1 - alpha) + col.b * alpha);
                }
              }
            }
          }
        }
      }
    }
  }
}

// ─── Write PPM (portable pixmap) — no dependencies ───────────────────────────
function writePpm(buffer, width, height, outputPath) {
  const header = Buffer.from(`P6\n${width} ${height}\n255\n`);
  const fd = fs.openSync(outputPath, 'w');
  fs.writeSync(fd, header);
  fs.writeSync(fd, Buffer.from(buffer));
  fs.closeSync(fd);
}

// ─── Convert PPM → JPEG using Python (always available) ──────────────────────
function convertPpmToJpeg(ppmPath, jpegPath) {
  const { execSync } = require('child_process');
  try {
    // Try Python PIL first
    execSync(`python3 -c "
from PIL import Image
img = Image.open('${ppmPath}')
img.save('${jpegPath}', 'JPEG', quality=95)
print('PIL OK')
"`, { stdio: 'pipe' });
    return true;
  } catch {
    try {
      // Try ImageMagick convert
      execSync(`convert "${ppmPath}" "${jpegPath}"`, { stdio: 'pipe' });
      return true;
    } catch {
      try {
        // Try sharp (Node.js)
        execSync(`node -e "require('sharp')('${ppmPath}').jpeg({quality:95}).toFile('${jpegPath}')"`, { stdio: 'pipe' });
        return true;
      } catch {
        return false;
      }
    }
  }
}

// ─── Main ──────────────────────────────────────────────────────────────────────
function main() {
  if (!fs.existsSync(resolvedAssetsDir)) {
    fs.mkdirSync(resolvedAssetsDir, { recursive: true });
  }

  const width = 1280, height = 720;
  
  const startColor = palette.gradientStart || palette.bg || '#E0F7FA';
  const endColor   = palette.gradientEnd   || palette.bg || '#E0F7FA';
  const direction  = palette.gradientDirection || 'to-bottom';
  const decorations = (dna._dna || dna).decorativeVocabulary || dna.decorativeVocabulary || [];

  console.log(`\n🎨 开始生成渐变底图...`);
  console.log(`   主题: ${themeKey} (${dna.themeName || ''})`);
  console.log(`   尺寸: ${width}×${height}`);
  console.log(`   渐变: ${startColor} → ${endColor} (${direction})`);
  console.log(`   装饰元素: ${decorations.length} 个\n`);

  // Generate pixel buffer
  const buffer = generateGradientBuffer(width, height, startColor, endColor, direction, decorations);

  // Write PPM temp file
  const tmpPpm = path.join(resolvedAssetsDir, `${themeKey}_bg.ppm`);
  writePpm(buffer, width, height, tmpPpm);

  // Convert to JPEG
  const bgPath      = path.join(resolvedAssetsDir, `${themeKey}_bg.jpg`);
  const bgSlide0Path = path.join(resolvedAssetsDir, `${themeKey}_bg_slide_0.jpg`);

  const converted = convertPpmToJpeg(tmpPpm, bgPath);

  if (converted) {
    // Copy for slide 0 (cover) — could be different in future versions
    fs.copyFileSync(bgPath, bgSlide0Path);
    // Clean up temp PPM
    fs.unlinkSync(tmpPpm);
    console.log(`✅ 渐变底图已生成:`);
    console.log(`   通用底图:  ${bgPath}`);
    console.log(`   封面底图:  ${bgSlide0Path}`);
  } else {
    // Keep PPM if conversion failed, but also try direct write as reference
    console.warn(`⚠️  JPEG 转换失败（PIL/ImageMagick/sharp 均不可用），保留 PPM 原始文件`);
    console.warn(`   PPM 文件: ${tmpPpm}`);
    console.warn(`   请手动安装 PIL: pip install Pillow 或 apt install imagemagick`);
    // Still write a placeholder
    fs.copyFileSync(tmpPpm, bgPath.replace('.jpg', '.ppm'));
  }

  console.log(`\n🎯 下一步：运行 DSL JSON 生成器并使用视觉 DNA 规则指导布局：`);
  console.log(`   参考 ${themeKey}_visual_dna.json 中的 decorativeVocabulary 和 layoutPattern 设计每页布局`);
  console.log(`   编译命令: node scripts/md2pptx_web.js <deck.json> <output.pptx> -t ${themeKey}`);
}

main();
