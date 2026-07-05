#!/usr/bin/env node
/**
 * extract_visual_dna.js
 * 
 * 视觉 DNA 提炼器 — 接收用户提供的图片模板路径，
 * 调用 Gemini Vision API 按 7 个标准维度提炼视觉规则，
 * 输出结构化 visual_dna.json，并自动同步到 themes.json。
 * 
 * Usage:
 *   node scripts/extract_visual_dna.js <image_path> <theme_key> [output_dir]
 *
 * Example:
 *   node scripts/extract_visual_dna.js /path/to/template.png jryw4 /path/to/project/
 *
 * Requires: GEMINI_API_KEY in env, Node.js 18+
 */

const fs = require('fs');
const path = require('path');
const https = require('https');

// ─── CLI Args ──────────────────────────────────────────────────────────────────
const [,, imagePath, themeKey, outputDir] = process.argv;

if (!imagePath || !themeKey) {
  console.error(`
Usage: node extract_visual_dna.js <image_path> <theme_key> [output_dir]

  image_path  — 图片模板的绝对路径 (PNG/JPG/WEBP)
  theme_key   — 主题键名，如 jryw4 / medical / mycompany (建议用图片名前5位)
  output_dir  — [可选] 视觉 DNA JSON 的输出目录，默认为图片所在目录

Example:
  node scripts/extract_visual_dna.js /mnt/c/Users/Administrator/Downloads/template.png jryw4 /mnt/c/Users/Administrator/Downloads/my_project/
`);
  process.exit(1);
}

const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
if (!GEMINI_API_KEY) {
  console.error('❌ 缺少环境变量 GEMINI_API_KEY，请先执行：export GEMINI_API_KEY="your-key"');
  process.exit(1);
}

// ─── Read image ────────────────────────────────────────────────────────────────
if (!fs.existsSync(imagePath)) {
  console.error(`❌ 图片文件不存在: ${imagePath}`);
  process.exit(1);
}

const imageBuffer = fs.readFileSync(imagePath);
const imageBase64 = imageBuffer.toString('base64');
const ext = path.extname(imagePath).toLowerCase().replace('.', '');
const mimeTypes = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', webp: 'image/webp' };
const mimeType = mimeTypes[ext] || 'image/png';

// ─── Prompt ────────────────────────────────────────────────────────────────────
const DNA_EXTRACTION_PROMPT = `你是一位顶级的 UI/UX 设计总监，专长于从设计图中提炼完整的视觉设计系统。

请对这张 PPT 幻灯片模板图片进行系统性的视觉 DNA 提炼，严格按以下 7 个维度分析，并以 JSON 格式返回结果。

**提炼维度：**

1. **palette（色彩语义角色）**：
   - bg: 背景主色（整体背景色调）
   - cardBg: 卡片/面板背景色（比背景略浅或深）
   - cardBorder: 卡片边框/分割线色
   - titleColor: 标题文字主色（最显眼的大字颜色）
   - textColor: 正文/辅助文字颜色
   - accentColor: 强调/高亮/交互色（如按钮、图标、进度条）
   - dividerColor: 横向分割线颜色
   - gradientStart: 渐变起始色（如果背景是渐变，否则等同于bg）
   - gradientEnd: 渐变结束色（如果背景是渐变，否则等同于bg）
   - gradientDirection: 渐变方向 ("to-right" | "to-bottom" | "diagonal" | "radial" | "none")

2. **typography（字型规则）**：
   - headingFont: 标题字体名称
   - bodyFont: 正文字体名称
   - headingWeight: 标题字重 ("bold" | "semibold" | "normal")
   - headingScale: 标题字号估算（px，如48）
   - bodyScale: 正文字号估算（px，如14）
   - letterSpacing: 字间距感觉 ("tight" | "normal" | "wide")
   - lineHeight: 行高感觉 ("tight" | "relaxed" | "spacious")

3. **spacing（间距节奏）**：
   - unit: 基础间距单位估算（px，如8或16）
   - pagePadding: 页面四边留白感觉 ("compact" | "normal" | "generous")
   - cardPadding: 卡片内边距感觉 ("compact" | "normal" | "generous")
   - elementGap: 元素间距估算（px，如12）

4. **componentGrammar（组件语法规则）**：
   - cardRadius: 卡片圆角感觉 ("none" | "small" | "medium" | "large" | "pill")
   - cardShadow: 卡片投影 ("none" | "subtle" | "medium" | "strong")
   - cardStyle: 卡片主视觉风格 ("flat" | "glassmorphism" | "neumorphism" | "outlined" | "filled" | "floating")
   - buttonStyle: 按钮/标签风格 ("rounded" | "square" | "pill" | "ghost")
   - iconStyle: 图标风格 ("outline" | "filled" | "duotone" | "emoji")

5. **decorativeVocabulary（装饰元素词汇）**：
   列出图中存在的所有装饰性元素，每个元素包含：
   - type: 装饰类型 (如 "geometric-circles" | "hexagons" | "diagonal-stripes" | "dots-grid" | "wave" | "glow-orb" | "book-icon" | "gradient-blob" 等)
   - color: 颜色
   - opacity: 透明度估算 (0.0-1.0)
   - placement: 位置 ("top-left" | "top-right" | "bottom-left" | "bottom-right" | "center" | "scattered")
   - size: 大小感觉 ("small" | "medium" | "large" | "fullwidth")

6. **layoutPattern（布局模式）**：
   - gridColumns: 大致列数 (如 2 | 3 | 4)
   - contentSplit: 内容区主要分割方式 ("left-heavy" | "right-heavy" | "balanced" | "centered" | "full-bleed")
   - headerStyle: 顶部标题区样式 ("minimal-bar" | "colored-band" | "no-header" | "full-width")
   - footerStyle: 底部样式 ("page-dots" | "page-number" | "logo-bar" | "none")
   - hasVerticalAccentBar: 是否有竖向强调色块 (true/false)
   - hasDiagonalCut: 是否有斜切分割 (true/false)

7. **mood（视觉情绪与适用场景）**：
   - tone: 整体情绪 ("professional" | "playful" | "academic" | "creative" | "luxury" | "tech" | "warm" | "minimal")
   - isDark: 是否深色主题 (true/false)
   - energyLevel: 活跃程度 ("calm" | "balanced" | "energetic" | "vibrant")
   - bestFor: 最适合的场景描述（中文，50字以内）

**输出格式（严格 JSON，不要任何 Markdown 包裹）：**
{
  "themeKey": "<<用户指定的主题键>>",
  "themeName": "<<给这个主题起一个中文名称，如：教育青绿渐变>>",
  "palette": { ... },
  "typography": { ... },
  "spacing": { ... },
  "componentGrammar": { ... },
  "decorativeVocabulary": [ ... ],
  "layoutPattern": { ... },
  "mood": { ... },
  "extractedAt": "<<ISO 8601 时间戳>>"
}

请仅返回 JSON，不要任何解释或包裹文本。`;

// ─── Call Gemini Vision API ────────────────────────────────────────────────────
function callGeminiVision(prompt, imageBase64, mimeType) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({
      contents: [{
        parts: [
          { text: prompt },
          { inline_data: { mime_type: mimeType, data: imageBase64 } }
        ]
      }],
      generationConfig: {
        temperature: 0.1,
        topP: 0.9,
        maxOutputTokens: 4096
      }
    });

    const options = {
      hostname: 'generativelanguage.googleapis.com',
      path: `/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.error) return reject(new Error(`Gemini API Error: ${json.error.message}`));
          const text = json.candidates?.[0]?.content?.parts?.[0]?.text;
          if (!text) return reject(new Error('Gemini 返回了空响应'));
          resolve(text);
        } catch (e) {
          reject(new Error(`解析 Gemini 响应失败: ${e.message}`));
        }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// ─── Parse DNA JSON from LLM response ─────────────────────────────────────────
function parseDnaJson(rawText) {
  // Strip markdown code fences if present
  let cleaned = rawText.trim();
  cleaned = cleaned.replace(/^```json\s*/i, '').replace(/\s*```$/, '');
  cleaned = cleaned.replace(/^```\s*/i, '').replace(/\s*```$/, '');
  return JSON.parse(cleaned);
}

// ─── Sync DNA to themes.json ───────────────────────────────────────────────────
function syncToThemes(dna, themeKey, themesJsonPath) {
  let themes = {};
  if (fs.existsSync(themesJsonPath)) {
    themes = JSON.parse(fs.readFileSync(themesJsonPath, 'utf8'));
  }

  const p = dna.palette;
  const t = dna.typography;
  const m = dna.mood;

  // Build theme entry from DNA
  themes[themeKey] = {
    name: dna.themeName || `${themeKey} 主题`,
    bg: p.bg ? p.bg.replace('#', '') : 'F8F9FA',
    cardBg: p.cardBg ? p.cardBg.replace('#', '') : 'FFFFFF',
    cardBorder: p.cardBorder ? p.cardBorder.replace('#', '') : 'E2E8F0',
    titleColor: p.titleColor ? p.titleColor.replace('#', '') : '0F172A',
    textColor: p.textColor ? p.textColor.replace('#', '') : '374151',
    accentColor: p.accentColor ? p.accentColor.replace('#', '') : '3B82F6',
    formulaColor: p.accentColor ? p.accentColor.replace('#', '') : '2563EB',
    dividerColor: p.dividerColor ? p.dividerColor.replace('#', '') : 'E2E8F0',
    fontFace: t.headingFont || 'Segoe UI',
    fontFamily: `"${t.headingFont || 'Segoe UI'}", "${t.bodyFont || 'Segoe UI'}", sans-serif`,
    isDark: m.isDark || false,
    bgImage: true,
    // Extended DNA fields (used by gradient generator and layout engine)
    _dna: {
      gradientStart: p.gradientStart || p.bg,
      gradientEnd: p.gradientEnd || p.bg,
      gradientDirection: p.gradientDirection || 'none',
      decorativeVocabulary: dna.decorativeVocabulary || [],
      layoutPattern: dna.layoutPattern || {},
      mood: dna.mood || {},
      componentGrammar: dna.componentGrammar || {},
      spacing: dna.spacing || {}
    }
  };

  fs.writeFileSync(themesJsonPath, JSON.stringify(themes, null, 2), 'utf8');
  return themes[themeKey];
}

// ─── Main ──────────────────────────────────────────────────────────────────────
async function main() {
  const resolvedImagePath = path.resolve(imagePath);
  const resolvedOutputDir = outputDir ? path.resolve(outputDir) : path.dirname(resolvedImagePath);
  const themesJsonPath = path.join(__dirname, 'themes.json');

  // Ensure output dir exists
  if (!fs.existsSync(resolvedOutputDir)) {
    fs.mkdirSync(resolvedOutputDir, { recursive: true });
  }

  console.log(`\n🔍 开始提炼视觉 DNA...`);
  console.log(`   📄 图片: ${resolvedImagePath}`);
  console.log(`   🔑 主题键: ${themeKey}`);
  console.log(`   📂 输出目录: ${resolvedOutputDir}`);
  console.log(`   🤖 调用 Gemini Vision API...\n`);

  let rawResponse;
  try {
    rawResponse = await callGeminiVision(DNA_EXTRACTION_PROMPT.replace('<<用户指定的主题键>>', themeKey), imageBase64, mimeType);
  } catch (e) {
    console.error(`❌ API 调用失败: ${e.message}`);
    process.exit(1);
  }

  let dna;
  try {
    dna = parseDnaJson(rawResponse);
    dna.themeKey = themeKey;
    dna.extractedAt = new Date().toISOString();
  } catch (e) {
    console.error(`❌ 解析 DNA JSON 失败: ${e.message}`);
    console.error('原始响应:\n', rawResponse);
    process.exit(1);
  }

  // Write visual_dna.json to output dir
  const dnaOutputPath = path.join(resolvedOutputDir, `${themeKey}_visual_dna.json`);
  fs.writeFileSync(dnaOutputPath, JSON.stringify(dna, null, 2), 'utf8');
  console.log(`✅ 视觉 DNA 已提炼并保存: ${dnaOutputPath}`);

  // Sync to themes.json
  const themeEntry = syncToThemes(dna, themeKey, themesJsonPath);
  console.log(`✅ themes.json 已同步更新 → 主题 "${themeKey}" (${themeEntry.name})`);

  // Print summary
  console.log(`\n📊 视觉 DNA 提炼摘要：`);
  console.log(`   🎨 色彩系统:`);
  console.log(`      背景色: ${dna.palette.bg}`);
  console.log(`      强调色: ${dna.palette.accentColor}`);
  console.log(`      标题色: ${dna.palette.titleColor}`);
  if (dna.palette.gradientDirection !== 'none') {
    console.log(`      渐变: ${dna.palette.gradientStart} → ${dna.palette.gradientEnd} (${dna.palette.gradientDirection})`);
  }
  console.log(`   🔤 字型:`);
  console.log(`      标题: ${dna.typography.headingFont} ${dna.typography.headingWeight}`);
  console.log(`      正文: ${dna.typography.bodyFont}`);
  console.log(`   🎭 装饰元素:`);
  (dna.decorativeVocabulary || []).forEach(d => {
    console.log(`      - ${d.type} (${d.placement}, opacity ${d.opacity})`);
  });
  console.log(`   🏗️  布局:`);
  console.log(`      内容分割: ${dna.layoutPattern.contentSplit}`);
  console.log(`      列数: ${dna.layoutPattern.gridColumns}`);
  console.log(`   💡 情绪定位: ${dna.mood.tone} / ${dna.mood.bestFor}`);
  console.log(`\n🎯 下一步：运行渐变底图生成器：`);
  console.log(`   node scripts/generate_gradient.js ${themeKey} "${resolvedOutputDir}"`);
}

main().catch(e => {
  console.error('❌ 未预期错误:', e);
  process.exit(1);
});
