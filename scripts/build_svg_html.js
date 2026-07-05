// scripts/build_svg_html.js
const fs = require('fs');
const path = require('path');

const projectPath = path.resolve('/home/administrator/.gemini/antigravity-cli/brain/7886a93f-fb92-435a-b44c-876dcbaaaa59/education_project');
const svgDir = path.join(projectPath, 'svg_output');
const htmlOutPath = path.join(projectPath, 'output_education_svg.html');

function buildHtml() {
    const files = fs.readdirSync(svgDir)
        .filter(f => f.endsWith('.svg'))
        .sort(); // 01_cover.svg, 02_trends.svg...

    let slidesHtml = '';
    files.forEach((file, idx) => {
        const svgContent = fs.readFileSync(path.join(svgDir, file), 'utf-8');
        // Clean up XML declaration if present
        const cleanedSvg = svgContent.replace(/<\?xml[^>]*\?>/g, '').trim();
        const activeClass = idx === 0 ? ' active' : '';
        slidesHtml += `      <div class="slide-item${activeClass}" id="slide-${idx}">
        ${cleanedSvg}
      </div>\n`;
    });

    const htmlContent = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Custom SVG Presentation Preview</title>
  <style>
    body {
      background-color: #08090d;
      color: #e8eaed;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      margin: 0;
      padding: 0;
      overflow: hidden;
      display: flex;
      align-items: center;
      justify-content: center;
      min-height: 100vh;
      width: 100vw;
    }
    .deck-shell {
      position: fixed;
      inset: 0;
      overflow: hidden;
      display: flex;
      align-items: center;
      justify-content: center;
      pointer-events: none;
    }
    .deck-stage {
      width: 1280px;
      height: 720px;
      position: relative;
      transform-origin: center center;
      box-shadow: 0 30px 80px rgba(0, 0, 0, 0.4);
      background-color: #E0F7FA;
      pointer-events: auto;
    }
    .slide-item {
      position: absolute;
      inset: 0;
      display: none;
      width: 100%;
      height: 100%;
    }
    .slide-item.active {
      display: block !important;
    }
    /* Floating Toolbar */
    .toolbar {
      position: fixed;
      bottom: 25px;
      left: 50%;
      transform: translateX(-50%);
      background: rgba(30, 41, 59, 0.85);
      backdrop-filter: blur(8px);
      border: 1px solid rgba(255, 255, 255, 0.1);
      padding: 8px 24px;
      border-radius: 30px;
      display: flex;
      gap: 15px;
      align-items: center;
      z-index: 1000;
      box-shadow: 0 10px 25px rgba(0,0,0,0.35);
    }
    .toolbar-title {
      font-size: 14px;
      font-weight: 600;
      color: #94A3B8;
    }
    .theme-btn {
      border: none;
      padding: 6px 16px;
      border-radius: 20px;
      cursor: pointer;
      font-weight: bold;
      font-size: 13px;
      transition: all 0.2s;
      background: #10B981;
      color: #FFFFFF;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      gap: 5px;
    }
    @media print {
      body { background: white !important; overflow: visible !important; display: block !important; min-height: auto !important; width: auto !important; }
      .deck-shell { position: relative !important; inset: auto !important; display: block !important; pointer-events: auto !important; }
      .deck-stage { transform: none !important; box-shadow: none !important; width: 1280px !important; height: 720px !important; position: relative !important; }
      .slide-item { display: block !important; position: relative !important; width: 1280px !important; height: 720px !important; page-break-after: always !important; }
      .toolbar { display: none !important; }
    }
  </style>
</head>
<body>
  <div class="deck-shell">
    <div class="deck-stage" id="stage">
${slidesHtml}
    </div>
  </div>

  <div class="toolbar">
    <span class="toolbar-title">动态排版模式: 自由SVG生成</span>
    <span style="color: rgba(255,255,255,0.2); margin: 0 5px;">|</span>
    <span class="toolbar-title">页码:</span>
    <span id="slide-counter" class="toolbar-title" style="background: rgba(255,255,255,0.08); padding: 4px 10px; border-radius: 12px; font-family: monospace; color: #FFFFFF;">1 / 1</span>
    <span style="color: rgba(255,255,255,0.2); margin: 0 5px;">|</span>
    <a href="exports/education_project.pptx" class="theme-btn">
      📥 下载 Native DrawingML PPTX
    </a>
  </div>

  <script>
    let currentSlideIdx = 0;
    const slideItems = Array.from(document.querySelectorAll('.slide-item'));
    
    function showSlide(idx) {
      if (idx < 0 || idx >= slideItems.length) return;
      slideItems.forEach((el, i) => {
        el.classList.toggle('active', i === idx);
      });
      currentSlideIdx = idx;
      updatePageCounter();
    }
    
    function updatePageCounter() {
      const counterEl = document.getElementById('slide-counter');
      if (counterEl) {
        counterEl.innerText = (currentSlideIdx + 1) + ' / ' + slideItems.length;
      }
    }

    window.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
        showSlide(Math.min(slideItems.length - 1, currentSlideIdx + 1));
        e.preventDefault();
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        showSlide(Math.max(0, currentSlideIdx - 1));
        e.preventDefault();
      } else if (e.key === 'Home') {
        showSlide(0);
        e.preventDefault();
      } else if (e.key === 'End') {
        showSlide(slideItems.length - 1);
        e.preventDefault();
      }
    });

    function fit() {
      const stage = document.getElementById('stage');
      if (!stage) return;
      const sw = 1280, sh = 720;
      const vw = window.innerWidth, vh = window.innerHeight;
      const scale = Math.min(vw / sw, vh / sh);
      stage.style.transform = 'scale(' + scale + ')';
    }
    
    window.addEventListener('resize', fit);
    window.addEventListener('load', () => {
      fit();
      updatePageCounter();
    });
  </script>
</body>
</html>`;

    fs.writeFileSync(htmlOutPath, htmlContent, 'utf-8');
    console.log(`✅ Responsive HTML preview from custom SVGs built at: ${htmlOutPath}`);
}

buildHtml();
