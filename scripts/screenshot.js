#!/usr/bin/env node

/**
 * screenshot.js
 * 运行 Windows 主机 Chrome/Edge 浏览器，对生成的 HTML 幻灯片进行 1280x720 的无损截图
 */

const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer-core');

// Find standard Chrome/Edge locations in Windows mount
let executablePath = '';
const paths = [
    '/mnt/c/Program Files/Google/Chrome/Application/chrome.exe',
    '/mnt/c/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'
];
for (const p of paths) {
    if (fs.existsSync(p)) {
        executablePath = p;
        break;
    }
}

if (!executablePath) {
    console.error("❌ Error: No local Chrome/Edge browser found at Windows default directories.");
    console.error("   Please ensure Google Chrome or Microsoft Edge is installed on your Windows host.");
    process.exit(1);
}

async function main() {
    const args = process.argv.slice(2);
    if (args.length < 2) {
        console.log("Usage: node scripts/screenshot.js <input_html> <output_dir>");
        process.exit(1);
    }

    const htmlPath = path.resolve(args[0]);
    const outputDir = path.resolve(args[1]);

    if (!fs.existsSync(htmlPath)) {
        console.error(`❌ File not found: ${htmlPath}`);
        process.exit(1);
    }

    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    console.log(`🚀 Launching browser: ${executablePath}`);
    const browser = await puppeteer.launch({
        executablePath: executablePath,
        headless: "new",
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    try {
        const page = await browser.newPage();
        await page.setViewport({ width: 1280, height: 720, deviceScaleFactor: 2 }); // High-DPI 2x scale for visual review

        // Translate WSL path to Windows path for Windows Chrome browser
        let winPath = htmlPath;
        if (winPath.startsWith('/mnt/')) {
            const drive = winPath.charAt(5).toUpperCase();
            winPath = `${drive}:/${winPath.substring(7)}`.replace(/\\/g, '/');
        }

        const fileUrl = 'file:///' + winPath;
        console.log(`🌐 Loading HTML: ${fileUrl}`);
        await page.goto(fileUrl, { waitUntil: 'networkidle2' });

        // Query all slide frames
        const frames = await page.$$('.slide-frame');
        console.log(`📸 Found ${frames.length} slides to capture.`);

        for (let i = 0; i < frames.length; i++) {
            const outPath = path.join(outputDir, `slide_${i}.png`);
            await frames[i].screenshot({ path: outPath });
            console.log(`   [Slide ${i}] -> ${outPath}`);
        }

        console.log("✅ All screenshots captured successfully!");
    } catch (e) {
        console.error("❌ Error during screenshot capture:", e.message);
        process.exit(1);
    } finally {
        await browser.close();
    }
}

main().catch(err => {
    console.error("Fatal exception:", err);
    process.exit(1);
});
