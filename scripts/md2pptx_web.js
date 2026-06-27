#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const { execSync } = require('child_process');

const args = process.argv.slice(2);
if (args.length === 0) {
    console.error("Usage: node md2pptx_web.js <input.md> [-o <output.pptx>]");
    process.exit(1);
}

let inputFile = args[0];
let outputFile = "output.pptx";

for (let i = 1; i < args.length; i++) {
    if (args[i] === '-o' || args[i] === '--output') {
        outputFile = args[i+1];
        i++;
    }
}

const mdText = fs.readFileSync(inputFile, 'utf-8');

let slidesData;
let isDslDeck = false;

if (inputFile.endsWith('.json')) {
    try {
        const deck = JSON.parse(mdText);
        const { validateDeck } = require('../tests/validate_dsl.js');
        const errors = validateDeck(deck);
        if (errors.length > 0) {
            console.error("❌ DSL Validation Failed:");
            errors.forEach(err => console.error("  - " + err));
            process.exit(1);
        }
        slidesData = deck.slides;
        slidesData.forEach(slide => {
            slide.isDsl = true;
            slide.dsl = slide;
        });
        isDslDeck = true;
    } catch (e) {
        console.error("❌ Failed to parse JSON slide deck:", e.message);
        process.exit(1);
    }
}


function getPngDimensions(filePath) {
    try {
        const fd = fs.openSync(filePath, 'r');
        const buffer = Buffer.alloc(8);
        fs.readSync(fd, buffer, 0, 8, 16);
        fs.closeSync(fd);
        const width = buffer.readInt32BE(0);
        const height = buffer.readInt32BE(4);
        if (width > 0 && height > 0) {
            return { width, height };
        }
    } catch (e) {
        // quiet fallback
    }
    return null;
}

function getImageDimensions(filePath) {
    const pngDims = getPngDimensions(filePath);
    if (pngDims) return pngDims;

    try {
        const output = execSync(`file "${filePath}"`, { encoding: 'utf-8' });
        const match = output.match(/(\d+)\s*x\s*(\d+)/);
        if (match) {
            return {
                width: parseInt(match[1], 10),
                height: parseInt(match[2], 10)
            };
        }
    } catch (e) {
        console.error("Failed to get image dimensions for", filePath, e);
    }
    return null;
}

function parseFormulaToObjects(formulaText, baseOptions = {}) {
    let result = [];
    let i = 0;
    let currentText = "";
    
    const flush = () => {
        if (currentText) {
            result.push({ text: currentText, options: { ...baseOptions } });
            currentText = "";
        }
    };
    
    while (i < formulaText.length) {
        if (formulaText[i] === '^') {
            flush();
            i++;
            if (formulaText[i] === '{') {
                let end = formulaText.indexOf('}', i);
                if (end !== -1) {
                    result.push({ text: formulaText.substring(i + 1, end), options: { ...baseOptions, superscript: true } });
                    i = end + 1;
                } else {
                    currentText += '^';
                }
            } else if (i < formulaText.length) {
                result.push({ text: formulaText[i], options: { ...baseOptions, superscript: true } });
                i++;
            } else {
                currentText += '^';
            }
        } else if (formulaText[i] === '_') {
            flush();
            i++;
            if (formulaText[i] === '{') {
                let end = formulaText.indexOf('}', i);
                if (end !== -1) {
                    result.push({ text: formulaText.substring(i + 1, end), options: { ...baseOptions, subscript: true } });
                    i = end + 1;
                } else {
                    currentText += '_';
                }
            } else if (i < formulaText.length) {
                result.push({ text: formulaText[i], options: { ...baseOptions, subscript: true } });
                i++;
            } else {
                currentText += '_';
            }
        } else {
            currentText += formulaText[i];
            i++;
        }
    }
    flush();
    return result;
}

function parseLineToSegments(lineText, lineOpts = {}) {
    let parts = lineText.split(/(\*\*.*?\*\*|\$.*?\$|\*.*?\*)/g);
    let segments = [];
    
    parts.forEach(part => {
        if (!part) return;
        
        let segOpts = { ...lineOpts };
        
        if (part.startsWith('**') && part.endsWith('**')) {
            segOpts.bold = true;
            segments.push({ text: part.slice(2, -2), options: segOpts });
        } else if (part.startsWith('$') && part.endsWith('$')) {
            let formulaText = part.slice(1, -1);
            let formulaRuns = parseFormulaToObjects(formulaText, { ...segOpts, italic: true, color: '003366' });
            segments.push(...formulaRuns);
        } else if (part.startsWith('*') && part.endsWith('*')) {
            segOpts.italic = true;
            segments.push({ text: part.slice(1, -1), options: segOpts });
        } else {
            segments.push({ text: part, options: segOpts });
        }
    });
    
    return segments;
}

function processLines(dataText, fontSize) {
    const lines = dataText.split('\n');
    let allSegments = [];
    const activeLines = lines.map(l => l.trimRight()).filter(l => l.length > 0);
    
    activeLines.forEach((line, idx) => {
        let bullet = false;
        let indentLevel = 0;
        let rawLine = line;
        
        const listMatch = line.match(/^(\s*)([-\*\+])\s+(.*)$/);
        if (listMatch) {
            bullet = true;
            const spaces = listMatch[1].length;
            indentLevel = Math.floor(spaces / 2);
            rawLine = listMatch[3];
        } else {
            const numMatch = line.match(/^(\s*)(\d+)\.\s+(.*)$/);
            if (numMatch) {
                bullet = true;
                const spaces = numMatch[1].length;
                indentLevel = Math.floor(spaces / 2);
                rawLine = numMatch[3];
            } else {
                const indentMatch = line.match(/^(\s+)(.*)$/);
                if (indentMatch) {
                    const spaces = indentMatch[1].length;
                    indentLevel = Math.floor(spaces / 2);
                    rawLine = indentMatch[2];
                }
            }
        }
        
        let lineSegments = parseLineToSegments(rawLine, { color: '333333', fontSize: fontSize });
        if (lineSegments.length > 0) {
            if (bullet) {
                lineSegments.forEach(seg => {
                    seg.options.bullet = true;
                });
            }
            if (indentLevel > 0) {
                lineSegments.forEach(seg => {
                    seg.options.indentLevel = indentLevel;
                });
            }
            if (idx < activeLines.length - 1) {
                lineSegments[lineSegments.length - 1].options.breakLine = true;
            }
            
            allSegments.push(...lineSegments);
        }
    });
    
    return allSegments;
}

function calculateDynamicFontSize(text, hasImage) {
    const textLength = text.length;
    if (hasImage) {
        if (textLength > 200) return 13;
        if (textLength > 150) return 14;
        if (textLength > 100) return 16;
        return 18;
    } else {
        if (textLength > 400) return 14;
        if (textLength > 200) return 16;
        return 18;
    }
}

function getSlideLayoutInfo(data, hasImage) {
    if (hasImage) {
        return { type: 'asymmetric' };
    }
    const lines = data.text.split('\n');
    const activeLines = lines.map(l => l.trimRight()).filter(l => l.length > 0);
    
    let introLines = [];
    let bulletLines = [];
    let inBullets = false;
    
    activeLines.forEach(line => {
        const isBullet = line.match(/^(\s*)([-\*\+])\s+(.*)$/) || line.match(/^(\s*)(\d+)\.\s+(.*)$/);
        if (isBullet) {
            inBullets = true;
        }
        if (inBullets) {
            bulletLines.push(line);
        } else {
            introLines.push(line);
        }
    });
    
    let minIndent = 999;
    bulletLines.forEach(line => {
        const match = line.match(/^(\s*)([-\*\+]|\d+\.)\s+/);
        if (match) {
            const indent = match[1].length;
            if (indent < minIndent) minIndent = indent;
        }
    });
    if (minIndent === 999) minIndent = 0;

    let cardItems = [];
    let currentCard = null;
    
    bulletLines.forEach(line => {
        const isBullet = line.match(/^(\s*)([-\*\+])\s+(.*)$/) || line.match(/^(\s*)(\d+)\.\s+(.*)$/);
        if (isBullet) {
            const indent = isBullet[1].length;
            if (indent <= minIndent) {
                const cleanContent = line.replace(/^\s*[-\*\+]?\s*/, '').replace(/^\s*\d+\.\s*/, '').trim();
                const boldMatch = cleanContent.match(/^\*\*(.*?)\*\*(.*)$/);
                if (boldMatch) {
                    currentCard = {
                        title: boldMatch[1].trim(),
                        body: boldMatch[2].replace(/^[:：\s]*/, '').trim()
                    };
                    cardItems.push(currentCard);
                } else {
                    const splitIdx = cleanContent.indexOf('：') !== -1 ? cleanContent.indexOf('：') : cleanContent.indexOf(':');
                    if (splitIdx !== -1) {
                        currentCard = {
                            title: cleanContent.substring(0, splitIdx).replace(/\*\*/g, '').trim(),
                            body: cleanContent.substring(splitIdx + 1).trim()
                        };
                        cardItems.push(currentCard);
                    } else {
                        currentCard = {
                            title: cleanContent.substring(0, 8).replace(/\*\*/g, '').trim() + "...",
                            body: cleanContent
                        };
                        cardItems.push(currentCard);
                    }
                }
            } else {
                if (currentCard) {
                    if (currentCard.body) {
                        currentCard.body += '\n' + line;
                    } else {
                        currentCard.body = line;
                    }
                }
            }
        } else {
            if (currentCard) {
                if (currentCard.body) {
                    currentCard.body += '\n' + line;
                } else {
                    currentCard.body = line;
                }
            }
        }
    });

    // Detect timeline/sequence
    let isTimeline = false;
    if (cardItems.length >= 3 && cardItems.length <= 5) {
        const slideTitleLower = (data.title || "").toLowerCase();
        const hasTimelineKeyword = slideTitleLower.includes('evolution') || 
                                    slideTitleLower.includes('timeline') || 
                                    slideTitleLower.includes('process') || 
                                    slideTitleLower.includes('stage') || 
                                    slideTitleLower.includes('step') || 
                                    slideTitleLower.includes('history') || 
                                    slideTitleLower.includes('workflow') || 
                                    slideTitleLower.includes('sequence') || 
                                    slideTitleLower.includes('演进') || 
                                    slideTitleLower.includes('流程') || 
                                    slideTitleLower.includes('阶段') || 
                                    slideTitleLower.includes('步骤') || 
                                    slideTitleLower.includes('历史') || 
                                    slideTitleLower.includes('发展');
                                    
        const allNumbered = bulletLines.filter(line => line.trim().match(/^(\s*)([-\*\+]|\d+\.)/)).every(line => line.trim().match(/^\d+\./));
        const titleHasSequence = cardItems.every((item, idx) => {
            const t = item.title.toLowerCase();
            const hasStepWord = t.includes('step') || t.includes('phase') || t.includes('stage') || t.includes('步骤') || t.includes('阶段');
            return hasStepWord || (t.match(/^\d+/) && hasTimelineKeyword);
        });
        
        if (hasTimelineKeyword || allNumbered || titleHasSequence) {
            if (cardItems.length === 4 && !hasTimelineKeyword) {
                isTimeline = false;
            } else {
                isTimeline = true;
            }
        }
    }

    if (isTimeline) {
        return { type: 'timeline', introLines, cardItems };
    }

    const canUseColumns = cardItems.length >= 2 && cardItems.length <= 4;
    if (canUseColumns) {
        return { type: 'grid', introLines, cardItems };
    }

    // Check for Centered Breathe (low density text, word count is small, no columns)
    const textTrimmed = data.text.trim();
    if (textTrimmed.length > 0 && textTrimmed.length < 120 && cardItems.length < 2) {
        return { type: 'centered-breathe' };
    }

    return { type: 'default' };
}

function parseSlides(markdown) {
    const rawSlides = markdown.split('\n---\n');
    let slides = [];
    
    rawSlides.forEach(raw => {
        const lines = raw.trim().split('\n');
        if (lines.length === 0 || !lines[0]) return;
        
        let slideData = {
            title: '',
            text: '',
            images: [],
            notes: ''
        };
        
        // Extract notes
        const notesMatch = raw.match(/<!--\s*notes?:?(.*?)\s*-->/is);
        if (notesMatch) {
            slideData.notes = notesMatch[1].trim();
        }
        
        let textLines = [];
        
        lines.forEach(line => {
            const originalLine = line;
            line = line.trim();
            if (line.match(/<!--\s*notes/i)) return;
            if (line === '-->') return;
            
            const imgMatch = line.match(/!\[(.*?)\]\((.*?)\)/);
            if (imgMatch) {
                slideData.images.push({ path: imgMatch[2], alt: imgMatch[1] });
                return;
            }
            
            if (line.startsWith('# ') || line.startsWith('## ')) {
                if (!slideData.title) {
                    slideData.title = line.replace(/^#+\s/, '');
                    return;
                }
            }
            
            textLines.push(originalLine.trimRight());
        });
        
        slideData.text = textLines.join('\n').trim();
        slides.push(slideData);
    });
    
    return slides;
}

function escapeHtml(text) {
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function segmentsToHtml(segments) {
    if (!segments || segments.length === 0) return '';
    
    // Group segments by paragraph
    let paragraphs = [];
    let currentParagraph = [];
    
    segments.forEach(seg => {
        currentParagraph.push(seg);
        if (seg.options.breakLine) {
            paragraphs.push(currentParagraph);
            currentParagraph = [];
        }
    });
    if (currentParagraph.length > 0) {
        paragraphs.push(currentParagraph);
    }
    
    let html = '';
    paragraphs.forEach(p => {
        const firstSeg = p[0];
        const isBullet = firstSeg.options.bullet;
        const indentLevel = firstSeg.options.indentLevel || 0;
        const align = firstSeg.options.align || 'left';
        
        if (isBullet) {
            let containerStyle = `margin: 0; padding: 0; text-align: ${align}; line-height: 1.5; display: flex; align-items: flex-start; margin-left: ${indentLevel * 24}px; margin-bottom: 8px; `;
            let bulletSize = firstSeg.options.fontSize || 14;
            html += `<div style="${containerStyle}">`;
            // Bullet marker column
            html += `<div style="flex-shrink: 0; width: 18px; text-align: left; font-size: ${bulletSize}pt; color: var(--accent-color); line-height: 1.4; padding-top: 1px;">•</div>`;
            // Content column
            html += `<div style="flex-grow: 1;">`;
            p.forEach(seg => {
                let spanStyle = `font-size: ${seg.options.fontSize || 14}pt; `;
                if (seg.options.bold) spanStyle += `font-weight: bold; `;
                if (seg.options.italic) spanStyle += `font-style: italic; `;
                if (seg.options.color) {
                    if (seg.options.color === '003366') {
                        spanStyle += `color: var(--formula-color); `;
                    } else if (seg.options.color === '333333') {
                        spanStyle += `color: var(--text-color); `;
                    } else {
                        spanStyle += `color: #${seg.options.color}; `;
                    }
                }
                if (seg.options.superscript) spanStyle += `vertical-align: super; font-size: 70%; `;
                if (seg.options.subscript) spanStyle += `vertical-align: sub; font-size: 70%; `;
                
                html += `<span style="${spanStyle}">${escapeHtml(seg.text)}</span>`;
            });
            html += `</div></div>`;
        } else {
            let pStyle = `margin: 0; padding: 0; text-align: ${align}; line-height: 1.5; margin-left: ${indentLevel * 24}px; margin-bottom: 8px; `;
            html += `<div style="${pStyle}">`;
            p.forEach(seg => {
                let spanStyle = `font-size: ${seg.options.fontSize || 14}pt; `;
                if (seg.options.bold) spanStyle += `font-weight: bold; `;
                if (seg.options.italic) spanStyle += `font-style: italic; `;
                if (seg.options.color) {
                    if (seg.options.color === '003366') {
                        spanStyle += `color: var(--formula-color); `;
                    } else if (seg.options.color === '333333') {
                        spanStyle += `color: var(--text-color); `;
                    } else {
                        spanStyle += `color: #${seg.options.color}; `;
                    }
                }
                if (seg.options.superscript) spanStyle += `vertical-align: super; font-size: 70%; `;
                if (seg.options.subscript) spanStyle += `vertical-align: sub; font-size: 70%; `;
                
                html += `<span style="${spanStyle}">${escapeHtml(seg.text)}</span>`;
            });
            html += `</div>`;
        }
    });
    
    return html;
}

function renderSlideListHtml(slidesList, globalSlidesData) {
    let content = "";
    slidesList.forEach((data, idx) => {
        content += `        <div class="slide-item">
            <div class="slide-frame">
`;
        if (data.isDsl) {
            const formatMarkdownTokens = (txt) => segmentsToHtml(processLines(txt, 11));
            const runningSubtitle = (slidesList[0] && slidesList[0].dsl && slidesList[0].dsl.elements)
                ? (slidesList[0].dsl.elements.find(el => el.id === 'title_card')?.content?.body || "Industry Overview")
                : "Industry Overview";
                
            const headerHtml = `
                    <!-- Spatial theme header panel -->
                    <div class="spatial-header-panel">
                        <div style="font-size: 24pt; font-weight: bold; color: var(--title-color); font-family: var(--font-face); flex: 0 0 1.8in; display: flex; align-items: center; gap: 8px;">
                            <span>E</span>
                            <span style="color: var(--text-color); font-size: 20pt; font-weight: normal;">&amp;</span>
                            <span>🧊</span>
                        </div>
                        <div style="flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center;">
                            <div style="font-size: 20pt; font-weight: bold; color: var(--title-color); font-family: var(--font-face); text-transform: uppercase; text-align: center; letter-spacing: 0.5px; line-height: 1.25;">${escapeHtml(data.title || '')}</div>
                            <div style="font-size: 11pt; color: #A0AEC0; font-family: var(--font-face); margin-top: 3px; font-weight: 500; text-align: center;">${escapeHtml(runningSubtitle)}</div>
                        </div>
                        <div style="flex: 0 0 1.8in;"></div>
                    </div>
                    
                    <!-- Cyberpunk theme header panel -->
                    <div class="cyberpunk-header-panel">
                        <div style="font-size: 11pt; font-weight: bold; color: var(--accent-color); font-family: var(--font-family); flex: 0 0 2.2in; display: flex; align-items: center; gap: 8px; letter-spacing: 1px;">
                            <span>[SYS_LOG // V2.1]</span>
                        </div>
                        <div style="flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center;">
                            <div style="font-size: 16pt; font-weight: bold; color: var(--title-color); font-family: var(--font-family); text-transform: uppercase; text-align: center; letter-spacing: 2px; line-height: 1.25; text-shadow: 0 0 10px rgba(0, 240, 255, 0.3);">${escapeHtml(data.title || '')}</div>
                        </div>
                        <div style="flex: 0 0 2.2in; text-align: right; font-family: var(--font-family); font-size: 9pt; color: var(--text-color); opacity: 0.7; letter-spacing: 0.5px;">
                            SEC_0${idx} // ACTIVE
                        </div>
                    </div>

                    <!-- Holodeck theme header panel -->
                    <div class="holodeck-header-panel">
                        <div style="font-size: 11pt; font-weight: bold; color: var(--accent-color); font-family: var(--font-family); flex: 0 0 2.2in; display: flex; align-items: center; gap: 8px; letter-spacing: 1px;">
                            <span>[HOLODECK MATRIX // V1.0]</span>
                        </div>
                        <div style="flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center;">
                            <div style="font-size: 16pt; font-weight: bold; color: var(--title-color); font-family: var(--font-family); text-transform: uppercase; text-align: center; letter-spacing: 2px; line-height: 1.25; text-shadow: 0 0 10px rgba(245, 158, 11, 0.3);">${escapeHtml(data.title || '')}</div>
                        </div>
                        <div style="flex: 0 0 2.2in; text-align: right; font-family: var(--font-family); font-size: 9pt; color: var(--text-color); opacity: 0.7; letter-spacing: 0.5px;">
                            GRID_0${idx} // EMULATED
                        </div>
                    </div>

                    <div class="slide-title-container">
                        <span class="slide-subtitle">${escapeHtml(runningSubtitle.toUpperCase())}</span>
                        <h2 class="slide-title">${escapeHtml(data.title || '')}</h2>
                    </div>
            `;

            content += `                <!-- Generative DSL Layout -->
                <div class="dsl-slide-layout" style="width:100%; height:100%; position:relative;">
                    ${headerHtml}
`;
            
            data.dsl.elements.forEach(el => {
                const x = el.bounds.x;
                const y = el.bounds.y;
                const w = el.bounds.w;
                const h = el.bounds.h;
                
                const posStyle = `position: absolute; left: ${x}; top: ${y}; width: ${w}; height: ${h}; box-sizing: border-box;`;
                
                if (el.type === 'text') {
                    const fontSize = el.style?.fontSize || 12;
                    const fontWeight = el.style?.fontWeight || 'normal';
                    const color = el.style?.color || 'text-main';
                    const align = el.style?.align || 'left';
                    const valign = el.style?.valign || 'top';
                    
                    const colorVal = (color === 'activeCyan' || color === 'activePurple' || color === 'text-main' || color === 'text-muted' || color === 'neon-green' || color === 'neon-orange')
                        ? `var(--${color.replace(/[A-Z]/g, m => '-' + m.toLowerCase())})`
                        : color;
                        
                    const displayStyle = `display: flex; flex-direction: column; justify-content: ${valign === 'middle' ? 'center' : valign === 'bottom' ? 'flex-end' : 'flex-start'}; align-items: ${align === 'center' ? 'center' : align === 'right' ? 'flex-end' : 'flex-start'}; text-align: ${align}; font-size: ${fontSize}pt; font-weight: ${fontWeight}; color: ${colorVal};`;
                    
                    content += `                    <div style="${posStyle} ${displayStyle} overflow: hidden;">
                        ${formatMarkdownTokens(el.content || '')}
                    </div>\n`;
                } else if (el.type === 'card') {
                    const themeClass = el.style?.theme || 'default';
                    const titleColor = el.style?.theme === 'accent' ? 'var(--active-cyan)' : 'var(--title-color)';
                    
                    content += `                    <div class="premium-card cyber-chassis ${themeClass}" style="${posStyle} overflow: hidden;">
                        <span class="tech-tag">${el.id ? el.id.toUpperCase().replace(/_/g, ' ') : 'COMPONENT'}</span>
                        <div class="premium-card-title" style="text-align: left; font-size: 13pt; color: ${titleColor};">${escapeHtml(el.content.title || '')}</div>
                        <div style="margin: 4px 0 6px 0; border-bottom: 1px solid var(--card-border-cyan); width: 100%;"></div>
                        <div class="premium-card-body" style="font-size: 10pt; line-height: 1.35; overflow: hidden; color: var(--text-main);">${formatMarkdownTokens(el.content.body || '')}</div>
                    </div>\n`;
                } else if (el.type === 'image') {
                    content += `                    <div class="image-card-frame cyber-chassis" style="${posStyle}">
                        <img src="${el.content.path}" style="width: 100%; height: 100%; object-fit: contain;">
                    </div>\n`;
                } else if (el.type === 'vector') {
                    if (el.name === 'radar-sweep') {
                        content += `                    <div class="radar-box" style="${posStyle}">
                            <svg class="radar-svg" viewBox="0 0 200 200">
                                <circle cx="100" cy="100" r="80" fill="none" stroke="rgba(var(--active-cyan-rgb), 0.1)" stroke-width="1" />
                                <circle cx="100" cy="100" r="60" fill="none" stroke="rgba(var(--active-cyan-rgb), 0.15)" stroke-width="1" />
                                <circle cx="100" cy="100" r="40" fill="none" stroke="rgba(var(--active-cyan-rgb), 0.15)" stroke-width="1" />
                                <circle cx="100" cy="100" r="20" fill="none" stroke="rgba(var(--active-cyan-rgb), 0.2)" stroke-width="1" stroke-dasharray="2 2"/>
                                <line x1="100" y1="100" x2="156" y2="44" stroke="var(--active-purple)" stroke-width="1.5" />
                                <path d="M100,100 L100,20 A80,80 0 0,1 156,44 Z" fill="rgba(var(--active-cyan-rgb), 0.15)" />
                            </svg>
                        </div>\n`;
                    } else if (el.name === 'spatial-sphere') {
                        content += `                    <div class="spatial-sphere-container" style="${posStyle}">
                            <div class="spatial-sphere">
                                <svg class="sphere-svg" viewBox="0 0 100 100">
                                    <circle cx="50" cy="50" r="42" fill="none" stroke="var(--active-cyan)" stroke-width="1" />
                                    <ellipse cx="50" cy="50" rx="42" ry="16" fill="none" stroke="var(--active-cyan)" stroke-width="0.8" stroke-dasharray="2 2" />
                                    <ellipse cx="50" cy="50" rx="16" ry="42" fill="none" stroke="var(--active-cyan)" stroke-width="0.8" stroke-dasharray="2 2" />
                                    <circle cx="50" cy="50" r="2.5" fill="var(--active-purple)" />
                                </svg>
                            </div>
                        </div>\n`;
                    } else if (el.name === 'viewfinder-reticle') {
                        content += `                    <div class="viewport-shutter-flash"></div>
                        <div class="viewfinder-screen" style="${posStyle} border: 1.5px solid var(--card-border-cyan); border-radius: 6px; overflow: hidden; display: flex; align-items: center; justify-content: center; cursor: pointer;">
                            <div class="viewfinder-hud-reticle" style="position: relative;"></div>
                        </div>\n`;
                    } else if (el.name === 'timeline-pulse') {
                        content += `                    <div class="timeline-neon-line" style="${posStyle}">
                            <div class="timeline-signal-pulse"></div>
                        </div>\n`;
                    } else if (el.name === 'multiview-diagnostic') {
                        content += `                    <div class="spatial-multiview-container" style="${posStyle}">
                            <span style="font-family: var(--font-display); font-size: 9px; color: var(--active-cyan); margin-bottom: 6px; letter-spacing: 1px; text-align: center; display: block;">3D SCAN RECONSTRUCTION</span>
                            <div class="multiview-grids" style="display: flex; gap: 10px; width: 100%; justify-content: center;">
                                <div class="multiview-box" style="flex: 1; background: rgba(0, 0, 0, 0.3); border: 1px solid rgba(var(--active-cyan-rgb), 0.2); border-radius: 4px; position: relative; height: 110px; display: flex; align-items: center; justify-content: center;">
                                    <span class="box-tag" style="position: absolute; top: 4px; left: 6px; font-family: var(--font-mono); font-size: 7px; color: var(--text-muted);">TOP_V</span>
                                    <svg class="box-svg" viewBox="0 0 100 100" style="width: 70px; height: 70px;">
                                        <line x1="10" y1="50" x2="90" y2="50" stroke="rgba(20,184,166,0.2)" stroke-width="0.5" />
                                        <line x1="50" y1="10" x2="50" y2="90" stroke="rgba(20,184,166,0.2)" stroke-width="0.5" />
                                        <circle cx="50" cy="50" r="30" fill="none" stroke="var(--active-cyan)" stroke-width="1.2" stroke-dasharray="4 2" />
                                        <circle cx="50" cy="50" r="10" fill="none" stroke="var(--active-purple)" stroke-width="1" />
                                        <line x1="50" y1="50" x2="71" y2="29" stroke="var(--active-cyan)" stroke-width="1.5" />
                                    </svg>
                                </div>
                                <div class="multiview-box" style="flex: 1; background: rgba(0, 0, 0, 0.3); border: 1px solid rgba(var(--active-cyan-rgb), 0.2); border-radius: 4px; position: relative; height: 110px; display: flex; align-items: center; justify-content: center;">
                                    <span class="box-tag" style="position: absolute; top: 4px; left: 6px; font-family: var(--font-mono); font-size: 7px; color: var(--text-muted);">SIDE_V</span>
                                    <svg class="box-svg" viewBox="0 0 100 100" style="width: 70px; height: 70px;">
                                        <line x1="10" y1="80" x2="90" y2="80" stroke="var(--active-cyan)" stroke-width="1.2" />
                                        <rect x="25" y="45" width="50" height="35" fill="rgba(20,184,166,0.05)" stroke="var(--active-cyan)" stroke-width="1" />
                                        <circle cx="35" cy="80" r="8" fill="none" stroke="var(--active-purple)" stroke-width="1" />
                                        <circle cx="65" cy="80" r="8" fill="none" stroke="var(--active-purple)" stroke-width="1" />
                                    </svg>
                                </div>
                            </div>
                        </div>\n`;
                    }
                } else if (el.type === 'timeline') {
                    const numCards = el.content.length;
                    const gap = 3;
                    
                    let timelineHtml = `                    <div style="${posStyle} display: flex; gap: ${gap}%;">`;
                    el.content.forEach((item, itemIdx) => {
                        const bodyHtml = segmentsToHtml(processLines(item.body, 11));
                        timelineHtml += `
                        <div class="premium-card cyber-chassis" style="flex: 1; display: flex; flex-direction: column; gap: 0.1in; height: 100%; box-sizing: border-box; padding: 20px 15px;">
                            <div style="font-family: var(--font-display); font-size: 16pt; font-weight: bold; color: var(--active-cyan);">0${itemIdx + 1}</div>
                            <div class="premium-card-title" style="text-align: left; font-size: 11pt; border-bottom: none; color: var(--title-color);">${escapeHtml(item.title || '')}</div>
                            <div style="flex: 1; overflow: hidden; font-size: 9.5pt; text-align: left; width: 100%; color: var(--text-main); margin-top: 5px;">${bodyHtml}</div>
                        </div>`;
                    });
                    timelineHtml += `                    </div>\n`;
                    content += timelineHtml;
                }
            });
            content += `                </div>`;
        } else {
            // 1. Cover Slide layout (First slide)
            if (idx === 0) {
                content += `                <div style="position: absolute; left: 0in; top: 0in; width: 4.0in; height: 7.5in; background-color: var(--accent-color); display: flex; align-items: center; justify-content: center;">
                        <div style="width: 3.0in; height: 6.5in; border: 1.5px solid rgba(255,255,255,0.4); border-radius: 4px;"></div>
                    </div>\n`;
                if (data.title) {
                    content += `                <div style="position: absolute; left: 4.8in; top: 2.0in; width: 7.5in; height: 2.2in; display: flex; align-items: flex-end; text-align: left; font-size: 36pt; font-weight: bold; color: var(--title-color); line-height: 1.25;">
                        ${escapeHtml(data.title)}
                    </div>\n`;
                }
                if (data.text) {
                    const formattedText = processLines(data.text, 18);
                    formattedText.forEach(seg => {
                        seg.options.align = "left";
                    });
                    const formattedHtml = segmentsToHtml(formattedText);
                    content += `                <div style="position: absolute; left: 4.8in; top: 4.5in; width: 7.5in; height: 2.0in; text-align: left;">
                        ${formattedHtml}
                    </div>\n`;
                }
            } else {
                // 2. Normal Content Slide layout
                const mainTitle = globalSlidesData[0] ? (globalSlidesData[0].title || 'PRESENTATION') : 'PRESENTATION';
                const runningSubtitle = (globalSlidesData[0] && globalSlidesData[0].text) 
                    ? globalSlidesData[0].text.split('\n')[0].trim() 
                    : "Industry Overview";
                const footerTitle = globalSlidesData[0] ? (globalSlidesData[0].title || 'EMBODIED AI & SPATIAL INTELLIGENCE') : 'EMBODIED AI & SPATIAL INTELLIGENCE';
                const totalContentSlides = globalSlidesData.length - 1;
                const totalContentSlidesStr = totalContentSlides < 10 ? '0' + totalContentSlides : totalContentSlides;

                content += `                <!-- Spatial theme background decorations -->
                <div class="spatial-decorations">
                    <svg style="position: absolute; left: 0; top: 0; width: 13.333in; height: 7.5in; pointer-events: none;">
                        <!-- Top left circuit line -->
                        <line x1="0.2in" y1="0.4in" x2="0.7in" y2="0.4in" stroke="#14B8A6" stroke-width="1.5" />
                        <line x1="0.7in" y1="0.4in" x2="0.9in" y2="0.6in" stroke="#14B8A6" stroke-width="1.5" />
                        <line x1="0.9in" y1="0.6in" x2="1.4in" y2="0.6in" stroke="#14B8A6" stroke-width="1.5" />
                        <!-- Top right horizontal line -->
                        <line x1="12.63in" y1="0.4in" x2="13.13in" y2="0.4in" stroke="#14B8A6" stroke-width="1.5" />
                        <!-- Bottom divider line -->
                        <line x1="0.8in" y1="7.0in" x2="12.5in" y2="7.0in" stroke="#163D3F" stroke-width="1.5" />
                    </svg>
                    <div style="position: absolute; left: 0.8in; top: 7.0in; width: 11.7in; height: 0.4in; display: flex; align-items: center; justify-content: center; font-size: 8pt; font-weight: bold; color: #5F7D81; font-family: var(--font-family); letter-spacing: 1px;">
                        ${escapeHtml(footerTitle.toUpperCase())} | INDUSTRY TRENDS | SLIDE ${idx < 10 ? '0' + idx : idx}/${totalContentSlidesStr} | 2024
                    </div>
                </div>

                <div class="slide-meta" style="position: absolute; left: 0.8in; top: 0.15in; width: 6.0in; height: 0.25in; font-size: 9pt; font-weight: bold; color: var(--text-color); opacity: 0.6; display: flex; align-items: center; font-family: var(--font-face);">
                    SLIDE ${idx < 10 ? '0' + idx : idx} | ${mainTitle.toUpperCase()}
                </div>
                <div style="position: absolute; right: 0.8in; top: 0.23in; display: flex; gap: 6px;">
                    <div style="width: 8px; height: 8px; border-radius: 50%; background-color: #FF5F56;"></div>
                    <div style="width: 8px; height: 8px; border-radius: 50%; background-color: #FFBD2E;"></div>
                    <div style="width: 8px; height: 8px; border-radius: 50%; background-color: #27C93F;"></div>
                </div>\n`;
                
                if (data.title) {
                    content += `                <!-- Spatial theme header panel -->
                    <div class="spatial-header-panel">
                        <div style="font-size: 24pt; font-weight: bold; color: var(--title-color); font-family: var(--font-face); flex: 0 0 1.8in; display: flex; align-items: center; gap: 8px;">
                            <span>E</span>
                            <span style="color: var(--text-color); font-size: 20pt; font-weight: normal;">&amp;</span>
                            <span>🧊</span>
                        </div>
                        <div style="flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center;">
                            <div style="font-size: 20pt; font-weight: bold; color: var(--title-color); font-family: var(--font-face); text-transform: uppercase; text-align: center; letter-spacing: 0.5px; line-height: 1.25;">${escapeHtml(data.title)}</div>
                            <div style="font-size: 11pt; color: #A0AEC0; font-family: var(--font-face); margin-top: 3px; font-weight: 500; text-align: center;">${escapeHtml(runningSubtitle)}</div>
                        </div>
                        <div style="flex: 0 0 1.8in;"></div>
                    </div>

                    <div class="title-indicator" style="position: absolute; left: 0.8in; top: 0.55in; width: 0.12in; height: 0.5in; background-color: var(--accent-color); border-radius: 2px;"></div>\n`;
                    content += `                <div class="slide-title" style="position: absolute; left: 1.05in; top: 0.4in; width: 11.45in; height: 0.8in; display: flex; align-items: center; font-size: 28pt; font-weight: bold; color: var(--title-color); font-family: var(--font-face);">
                        ${escapeHtml(data.title)}
                    </div>
                    <div class="slide-subtitle" style="position: absolute; left: 0.8in; top: 0.9in; width: 11.7in; height: 0.3in; display: none; align-items: center; justify-content: center; font-size: 13pt; color: #A0AEC0; font-family: var(--font-face);">
                        ${escapeHtml(runningSubtitle)}
                    </div>
                    <div class="slide-divider" style="position: absolute; left: 0.8in; top: 1.25in; width: 11.7in; border-bottom: 1.5px solid var(--divider-color);"></div>\n`;
                }
                
                const hasImage = data.images.length > 0;
                const layout = getSlideLayoutInfo(data, hasImage);
                
                if (layout.type === 'asymmetric') {
                    const dynamicFontSize = calculateDynamicFontSize(data.text, true);
                    const formattedHtml = segmentsToHtml(processLines(data.text, dynamicFontSize));
                    
                    if (formattedHtml) {
                        content += `                <div class="text-card">
                            ${formattedHtml}
                        </div>\n`;
                    }
                    
                    let imgObj = data.images[0];
                    let imgPath = imgObj.path;
                    let imgAlt = imgObj.alt || "";
                    if (!path.isAbsolute(imgPath)) {
                        imgPath = path.resolve(path.dirname(inputFile), imgPath);
                    }
                    
                    if (fs.existsSync(imgPath)) {
                        const maxW = 5.7; // inches
                        const maxH = 4.8; // inches
                        const targetX = 6.8;
                        
                        let imgW = maxW;
                        let imgH = maxH;
                        
                        const dims = getImageDimensions(imgPath);
                        if (dims) {
                            const aspect = dims.width / dims.height;
                            const targetAspect = maxW / maxH;
                            if (aspect > targetAspect) {
                                imgW = maxW;
                                imgH = maxW / aspect;
                            } else {
                                imgH = maxH;
                                imgW = maxH * aspect;
                            }
                        }
                        
                        const targetY = 1.6 + (4.8 - imgH) / 2;
                        
                        let fileUrl = 'file://' + imgPath.replace(/\\/g, '/');
                        fileUrl = fileUrl.replace(/file:\/\/\/mnt\/([a-zA-Z])\//, 'file:///$1:/');
                        content += `                <div class="image-card-frame" style="left: ${targetX}in; top: ${targetY}in; width: ${imgW}in; height: ${imgH}in;">
                            <img src="${fileUrl}" style="width: 100%; height: 100%; object-fit: contain;">
                        </div>\n`;
                        
                        const captionH = 6.4 - (targetY + imgH) - 0.15;
                        if (captionH >= 0.8 && imgAlt) {
                            content += `                <div class="image-card-frame" style="left: ${targetX}in; top: ${targetY + imgH + 0.15}in; width: ${imgW}in; height: ${captionH}in; display: flex; align-items: center; justify-content: center; background-color: var(--card-bg); border: 1.5px solid var(--card-border); border-top: none; padding: 0.1in;">
                                <div style="font-size: 11pt; font-style: italic; text-align: center; color: var(--text-color); opacity: 0.8; line-height: 1.3;">${escapeHtml(imgAlt)}</div>
                            </div>\n`;
                        }
                    }
                } else if (layout.type === 'timeline') {
                    const { introLines, cardItems } = layout;
                    if (introLines.length > 0) {
                        const introText = introLines.join('\n');
                        const formattedHtml = segmentsToHtml(processLines(introText, 16));
                        content += `                <div style="position: absolute; left: 0.8in; top: 1.35in; width: 11.7in; height: 0.5in; display: flex; align-items: center;">
                            ${formattedHtml}
                        </div>\n`;
                    }
                    
                    content += `                <div style="position: absolute; left: 1.0in; top: 2.2in; width: 11.3in; height: 2px; border-bottom: 2px dashed var(--accent-color); opacity: 0.6;"></div>\n`;
                    
                    const numCards = cardItems.length;
                    const totalW = 11.7;
                    const gap = 0.4;
                    const cardW = (totalW - (numCards - 1) * gap) / numCards;
                    const startX = 0.8;
                    
                    cardItems.forEach((item, cardIdx) => {
                        const cardX = startX + cardIdx * (cardW + gap);
                        const cx = cardX + cardW / 2;
                        
                        content += `                <div style="position: absolute; left: ${cx - 0.15}in; top: 2.05in; width: 0.3in; height: 0.3in; background-color: var(--accent-color); border-radius: 50%; box-sizing: border-box; display: flex; align-items: center; justify-content: center; border: 3px solid var(--bg-color);"></div>\n`;
                        content += `                <div style="position: absolute; left: ${cx - 0.5}in; top: 1.4in; width: 1.0in; height: 0.5in; display: flex; align-items: flex-end; justify-content: center; font-size: 18pt; font-weight: bold; color: var(--accent-color); font-family: var(--font-family);">0\Ref{0}${cardIdx + 1}</div>\n`;
                        
                        const bodyHtml = segmentsToHtml(processLines(item.body, 11));
                        content += `                <div class="premium-card" style="position: absolute; left: ${cardX}in; top: 2.6in; width: ${cardW}in; height: 3.8in; padding: 0.2in 0.15in; display: flex; flex-direction: column; gap: 0.1in; align-items: center; text-align: center;">
                            <div class="premium-card-title" style="font-size: 13pt; text-align: center; border-bottom: none;">${escapeHtml(item.title)}</div>
                            <div style="margin-top: 5px; flex: 1; overflow: hidden; font-size: 11pt; text-align: left; width: 100%;">${bodyHtml}</div>
                        </div>\n`;
                    });
                } else if (layout.type === 'grid') {
                    const { introLines, cardItems } = layout;
                    if (introLines.length > 0) {
                        const introText = introLines.join('\n');
                        const formattedHtml = segmentsToHtml(processLines(introText, 16));
                        content += `                <div style="position: absolute; left: 0.8in; top: 1.5in; width: 11.7in; height: 0.6in; display: flex; align-items: center;">
                            ${formattedHtml}
                        </div>\n`;
                    }
                    
                    const cardY = introLines.length > 0 ? 2.3 : 1.6;
                    const cardH = introLines.length > 0 ? 4.1 : 4.8;
                    
                    const numCards = cardItems.length;
                    if (numCards === 4) {
                        content += `                <div class="grid-container-2x2" style="position: absolute; left: 0.8in; top: ${cardY}in; width: 11.7in; height: ${cardH}in; display: grid; grid-template-columns: 1fr 1fr; grid-template-rows: 1fr 1fr; gap: 0.3in;">\n`;
                        cardItems.forEach((item, cardIdx) => {
                            const bodyHtml = segmentsToHtml(processLines(item.body, 10));
                            const icon = getIconForTitle(item.title);
                            content += `                    <div class="premium-card" style="display: flex; flex-direction: row; align-items: center; gap: 0.2in; padding: 0.2in; height: auto;">
                                <div class="premium-card-icon" style="font-size: 28pt; flex: 0 0 0.8in; text-align: center; line-height: 1; display: flex; align-items: center; justify-content: center;">
                                    <span class="emoji-icon">${icon}</span>
                                    <img class="image-icon" src="assets/spatial_icon_${cardIdx + 1}.svg" style="display: none; width: 100%; height: 100%; object-fit: contain;">
                                </div>
                                <div style="flex: 1; display: flex; flex-direction: column; text-align: left; gap: 0.02in; overflow: hidden;">
                                    <div class="premium-card-title" style="text-align: left; font-size: 13pt;">${escapeHtml(item.title)}</div>
                                    <div style="margin: 4px 0 6px 0; border-bottom: 1px solid var(--divider-color); width: 100%;"></div>
                                    <div class="premium-card-body" style="font-size: 10pt; line-height: 1.35;">${bodyHtml}</div>
                                </div>
                            </div>\n`;
                        });
                        content += `                </div>\n`;
                    } else {
                        const gap = 0.4;
                        const cardW = (11.7 - (numCards - 1) * gap) / numCards;
                        content += `                <div style="position: absolute; left: 0.8in; top: \Ref{cardY}in; width: 11.7in; height: ${cardH}in; display: flex; gap: ${gap}in;">\n`;
                        cardItems.forEach(item => {
                            const bodyHtml = segmentsToHtml(processLines(item.body, 11));
                            const icon = getIconForTitle(item.title);
                            content += `                    <div class="premium-card" style="flex: 0 0 ${cardW}in; height: ${cardH}in; display: flex; flex-direction: column; align-items: center; text-align: center; gap: 0.1in;">
                                <div style="font-size: 24pt; margin-top: 10px; line-height: 1;">\Ref{icon}</div>
                                <div class="premium-card-title">${escapeHtml(item.title)}</div>
                                <div class="premium-card-body" style="font-size: 11pt; width: 100%; text-align: left; margin-top: 5px;">${bodyHtml}</div>
                            </div>\n`;
                        });
                        content += `                </div>\n`;
                    }
                } else if (layout.type === 'centered-breathe') {
                    const dynamicFontSize = calculateDynamicFontSize(data.text, false) + 4;
                    const formattedHtml = segmentsToHtml(processLines(data.text, dynamicFontSize));
                    content += `                <div style="position: absolute; left: 3.16in; top: 2.2in; width: 7.0in; height: 3.6in; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; background-color: var(--card-bg); border: 1.5px solid var(--card-border); border-top: 4px solid var(--accent-color); border-radius: 6px; padding: 0.35in; box-sizing: border-box; overflow: hidden; line-height: 1.5;">
                        <div style="width: 100%; text-align: center;">${formattedHtml}</div>
                    </div>\n`;
                } else {
                    const dynamicFontSize = calculateDynamicFontSize(data.text, false);
                    const formattedHtml = segmentsToHtml(processLines(data.text, dynamicFontSize));
                    content += `                <div class="centered-card">
                        ${formattedHtml}
                    </div>\n`;
                }
            }
        }
        content += `            </div>
            <div class="slide-num">Slide ${idx + 1}</div>
        </div>\n`;
    });
    return content;
}

function generateHtmlPreview(slidesData, pptxPath, themeMode = "all") {

    const htmlPath = pptxPath.replace(/\.pptx$/i, '.html');
    const basePptxName = path.basename(pptxPath);
    const baseNameWithoutExt = basePptxName.replace(/\.pptx$/i, '');
    
    let htmlContent = `<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Presentation Preview: ${basePptxName}</title>
    <style>
        :root {
            --bg-color: #f8f9fa;
            --card-bg: #ffffff;
            --card-border: #E2E8F0;
            --title-color: #0F172A;
            --text-color: #334155;
            --accent-color: #2563EB;
            --formula-color: #2563EB;
            --divider-color: #E2E8F0;
            --font-family: "Segoe UI", -apple-system, sans-serif;
        }
        body.theme-light {
            --bg-color: #f8f9fa;
            --card-bg: #ffffff;
            --card-border: #E2E8F0;
            --title-color: #0F172A;
            --text-color: #334155;
            --accent-color: #2563EB;
            --formula-color: #2563EB;
            --divider-color: #E2E8F0;
            --font-family: "Segoe UI", -apple-system, sans-serif;
        }
        body.theme-dark {
            --bg-color: #0F172A;
            --card-bg: #1E293B;
            --card-border: #334155;
            --title-color: #38BDF8;
            --text-color: #E2E8F0;
            --accent-color: #38BDF8;
            --formula-color: #0ea5e9;
            --divider-color: #334155;
            --font-family: "Segoe UI", "Trebuchet MS", -apple-system, sans-serif;
        }
        body.theme-warm {
            --bg-color: #FAF6F0;
            --card-bg: #FFFDF9;
            --card-border: #E6DFD3;
            --title-color: #3F2E2C;
            --text-color: #4A3E3D;
            --accent-color: #D97706;
            --formula-color: #B45309;
            --divider-color: #E6DFD3;
            --font-family: "Georgia", serif;
        }
        body.theme-aurora {
            --bg-color: #FAF5FF;
            --card-bg: #FFFFFF;
            --card-border: #E9D5FF;
            --title-color: #7C3AED;
            --text-color: #4B5563;
            --accent-color: #D946EF;
            --formula-color: #C084FC;
            --divider-color: #E9D5FF;
            --font-family: "Trebuchet MS", "Arial", sans-serif;
        }
        body.theme-forest {
            --bg-color: #F4F7F5;
            --card-bg: #FFFFFF;
            --card-border: #D1DDD4;
            --title-color: #1C3F24;
            --text-color: #2F3E32;
            --accent-color: #10B981;
            --formula-color: #059669;
            --divider-color: #D1DDD4;
            --font-family: "Segoe UI", -apple-system, sans-serif;
        }
        body.theme-ocean {
            --bg-color: #F0F7FF;
            --card-bg: #FFFFFF;
            --card-border: #C7D2FE;
            --title-color: #1E3A8A;
            --text-color: #374151;
            --accent-color: #3B82F6;
            --formula-color: #2563EB;
            --divider-color: #C7D2FE;
            --font-family: "Calibri", "Segoe UI", -apple-system, sans-serif;
        }
        body.theme-spatial {
            --bg-color: #060E11;
            --card-bg: #0E1A1E;
            --card-border: #163D3F;
            --title-color: #2DD4BF;
            --text-color: #D1E2E4;
            --accent-color: #14B8A6;
            --formula-color: #2DD4BF;
            --divider-color: #163D3F;
            --font-family: "Trebuchet MS", "Segoe UI", -apple-system, sans-serif;
        }
        .theme-spatial .title-indicator {
            display: none !important;
        }
        .theme-spatial .slide-title {
            display: none !important;
        }
        .theme-spatial .slide-divider {
            display: none !important;
        }
        .theme-spatial .slide-subtitle {
            display: none !important;
        }
        .theme-spatial .slide-meta {
            display: none !important;
        }
        .theme-spatial .text-card {
            border-top: 1.5px solid var(--card-border) !important;
        }
        .spatial-decorations {
            display: none;
        }
        .theme-spatial .spatial-decorations {
            display: block !important;
        }
        .spatial-header-panel {
            display: none;
        }
        .theme-spatial .spatial-header-panel {
            display: flex !important;
            position: absolute;
            left: 0.4in;
            top: 0.25in;
            width: 12.53in;
            height: 1.1in;
            background-color: var(--card-bg);
            border: 1.5px solid var(--card-border);
            border-radius: 8px;
            box-sizing: border-box;
            align-items: center;
            padding: 0 0.3in;
        }
        .theme-spatial .grid-container-2x2 {
            left: 0.4in !important;
            top: 1.6in !important;
            width: 12.53in !important;
            height: 5.25in !important;
            gap: 0.35in 0.53in !important;
        }
        .theme-spatial .grid-container-2x2 .premium-card {
            padding: 0.25in 0.2in !important;
            border-radius: 8px !important;
            display: flex !important;
            flex-direction: row !important;
            align-items: center !important;
            gap: 0.2in !important;
            height: auto !important;
        }
        .theme-spatial .grid-container-2x2 .premium-card-icon {
            font-size: 44pt !important;
            flex: 0 0 1.2in !important;
            text-align: center !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            width: 1.2in !important;
            height: 1.2in !important;
        }
        .theme-spatial .slide-frame {
            background-image: url('assets/spatial_bg.jpg') !important;
            background-size: cover !important;
            background-position: center !important;
        }
        body.theme-cyberpunk {
            --bg-color: #030611;
            --card-bg: #060B1A;
            --card-border: #162C3F;
            --title-color: #00F0FF;
            --text-color: #F1F5F9;
            --accent-color: #BC39FA;
            --formula-color: #00F0FF;
            --divider-color: #00F0FF;
            --font-family: "Share Tech Mono", "Courier New", monospace;
        }
        .theme-cyberpunk .title-indicator {
            display: none !important;
        }
        .theme-cyberpunk .slide-title {
            display: none !important;
        }
        .theme-cyberpunk .slide-divider {
            display: none !important;
        }
        .theme-cyberpunk .slide-subtitle {
            display: none !important;
        }
        .theme-cyberpunk .slide-meta {
            display: none !important;
        }
        .theme-cyberpunk .slide-frame {
            background-image: url('assets/cyberpunk_bg.jpg') !important;
            background-size: cover !important;
            background-position: center !important;
        }
        .cyberpunk-header-panel {
            display: none;
        }
        .theme-cyberpunk .cyberpunk-header-panel {
            display: flex !important;
            position: absolute;
            left: 0.4in;
            top: 0.25in;
            width: 12.53in;
            height: 0.8in;
            background-color: var(--card-bg);
            border: 1.5px solid var(--card-border);
            border-left: 4px solid var(--accent-color);
            border-right: 4px solid var(--accent-color);
            box-sizing: border-box;
            align-items: center;
            padding: 0 0.3in;
        }
        body.theme-holodeck {
            --bg-color: #080806;
            --card-bg: #12110E;
            --card-border: #2A2215;
            --title-color: #F59E0B;
            --text-color: #E5E5E5;
            --accent-color: #F59E0B;
            --formula-color: #FBBF24;
            --divider-color: #2A2215;
            --font-family: "Share Tech Mono", "Courier New", monospace;
        }
        .theme-holodeck .title-indicator {
            display: none !important;
        }
        .theme-holodeck .slide-title {
            display: none !important;
        }
        .theme-holodeck .slide-divider {
            display: none !important;
        }
        .theme-holodeck .slide-subtitle {
            display: none !important;
        }
        .theme-holodeck .slide-meta {
            display: none !important;
        }
        .theme-holodeck .slide-frame {
            background-image: url('assets/holodeck_bg.jpg') !important;
            background-size: cover !important;
            background-position: center !important;
        }
        .holodeck-header-panel {
            display: none;
        }
        .theme-holodeck .holodeck-header-panel {
            display: flex !important;
            position: absolute;
            left: 0.4in;
            top: 0.25in;
            width: 12.53in;
            height: 0.8in;
            background-color: var(--card-bg);
            border: 1.5px solid var(--card-border);
            border-top: 4px solid var(--accent-color);
            border-bottom: 4px solid var(--accent-color);
            box-sizing: border-box;
            align-items: center;
            padding: 0 0.3in;
        }
        body.theme-ghibli {
            --bg-color: #EAF3E1;
            --card-bg: #FBF7EE;
            --card-border: #C9D9C0;
            --title-color: #2F5D3A;
            --text-color: #4A3E2D;
            --accent-color: #D9A441;
            --formula-color: #B45309;
            --divider-color: #C9D9C0;
            --font-family: "Georgia", "Segoe UI", serif;
        }
        .theme-ghibli .title-indicator { display: none !important; }
        .theme-ghibli .slide-meta { display: none !important; }
        .theme-ghibli .slide-frame {
            background-image: url('assets/ghibli_bg.jpg') !important;
            background-size: cover !important;
            background-position: center !important;
        }
        .theme-ghibli .slide-title-container {
            position: absolute; left: 0.8in; top: 0.3in; width: 11.7in; height: 0.95in;
            display: flex; flex-direction: column; align-items: center; justify-content: center;
            text-align: center;
        }
        .theme-ghibli .slide-subtitle {
            font-size: 11pt; color: var(--accent-color); font-family: var(--font-family);
            font-style: italic; margin: 0 0 4px 0; letter-spacing: 1px;
        }
        .theme-ghibli .slide-title {
            font-size: 26pt; font-weight: bold; color: var(--title-color);
            font-family: var(--font-family); margin: 0; line-height: 1.2;
        }
        .theme-ghibli .premium-card {
            box-shadow: 0 6px 18px rgba(80, 60, 30, 0.10);
        }
        .premium-card-icon .emoji-icon {
            display: inline-block;
        }
        .premium-card-icon .image-icon {
            display: none;
        }
        .theme-spatial .premium-card-icon .emoji-icon {
            display: none !important;
        }
        .theme-spatial .premium-card-icon .image-icon {
            display: block !important;
            width: 100% !important;
            height: 100% !important;
            object-fit: contain !important;
        }
        
        body {
            background-color: #202124;
            color: #e8eaed;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            margin: 0;
            padding: 80px 20px 20px 20px;
            display: flex;
            flex-direction: column;
            align-items: center;
            transition: background-color 0.3s;
        }
        h1 {
            font-size: 24px;
            margin-bottom: 20px;
        }
        .slides-container {
            display: flex;
            flex-direction: column;
            gap: 40px;
            width: 13.333in;
        }
        .slide-frame {
            background-color: var(--bg-color);
            color: var(--text-color);
            font-family: var(--font-family);
            width: 13.333in;
            height: 7.5in;
            position: relative;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.5);
            border-radius: 4px;
            overflow: hidden;
            box-sizing: border-box;
            transition: background-color 0.3s, color 0.3s;
        }
        .slide-num {
            color: #9aa0a6;
            font-size: 14px;
            margin-top: 8px;
            align-self: flex-start;
        }
        .slide-item {
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        .text-card {
            background-color: var(--card-bg);
            border: 1.5px solid var(--card-border);
            border-top: 4px solid var(--accent-color);
            border-radius: 6px;
            position: absolute;
            left: 0.8in;
            top: 1.6in;
            width: 5.2in;
            height: 4.8in;
            padding: 0.3in;
            box-sizing: border-box;
            overflow: hidden;
            transition: background-color 0.3s, border-color 0.3s;
        }
        .image-card-frame {
            background-color: var(--card-bg);
            border: 1.5px solid var(--card-border);
            border-radius: 6px;
            position: absolute;
            box-sizing: border-box;
            padding: 0.08in;
            overflow: hidden;
            transition: background-color 0.3s, border-color 0.3s;
        }
        .card-container {
            display: flex;
            gap: 0.4in;
            width: 11.7in;
            position: absolute;
            left: 0.8in;
            top: 1.6in;
            height: 4.8in;
        }
        .premium-card {
            background-color: var(--card-bg);
            border: 1.5px solid var(--card-border);
            border-radius: 8px;
            flex: 1;
            padding: 0.25in 0.2in;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
            gap: 0.1in;
            height: 4.8in;
            overflow: hidden;
            transition: background-color 0.3s, border-color 0.3s;
            box-shadow: 0 4px 12px rgba(0,0,0,0.02);
        }
        .premium-card-title {
            font-size: 14pt;
            font-weight: bold;
            color: var(--title-color);
            text-align: center;
            margin: 0;
            transition: color 0.3s;
        }
        .premium-card-body {
            margin: 0;
            flex: 1;
            overflow: hidden;
        }
        .centered-card {
            background-color: var(--card-bg);
            border: 1.5px solid var(--card-border);
            border-radius: 8px;
            position: absolute;
            left: 1.91in;
            top: 1.6in;
            width: 9.5in;
            height: 4.8in;
            padding: 0.3in;
            box-sizing: border-box;
            overflow: hidden;
            transition: background-color 0.3s, border-color 0.3s;
        }
        
        /* Floating Toolbar */
        .toolbar {
            position: fixed;
            top: 15px;
            background: rgba(30, 41, 59, 0.85);
            backdrop-filter: blur(8px);
            border: 1px solid rgba(255, 255, 255, 0.1);
            padding: 10px 20px;
            border-radius: 30px;
            display: flex;
            gap: 15px;
            align-items: center;
            z-index: 1000;
            box-shadow: 0 10px 25px rgba(0,0,0,0.3);
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
            background: rgba(255, 255, 255, 0.1);
            color: #E2E8F0;
        }
        .theme-btn:hover {
            background: rgba(255, 255, 255, 0.2);
        }
        .theme-btn.active {
            background: #38BDF8;
            color: #0F172A;
            box-shadow: 0 0 10px rgba(56, 189, 248, 0.5);
        }
    </style>
</head>
<body>`;
    htmlContent += `    <h1>Presentation Preview: ${escapeHtml(basePptxName)}</h1>`;
    
    // Check if we have spatial, cyberpunk and holodeck decks
    const inputDir = path.dirname(inputFile);
    let spatialSlides = null;
    let cyberpunkSlides = null;
    let holodeckSlides = null;
    
    const spatialPath = path.resolve(inputDir, 'spatial_deck.json');
    const cyberpunkPath = path.resolve(inputDir, 'cyberpunk_deck.json');
    const holodeckPath = path.resolve(inputDir, 'holodeck_deck.json');
    
    if (fs.existsSync(spatialPath)) {
        try {
            spatialSlides = JSON.parse(fs.readFileSync(spatialPath, 'utf-8')).slides;
            spatialSlides.forEach(s => { s.isDsl = true; s.dsl = s; });
        } catch(e) {}
    }
    if (fs.existsSync(cyberpunkPath)) {
        try {
            cyberpunkSlides = JSON.parse(fs.readFileSync(cyberpunkPath, 'utf-8')).slides;
            cyberpunkSlides.forEach(s => { s.isDsl = true; s.dsl = s; });
        } catch(e) {}
    }
    if (fs.existsSync(holodeckPath)) {
        try {
            holodeckSlides = JSON.parse(fs.readFileSync(holodeckPath, 'utf-8')).slides;
            holodeckSlides.forEach(s => { s.isDsl = true; s.dsl = s; });
        } catch(e) {}
    }

    if (spatialSlides || cyberpunkSlides || holodeckSlides) {
        if (spatialSlides) {
            htmlContent += `    <div class="slides-container theme-layout-spatial" style="display: flex; flex-direction: column; gap: 40px;">\n`;
            htmlContent += renderSlideListHtml(spatialSlides, slidesData);
            htmlContent += `    </div>\n`;
        }
        if (cyberpunkSlides) {
            htmlContent += `    <div class="slides-container theme-layout-cyberpunk" style="display: none; flex-direction: column; gap: 40px;">\n`;
            htmlContent += renderSlideListHtml(cyberpunkSlides, slidesData);
            htmlContent += `    </div>\n`;
        }
        if (holodeckSlides) {
            htmlContent += `    <div class="slides-container theme-layout-holodeck" style="display: none; flex-direction: column; gap: 40px;">\n`;
            htmlContent += renderSlideListHtml(holodeckSlides, slidesData);
            htmlContent += `    </div>\n`;
        }
    } else {
        htmlContent += `    <div class="slides-container theme-layout-default" style="display: flex; flex-direction: column; gap: 40px;">\n`;
        htmlContent += renderSlideListHtml(slidesData, slidesData);
        htmlContent += `    </div>\n`;
    }
    
    // Build switcher HTML toolbar
    let switcherHtml = '';
    if (themeMode === "all") {
        switcherHtml = `
    <div class="toolbar">
        <span class="toolbar-title">切换主题预览:</span>
        <button class="theme-btn" onclick="setTheme('light', this)">极简浅色</button>
        <button class="theme-btn" onclick="setTheme('spatial', this)">动漫风格</button>
        <button class="theme-btn" onclick="setTheme('cyberpunk', this)">赛博朋克</button>
        <button class="theme-btn" onclick="setTheme('holodeck', this)">全息甲板</button>
        <button class="theme-btn" onclick="setTheme('dark', this)">科技深色</button>
        <button class="theme-btn" onclick="setTheme('warm', this)">优雅沙滩</button>
        <button class="theme-btn" onclick="setTheme('aurora', this)">极光幻彩</button>
        <button class="theme-btn" onclick="setTheme('forest', this)">清新森林</button>
        <button class="theme-btn" onclick="setTheme('ocean', this)">深邃海洋</button>
        <button class="theme-btn" onclick="setTheme('ghibli', this)">吉卜力</button>
        <span style="color: rgba(255,255,255,0.2); margin: 0 5px;">|</span>
        <span class="toolbar-title">下载对应的 PPTX:</span>
        <a id="download-link" href="#" class="theme-btn" style="background: #10B981; color: #FFFFFF; text-decoration: none; display: inline-flex; align-items: center; gap: 5px;" download>
            📥 下载当前版本
        </a>
    </div>
    
    <script>
        const baseName = '${baseNameWithoutExt}';
        function setTheme(theme, btn) {
            document.body.className = 'theme-' + theme;
            document.querySelectorAll('.theme-btn').forEach(b => b.classList.remove('active'));
            if (btn) btn.classList.add('active');
            else {
                const btns = document.querySelectorAll('.theme-btn');
                btns.forEach(b => {
                    if (b.getAttribute('onclick').includes(theme)) b.classList.add('active');
                });
            }
            
            // Toggle layout containers
            const spatialContainer = document.querySelector('.theme-layout-spatial');
            const cyberpunkContainer = document.querySelector('.theme-layout-cyberpunk');
            const holodeckContainer = document.querySelector('.theme-layout-holodeck');
            
            if (spatialContainer) spatialContainer.style.display = (theme === 'spatial' || (theme !== 'cyberpunk' && theme !== 'holodeck')) ? 'flex' : 'none';
            if (cyberpunkContainer) cyberpunkContainer.style.display = (theme === 'cyberpunk') ? 'flex' : 'none';
            if (holodeckContainer) holodeckContainer.style.display = (theme === 'holodeck') ? 'flex' : 'none';
            
            // Update download link
            const downloadLink = document.getElementById('download-link');
            downloadLink.href = baseName + '_' + theme + '.pptx';
            downloadLink.innerText = '📥 下载 ' + getThemeName(theme) + ' PPTX';
        }
        function getThemeName(theme) {
            switch(theme) {
                case 'light': return '极简浅色';
                case 'spatial': return '动漫风格';
                case 'cyberpunk': return '赛博朋克';
                case 'holodeck': return '全息甲板';
                case 'dark': return '科技深色';
                case 'warm': return '优雅沙滩';
                case 'aurora': return '极光幻彩';
                case 'forest': return '清新森林';
                case 'ocean': return '深邃海洋';
                case 'ghibli': return '吉卜力';
                default: return '当前';
            }
        }
        // Initialize
        const defaultInitTheme = '${inputFile.toLowerCase().includes('cyberpunk') ? 'cyberpunk' : inputFile.toLowerCase().includes('holodeck') ? 'holodeck' : inputFile.toLowerCase().includes('ghibli') ? 'ghibli' : 'spatial'}';
        setTheme(defaultInitTheme);
    </script>
`;
    } else {
        switcherHtml = `
    <div class="toolbar">
        <span class="toolbar-title">当前主题: ${themeMode === 'light' ? '极简浅色' : themeMode === 'spatial' ? '动漫风格' : themeMode === 'cyberpunk' ? '赛博朋克' : themeMode === 'holodeck' ? '全息甲板' : themeMode === 'dark' ? '科技深色' : themeMode === 'warm' ? '优雅沙滩' : themeMode === 'aurora' ? '极光幻彩' : themeMode === 'forest' ? '清新森林' : themeMode === 'ghibli' ? '吉卜力' : '深邃海洋'}</span>
        <span style="color: rgba(255,255,255,0.2); margin: 0 5px;">|</span>
        <a href="${basePptxName}" class="theme-btn" style="background: #10B981; color: #FFFFFF; text-decoration: none; display: inline-flex; align-items: center; gap: 5px;" download>
            📥 下载 PPTX 文件
        </a>
    </div>
    <script>
        document.body.className = 'theme-${themeMode}';
        const spatialContainer = document.querySelector('.theme-layout-spatial');
        const cyberpunkContainer = document.querySelector('.theme-layout-cyberpunk');
        const holodeckContainer = document.querySelector('.theme-layout-holodeck');
        
        if (spatialContainer) spatialContainer.style.display = ('${themeMode}' === 'spatial' || ('${themeMode}' !== 'cyberpunk' && '${themeMode}' !== 'holodeck')) ? 'flex' : 'none';
        if (cyberpunkContainer) cyberpunkContainer.style.display = ('${themeMode}' === 'cyberpunk') ? 'flex' : 'none';
        if (holodeckContainer) holodeckContainer.style.display = ('${themeMode}' === 'holodeck') ? 'flex' : 'none';
    </script>
`;
    }
    
    htmlContent += `
    ${switcherHtml}
</body>
</html>
`;
    
    fs.writeFileSync(htmlPath, htmlContent, 'utf-8');
    console.log(`\u2705 Interactive HTML preview saved to ${htmlPath}`);
    
    // Copy skill assets to output directory next to HTML
    try {
        const skillAssetsDir = path.resolve(__dirname, '../assets');
        const htmlAssetsDir = path.join(path.dirname(htmlPath), 'assets');
        if (fs.existsSync(skillAssetsDir)) {
            if (!fs.existsSync(htmlAssetsDir)) {
                fs.mkdirSync(htmlAssetsDir, { recursive: true });
            }
            const files = fs.readdirSync(skillAssetsDir);
            files.forEach(file => {
                const srcFile = path.join(skillAssetsDir, file);
                const destFile = path.join(htmlAssetsDir, file);
                if (fs.statSync(srcFile).isFile()) {
                    fs.copyFileSync(srcFile, destFile);
                }
            });
        }
    } catch (e) {
        console.error("Failed to copy assets for HTML preview:", e.message);
    }
}

function getIconForTitle(title) {
    const t = (title || "").toLowerCase();
    if (t.includes("threat") || t.includes("security") || t.includes("safety") || t.includes("安全") || t.includes("防御")) return "🛡️";
    if (t.includes("intelligence") || t.includes("ai") || t.includes("analysis") || t.includes("智能") || t.includes("分析")) return "💡";
    if (t.includes("real-time") || t.includes("monitoring") || t.includes("speed") || t.includes("实时") || t.includes("监测")) return "⏱️";
    if (t.includes("protection") || t.includes("lock") || t.includes("defense") || t.includes("保护") || t.includes("锁定")) return "🔒";
    if (t.includes("market") || t.includes("growth") || t.includes("revenue") || t.includes("增长") || t.includes("市场")) return "📈";
    if (t.includes("product") || t.includes("innovation") || t.includes("design") || t.includes("创新") || t.includes("设计")) return "⚙️";
    if (t.includes("customer") || t.includes("success") || t.includes("team") || t.includes("用户") || t.includes("客户") || t.includes("团队")) return "👥";
    if (t.includes("sustainability") || t.includes("energy") || t.includes("green") || t.includes("环保") || t.includes("绿色") || t.includes("可持续")) return "🌱";
    if (t.includes("water") || t.includes("conservation") || t.includes("liquid") || t.includes("水")) return "💧";
    if (t.includes("carbon") || t.includes("offset") || t.includes("emission") || t.includes("碳")) return "🍃";
    if (t.includes("trend") || t.includes("direction") || t.includes("future") || t.includes("趋势")) return "🎯";
    if (t.includes("architecture") || t.includes("model") || t.includes("structure") || t.includes("架构") || t.includes("模型")) return "🏗️";
    return "🔹";
}

const THEMES = {
    cyberpunk: {
        bg: "030611",
        cardBg: "060B1A",
        cardBorder: "162C3F",
        titleColor: "00F0FF",
        textColor: "F1F5F9",
        accentColor: "BC39FA",
        formulaColor: "00F0FF",
        dividerColor: "00F0FF",
        fontFace: "Share Tech Mono",
        isDark: true
    },
    light: {
        bg: "F8F9FA",
        cardBg: "FFFFFF",
        cardBorder: "E2E8F0",
        titleColor: "0F172A",
        textColor: "334155",
        accentColor: "2563EB",
        formulaColor: "2563EB",
        dividerColor: "E2E8F0",
        fontFace: "Segoe UI",
        isDark: false
    },
    spatial: {
        bg: "060E11",
        cardBg: "0E1A1E",
        cardBorder: "163D3F",
        titleColor: "2DD4BF",
        textColor: "D1E2E4",
        accentColor: "14B8A6",
        formulaColor: "2DD4BF",
        dividerColor: "163D3F",
        fontFace: "Trebuchet MS",
        isDark: true
    },
    holodeck: {
        bg: "080806",
        cardBg: "12110E",
        cardBorder: "2A2215",
        titleColor: "F59E0B",
        textColor: "E5E5E5",
        accentColor: "F59E0B",
        formulaColor: "FBBF24",
        dividerColor: "2A2215",
        fontFace: "Share Tech Mono",
        isDark: true
    },
    dark: {
        bg: "0F172A",
        cardBg: "1E293B",
        cardBorder: "334155",
        titleColor: "38BDF8",
        textColor: "E2E8F0",
        accentColor: "38BDF8",
        formulaColor: "0EA5E9",
        dividerColor: "334155",
        fontFace: "Trebuchet MS",
        isDark: true
    },
    warm: {
        bg: "FAF6F0",
        cardBg: "FFFDF9",
        cardBorder: "E6DFD3",
        titleColor: "3F2E2C",
        textColor: "4A3E3D",
        accentColor: "D97706",
        formulaColor: "B45309",
        dividerColor: "E6DFD3",
        fontFace: "Georgia",
        isDark: false
    },
    aurora: {
        bg: "FAF5FF",
        cardBg: "FFFFFF",
        cardBorder: "E9D5FF",
        titleColor: "7C3AED",
        textColor: "4B5563",
        accentColor: "D946EF",
        formulaColor: "C084FC",
        dividerColor: "E9D5FF",
        fontFace: "Trebuchet MS",
        isDark: false
    },
    forest: {
        bg: "F4F7F5",
        cardBg: "FFFFFF",
        cardBorder: "D1DDD4",
        titleColor: "1C3F24",
        textColor: "2F3E32",
        accentColor: "10B981",
        formulaColor: "059669",
        dividerColor: "D1DDD4",
        fontFace: "Segoe UI",
        isDark: false
    },
    ocean: {
        bg: "F0F7FF",
        cardBg: "FFFFFF",
        cardBorder: "C7D2FE",
        titleColor: "1E3A8A",
        textColor: "374151",
        accentColor: "3B82F6",
        formulaColor: "2563EB",
        dividerColor: "C7D2FE",
        fontFace: "Segoe UI",
        isDark: false
    },
    ghibli: {
        bg: "EAF3E1",
        cardBg: "FBF7EE",
        cardBorder: "C9D9C0",
        titleColor: "2F5D3A",
        textColor: "4A3E2D",
        accentColor: "D9A441",
        formulaColor: "B45309",
        dividerColor: "C9D9C0",
        fontFace: "Georgia",
        isDark: false
    }
};

function addPremiumCard(slide, x, y, w, h, theme, hasAccentBar = true) {
    slide.addShape('roundRect', {
        x: x, y: y, w: w, h: h,
        fill: { color: theme.cardBg },
        line: { color: theme.cardBorder, width: 1.5 }
    });
    // For spatial theme (bg is 060E11), we do not add the thick accent top bar to align with the mockup
    const isSpatial = (theme.bg === '060E11');
    if (hasAccentBar && !isSpatial) {
        slide.addShape('rect', {
            x: x, y: y, w: w, h: 0.08,
            fill: { color: theme.accentColor }
        });
    }
}

function generatePptxForTheme(themeKey, outFile) {
    const theme = THEMES[themeKey];
    const pres = new pptxgen();
    pres.layout = 'LAYOUT_WIDE';
    
    // Check if we have a theme-specific DSL deck
    let themeSlidesData = slidesData;
    const inputDir = path.dirname(inputFile);
    const dslPath = path.resolve(inputDir, `${themeKey}_deck.json`);
    if (fs.existsSync(dslPath)) {
        try {
            themeSlidesData = JSON.parse(fs.readFileSync(dslPath, 'utf-8')).slides;
            themeSlidesData.forEach(s => { s.isDsl = true; s.dsl = s; });
        } catch (e) {
            console.error(`Failed to parse DSL file for theme ${themeKey}:`, e.message);
        }
    }
    
    // Helper to format line colors based on theme
    const processLinesForTheme = (dataText, fontSize) => {
        const segments = processLines(dataText, fontSize);
        let isNewPara = true;
        segments.forEach(seg => {
            // Map hardcoded colors to theme colors
            if (seg.options.color === '003366') {
                seg.options.color = theme.formulaColor;
            } else if (seg.options.color === '333333') {
                seg.options.color = theme.textColor;
            }
            // Set theme font Face
            seg.options.fontFace = theme.fontFace;

            if (isNewPara) {
                isNewPara = false;
            } else {
                // Remove bullet and indent options from subsequent segments of same paragraph to avoid splitting them into separate bullets in pptxgenjs
                if (seg.options.bullet) {
                    delete seg.options.bullet;
                }
                if (seg.options.indentLevel) {
                    delete seg.options.indentLevel;
                }
            }

            if (seg.options.breakLine) {
                isNewPara = true;
            }
        });
        return segments;
    };
    
    themeSlidesData.forEach((data, idx) => {
        let slide = pres.addSlide();
        if (data.isDsl) {
            const slideBgName = `${themeKey}_bg_slide_${idx}.jpg`;
            const slideBgPath = path.resolve(__dirname, `../assets/${slideBgName}`);
            const themeBgName = `${themeKey}_bg.jpg`;
            const themeBgPath = path.resolve(__dirname, `../assets/${themeBgName}`);
            
            if (fs.existsSync(slideBgPath)) {
                slide.background = { path: slideBgPath };
            } else if (fs.existsSync(themeBgPath)) {
                slide.background = { path: themeBgPath };
            } else {
                slide.background = { fill: theme.bg };
            }

            // 1. Cover Slide vs Content Slide Headers
            if (idx === 0) {
                // Cover background details are handled by specific backgrounds
            } else {
                if (themeKey === 'spatial') {
                    // Top Header Panel Rounded Rectangle
                    slide.addShape('roundRect', {
                        x: 0.4, y: 0.25, w: 12.53, h: 1.1,
                        fill: { color: theme.cardBg },
                        line: { color: theme.cardBorder, width: 1.5 }
                    });
                    
                    // Logo
                    slide.addText([
                        { text: "E", options: { bold: true, color: theme.titleColor, fontSize: 24 } },
                        { text: " & ", options: { color: theme.textColor, fontSize: 20 } },
                        { text: "🧊", options: { fontSize: 20 } }
                    ], {
                        x: 0.6, y: 0.4, w: 1.8, h: 0.8,
                        fontFace: theme.fontFace,
                        valign: "middle", align: "left"
                    });
                    
                    // Centered Title and Subtitle inside the header panel
                    slide.addText(data.title || '', {
                        x: 2.2, y: 0.35, w: 8.93, h: 0.45, 
                        fontSize: 20, bold: true, color: theme.titleColor,
                        fontFace: theme.fontFace,
                        align: "center", valign: "middle",
                        margin: 0
                    });
                    
                    const runningSubtitle = (slidesData[0] && slidesData[0].dsl && slidesData[0].dsl.elements)
                        ? (slidesData[0].dsl.elements.find(el => el.id === 'title_card')?.content?.body || "Industry Overview")
                        : "Industry Overview";
                    slide.addText(runningSubtitle, {
                        x: 2.2, y: 0.85, w: 8.93, h: 0.3, 
                        fontSize: 11, color: "A0AEC0",
                        fontFace: theme.fontFace,
                        align: "center", valign: "middle",
                        margin: 0
                    });
                } else {
                    // Accent indicator
                    slide.addShape('rect', {
                        x: 0.8, y: 0.55, w: 0.12, h: 0.5,
                        fill: { color: theme.accentColor }
                    });
                    
                    slide.addText(data.title || '', {
                        x: 1.05, y: 0.4, w: 11.45, h: 0.8, 
                        fontSize: 28, bold: true, color: theme.titleColor,
                        fontFace: theme.fontFace,
                        valign: "middle",
                        margin: 0
                    });
                    
                    slide.addShape('line', {
                        x: 0.8, y: 1.25, w: 11.7, h: 0,
                        line: { color: theme.dividerColor, width: 1.5 }
                    });
                }
            }

            // 2. Render each element in DSL
            const parsePct = (val, maxVal) => {
                if (typeof val === 'number') return val;
                if (typeof val === 'string') {
                    if (val.endsWith('%')) {
                        return (parseFloat(val) / 100) * maxVal;
                    }
                    return parseFloat(val);
                }
                return 0;
            };

            data.dsl.elements.forEach(el => {
                const x_inch = parsePct(el.bounds.x, 13.333);
                const y_inch = parsePct(el.bounds.y, 7.5);
                const w_inch = parsePct(el.bounds.w, 13.333);
                const h_inch = parsePct(el.bounds.h, 7.5);

                if (el.type === 'text') {
                    const fontSize = el.style?.fontSize || 12;
                    const fontWeight = el.style?.fontWeight || 'normal';
                    const align = el.style?.align || 'left';
                    const valign = el.style?.valign || 'top';
                    const color = el.style?.color || 'text-main';

                    let colorVal = theme.textColor;
                    if (color === 'activeCyan' || color === 'active-cyan') colorVal = theme.accentColor;
                    else if (color === 'activePurple' || color === 'active-purple') colorVal = 'A78BFA';
                    else if (color === 'text-main') colorVal = theme.textColor;
                    else if (color === 'text-muted') colorVal = 'A0AEC0';
                    else if (color === 'neon-green') colorVal = '10B981';
                    else if (color === 'neon-orange') colorVal = 'F59E0B';
                    else if (color && color !== 'default') colorVal = color;

                    const runs = processLinesForTheme(el.content || '', fontSize);
                    runs.forEach(run => {
                        run.options.fontFace = theme.fontFace;
                        if (run.options.bold === undefined && fontWeight === 'bold') run.options.bold = true;
                        if (!run.options.color) run.options.color = colorVal;
                    });

                    slide.addText(runs, {
                        x: x_inch, y: y_inch, w: w_inch, h: h_inch,
                        align: align, valign: valign,
                        margin: 0
                    });
                } else if (el.type === 'card') {
                    const hasAccentBar = el.style?.theme === 'accent';
                    addPremiumCard(slide, x_inch, y_inch, w_inch, h_inch, theme, hasAccentBar);

                    // Component Header Tag
                    slide.addText(el.id ? el.id.toUpperCase().replace(/_/g, ' ') : 'COMPONENT', {
                        x: x_inch + 0.15, y: y_inch + 0.1, w: w_inch - 0.3, h: 0.25,
                        fontSize: 8, bold: true, color: theme.textColor, fontFace: theme.fontFace, margin: 0
                    });

                    // Title
                    slide.addText(el.content.title || '', {
                        x: x_inch + 0.15, y: y_inch + 0.35, w: w_inch - 0.3, h: 0.3,
                        fontSize: 13, bold: true, color: el.style?.theme === 'accent' ? theme.accentColor : theme.titleColor,
                        fontFace: theme.fontFace, margin: 0
                    });

                    // Divider line
                    slide.addShape('line', {
                        x: x_inch + 0.15, y: y_inch + 0.7, w: w_inch - 0.3, h: 0,
                        line: { color: theme.dividerColor, width: 1 }
                    });

                    // Body
                    const bodyRuns = processLinesForTheme(el.content.body || '', 10);
                    bodyRuns.forEach(run => {
                        run.options.fontFace = theme.fontFace;
                        if (!run.options.color) run.options.color = theme.textColor;
                    });
                    slide.addText(bodyRuns, {
                        x: x_inch + 0.15, y: y_inch + 0.8, w: w_inch - 0.3, h: h_inch - 0.95,
                        valign: "top", margin: 0
                    });
                } else if (el.type === 'image') {
                    addPremiumCard(slide, x_inch, y_inch, w_inch, h_inch, theme, false);
                    let imgPath = el.content.path;
                    if (!path.isAbsolute(imgPath)) {
                        imgPath = path.resolve(path.dirname(inputFile), imgPath);
                    }
                    if (fs.existsSync(imgPath)) {
                        slide.addImage({
                            path: imgPath,
                            x: x_inch + 0.05, y: y_inch + 0.05,
                            w: w_inch - 0.1, h: h_inch - 0.1,
                            sizing: { type: 'contain' }
                        });
                    }
                } else if (el.type === 'vector') {
                    if (el.name === 'radar-sweep') {
                        const radarSize = Math.min(w_inch, h_inch);
                        const cx = x_inch + (w_inch - radarSize) / 2;
                        const cy = y_inch + (h_inch - radarSize) / 2;
                        [1.0, 0.75, 0.5, 0.25].forEach((ratio, rIdx) => {
                            const size = radarSize * ratio;
                            const rx = cx + (radarSize - size) / 2;
                            const ry = cy + (radarSize - size) / 2;
                            slide.addShape('oval', {
                                x: rx, y: ry, w: size, h: size,
                                fill: { color: 'FFFFFF', transparency: 100 },
                                line: {
                                    color: theme.accentColor,
                                    width: 1,
                                    transparency: 100 - ratio * 40,
                                    dashType: rIdx === 3 ? 'dash' : 'solid'
                                }
                            });
                        });
                        slide.addShape('line', {
                            x: cx + radarSize/2, y: cy + radarSize/2,
                            w: (radarSize / 2) * Math.cos(-Math.PI/4),
                            h: -(radarSize / 2) * Math.sin(-Math.PI/4),
                            line: { color: 'A78BFA', width: 1.5 }
                        });
                    } else if (el.name === 'spatial-sphere') {
                        const sphSize = Math.min(w_inch, h_inch);
                        const sx = x_inch + (w_inch - sphSize) / 2;
                        const sy = y_inch + (h_inch - sphSize) / 2;
                        slide.addShape('oval', {
                            x: sx, y: sy, w: sphSize, h: sphSize,
                            fill: { color: 'FFFFFF', transparency: 100 },
                            line: { color: theme.accentColor, width: 1 }
                        });
                        slide.addShape('oval', {
                            x: sx, y: sy + sphSize * 0.3, w: sphSize, h: sphSize * 0.4,
                            fill: { color: 'FFFFFF', transparency: 100 },
                            line: { color: theme.accentColor, width: 0.8, dashType: 'dash' }
                        });
                        slide.addShape('oval', {
                            x: sx + sphSize * 0.3, y: sy, w: sphSize * 0.4, h: sphSize,
                            fill: { color: 'FFFFFF', transparency: 100 },
                            line: { color: theme.accentColor, width: 0.8, dashType: 'dash' }
                        });
                        slide.addShape('oval', {
                            x: sx + sphSize/2 - 0.05, y: sy + sphSize/2 - 0.05, w: 0.1, h: 0.1,
                            fill: { color: 'A78BFA' }
                        });
                    } else if (el.name === 'viewfinder-reticle') {
                        slide.addShape('roundRect', {
                            x: x_inch, y: y_inch, w: w_inch, h: h_inch,
                            fill: { color: 'FFFFFF', transparency: 100 },
                            line: { color: theme.cardBorder, width: 1.5 }
                        });
                        const ccx = x_inch + w_inch/2;
                        const ccy = y_inch + h_inch/2;
                        slide.addShape('line', { x: ccx - 0.3, y: ccy, w: 0.6, h: 0, line: { color: theme.accentColor, width: 1 } });
                        slide.addShape('line', { x: ccx, y: ccy - 0.3, w: 0, h: 0.6, line: { color: theme.accentColor, width: 1 } });
                        slide.addShape('oval', { x: ccx - 0.1, y: ccy - 0.1, w: 0.2, h: 0.2, fill: { color: 'FFFFFF', transparency: 100 }, line: { color: theme.accentColor, width: 1 } });
                    } else if (el.name === 'timeline-pulse') {
                        slide.addShape('line', {
                            x: x_inch, y: y_inch + h_inch/2, w: w_inch, h: 0,
                            line: { color: theme.accentColor, width: 2, dashType: 'dash' }
                        });
                        slide.addShape('oval', {
                            x: x_inch + w_inch/2 - 0.1, y: y_inch + h_inch/2 - 0.1, w: 0.2, h: 0.2,
                            fill: { color: theme.accentColor }
                        });
                    } else if (el.name === 'multiview-diagnostic') {
                        slide.addText("3D SCAN RECONSTRUCTION", {
                            x: x_inch, y: y_inch, w: w_inch, h: 0.25,
                            fontSize: 8, color: theme.accentColor, fontFace: theme.fontFace, align: "center", bold: true
                        });
                        
                        const subW = (w_inch - 0.2) / 2;
                        const subH = h_inch - 0.3;
                        const subY = y_inch + 0.3;
                        
                        slide.addShape('roundRect', { x: x_inch, y: subY, w: subW, h: subH, fill: { color: '000000', transparency: 70 }, line: { color: theme.cardBorder, width: 1 } });
                        slide.addText("TOP_V", { x: x_inch + 0.05, y: subY + 0.05, w: subW - 0.1, h: 0.2, fontSize: 6, color: 'A0AEC0', fontFace: theme.fontFace });
                        
                        const cx1 = x_inch + subW/2;
                        const cy1 = subY + subH/2;
                        const rsize1 = Math.min(subW, subH) * 0.6;
                        slide.addShape('oval', { x: cx1 - rsize1/2, y: cy1 - rsize1/2, w: rsize1, h: rsize1, fill: { color: 'FFFFFF', transparency: 100 }, line: { color: theme.accentColor, width: 1, dashType: 'dash' } });
                        slide.addShape('oval', { x: cx1 - rsize1/6, y: cy1 - rsize1/6, w: rsize1/3, h: rsize1/3, fill: { color: 'FFFFFF', transparency: 100 }, line: { color: 'A78BFA', width: 1 } });
                        slide.addShape('line', { x: cx1, y: cy1, w: rsize1/2 * 0.7, h: -rsize1/2 * 0.7, line: { color: theme.accentColor, width: 1.5 } });
                        
                        const bx2 = x_inch + subW + 0.2;
                        slide.addShape('roundRect', { x: bx2, y: subY, w: subW, h: subH, fill: { color: '000000', transparency: 70 }, line: { color: theme.cardBorder, width: 1 } });
                        slide.addText("SIDE_V", { x: bx2 + 0.05, y: subY + 0.05, w: subW - 0.1, h: 0.2, fontSize: 6, color: 'A0AEC0', fontFace: theme.fontFace });
                        
                        const cx2 = bx2 + subW/2;
                        const cy2 = subY + subH/2;
                        slide.addShape('line', { x: bx2 + 0.1, y: cy2 + 0.3, w: subW - 0.2, h: 0, line: { color: theme.accentColor, width: 1 } });
                        slide.addShape('rect', { x: cx2 - 0.4, y: cy2 - 0.2, w: 0.8, h: 0.4, fill: { color: theme.accentColor, transparency: 90 }, line: { color: theme.accentColor, width: 1 } });
                        slide.addShape('oval', { x: cx2 - 0.3, y: cy2 + 0.2, w: 0.15, h: 0.15, fill: { color: 'FFFFFF', transparency: 100 }, line: { color: 'A78BFA', width: 1 } });
                        slide.addShape('oval', { x: cx2 + 0.15, y: cy2 + 0.2, w: 0.15, h: 0.15, fill: { color: 'FFFFFF', transparency: 100 }, line: { color: 'A78BFA', width: 1 } });
                    }
                } else if (el.type === 'timeline') {
                    const numCards = el.content.length;
                    const gap = 0.25;
                    const cardW = (w_inch - (numCards - 1) * gap) / numCards;
                    
                    el.content.forEach((item, itemIdx) => {
                        const itemX = x_inch + itemIdx * (cardW + gap);
                        addPremiumCard(slide, itemX, y_inch, cardW, h_inch, theme, true);
                        
                        slide.addText(`0${itemIdx + 1}`, {
                            x: itemX + 0.15, y: y_inch + 0.1, w: cardW - 0.3, h: 0.3,
                            fontSize: 16, bold: true, color: theme.accentColor, fontFace: theme.fontFace, margin: 0
                        });
                        
                        slide.addText(item.title || '', {
                            x: itemX + 0.15, y: y_inch + 0.45, w: cardW - 0.3, h: 0.25,
                            fontSize: 11, bold: true, color: theme.titleColor, fontFace: theme.fontFace, margin: 0
                        });
                        
                        const itemBodyRuns = processLinesForTheme(item.body || '', 9.5);
                        itemBodyRuns.forEach(run => {
                            run.options.fontFace = theme.fontFace;
                            if (!run.options.color) run.options.color = theme.textColor;
                        });
                        slide.addText(itemBodyRuns, {
                            x: itemX + 0.15, y: y_inch + 0.75, w: cardW - 0.3, h: h_inch - 0.9,
                            valign: "top", margin: 0
                        });
                    });
                }
            });

            if (data.notes) {
                slide.addNotes(data.notes);
            }
        } else {

        
        if (themeKey === 'spatial' || themeKey === 'cyberpunk' || themeKey === 'holodeck') {
            const bgPath = path.resolve(__dirname, `../assets/${themeKey}_bg.jpg`);
            if (fs.existsSync(bgPath)) {
                slide.background = { path: bgPath };
            } else {
                slide.background = { fill: theme.bg };
            }
        } else {
            slide.background = { fill: theme.bg };
        }
        
        // 1. Cover Slide layout
        if (idx === 0) {
            // Left accent color block
            slide.addShape('rect', {
                x: 0, y: 0, w: 4.0, h: 7.5,
                fill: { color: theme.accentColor }
            });
            // Left accent border box (inset)
            slide.addShape('rect', {
                x: 0.5, y: 0.5, w: 3.0, h: 6.5,
                fill: { color: 'FFFFFF', transparency: 100 },
                line: { color: 'FFFFFF', width: 1.5 }
            });
            
            if (data.title) {
                slide.addText(data.title, {
                    x: 4.8, y: 2.0, w: 7.5, h: 2.2, 
                    fontSize: 36, bold: true, color: theme.titleColor,
                    fontFace: theme.fontFace,
                    align: "left", valign: "bottom",
                    margin: 0
                });
            }
            if (data.text) {
                const formattedText = processLinesForTheme(data.text, 18);
                formattedText.forEach(seg => {
                    seg.options.align = "left";
                });
                slide.addText(formattedText, {
                    x: 4.8, y: 4.5, w: 7.5, h: 2.0,
                    valign: "top",
                    margin: 0,
                    paraSpaceAfter: 8
                });
            }
            if (data.notes) {
                slide.addNotes(data.notes);
            }
            return;
        }
        
        // 2. Normal Content Slide layout
        const mainTitle = slidesData[0] ? (slidesData[0].title || 'PRESENTATION') : 'PRESENTATION';
        // Top bar meta text
        if (themeKey !== 'spatial') {
            slide.addText(`SLIDE ${idx < 10 ? '0' + idx : idx} | ${mainTitle.toUpperCase()}`, {
                x: 0.8, y: 0.15, w: 6.0, h: 0.25,
                fontSize: 9, bold: true, color: theme.textColor,
                fontFace: theme.fontFace,
                align: "left", valign: "middle",
                margin: 0
            });
        }
        
        // macOS style window dots
        const dotSize = 0.08;
        const dotY = 0.23;
        slide.addShape('oval', {
            x: 12.1, y: dotY, w: dotSize, h: dotSize,
            fill: { color: 'FF5F56' }
        });
        slide.addShape('oval', {
            x: 12.22, y: dotY, w: dotSize, h: dotSize,
            fill: { color: 'FFBD2E' }
        });
        slide.addShape('oval', {
            x: 12.34, y: dotY, w: dotSize, h: dotSize,
            fill: { color: '27C93F' }
        });

        // Background decorations and footer for spatial theme
        if (themeKey === 'spatial') {
            const footerTitle = slidesData[0] ? (slidesData[0].title || 'EMBODIED AI & SPATIAL INTELLIGENCE') : 'EMBODIED AI & SPATIAL INTELLIGENCE';
            const totalContentSlides = slidesData.length - 1;
            const totalContentSlidesStr = totalContentSlides < 10 ? '0' + totalContentSlides : totalContentSlides;

            slide.addShape('line', { x: 0.2, y: 0.4, w: 0.5, h: 0, line: { color: "14B8A6", width: 1 } });
            slide.addShape('line', { x: 0.7, y: 0.4, w: 0.2, h: 0.2, line: { color: "14B8A6", width: 1 } });
            slide.addShape('line', { x: 0.9, y: 0.6, w: 0.5, h: 0, line: { color: "14B8A6", width: 1 } });
            slide.addShape('line', { x: 12.63, y: 0.4, w: 0.5, h: 0, line: { color: "14B8A6", width: 1 } });
            slide.addShape('line', { x: 0.8, y: 7.0, w: 11.7, h: 0, line: { color: "163D3F", width: 1 } });
            
            slide.addText(`${footerTitle.toUpperCase()} | INDUSTRY TRENDS | SLIDE ${idx < 10 ? '0'+idx : idx}/${totalContentSlidesStr} | 2024`, {
                x: 0.8, y: 7.1, w: 11.7, h: 0.3,
                fontSize: 8, bold: true, color: "5F7D81",
                fontFace: theme.fontFace,
                align: "center", valign: "middle"
            });
        }

        if (data.title) {
            if (themeKey === 'spatial') {
                // Top Header Panel Rounded Rectangle
                slide.addShape('roundRect', {
                    x: 0.4, y: 0.25, w: 12.53, h: 1.1,
                    fill: { color: theme.cardBg },
                    line: { color: theme.cardBorder, width: 1.5 }
                });
                
                // Logo on the left inside the header panel
                slide.addText([
                    { text: "E", options: { bold: true, color: theme.titleColor, fontSize: 24 } },
                    { text: " & ", options: { color: theme.textColor, fontSize: 20 } },
                    { text: "🧊", options: { fontSize: 20 } }
                ], {
                    x: 0.6, y: 0.4, w: 1.8, h: 0.8,
                    fontFace: theme.fontFace,
                    valign: "middle", align: "left"
                });
                
                // Centered Title and Subtitle inside the header panel
                slide.addText(data.title, {
                    x: 2.2, y: 0.35, w: 8.93, h: 0.45, 
                    fontSize: 20, bold: true, color: theme.titleColor,
                    fontFace: theme.fontFace,
                    align: "center", valign: "middle",
                    margin: 0
                });
                
                const runningSubtitle = (slidesData[0] && slidesData[0].text) 
                    ? slidesData[0].text.split('\n')[0].trim() 
                    : "Industry Overview";
                slide.addText(runningSubtitle, {
                    x: 2.2, y: 0.85, w: 8.93, h: 0.3, 
                    fontSize: 11, color: "A0AEC0",
                    fontFace: theme.fontFace,
                    align: "center", valign: "middle",
                    margin: 0
                });
            } else {
                // Accent indicator
                slide.addShape('rect', {
                    x: 0.8, y: 0.55, w: 0.12, h: 0.5,
                    fill: { color: theme.accentColor }
                });
                
                slide.addText(data.title, {
                    x: 1.05, y: 0.4, w: 11.45, h: 0.8, 
                    fontSize: 28, bold: true, color: theme.titleColor,
                    fontFace: theme.fontFace,
                    valign: "middle",
                    margin: 0
                });
                
                slide.addShape('line', {
                    x: 0.8, y: 1.25, w: 11.7, h: 0,
                    line: { color: theme.dividerColor, width: 1.5 }
                });
            }
        }

        const hasImage = data.images.length > 0;
        const layout = getSlideLayoutInfo(data, hasImage);
        
        if (layout.type === 'asymmetric') {
            const dynamicFontSize = calculateDynamicFontSize(data.text, true);
            const formattedText = processLinesForTheme(data.text, dynamicFontSize);
            
            if (formattedText.length > 0) {
                addPremiumCard(slide, 0.8, 1.6, 5.2, 4.8, theme, true);
                slide.addText(formattedText, {
                    x: 1.1, y: 1.9, w: 4.6, h: 4.2,
                    valign: "top",
                    margin: 0,
                    paraSpaceAfter: 6
                });
            }
            
            let imgObj = data.images[0];
            let imgPath = imgObj.path;
            let imgAlt = imgObj.alt || "";
            if (!path.isAbsolute(imgPath)) {
                imgPath = path.resolve(path.dirname(inputFile), imgPath);
            }
            
            if (fs.existsSync(imgPath)) {
                const maxW = 5.7;
                const maxH = 4.8;
                const targetX = 6.8;
                
                let imgW = maxW;
                let imgH = maxH;
                
                const dims = getImageDimensions(imgPath);
                if (dims) {
                    const aspect = dims.width / dims.height;
                    const targetAspect = maxW / maxH;
                    if (aspect > targetAspect) {
                        imgW = maxW;
                        imgH = maxW / aspect;
                    } else {
                        imgH = maxH;
                        imgW = maxH * aspect;
                    }
                }
                
                const targetY = 1.6 + (4.8 - imgH) / 2;
                
                addPremiumCard(slide, targetX, targetY, imgW, imgH, theme, false);
                
                const pad = 0.08;
                slide.addImage({
                    path: imgPath,
                    x: targetX + pad, y: targetY + pad, w: imgW - pad*2, h: imgH - pad*2
                });
                
                // If image is flat, render a grey caption card underneath
                const captionH = 6.4 - (targetY + imgH) - 0.15;
                if (captionH >= 0.8 && imgAlt) {
                    addPremiumCard(slide, targetX, targetY + imgH + 0.15, imgW, captionH, theme, false);
                    slide.addText(imgAlt, {
                        x: targetX + 0.1, y: targetY + imgH + 0.25, w: imgW - 0.2, h: captionH - 0.2,
                        fontSize: 11, italic: true, color: theme.textColor,
                        fontFace: theme.fontFace,
                        align: "center", valign: "middle",
                        margin: 0
                    });
                }
            }
        } else if (layout.type === 'timeline') {
            const { introLines, cardItems } = layout;
            if (introLines.length > 0) {
                const introText = introLines.join('\n');
                const formattedIntro = processLinesForTheme(introText, 16);
                slide.addText(formattedIntro, {
                    x: 0.8, y: 1.35, w: 11.7, h: 0.5,
                    valign: "middle",
                    margin: 0,
                    paraSpaceAfter: 4
                });
            }
            
            // Draw Timeline dashed line
            slide.addShape('line', {
                x: 1.0, y: 2.2, w: 11.3, h: 0,
                line: { color: theme.accentColor, width: 2, dashType: "dash" }
            });
            
            const numCards = cardItems.length;
            const totalW = 11.7;
            const gap = 0.4;
            const cardW = (totalW - (numCards - 1) * gap) / numCards;
            const startX = 0.8;
            
            cardItems.forEach((item, cardIdx) => {
                const cardX = startX + cardIdx * (cardW + gap);
                const cx = cardX + cardW / 2;
                
                // Oval dot
                slide.addShape('oval', {
                    x: cx - 0.15, y: 2.05, w: 0.3, h: 0.3,
                    fill: { color: theme.accentColor },
                    line: { color: theme.bg, width: 2 }
                });
                
                // Step number above timeline
                slide.addText("0" + (cardIdx + 1), {
                    x: cx - 0.5, y: 1.4, w: 1.0, h: 0.5,
                    fontSize: 18, bold: true, color: theme.accentColor,
                    fontFace: theme.fontFace,
                    align: "center", valign: "bottom",
                    margin: 0
                });
                
                // Card below timeline
                addPremiumCard(slide, cardX, 2.6, cardW, 3.8, theme, true);
                
                // Card Title (Centered)
                slide.addText(item.title, {
                    x: cardX + 0.15, y: 2.75, w: cardW - 0.3, h: 0.4,
                    fontSize: 13, bold: true, color: theme.titleColor,
                    fontFace: theme.fontFace,
                    align: "center", valign: "middle",
                    margin: 0
                });
                
                // Card Body (No divider, starts right after title)
                const bodySegments = processLinesForTheme(item.body, 11);
                slide.addText(bodySegments, {
                    x: cardX + 0.15, y: 3.25, w: cardW - 0.3, h: 3.0,
                    valign: "top",
                    margin: 0,
                    paraSpaceAfter: 4
                });
            });
        } else if (layout.type === 'grid') {
            const { introLines, cardItems } = layout;
            if (introLines.length > 0) {
                const introText = introLines.join('\n');
                const formattedIntro = processLinesForTheme(introText, 16);
                slide.addText(formattedIntro, {
                    x: 0.8, y: 1.5, w: 11.7, h: 0.6,
                    valign: "middle",
                    margin: 0,
                    paraSpaceAfter: 4
                });
            }
            
            const cardY = introLines.length > 0 ? 2.3 : 1.6;
            const cardH = introLines.length > 0 ? 4.1 : 4.8;
            
            const numCards = cardItems.length;
            const totalW = 11.7;
            const gap = 0.4;
            const cardW = (totalW - (numCards - 1) * gap) / numCards;
            const startX = 0.8;
            
            if (numCards === 4) {
                // 2x2 matrix layout as seen in spatial_ai_preview.jpg
                const startX = 0.4;
                const cardY = 1.6;
                const cardW_matrix = 6.0;
                const cardH_matrix = 2.45;
                const gapX = 0.53;
                const gapY = 0.35;
                
                cardItems.forEach((item, cardIdx) => {
                    const row = Math.floor(cardIdx / 2);
                    const col = cardIdx % 2;
                    const cardX = startX + col * (cardW_matrix + gapX);
                    const cardY_matrix = cardY + row * (cardH_matrix + gapY);
                    
                    addPremiumCard(slide, cardX, cardY_matrix, cardW_matrix, cardH_matrix, theme, true);
                    
                    const icon = getIconForTitle(item.title);
                    
                    // Left Column: Icon
                    if (themeKey === 'spatial') {
                        let iconPath = path.resolve(__dirname, `../assets/spatial_icon_${cardIdx + 1}.svg`);
                        if (!fs.existsSync(iconPath)) {
                            iconPath = path.resolve(__dirname, `../assets/spatial_icon_${cardIdx + 1}.png`);
                        }
                        if (fs.existsSync(iconPath)) {
                            slide.addImage({
                                path: iconPath,
                                x: cardX + 0.15, y: cardY_matrix + 0.3, w: 1.3, h: 1.85,
                                sizing: { type: "contain" }
                            });
                        } else {
                            slide.addText(icon, {
                                x: cardX + 0.15, y: cardY_matrix + 0.3, w: 1.3, h: 1.85,
                                fontSize: 44,
                                align: "center", valign: "middle"
                            });
                        }
                    } else {
                        slide.addText(icon, {
                            x: cardX + 0.15, y: cardY_matrix + 0.3, w: 1.3, h: 1.85,
                            fontSize: 44,
                            align: "center", valign: "middle"
                        });
                    }
                    
                    // Right Column: Title and Body
                    // Title
                    slide.addText(item.title, {
                        x: cardX + 1.6, y: cardY_matrix + 0.25, w: cardW_matrix - 1.8, h: 0.35,
                        fontSize: 13, bold: true, color: theme.titleColor,
                        fontFace: theme.fontFace,
                        align: "left", valign: "middle",
                        margin: 0
                    });
                    
                    // Card Title Divider (Right Column)
                    slide.addShape('line', {
                        x: cardX + 1.6, y: cardY_matrix + 0.65, w: cardW_matrix - 1.9, h: 0,
                        line: { color: theme.dividerColor, width: 1 }
                    });
                    
                    // Body
                    const bodySegments = processLinesForTheme(item.body, 9.5);
                    slide.addText(bodySegments, {
                        x: cardX + 1.6, y: cardY_matrix + 0.75, w: cardW_matrix - 1.8, h: cardH_matrix - 0.9,
                        valign: "top",
                        margin: 0,
                        paraSpaceAfter: 3
                    });
                });
            } else {
                cardItems.forEach((item, cardIdx) => {
                    const cardX = startX + cardIdx * (cardW + gap);
                    
                    addPremiumCard(slide, cardX, cardY, cardW, cardH, theme, true);
                    
                    const icon = getIconForTitle(item.title);
                    
                    // Centered Icon
                    slide.addText(icon, {
                        x: cardX, y: cardY + 0.25, w: cardW, h: 0.5,
                        fontSize: 24,
                        align: "center", valign: "middle"
                    });
                    
                    // Centered Title
                    slide.addText(item.title, {
                        x: cardX + 0.1, y: cardY + 0.8, w: cardW - 0.2, h: 0.4,
                        fontSize: 14, bold: true, color: theme.titleColor,
                        fontFace: theme.fontFace,
                        align: "center", valign: "middle",
                        margin: 0
                    });
                    
                    // Body content (Left-aligned, starts right after title)
                    const bodySegments = processLinesForTheme(item.body, 11);
                    slide.addText(bodySegments, {
                        x: cardX + 0.2, y: cardY + 1.35, w: cardW - 0.4, h: cardH - 1.5,
                        valign: "top",
                        margin: 0,
                        paraSpaceAfter: 4
                    });
                });
            }
        } else if (layout.type === 'centered-breathe') {
            addPremiumCard(slide, 3.16, 2.2, 7.0, 3.6, theme, true);
            
            const dynamicFontSize = calculateDynamicFontSize(data.text, false) + 4;
            const formattedText = processLinesForTheme(data.text, dynamicFontSize);
            
            slide.addText(formattedText, {
                x: 3.41, y: 2.45, w: 6.5, h: 3.1,
                valign: "middle",
                align: "center",
                margin: 0,
                paraSpaceAfter: 6
            });
        } else {
            addPremiumCard(slide, 1.91, 1.6, 9.5, 4.8, theme, true);
            
            const dynamicFontSize = calculateDynamicFontSize(data.text, false);
            const formattedText = processLinesForTheme(data.text, dynamicFontSize);
            
            slide.addText(formattedText, {
                x: 2.21, y: 1.9, w: 8.9, h: 4.2,
                valign: "top",
                margin: 0,
                paraSpaceAfter: 6
            });
        }
        
        if (data.notes) {
            slide.addNotes(data.notes);
        }
    
        }
});
    
    return pres.writeFile({ fileName: outFile });
}

let slidesDataParsed = [];
if (isDslDeck) {
    slidesDataParsed = slidesData;
} else {
    slidesDataParsed = parseSlides(mdText);
    slidesDataParsed.forEach(slide => {
        if (slide.text.trim().startsWith('```json') && slide.text.trim().endsWith('```')) {
            try {
                const jsonText = slide.text.trim().replace(/^```json/, '').replace(/```$/, '').trim();
                const dsl = JSON.parse(jsonText);
                slide.isDsl = true;
                slide.dsl = dsl;
            } catch (e) {
                console.error("Failed to parse inline JSON DSL slide:", e.message);
            }
        }
    });
}
slidesData = slidesDataParsed;


let selectedTheme = "all";
for (let i = 1; i < args.length; i++) {
    if (args[i] === '-t' || args[i] === '--theme') {
        selectedTheme = args[i+1].toLowerCase();
        i++;
    }
}

if (selectedTheme === "all") {
    const promises = Object.keys(THEMES).map(themeKey => {
        let themeOutFile = outputFile.replace(/\.pptx$/i, `_${themeKey}.pptx`);
        if (themeOutFile === outputFile) {
            themeOutFile = `${outputFile}_${themeKey}.pptx`;
        }
        return generatePptxForTheme(themeKey, themeOutFile);
    });
    Promise.all(promises).then(() => {
        console.log(`✅ Generated all ${Object.keys(THEMES).length} theme versions!`);
        generateHtmlPreview(slidesData, outputFile, "all");
    }).catch(err => {
        console.error("Error generating presentations:", err);
    });
} else {
    if (!THEMES[selectedTheme]) {
        console.error(`Unknown theme: ${selectedTheme}. Available: ${Object.keys(THEMES).join(', ')}`);
        process.exit(1);
    }
    generatePptxForTheme(selectedTheme, outputFile).then(() => {
        console.log(`✅ Generated theme: ${selectedTheme}`);
        generateHtmlPreview(slidesData, outputFile, selectedTheme);
    }).catch(err => {
        console.error("Error generating presentation:", err);
    });
}

