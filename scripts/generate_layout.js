#!/usr/bin/env node

/**
 * generate_layout.js
 * AIGC 自动化排版闭环管线脚本 (Layout Generation Pipeline)
 * 流程：读取 Markdown -> 调用 LLM 生成 Layout JSON -> 运行 validate_dsl 校验 -> 反馈纠错自愈 -> 生成 PPTX 与 HTML 预览
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// system instruction for LLM layout designer
const SYSTEM_INSTRUCTION = `
You are a Senior Visual Designer and AIGC Presentation expert. Your role is to convert a raw Markdown presentation into a high-fidelity "markdown-to-pptx" Layout DSL JSON representation.

Your output must follow these design guidelines:
1. CANVAS PAINTER MINDSET: Break grid uniformity. Choose a distinct rhythm ("anchor", "dense", "breathing") for each slide. Use asymmetrical column layouts, nested decorative shapes, or diagonal cuts.
2. DNA RULE-DRIVEN LAYOUT: When a theme's visual DNA is supplied, you MUST follow its rules for spacing rhythm, typography scale, color roles, and decorative elements:
   - Use the visual spacing rhythm (page Padding, element Gap) to set element bounds coordinates.
   - Use the component style (e.g., cardStyle: "glassmorphism" -> card.style.variant: "glass") for cards.
   - Place the decorative vocabulary elements (e.g., circle-ring, dots-grid, hexagons) in the elements array at appropriate coordinates (bounds x/y/w/h) corresponding to their placement rules.
   - Align layouts to the layoutPattern (e.g. contentSplit, gridColumns).
3. NO DUPLICATE TITLES (CRITICAL): The compiler automatically renders H2 slide titles and top-left dividers for content slides. DO NOT manually create a "text" element at y < 15% that repeats the slide title. Reviewer will block it!
4. HIGH VISUAL DENSITY: Add 2-4 decorative shape elements (type: "decoration") per slide to enhance detail based on the theme's decorative vocabulary, such as:
   - "glow-spot": backdrop radial glow (transparency built-in)
   - "cross-marker": HUD crosshairs for diagnostic/scientific looks
   - "separator": dashed connectors or dividers
   - "diagonal-split": color split panels
5. IN-CARD WIDGET SYSTEM: Instead of calculating progress bar coordinates manually, embed them in cards style.widgets:
   - [{"type": "progress-bar", "value": 0.85}, {"type": "icon", "symbol": "❤"}]
6. SCHEMA RULES:
   - Coordinates are percentage strings, e.g., "x": "10%", "y": "20%", "w": "40%", "h": "60%".
   - Top-level metadata is required: {"metadata": {"docTitle": "Title Text", "showPageNumber": true}, "slides": [...]}
   - Slide index matching: slideIndex matches its position.

Output strictly valid JSON matching this schema, without markdown formatting blocks.
`;

function formatDnaPrompt(dna, themeName) {
    if (!dna) return "";
    
    const palette = dna.palette || {};
    const typography = dna.typography || {};
    const spacing = dna.spacing || {};
    const grammar = dna.componentGrammar || {};
    const decors = dna.decorativeVocabulary || (dna._dna && dna._dna.decorativeVocabulary) || [];
    const layout = dna.layoutPattern || {};
    const mood = dna.mood || {};
    
    return `
=== VISUAL DNA RULES FOR THEME "${themeName}" ===
Use the following visual rules extracted from the theme's design DNA to determine card layout bounds, decorative placements, and structural configurations. DO NOT ignore these rules!

1. COLOR ROLES (色彩角色系统):
   - Background Color: ${palette.bg || 'Default'}
   - Card Background: ${palette.cardBg || 'Default'}
   - Card Border / Separator: ${palette.cardBorder || 'Default'}
   - Slide Title Color: ${palette.titleColor || 'Default'}
   - Main Body Text Color: ${palette.textColor || 'Default'}
   - Accent Highlight Color (e.g. badges, progress bars, accents): ${palette.accentColor || 'Default'}
   - Gradient Start/End: ${palette.gradientStart || 'none'} -> ${palette.gradientEnd || 'none'} (${palette.gradientDirection || 'none'})

2. TYPOGRAPHY SCALE (字型与比例规范):
   - Font Face/Family: ${typography.headingFont || 'Default'} / ${typography.bodyFont || 'Default'}
   - Title Font Size / Weight: ${typography.headingScale ? typography.headingScale + 'px' : 'Default'} (${typography.headingWeight || 'bold'})
   - Body Font Size: ${typography.bodyScale ? typography.bodyScale + 'px' : 'Default'}
   - Spacing & Line Height: letter spacing is ${typography.letterSpacing || 'normal'}, line height is ${typography.lineHeight || 'normal'}

3. SPACING RHYTHM (间距节奏系统):
   - Spacing unit: ${spacing.unit ? spacing.unit + 'px' : 'Default'}
   - Page Outer Padding: ${spacing.pagePadding || 'normal'}
   - Card Inner Padding: ${spacing.cardPadding || 'normal'}
   - Gap Between Elements: ${spacing.elementGap ? spacing.elementGap + 'px' : 'Default'}

4. COMPONENT GRAMMAR (组件视觉规范):
   - Card Corner Radius: ${grammar.cardRadius || 'medium'}
   - Card Border / Style: ${grammar.cardStyle || 'default'} (e.g., flat, glassmorphism, outlined, filled)
   - Accent Style: Use variant "${grammar.cardStyle === 'glassmorphism' ? 'glass' : 'default'}" for card components.

5. LAYOUT PATTERNS (页面布局模式):
   - Grid Columns recommendation: ${layout.gridColumns || 3} columns.
   - Content Division Style: ${layout.contentSplit || 'balanced'} (e.g. left-heavy, right-heavy, centered)
   - Layout Elements:
     ${layout.hasDiagonalCut ? '- Apply diagonal split or polygon decoration elements for section transitions.' : ''}
     ${layout.hasVerticalAccentBar ? '- Include vertical accent blocks or accent-line decorations next to text card groups.' : ''}

6. DECORATIVE VOCABULARY (装饰元素应用规则) - CRITICAL:
   You MUST actively place the following decoration elements ("type": "decoration") in corresponding slide layout coordinates (bounds x/y/w/h) to match the visual theme design. Do NOT generate empty or plain pages!
   ${decors.length > 0 ? decors.map((d, i) => `   - Decorative Item ${i+1}:
     * name: "${d.type}"
     * color: "${d.color || palette.accentColor || 'accentColor'}"
     * opacity: ${d.opacity || 0.15}
     * default placement: "${d.placement || 'scattered'}"
     * size category: "${d.size || 'medium'}"`).join('\n') : '   - None specified. Feel free to use subtle cross-markers or glow-spots.'}

7. VISUAL TONE & MOOD (视觉情绪):
   - Tone: ${mood.tone || 'professional'}
   - Theme Light/Dark Mode: ${mood.isDark ? 'Dark Theme' : 'Light Theme'} (Make sure to set element colors/text contrast accordingly!)
   - Recommended Scenario: ${mood.bestFor || 'General'}
=================================================
`;
}

async function main() {
    const args = process.argv.slice(2);
    if (args.length < 3) {
        console.log("Usage: node scripts/generate_layout.js <input_markdown> <output_pptx> <theme_name>");
        console.log("Required Env Vars: GEMINI_API_KEY or OPENAI_API_KEY");
        process.exit(1);
    }

    const inputMdPath = path.resolve(args[0]);
    const outputPptxPath = path.resolve(args[1]);
    const themeName = args[2].toLowerCase();

    if (!fs.existsSync(inputMdPath)) {
        console.error(`❌ Input Markdown not found: ${inputMdPath}`);
        process.exit(1);
    }

    const mdContent = fs.readFileSync(inputMdPath, 'utf-8');
    const inputDir = path.dirname(inputMdPath);
    const targetDslPath = path.resolve(inputDir, `${themeName}_deck.json`);

    // Choose LLM Provider
    let callLLM;
    if (process.env.GEMINI_API_KEY) {
        console.log("🤖 Using Gemini API Provider...");
        callLLM = callGemini;
    } else if (process.env.OPENAI_API_KEY) {
        console.log("🤖 Using OpenAI API Provider...");
        callLLM = callOpenAI;
    } else {
        console.error("❌ Error: No API key found in environment (GEMINI_API_KEY or OPENAI_API_KEY).");
        console.log("💡 Tip: If you are running this within the Antigravity Agent Chat, you don't need API keys!");
        console.log("   Simply tell the agent in chat: \\\"Please generate layout and compile medical_report.md for me.\\\"");
        console.log("   The agent will automatically run the Agent-Driven self-correction loop and compile it for you!");
        process.exit(1);
    }

    // Load visual DNA rules if available
    let themeDna = null;
    const themesJsonPath = path.join(__dirname, 'themes.json');
    if (fs.existsSync(themesJsonPath)) {
        try {
            const themes = JSON.parse(fs.readFileSync(themesJsonPath, 'utf8'));
            const themeEntry = themes[themeName];
            if (themeEntry) {
                if (themeEntry._dna) {
                    themeDna = {
                        themeKey: themeName,
                        themeName: themeEntry.name,
                        palette: {
                            bg: themeEntry.bg,
                            cardBg: themeEntry.cardBg,
                            cardBorder: themeEntry.cardBorder,
                            titleColor: themeEntry.titleColor,
                            textColor: themeEntry.textColor,
                            accentColor: themeEntry.accentColor,
                            dividerColor: themeEntry.dividerColor,
                            gradientStart: themeEntry._dna.gradientStart,
                            gradientEnd: themeEntry._dna.gradientEnd,
                            gradientDirection: themeEntry._dna.gradientDirection
                        },
                        typography: {
                            headingFont: themeEntry.fontFace,
                            bodyFont: themeEntry.fontFace,
                            headingWeight: "bold"
                        },
                        decorativeVocabulary: themeEntry._dna.decorativeVocabulary || [],
                        layoutPattern: themeEntry._dna.layoutPattern || {},
                        mood: themeEntry._dna.mood || {},
                        componentGrammar: themeEntry._dna.componentGrammar || {},
                        spacing: themeEntry._dna.spacing || {}
                    };
                } else {
                    themeDna = {
                        themeKey: themeName,
                        themeName: themeEntry.name,
                        palette: {
                            bg: themeEntry.bg,
                            cardBg: themeEntry.cardBg,
                            cardBorder: themeEntry.cardBorder,
                            titleColor: themeEntry.titleColor,
                            textColor: themeEntry.textColor,
                            accentColor: themeEntry.accentColor,
                            dividerColor: themeEntry.dividerColor
                        },
                        typography: {
                            headingFont: themeEntry.fontFace,
                            bodyFont: themeEntry.fontFace
                        }
                    };
                }
            }
        } catch (e) {
            console.warn(`⚠️ Warning: Failed to load theme ${themeName} from themes.json:`, e.message);
        }
    }

    // Try to load direct visual_dna.json file in the same directory as input Markdown, or script directory
    const dnaPaths = [
        path.resolve(inputDir, `${themeName}_visual_dna.json`),
        path.resolve(inputDir, '..', `${themeName}_visual_dna.json`),
        path.join(__dirname, `${themeName}_visual_dna.json`)
    ];
    for (const dp of dnaPaths) {
        if (fs.existsSync(dp)) {
            try {
                themeDna = JSON.parse(fs.readFileSync(dp, 'utf8'));
                console.log(`✅ Loaded visual DNA from ${dp}`);
                break;
            } catch (e) {
                console.warn(`⚠️ Warning: Failed to parse visual DNA from ${dp}:`, e.message);
            }
        }
    }

    const dnaPrompt = formatDnaPrompt(themeDna, themeName);

    // 1. Initial Generation
    console.log("🚀 Generating initial Layout DSL JSON...");
    let prompt = `Convert the following Markdown presentation into Layout DSL JSON.
${dnaPrompt}

Markdown Content:
${mdContent}`;
    let dslJsonText = "";
    
    try {
        dslJsonText = await callLLM(prompt, SYSTEM_INSTRUCTION);
    } catch (e) {
        console.error("❌ LLM API call failed:", e.message);
        process.exit(1);
    }

    // 2. Clean LLM JSON block markdown output
    dslJsonText = cleanJsonBlock(dslJsonText);

    // 3. Self-Correction Loop
    let attempt = 1;
    const maxAttempts = 3;
    let validated = false;

    while (attempt <= maxAttempts && !validated) {
        console.log(`🔍 Validation Attempt ${attempt}/${maxAttempts}...`);
        fs.writeFileSync(targetDslPath, dslJsonText, 'utf-8');

        try {
            // Run validator script
            execSync(`node tests/validate_dsl.js ${targetDslPath}`, { stdio: 'pipe' });
            console.log("✅ Validation Passed!");
            validated = true;
        } catch (error) {
            const errorOutput = error.stderr ? error.stderr.toString() : error.stdout.toString();
            console.warn(`⚠️ Validation Failed:\n${errorOutput}`);

            if (attempt === maxAttempts) {
                console.error("❌ Max validation correction attempts reached. Exiting compiler pipeline.");
                process.exit(1);
            }

            console.log("🔄 Validation failed. Requesting self-correction from LLM...");
            const correctionPrompt = `Your previous JSON output failed validation with the following errors/warnings:\n\n${errorOutput}\n\nPlease fix the coordinate overlaps, duplicate slide titles (R7), or safe margins and output the fully corrected valid JSON object. Keep in mind the visual DNA rules for theme "${themeName}":\n${dnaPrompt}`;
            
            try {
                dslJsonText = await callLLM(correctionPrompt, SYSTEM_INSTRUCTION);
                dslJsonText = cleanJsonBlock(dslJsonText);
                attempt++;
            } catch (e) {
                console.error("❌ LLM Correction call failed:", e.message);
                process.exit(1);
            }
        }
    }

    // 4. Run the Compiler (Render PPTX and HTML)
    console.log("✨ Compilation Pipeline Triggered...");
    try {
        const compileOut = execSync(`node scripts/md2pptx_web.js ${inputMdPath} -o ${outputPptxPath} -t ${themeName}`, { stdio: 'inherit' });
        console.log(`🎉 Pipeline succeeded! Final files generated:`);
        console.log(`   - PPTX: ${outputPptxPath}`);
        console.log(`   - HTML: ${outputPptxPath.replace(/\.pptx$/i, '.html')}`);
        console.log(`   - DSL: ${targetDslPath}`);
    } catch (compileError) {
        console.error("❌ Compiler failed to render final output:", compileError.message);
        process.exit(1);
    }
}

function cleanJsonBlock(text) {
    let clean = text.trim();
    if (clean.startsWith('```')) {
        clean = clean.replace(/^```json/, '').replace(/^```/, '').replace(/```$/, '').trim();
    }
    return clean;
}

async function callGemini(prompt, systemInstruction) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${process.env.GEMINI_API_KEY}`;
    const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            systemInstruction: { parts: [{ text: systemInstruction }] },
            generationConfig: {
                responseMimeType: "application/json"
            }
        })
    });
    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Gemini API error: ${response.status} - ${errText}`);
    }
    const data = await response.json();
    return data.candidates[0].content.parts[0].text;
}

async function callOpenAI(prompt, systemInstruction) {
    const baseUrl = process.env.OPENAI_BASE_URL || 'https://api.openai.com/v1';
    const url = `${baseUrl}/chat/completions`;
    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${process.env.OPENAI_API_KEY}`
        },
        body: JSON.stringify({
            model: process.env.OPENAI_MODEL || 'gpt-4o',
            messages: [
                { role: 'system', content: systemInstruction },
                { role: 'user', content: prompt }
            ],
            response_format: { type: "json_object" }
        })
    });
    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`OpenAI API error: ${response.status} - ${errText}`);
    }
    const data = await response.json();
    return data.choices[0].message.content;
}

main().catch(err => {
    console.error("Fatal exception in pipeline:", err);
    process.exit(1);
});
