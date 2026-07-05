const fs = require('fs');
const path = require('path');

/**
 * Validates a single slide DSL object
 */
function validateSlideDsl(slide, slideIdx, baseDir) {
    const errors = [];
    const warnings = [];
    
    if (typeof slide !== 'object' || slide === null) {
        return [`Slide ${slideIdx}: must be an object.`];
    }
    
    // Rhythm check
    const rhythm = slide.rhythm || 'dense';
    if (slide.rhythm && !['anchor', 'dense', 'breathing'].includes(slide.rhythm)) {
        errors.push(`Slide ${slideIdx}: invalid rhythm "${slide.rhythm}". Valid: anchor, dense, breathing.`);
    }

    // Validate background (optional)
    if (slide.background) {
        if (typeof slide.background !== 'object') {
            errors.push(`Slide ${slideIdx}: background must be an object.`);
        } else {
            if (slide.background.image && typeof slide.background.image !== 'string') {
                errors.push(`Slide ${slideIdx}: background.image must be a string.`);
            }
            if (slide.background.fallbackColor && typeof slide.background.fallbackColor !== 'string') {
                errors.push(`Slide ${slideIdx}: background.fallbackColor must be a string.`);
            }
        }
    }
    
    // Validate elements
    if (!Array.isArray(slide.elements)) {
        errors.push(`Slide ${slideIdx}: elements must be an array.`);
        return errors;
    }

    const parsedElements = [];
    
    slide.elements.forEach((el, elIdx) => {
        const pathPrefix = `Slide ${slideIdx}, Element ${elIdx} (${el.id || 'unnamed'})`;
        
        if (!el.type) {
            errors.push(`${pathPrefix}: missing "type".`);
            return;
        }
        
        const validTypes = ['text', 'card', 'image', 'vector', 'timeline', 'decoration', 'chart'];
        if (!validTypes.includes(el.type)) {
            errors.push(`${pathPrefix}: invalid type "${el.type}". Valid: ${validTypes.join(', ')}.`);
        }
        
        // Validate bounds
        let hasValidBounds = true;
        if (!el.bounds) {
            errors.push(`${pathPrefix}: missing "bounds".`);
            hasValidBounds = false;
        } else {
            const boundsKeys = ['x', 'y', 'w', 'h'];
            boundsKeys.forEach(k => {
                if (typeof el.bounds[k] !== 'string' || !el.bounds[k].endsWith('%')) {
                    errors.push(`${pathPrefix}: bounds.${k} must be a percentage string (e.g. "10%").`);
                    hasValidBounds = false;
                } else {
                    const val = parseFloat(el.bounds[k]);
                    const isBleedingAllowed = (el.type === 'decoration' || el.type === 'vector');
                    const minVal = isBleedingAllowed ? -50 : 0;
                    const maxVal = isBleedingAllowed ? 200 : 100;
                    if (isNaN(val) || val < minVal || val > maxVal) {
                        errors.push(`${pathPrefix}: bounds.${k} value "${el.bounds[k]}" must be between ${minVal}% and ${maxVal}%.`);
                        hasValidBounds = false;
                    }
                }
            });
        }

        if (hasValidBounds) {
            const xe = parseFloat(el.bounds.x);
            const ye = parseFloat(el.bounds.y);
            const we = parseFloat(el.bounds.w);
            const he = parseFloat(el.bounds.h);

            // Edge checks
            const isBleedingAllowed = (el.type === 'decoration' || el.type === 'vector');
            if (!isBleedingAllowed) {
                if (xe + we > 100.5) {
                    errors.push(`${pathPrefix}: extends beyond right slide edge (x=${el.bounds.x} + w=${el.bounds.w} = ${(xe+we).toFixed(1)}% > 100%).`);
                }
                if (ye + he > 100.5) {
                    errors.push(`${pathPrefix}: extends beyond bottom slide edge (y=${el.bounds.y} + h=${el.bounds.h} = ${(ye+he).toFixed(1)}% > 100%).`);
                }
            }

            // Save for overlap/margin checks
            parsedElements.push({
                idx: elIdx,
                id: el.id || `el_${elIdx}`,
                type: el.type,
                x: xe, y: ye, w: we, h: he,
                area: we * he,
                raw: el
            });
        }
        
        // Type-specific validations
        if (el.type === 'text') {
            if (typeof el.content !== 'string' && typeof el.content !== 'object') {
                errors.push(`${pathPrefix}: text elements must have a string or object in "content".`);
            }
        } else if (el.type === 'card') {
            if (!el.content || typeof el.content !== 'object') {
                errors.push(`${pathPrefix}: card elements must have an object "content" containing "title" and "body".`);
            } else {
                if (typeof el.content.title !== 'string') {
                    errors.push(`${pathPrefix}: card.content.title must be a string.`);
                }
                if (typeof el.content.body !== 'string') {
                    errors.push(`${pathPrefix}: card.content.body must be a string.`);
                }
            }
            // Check style.variant if provided
            if (el.style?.variant && !['default', 'filled', 'glass', 'naked', 'stat'].includes(el.style.variant)) {
                errors.push(`${pathPrefix}: invalid card variant "${el.style.variant}".`);
            }
        } else if (el.type === 'image') {
            if (!el.content || typeof el.content.path !== 'string') {
                errors.push(`${pathPrefix}: image elements must have content.path string.`);
            } else if (baseDir) {
                const imgPath = path.isAbsolute(el.content.path) ? el.content.path : path.resolve(baseDir, el.content.path);
                if (!fs.existsSync(imgPath)) {
                    errors.push(`${pathPrefix}: image file not found: "${el.content.path}" (resolved to ${imgPath}).`);
                }
            }
            if (el.content.alt !== undefined && typeof el.content.alt !== 'string') {
                errors.push(`${pathPrefix}: image.content.alt must be a string if provided.`);
            }
        } else if (el.type === 'vector') {
            if (typeof el.name !== 'string') {
                errors.push(`${pathPrefix}: vector elements must have a "name" string.`);
            }
        } else if (el.type === 'timeline') {
            if (!Array.isArray(el.content)) {
                errors.push(`${pathPrefix}: timeline elements must have an array "content".`);
            } else {
                el.content.forEach((item, itemIdx) => {
                    if (typeof item.title !== 'string' || typeof item.body !== 'string') {
                        errors.push(`${pathPrefix}: timeline.content[${itemIdx}] must contain title and body strings.`);
                    }
                });
            }
        } else if (el.type === 'decoration') {
            if (typeof el.name !== 'string') {
                errors.push(`${pathPrefix}: decoration elements must have a "name" string.`);
            }
        }
    });

        // Detect duplicate title rendering in elements (R7: Header Title Duplicate Check)
    if (slide.title && slide.elements) {
        slide.elements.forEach((el, elIdx) => {
            if (el.type === 'text' && typeof el.content === 'string') {
                const cleanElContent = el.content.replace(/[\s\n\#\-\*]/g, '').toLowerCase();
                const cleanSlideTitle = slide.title.replace(/[\s\n\#\-\*]/g, '').toLowerCase();
                if (cleanElContent === cleanSlideTitle && parseFloat(el.bounds.y) < 15) {
                    errors.push(`Slide ${slideIdx}: Duplicate title rendering detected! Element "${el.id || 'unnamed'}" renders the slide title "${el.content}" manually inside the header zone (y=${el.bounds.y}). The compiler automatically renders H2 titles; remove this manual text element.`);
                }
            }
        });
    }


    // --- R1: Overlap Detection ---
    for (let i = 0; i < parsedElements.length; i++) {
        for (let j = i + 1; j < parsedElements.length; j++) {
            const a = parsedElements[i];
            const b = parsedElements[j];
            
            // Skip overlap checks involving decorations and vectors (decorations and vectors are allowed to overlay/underlay)
            if (a.type === 'decoration' || b.type === 'decoration' || a.type === 'vector' || b.type === 'vector') continue;
            
            // Calculate overlap rectangle
            const overlapX = Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x));
            const overlapY = Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y));
            const overlapArea = overlapX * overlapY;
            const smallerArea = Math.min(a.area, b.area);
            
            if (smallerArea > 0 && (overlapArea / smallerArea) > 0.05) {
                const pct = ((overlapArea / smallerArea) * 100).toFixed(0);
                errors.push(`Slide ${slideIdx}: Overlap detected! Element "${a.id}" and "${b.id}" overlap by ${pct}% of their smaller area.`);
            }
        }
    }

    // --- R2: Content Overflow Risk (for Cards) ---
    parsedElements.forEach(el => {
        if (el.type === 'card' && el.raw.content) {
            const textLength = (el.raw.content.title || '').length + (el.raw.content.body || '').length;
            const w_in = (el.w / 100) * 13.333;
            const h_in = (el.h / 100) * 7.5;
            const area_sq_in = w_in * h_in;
            const charLimit = Math.round(area_sq_in * 90); // ~90 chars/sq inch threshold
            if (textLength > charLimit && el.raw.style?.variant !== 'stat') {
                warnings.push(`Slide ${slideIdx}, Element "${el.id}": Text content (${textLength} chars) may overflow physical card space (${area_sq_in.toFixed(1)} sq inches, threshold is ~${charLimit} chars).`);
            }
        }
    });

    // --- R3: Rhythm Compliance ---
    if (rhythm === 'breathing') {
        const cardCount = parsedElements.filter(e => e.type === 'card').length;
        if (cardCount > 1) {
            errors.push(`Slide ${slideIdx}: "breathing" rhythm slide cannot contain more than 1 card element. Found: ${cardCount} cards.`);
        }
    }

    // --- R4: Safe Margins ---
    parsedElements.forEach(el => {
        if (el.type === 'decoration') return; // Skip decorations
        if (el.x < 4.0) {
            warnings.push(`Slide ${slideIdx}, Element "${el.id}": x coordinate (${el.x}%) is too close to left edge (safety margin is 4%).`);
        }
        if (el.x + el.w > 96.0) {
            warnings.push(`Slide ${slideIdx}, Element "${el.id}": x+w coordinate (${(el.x+el.w).toFixed(1)}%) is too close to right edge (safety margin is 96%).`);
        }
        if (slideIdx > 0 && rhythm !== 'anchor') { // Content pages only
            if (el.y < 12.0) {
                warnings.push(`Slide ${slideIdx}, Element "${el.id}": y coordinate (${el.y}%) is too close to top edge (safety margin is 12% to avoid title overlap).`);
            }
            if (el.y + el.h > 93.0) {
                warnings.push(`Slide ${slideIdx}, Element "${el.id}": y+h coordinate (${(el.y+el.h).toFixed(1)}%) is too close to bottom edge (safety margin is 93% to avoid footer overlap).`);
            }
        }
    });

    // --- R6: Visual Balance ---
    if (parsedElements.length > 0) {
        const contentElements = parsedElements.filter(e => e.type !== 'decoration');
        if (contentElements.length > 0) {
            let totalWeight = 0;
            let weightedSumX = 0;
            contentElements.forEach(el => {
                totalWeight += el.area;
                weightedSumX += (el.x + el.w / 2) * el.area;
            });
            const cx = weightedSumX / totalWeight;
            if (cx < 25 || cx > 75) {
                warnings.push(`Slide ${slideIdx}: Visual gravity center is highly unbalanced (cx = ${cx.toFixed(0)}%). Consider centering your layout.`);
            }
        }
    }

    // Log warnings to stderr
    if (warnings.length > 0) {
        warnings.forEach(warn => console.warn(`⚠️ Warning: ${warn}`));
    }
    
    return errors;
}

/**
 * Validates a full deck JSON representation
 */
function validateDeck(deck, baseDir) {
    const errors = [];
    if (typeof deck !== 'object' || deck === null || !Array.isArray(deck.slides)) {
        return ["Deck must contain a 'slides' array at root."];
    }
    
    deck.slides.forEach((slide, idx) => {
        const slideErrors = validateSlideDsl(slide, idx, baseDir);
        errors.push(...slideErrors);
    });
    
    return errors;
}

// If run from command line
if (require.main === module) {
    const args = process.argv.slice(2);
    if (args.length === 0) {
        console.log("Usage: node tests/validate_dsl.js <path_to_dsl.json>");
        process.exit(1);
    }
    
    try {
        const fileContent = fs.readFileSync(args[0], 'utf-8');
        const json = JSON.parse(fileContent);
        const errors = validateDeck(json, path.dirname(args[0]));
        
        if (errors.length > 0) {
            console.error("❌ DSL Validation Failed:");
            errors.forEach(err => console.error("  - " + err));
            process.exit(1);
        } else {
            console.log("✅ DSL JSON matches schema perfectly!");
            process.exit(0);
        }
    } catch (err) {
        console.error("❌ Failed to parse JSON file:", err.message);
        process.exit(1);
    }
}

module.exports = {
    validateSlideDsl,
    validateDeck
};
