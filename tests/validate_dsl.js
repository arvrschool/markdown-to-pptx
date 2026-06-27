const fs = require('fs');
const path = require('path');

/**
 * Validates a single slide DSL object
 */
function validateSlideDsl(slide, slideIdx, baseDir) {
    const errors = [];
    
    if (typeof slide !== 'object' || slide === null) {
        return [`Slide ${slideIdx}: must be an object.`];
    }
    
    // Validate background (optional)
    if (slide.background) {
        if (typeof slide.background !== 'object') {
            errors.push(`Slide ${slideIdx}: background must be an object.`);
        } else {
            if (slide.background.image && typeof slide.background.image !== 'string') {
                errors.push(`Slide ${slideIdx}: background.image must be a string.`);
            }
            if (slide.background.color && typeof slide.background.color !== 'string') {
                errors.push(`Slide ${slideIdx}: background.color must be a string.`);
            }
        }
    }
    
    // Validate elements
    if (!Array.isArray(slide.elements)) {
        errors.push(`Slide ${slideIdx}: elements must be an array.`);
        return errors;
    }
    
    slide.elements.forEach((el, elIdx) => {
        const pathPrefix = `Slide ${slideIdx}, Element ${elIdx} (${el.id || 'unnamed'})`;
        
        if (!el.type) {
            errors.push(`${pathPrefix}: missing "type".`);
            return;
        }
        
        const validTypes = ['text', 'card', 'image', 'vector', 'timeline'];
        if (!validTypes.includes(el.type)) {
            errors.push(`${pathPrefix}: invalid type "${el.type}". Valid: ${validTypes.join(', ')}.`);
        }
        
        // Validate bounds
        if (!el.bounds) {
            errors.push(`${pathPrefix}: missing "bounds".`);
        } else {
            const boundsKeys = ['x', 'y', 'w', 'h'];
            boundsKeys.forEach(k => {
                if (typeof el.bounds[k] !== 'string' || !el.bounds[k].endsWith('%')) {
                    errors.push(`${pathPrefix}: bounds.${k} must be a percentage string (e.g. "10%").`);
                } else {
                    const val = parseFloat(el.bounds[k]);
                    if (isNaN(val) || val < 0 || val > 100) {
                        errors.push(`${pathPrefix}: bounds.${k} value "${el.bounds[k]}" must be between 0% and 100%.`);
                    }
                }
            });
            // Overflow check: element extending beyond slide edge (>100.5% to allow floating rounding)
            const _xe = parseFloat(el.bounds.x), _ye = parseFloat(el.bounds.y);
            const _we = parseFloat(el.bounds.w), _he = parseFloat(el.bounds.h);
            if (!isNaN(_xe) && !isNaN(_we) && _xe + _we > 100.5) {
                errors.push(`${pathPrefix}: extends beyond right slide edge (x=${el.bounds.x} + w=${el.bounds.w} = ${(_xe+_we).toFixed(1)}% > 100%).`);
            }
            if (!isNaN(_ye) && !isNaN(_he) && _ye + _he > 100.5) {
                errors.push(`${pathPrefix}: extends beyond bottom slide edge (y=${el.bounds.y} + h=${el.bounds.h} = ${(_ye+_he).toFixed(1)}% > 100%).`);
            }
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
        }
    });
    
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
