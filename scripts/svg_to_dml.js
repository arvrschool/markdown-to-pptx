// scripts/svg_to_dml.js
const fs = require('fs');
const path = require('path');

/**
 * Translates SVG XML-like structural properties (or JSON Layout DSL)
 * into high-fidelity pptxgenjs slide layout render options.
 */
class DmlTranslator {
    constructor(theme) {
        this.theme = theme;
    }

    /**
     * Converts percentage coordinates to inches for 16:9 widescreen canvas (13.333 in * 7.5 in).
     */
    parsePct(val, maxVal) {
        if (typeof val === 'number') return val;
        if (typeof val === 'string') {
            if (val.endsWith('%')) {
                return (parseFloat(val) / 100) * maxVal;
            }
            return parseFloat(val);
        }
        return 0;
    }

    /**
     * Translates a DSL element into pptxgenjs shape/text commands.
     */
    translateElement(slide, el, theme) {
        const x = this.parsePct(el.bounds.x, 13.333);
        const y = this.parsePct(el.bounds.y, 7.5);
        const w = this.parsePct(el.bounds.w, 13.333);
        const h = this.parsePct(el.bounds.h, 7.5);

        // Map colors helper
        const getColorVal = (color) => {
            if (color === 'activeCyan' || color === 'active-cyan') return theme.accentColor;
            if (color === 'activePurple' || color === 'active-purple') return 'A78BFA';
            if (color === 'text-main') return theme.textColor;
            if (color === 'text-muted') return 'A0AEC0';
            if (color === 'neon-green') return '10B981';
            if (color === 'neon-orange') return 'F59E0B';
            if (color && color !== 'default') return color.replace('#', '');
            return theme.textColor;
        };

        if (el.type === 'text') {
            const fontSize = el.style?.fontSize || 12;
            const fontWeight = el.style?.fontWeight || 'normal';
            const align = el.style?.align || 'left';
            const valign = el.style?.valign || 'top';
            const color = getColorVal(el.style?.color);

            // Renders standard editable text
            slide.addText(el.content || '', {
                x: x, y: y, w: w, h: h,
                fontSize: fontSize,
                bold: fontWeight === 'bold',
                color: color,
                align: align,
                valign: valign,
                fontFace: theme.fontFace,
                margin: 0
            });
        } 
        else if (el.type === 'card') {
            const variant = el.style?.variant || 'default';
            const isAccent = el.style?.theme === 'accent' || variant === 'filled';
            
            // Premium Card styling: Renders native soft shadows
            const cardOpts = {
                x: x, y: y, w: w, h: h,
                fill: { color: variant === 'filled' ? theme.accentColor : theme.cardBg },
                line: { color: isAccent ? theme.accentColor : theme.cardBorder, width: 1.5 },
                shadow: {
                    type: 'outer',
                    color: '1E293B',
                    blur: 12,
                    offset: 4,
                    angle: 90,
                    opacity: 0.06
                }
            };
            slide.addShape('roundRect', cardOpts);

            // Premium Accent Top Bar (for non-spatial themes)
            const isSpatial = (theme.bg === '060E11' || theme.bg === '060e11');
            if (variant === 'default' && !isSpatial) {
                slide.addShape('rect', {
                    x: x, y: y, w: w, h: 0.08,
                    fill: { color: isAccent ? theme.accentColor : theme.accentColor }
                });
            }

            // Component Label/ID
            slide.addText(el.id ? el.id.toUpperCase().replace(/_/g, ' ') : 'COMPONENT', {
                x: x + 0.15, y: y + 0.1, w: w - 0.3, h: 0.25,
                fontSize: 8, bold: true, color: theme.textColor, fontFace: theme.fontFace, margin: 0
            });

            // Card Title
            slide.addText(el.content?.title || '', {
                x: x + 0.15, y: y + 0.35, w: w - 0.3, h: 0.3,
                fontSize: 13, bold: true, color: isAccent ? (variant === 'filled' ? 'FFFFFF' : theme.accentColor) : theme.titleColor,
                fontFace: theme.fontFace, margin: 0
            });

            // Card Divider Line
            slide.addShape('line', {
                x: x + 0.15, y: y + 0.7, w: w - 0.3, h: 0,
                line: { color: theme.dividerColor, width: 1 }
            });

            // Card Body Text
            slide.addText(el.content?.body || '', {
                x: x + 0.15, y: y + 0.8, w: w - 0.3, h: h - 0.95,
                fontSize: 10, color: variant === 'filled' ? 'FFFFFF' : theme.textColor,
                fontFace: theme.fontFace, valign: 'top', margin: 0
            });
        }
        else if (el.type === 'decoration') {
            const decType = el.name;
            const fillHex = getColorVal(el.style?.fill || '#FFFFFF');
            const opacity = el.style?.opacity !== undefined ? el.style.opacity : 0.15;
            const trans = Math.round((1 - opacity) * 100);

            if (decType === 'circle-ring') {
                // Large double rings
                slide.addShape('oval', {
                    x: x, y: y, w: w, h: h,
                    fill: { color: 'FFFFFF', transparency: 100 },
                    line: { color: fillHex, width: 1.5, transparency: trans }
                });
                slide.addShape('oval', {
                    x: x + w * 0.1, y: y + h * 0.1, w: w * 0.8, h: h * 0.8,
                    fill: { color: 'FFFFFF', transparency: 100 },
                    line: { color: fillHex, width: 0.8, transparency: Math.min(100, trans + 10) }
                });
            }
            else if (decType === 'hexagons') {
                // Hexagon shape groups
                slide.addShape('hexagon', {
                    x: x, y: y, w: w, h: h,
                    fill: { color: 'FFFFFF', transparency: 100 },
                    line: { color: fillHex, width: 1.2, transparency: trans }
                });
                slide.addShape('hexagon', {
                    x: x + w * 0.2, y: y + h * 0.1, w: w * 0.6, h: h * 0.8,
                    fill: { color: 'FFFFFF', transparency: 100 },
                    line: { color: fillHex, width: 0.8, transparency: Math.min(100, trans + 15) }
                });
            }
            else if (decType === 'glow-spot') {
                // Soft gradient radial glow emulator
                slide.addShape('oval', {
                    x: x, y: y, w: w, h: h,
                    fill: { color: fillHex, transparency: Math.max(0, Math.min(100, trans)) },
                    line: { color: 'FFFFFF', transparency: 100 }
                });
            }
            else if (decType === 'diagonal-split') {
                // Diagonal full block division
                slide.addShape('rtTriangle', {
                    x: x, y: y, w: w, h: h,
                    fill: { color: fillHex, transparency: trans },
                    line: { color: 'FFFFFF', transparency: 100 },
                    flipH: true, flipV: true
                });
            }
            else if (decType === 'cross-marker') {
                // Target reticle
                const cx = x + w / 2;
                const cy = y + h / 2;
                slide.addShape('line', { x: cx - 0.2, y: cy, w: 0.4, h: 0, line: { color: fillHex, width: 1, transparency: trans } });
                slide.addShape('line', { x: cx, y: cy - 0.2, w: 0, h: 0.4, line: { color: fillHex, width: 1, transparency: trans } });
            }
            else if (decType === 'separator') {
                // Dashed separator line
                slide.addShape('line', {
                    x: x, y: y, w: w, h: h,
                    line: { color: fillHex, width: 1.2, dashType: el.style?.dashType || 'dash', transparency: trans }
                });
            }
        }
    }
}

module.exports = DmlTranslator;
