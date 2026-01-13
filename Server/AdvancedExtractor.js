/**
 * @fileoverview Advanced Slides API Extractor
 * Uses Slides.Presentations.get() to extract presentations with fully resolved colors.
 */

// ============================================================================
// CONSTANTS & LOGGING
// ============================================================================

const ADVANCED_EMU_PER_PT = 12700; // EMUs per point

// Verbose logging flag - set to true for detailed extraction logs
const VERBOSE_LOGGING = true;

function log(message, level) {
    level = level || 'INFO';
    if (VERBOSE_LOGGING || level === 'ERROR' || level === 'WARN') {
        Logger.log('[' + level + '] ' + message);
    }
}

// Global theme color map (populated during extraction)
let _themeColorMap = {};

// Global placeholder styles cache (Layout -> Master -> Defaults)
let _placeholderStyles = {};

// Global background index: layoutId -> background color, plus 'master' key for master background
let _backgroundIndex = {};

// Global cached page dimensions for background shape detection
let _pageWidth = 9144000;
let _pageHeight = 5143500;

/**
 * Helper to check if a shape covers most of the slide (background shape)
 * Must be called after buildBackgroundIndex sets _pageWidth/_pageHeight
 */
function isBackgroundShape(element) {
    if (!element.shape) return false;
    const transform = element.transform || {};
    const size = element.size || {};

    const scaleX = transform.scaleX !== undefined ? Math.abs(transform.scaleX) : 1;
    const scaleY = transform.scaleY !== undefined ? Math.abs(transform.scaleY) : 1;
    const width = (size.width?.magnitude || 0) * scaleX;
    const height = (size.height?.magnitude || 0) * scaleY;
    const x = transform.translateX || 0;
    const y = transform.translateY || 0;

    // Shape is a background if it covers >90% of the slide and starts near origin
    const coversWidth = width >= _pageWidth * 0.9;
    const coversHeight = height >= _pageHeight * 0.9;
    const nearOrigin = Math.abs(x) < _pageWidth * 0.1 && Math.abs(y) < _pageHeight * 0.1;

    return coversWidth && coversHeight && nearOrigin;
}

/**
 * Helper to extract background from a page's elements
 * Looks for large shapes that cover the slide background
 */
function extractBackgroundFromShapes(pageElements) {
    for (const el of (pageElements || [])) {
        if (isBackgroundShape(el)) {
            const fill = el.shape?.shapeProperties?.shapeBackgroundFill;
            if (fill && fill.propertyState !== 'NOT_RENDERED') {
                const color = extractFillAdvanced(fill);
                if (color && color !== 'transparent') {
                    return color;
                }
            }
        }
    }
    return null;
}

/**
 * Build an index of background colors from masters and layouts
 * This enables resolving inherited backgrounds when they're not set on a slide
 */
function buildBackgroundIndex(presentation) {
    _backgroundIndex = {};

    // Cache page dimensions for background shape detection
    _pageWidth = presentation.pageSize?.width?.magnitude || 9144000;
    _pageHeight = presentation.pageSize?.height?.magnitude || 5143500;

    // Extract master background (fallback for all slides)
    const masters = presentation.masters || [];
    if (masters.length > 0) {
        const master = masters[0];
        // Try pageBackgroundFill first
        const masterBg = master.pageProperties?.pageBackgroundFill;
        if (masterBg && masterBg.propertyState !== 'NOT_RENDERED') {
            _backgroundIndex['master'] = extractFillAdvanced(masterBg);
        }
        // Then try background shapes
        if (!_backgroundIndex['master'] || _backgroundIndex['master'] === 'transparent') {
            const shapeBg = extractBackgroundFromShapes(master.pageElements);
            if (shapeBg) {
                _backgroundIndex['master'] = shapeBg;
                Logger.log('Master BG from shape: ' + shapeBg);
            }
        }
    }

    // Extract layout backgrounds (keyed by layout objectId)
    const layouts = presentation.layouts || [];
    layouts.forEach(layout => {
        const key = 'layout:' + layout.objectId;

        // Try pageBackgroundFill first
        const layoutBg = layout.pageProperties?.pageBackgroundFill;
        if (layoutBg && layoutBg.propertyState !== 'NOT_RENDERED' && layoutBg.propertyState !== 'INHERIT') {
            const color = extractFillAdvanced(layoutBg);
            if (color && color !== 'transparent') {
                _backgroundIndex[key] = color;
            }
        }

        // Then try background shapes on the layout
        if (!_backgroundIndex[key]) {
            const shapeBg = extractBackgroundFromShapes(layout.pageElements);
            if (shapeBg) {
                _backgroundIndex[key] = shapeBg;
                Logger.log('Layout ' + layout.objectId + ' BG from shape: ' + shapeBg);
            }
        }
    });

    Logger.log('Background index built: ' + JSON.stringify(_backgroundIndex));
}

/**
 * Resolve a slide's background color, checking slide -> layout -> master hierarchy
 * Note: Google Slides API returns resolved values, so even inherited backgrounds
 * should have their actual color data in solidFill.
 */
function resolveSlideBackground(slide) {
    // pageBackgroundFill is under pageProperties, NOT slideProperties
    // slideProperties contains layoutObjectId, masterObjectId, notesPage, etc.
    const slideBg = slide.pageProperties?.pageBackgroundFill;
    const layoutId = slide.slideProperties?.layoutObjectId;

    // Log for debugging
    Logger.log('=== SLIDE BACKGROUND DEBUG ===');
    Logger.log('pageProperties keys: ' + (slide.pageProperties ? Object.keys(slide.pageProperties).join(', ') : 'undefined'));
    Logger.log('pageBackgroundFill: ' + JSON.stringify(slideBg));
    Logger.log('Layout ID: ' + layoutId);

    // 1. Try to extract from slide's pageBackgroundFill (even if inherited)
    // The API returns resolved values, so solidFill should have actual color
    if (slideBg && slideBg.propertyState !== 'NOT_RENDERED') {
        // Try solidFill first
        if (slideBg.solidFill?.color?.rgbColor) {
            const color = rgbToHexAdvanced(slideBg.solidFill.color.rgbColor);
            Logger.log('Slide BG from solidFill RGB: ' + color);
            return color;
        }
        if (slideBg.solidFill?.color?.themeColor) {
            const color = _themeColorMap[slideBg.solidFill.color.themeColor];
            if (color) {
                Logger.log('Slide BG from solidFill theme: ' + color);
                return color;
            }
        }
    }

    // 1.5. Check for background shape on the slide itself
    // This handles slides where a large shape is used as the background instead of pageBackgroundFill
    const shapeBg = extractBackgroundFromShapes(slide.pageElements);
    if (shapeBg) {
        Logger.log('Slide BG from slide shape: ' + shapeBg);
        return shapeBg;
    }

    // 2. Check layout background from our index
    if (layoutId && _backgroundIndex['layout:' + layoutId]) {
        const color = _backgroundIndex['layout:' + layoutId];
        if (color && color !== 'transparent') {
            Logger.log('Slide BG from layout index: ' + color);
            return color;
        }
    }

    // 3. Check master background from our index
    if (_backgroundIndex['master']) {
        const color = _backgroundIndex['master'];
        if (color && color !== 'transparent') {
            // Skip black master backgrounds as they're usually not intended
            Logger.log('Slide BG from master index: ' + color);
            return color;
        }
    }

    // 4. Default to white
    Logger.log('Slide BG defaulting to white');
    return '#ffffff';
}

/**
 * Build an index of placeholder styles from masters and layouts
 * This enables resolving inherited properties when they're undefined on an element
 */
function buildPlaceholderIndex(presentation) {
    _placeholderStyles = {};

    // Index master placeholders (lowest priority)
    const masters = presentation.masters || [];
    masters.forEach(master => {
        (master.pageElements || []).forEach(el => {
            if (el.shape?.placeholder?.type) {
                const key = 'master:' + el.shape.placeholder.type;
                _placeholderStyles[key] = extractPlaceholderDefaults(el);
            }
        });
    });

    // Index layout placeholders (override master)
    const layouts = presentation.layouts || [];
    layouts.forEach(layout => {
        (layout.pageElements || []).forEach(el => {
            if (el.shape?.placeholder?.type) {
                const key = 'layout:' + layout.objectId + ':' + el.shape.placeholder.type;
                _placeholderStyles[key] = extractPlaceholderDefaults(el);
            }
        });
    });

    Logger.log('Placeholder index built: ' + Object.keys(_placeholderStyles).length + ' entries');
}

/**
 * Extract default styles from a placeholder element
 */
function extractPlaceholderDefaults(element) {
    const shape = element.shape;
    const textElements = shape?.text?.textElements || [];
    const firstRun = textElements.find(e => e.textRun)?.textRun?.style || {};

    return {
        fontSize: firstRun.fontSize?.magnitude,
        fontFamily: firstRun.fontFamily,
        color: firstRun.foregroundColor ? extractFillAdvanced(firstRun.foregroundColor) : undefined,
        bold: firstRun.bold,
        italic: firstRun.italic
    };
}

/**
 * Resolve an inherited property using the Layout -> Master -> Default chain
 * In Raw Mode: Only returns explicit value, no inheritance lookup
 * @param {Object} element - The page element
 * @param {string} property - The property name (fontSize, fontFamily, etc.)
 * @param {string} layoutObjectId - The layout ID for this slide (optional)
 * @returns {*} The resolved value
 */
function resolveInheritedProperty(element, property, layoutObjectId) {
    // Layer 1: Explicit on element - check text runs
    const textElements = element.shape?.text?.textElements || [];
    const firstRun = textElements.find(e => e.textRun)?.textRun?.style || {};

    const explicitValue = getPropertyFromStyle(firstRun, property);
    if (explicitValue !== undefined && explicitValue !== null) {
        return explicitValue;
    }

    // In Raw Mode: Skip inheritance, return null to force explicit default
    if (_extractionRawMode) {
        return null; // Will be handled by caller with explicit defaults
    }

    // Layer 2 & 3: Check placeholder inheritance (only in non-raw mode)
    const placeholderType = element.shape?.placeholder?.type;
    if (placeholderType) {
        // Try layout placeholder first
        if (layoutObjectId) {
            const layoutKey = 'layout:' + layoutObjectId + ':' + placeholderType;
            const layoutStyle = _placeholderStyles[layoutKey];
            if (layoutStyle?.[property] !== undefined) {
                return layoutStyle[property];
            }
        }

        // Fall back to master placeholder
        const masterKey = 'master:' + placeholderType;
        const masterStyle = _placeholderStyles[masterKey];
        if (masterStyle?.[property] !== undefined) {
            return masterStyle[property];
        }
    }

    // Layer 4: Hardcoded defaults (only reached in non-raw mode)
    const defaults = {
        fontSize: 12,
        fontFamily: 'Arial',
        color: '#000000',
        bold: false,
        italic: false
    };
    return defaults[property];
}

/**
 * Helper to extract a property from a text style object
 */
function getPropertyFromStyle(style, property) {
    switch (property) {
        case 'fontSize':
            return style.fontSize?.magnitude;
        case 'fontFamily':
            return style.fontFamily;
        case 'color':
            if (style.foregroundColor?.opaqueColor?.rgbColor) {
                return rgbToHexAdvanced(style.foregroundColor.opaqueColor.rgbColor);
            }
            return undefined;
        case 'bold':
            return style.bold;
        case 'italic':
            return style.italic;
        default:
            return undefined;
    }
}

// ============================================================================
// PUBLIC API
// ============================================================================

/**
 * Extract a presentation using the Advanced Slides API
 * This provides fully resolved colors including for inherited/placeholder styles
 * @param {string} presentationIdOrUrl
 * @param {Object} options - { rawMode: boolean }
 * @returns {Object} JSON schema
 */
function extractPresentationAdvanced(presentationIdOrUrl, options) {
    options = options || {};
    const rawMode = options.rawMode || false;

    log('=== EXTRACTION START [BG-INHERIT-FIX] ===');
    log('Input: ' + presentationIdOrUrl);
    log('Raw Mode: ' + rawMode);

    let presentationId = presentationIdOrUrl;
    if (presentationIdOrUrl.includes('docs.google.com')) {
        const match = presentationIdOrUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
        if (match && match[1]) {
            presentationId = match[1];
        }
    }
    log('Presentation ID: ' + presentationId);

    const presentation = Slides.Presentations.get(presentationId);
    log('Title: ' + presentation.title);
    log('Slides: ' + (presentation.slides || []).length);
    log('Masters: ' + (presentation.masters || []).length);
    log('Layouts: ' + (presentation.layouts || []).length);

    // In Raw Mode: Skip placeholder inheritance, just use explicit styles
    if (!rawMode) {
        buildPlaceholderIndex(presentation);
    } else {
        _placeholderStyles = {}; // Clear any existing
        log('Raw Mode: Skipping placeholder inheritance');
    }

    // Always build background index (needed to resolve inherited backgrounds)
    _themeColorMap = extractThemeColorMap(presentation);
    buildBackgroundIndex(presentation);
    log('Theme colors extracted: ' + Object.keys(_themeColorMap).join(', '));

    const themeColors = {
        text: _themeColorMap['DARK1'] || '#000000',
        textLight: _themeColorMap['DARK2'] || '#595959',
        background: _themeColorMap['LIGHT1'] || '#ffffff',
        surface: _themeColorMap['LIGHT2'] || '#eeeeee',
        primary: _themeColorMap['ACCENT1'] || '#4285f4',
        secondary: _themeColorMap['ACCENT2'] || '#34a853',
        accent: _themeColorMap['ACCENT3'] || '#fbbc05',
        error: _themeColorMap['ACCENT4'] || '#ea4335',
        accent5: _themeColorMap['ACCENT5'] || '#46bdc6',
        accent6: _themeColorMap['ACCENT6'] || '#7baaf7',
        hyperlink: _themeColorMap['HYPERLINK'] || '#1a73e8',
        followedHyperlink: _themeColorMap['FOLLOWED_HYPERLINK'] || '#681da8'
    };

    // Store rawMode flag globally for use in sub-functions
    _extractionRawMode = rawMode;

    log('Processing slides...');
    const slides = (presentation.slides || []).map((slide, idx) => {
        log('--- Slide ' + (idx + 1) + ' ---');
        return extractSlideAdvanced(slide);
    });

    const json = {
        config: {
            title: presentation.title || 'input ' + presentationId,
            rawMode: rawMode,
            theme: {
                colors: themeColors,
                fonts: {
                    heading: 'Montserrat',
                    body: 'Open Sans'
                }
            }
        },
        slides: slides
    };

    return json;
}

// Global flag for raw mode during extraction
let _extractionRawMode = false;

// ============================================================================
// THEME EXTRACTION
// ============================================================================

function extractThemeColorMap(presentation) {
    const colorMap = {};
    try {
        const masters = presentation.masters;
        if (!masters || masters.length === 0) return colorMap;

        const master = masters[0];
        const colorScheme = master.masterProperties?.colorScheme?.colors || [];

        colorScheme.forEach(c => {
            if (c.type && c.color && c.color.rgbColor) {
                colorMap[c.type] = rgbToHexAdvanced(c.color.rgbColor);
            }
        });
    } catch (e) {
        Logger.log('Error extracting theme colors: ' + e.message);
    }
    return colorMap;
}

function resolveThemeColor(color) {
    if (!color) return '#000000';
    if (typeof color === 'string' && color.startsWith('#')) return color;
    if (typeof color === 'string' && color.startsWith('theme:')) {
        const themeKey = color.substring(6);
        return _themeColorMap[themeKey] || '#000000';
    }
    return color;
}

// ============================================================================
// SLIDE EXTRACTION
// ============================================================================

function extractSlideAdvanced(slide) {
    const slideData = {
        elements: [],
        background: resolveSlideBackground(slide)
    };

    log('  Background: ' + slideData.background);

    if (slide.slideProperties?.notesPage?.pageElements) {
        const notesShape = slide.slideProperties.notesPage.pageElements.find(
            el => el.shape?.placeholder?.type === 'BODY'
        );
        if (notesShape?.shape?.text?.textElements) {
            const notesText = extractPlainTextAdvanced(notesShape.shape.text.textElements);
            if (notesText.trim()) {
                slideData.speakerNotes = notesText.trim();
                log('  Speaker notes: ' + slideData.speakerNotes.substring(0, 50) + '...');
            }
        }
    }

    // Capture layout ID for inheritance resolution
    const layoutId = slide.slideProperties?.layoutObjectId;
    log('  Layout ID: ' + (layoutId || 'none'));

    const elements = slide.pageElements || [];
    log('  Page elements: ' + elements.length);

    elements.forEach((element, idx) => {
        // Pass layoutId down to elements
        const extracted = extractElementAdvanced(element, null, layoutId);
        if (extracted) {
            if (Array.isArray(extracted)) {
                log('    [' + idx + '] Group with ' + extracted.length + ' children');
                slideData.elements.push(...extracted);
            } else {
                log('    [' + idx + '] ' + extracted.type + ' "' + (extracted.text || extracted.objectId || '').substring(0, 30) + '"');
                slideData.elements.push(extracted);
            }
        }
    });

    log('  Total extracted: ' + slideData.elements.length + ' elements');
    return slideData;
}

// ============================================================================
// ELEMENT EXTRACTION
// ============================================================================

function composeTransforms(parent, child) {
    if (!parent) return child;
    if (!child) return parent;

    const pScaleX = parent.scaleX !== undefined ? parent.scaleX : 1;
    const pScaleY = parent.scaleY !== undefined ? parent.scaleY : 1;
    const pTranslateX = parent.translateX || 0;
    const pTranslateY = parent.translateY || 0;

    const cScaleX = child.scaleX !== undefined ? child.scaleX : 1;
    const cScaleY = child.scaleY !== undefined ? child.scaleY : 1;
    const cTranslateX = child.translateX || 0;
    const cTranslateY = child.translateY || 0;

    return {
        scaleX: pScaleX * cScaleX,
        scaleY: pScaleY * cScaleY,
        translateX: pTranslateX + cTranslateX * pScaleX,
        translateY: pTranslateY + cTranslateY * pScaleY,
        rotation: (parent.rotation || 0) + (child.rotation || 0)
    };
}

function extractElementAdvanced(element, parentTransform, layoutId) {
    const elementTransform = element.transform || {};
    const transform = parentTransform
        ? composeTransforms(parentTransform, elementTransform)
        : elementTransform;

    const scaleX = transform.scaleX !== undefined ? transform.scaleX : 1;
    const scaleY = transform.scaleY !== undefined ? transform.scaleY : 1;
    const translateX = transform.translateX || 0;
    const translateY = transform.translateY || 0;

    const baseWidth = element.size?.width?.magnitude || 0;
    const baseHeight = element.size?.height?.magnitude || 0;

    const actualWidth = baseWidth * Math.abs(scaleX);
    const actualHeight = baseHeight * Math.abs(scaleY);

    const base = {
        objectId: element.objectId,
        x: emuToPt(translateX),
        y: emuToPt(translateY),
        w: emuToPt(actualWidth),
        h: emuToPt(actualHeight),
        rotation: transform.rotation || 0
    };

    try {
        if (element.shape) {
            return extractShapeAdvanced(element, base, layoutId);
        }
        if (element.image) {
            return extractImageAdvanced(element, base);
        }
        if (element.table) {
            return extractTableAdvanced(element, base);
        }
        if (element.line) {
            return extractLineAdvanced(element, base);
        }
        if (element.sheetsChart) {
            return extractChartAdvanced(element, base);
        }
        if (element.elementGroup) {
            const children = element.elementGroup.children || [];
            const extracted = children
                .map(c => extractElementAdvanced(c, transform, layoutId)) // Pass layoutId recursively
                .filter(Boolean)
                .flat();
            return extracted.length > 0 ? extracted : null;
        }
        return null;
    } catch (e) {
        Logger.log('Error extracting element: ' + e.message);
        return null;
    }
}

// ============================================================================
// SHAPE EXTRACTION
// ============================================================================

function extractShapeAdvanced(element, base, layoutId) {
    const shape = element.shape;
    const shapeType = shape.shapeType || 'RECTANGLE';
    const isTextBox = shapeType === 'TEXT_BOX';

    // Extract text content
    let textContent = '';
    let textRuns = [];
    let textStyle = null;

    if (shape.text && shape.text.textElements) {
        const textData = extractTextAdvanced(shape.text.textElements);
        textContent = textData.plainText;
        textRuns = textData.runs;
        textStyle = textData.firstRunStyle;
    }

    const fill = shape.shapeProperties?.shapeBackgroundFill;
    const outline = shape.shapeProperties?.outline;

    // Resolve inherited properties if explicit ones are missing
    const resolvedFontSize = textStyle?.fontSize || resolveInheritedProperty(element, 'fontSize', layoutId);
    const resolvedFontFamily = textStyle?.fontFamily || resolveInheritedProperty(element, 'fontFamily', layoutId);
    const resolvedColor = textStyle?.color || resolveInheritedProperty(element, 'color', layoutId);
    const resolvedBold = textStyle?.bold !== undefined ? textStyle.bold : resolveInheritedProperty(element, 'bold', layoutId);

    if (isTextBox && textContent.length > 0) {
        const result = {
            type: 'text',
            ...base,
            text: textContent
        };

        result.color = resolveThemeColor(resolvedColor) || '#000000';
        result.fontSize = resolvedFontSize || 12;
        result.fontFamily = resolvedFontFamily || 'Arial';
        result.bold = resolvedBold || false;
        result.italic = textStyle?.italic || false;
        result.underline = textStyle?.underline || false;
        result.strikethrough = textStyle?.strikethrough || false;
        result.smallCaps = textStyle?.smallCaps || false;
        result.baselineOffset = textStyle?.baselineOffset || 'NONE';

        // Paragraph styles (alignment, bullets, indent) are now on individual textRuns
        // We can optionally set a default alignment if all runs match, but for now relies on runs.
        if (textRuns.length > 0 && textRuns[0].paragraphStyle) {
            result.align = textRuns[0].paragraphStyle.align || 'left';
            // Also helpful to set default specific props for single-para shapes
            result.indentStart = textRuns[0].paragraphStyle.indentStart || 0;
            result.lineSpacing = textRuns[0].paragraphStyle.lineSpacing || 100;
        }

        if (textRuns.length > 1) {
            result.textRuns = textRuns.map(run => ({
                ...run,
                color: resolveThemeColor(run.color)
            }));
        }

        const autofit = shape.shapeProperties?.autofit;
        if (autofit) {
            result.paddingTop = (autofit.topOffset || 0) / ADVANCED_EMU_PER_PT;
            result.paddingBottom = (autofit.bottomOffset || 0) / ADVANCED_EMU_PER_PT;
            result.paddingLeft = (autofit.leftOffset || 0) / ADVANCED_EMU_PER_PT;
            result.paddingRight = (autofit.rightOffset || 0) / ADVANCED_EMU_PER_PT;
        }

        return result;
    } else {
        const result = {
            type: 'shape',
            ...base,
            shape: shapeType,
            fillColor: resolveThemeColor(fill ? extractFillAdvanced(fill) : 'transparent'),
            borderColor: resolveThemeColor(outline ? extractOutlineColorAdvanced(outline) : 'none'),
            borderWidth: outline?.weight?.magnitude ? outline.weight.magnitude / ADVANCED_EMU_PER_PT : 0
        };

        if (textContent.length > 0) {
            result.text = textContent;
            // Also apply resolved styles to shape text
            result.fontSize = resolvedFontSize;
            result.fontFamily = resolvedFontFamily;
            result.color = resolveThemeColor(resolvedColor);

            if (textRuns.length > 1) {
                result.textRuns = textRuns.map(run => ({
                    ...run,
                    color: resolveThemeColor(run.color)
                }));
            }
        }

        const autofit = shape.shapeProperties?.autofit;
        if (autofit) {
            result.paddingTop = (autofit.topOffset || 0) / ADVANCED_EMU_PER_PT;
            result.paddingBottom = (autofit.bottomOffset || 0) / ADVANCED_EMU_PER_PT;
            result.paddingLeft = (autofit.leftOffset || 0) / ADVANCED_EMU_PER_PT;
            result.paddingRight = (autofit.rightOffset || 0) / ADVANCED_EMU_PER_PT;
        }

        return result;
    }
}

// ============================================================================
// TEXT EXTRACTION
// ============================================================================

/**
 * Extract text runs with support for paragraph styles (bullets, alignment)
 *
 * CRITICAL: In Google Slides API, text structure is:
 * - paragraphMarker elements PRECEDE their text content and represent
 *   "the start of a new paragraph" (per Google's official docs)
 * - The marker's index range spans the full paragraph including the newline
 * - textRun elements follow and contain the actual text content
 *
 * Example structure: [paragraphMarker, textRun "Header\n", paragraphMarker, textRun "Item\n"]
 * The first marker describes "Header\n", the second marker describes "Item\n"
 */
function extractTextAdvanced(textElements) {
    const runs = [];
    let plainText = '';
    let firstRunStyle = null;

    // CORRECTED APPROACH (marker-first):
    // 1. When we see a paragraphMarker, save it as "pending" for upcoming text runs
    // 2. Collect textRuns into currentParagraphRuns
    // 3. When we see the next paragraphMarker, apply pending marker to collected runs
    // 4. At the end, apply any pending marker to remaining runs

    // Runs in the current paragraph (will receive the pending marker)
    let currentParagraphRuns = [];

    // Pending paragraph marker to apply to the next text runs
    let pendingMarker = null;

    // State for inheritance across paragraphs
    let currentIndentStart = undefined;
    let currentIndentFirstLine = undefined;

    // Helper to convert units to points
    const getPt = (prop) => {
        if (!prop || prop.magnitude === undefined) return undefined;
        return (prop.unit === 'PT') ? prop.magnitude : (prop.magnitude / ADVANCED_EMU_PER_PT);
    };

    // Helper to build paragraph style and bullet data from a marker
    const buildParagraphData = (marker) => {
        const ps = marker.style || {};
        const bullet = marker.bullet;

        // Update inheritance state
        const explicitIndentStart = getPt(ps.indentStart);
        if (explicitIndentStart !== undefined) {
            currentIndentStart = explicitIndentStart;
        }

        const explicitIndentFirst = getPt(ps.indentFirstLine);
        if (explicitIndentFirst !== undefined) {
            currentIndentFirstLine = explicitIndentFirst;
        }

        // Adjust firstLine for non-bullet paragraphs with inherited hanging indent
        let effectiveFirstLine = currentIndentFirstLine;
        if (!bullet && (currentIndentStart > currentIndentFirstLine)) {
            effectiveFirstLine = currentIndentStart;
        }

        const pStyle = {
            align: mapAlignmentAdvanced(ps.alignment),
            direction: ps.direction,
            spacingMode: ps.spacingMode,
            spaceAbove: getPt(ps.spaceAbove),
            spaceBelow: getPt(ps.spaceBelow),
            lineSpacing: ps.lineSpacing || 100,
            indentStart: currentIndentStart,
            indentFirstLine: effectiveFirstLine
        };

        let bulletData = null;
        if (bullet) {
            bulletData = {
                listId: bullet.listId,
                nestingLevel: bullet.nestingLevel || 0,
                glyph: bullet.glyph
            };
        }

        return { pStyle, bulletData, bullet };
    };

    // Helper to apply marker data to runs
    const applyMarkerToRuns = (marker, runsToApply) => {
        if (!marker || runsToApply.length === 0) return;

        const { pStyle, bulletData, bullet } = buildParagraphData(marker);

        // DEBUG: Log for troubleshooting
        const firstText = runsToApply[0].text.substring(0, 20).replace(/\n/g, '\\n');
        log('    Para [' + firstText + '...]: bullet=' + (bullet ? 'YES(' + (bullet.glyph || 'no-glyph') + ')' : 'NO'));

        for (const run of runsToApply) {
            run.paragraphStyle = pStyle;
            run.bullet = bulletData;
        }
    };

    for (const elem of textElements) {
        if (elem.paragraphMarker) {
            // A new paragraph is starting.
            // First, finalize the previous paragraph (if any runs were collected).
            if (currentParagraphRuns.length > 0 && pendingMarker) {
                applyMarkerToRuns(pendingMarker, currentParagraphRuns);
            } else if (currentParagraphRuns.length > 0 && !pendingMarker) {
                // Runs without a marker (shouldn't happen, but handle gracefully)
                for (const run of currentParagraphRuns) {
                    run.paragraphStyle = { align: 'left' };
                    run.bullet = null;
                }
            }

            // Save this marker for the upcoming text runs
            pendingMarker = elem.paragraphMarker;
            currentParagraphRuns = [];
        }
        else if (elem.textRun) {
            const content = elem.textRun.content || '';
            if (content.length === 0) continue;

            plainText += content;

            const style = elem.textRun.style || {};
            let color = '#000000';
            if (style.foregroundColor?.opaqueColor?.rgbColor) {
                color = rgbToHexAdvanced(style.foregroundColor.opaqueColor.rgbColor);
            } else if (style.foregroundColor?.opaqueColor?.themeColor) {
                color = _themeColorMap[style.foregroundColor.opaqueColor.themeColor] || '#000000';
            }

            const fontSize = style.fontSize?.magnitude
                ? Math.round(style.fontSize.magnitude * 10) / 10
                : undefined;

            const runData = {
                text: content,
                color: color,
                fontSize: fontSize,
                fontFamily: style.fontFamily,
                bold: style.bold,
                italic: style.italic,
                underline: style.underline || false,
                strikethrough: style.strikethrough || false,
                smallCaps: style.smallCaps || false,
                paragraphStyle: null,
                bullet: null
            };

            if (style.link?.url) {
                runData.link = { url: style.link.url };
            }

            runs.push(runData);
            currentParagraphRuns.push(runData);

            if (!firstRunStyle) {
                firstRunStyle = runData;
            }
        }
    }

    // Finalize any remaining runs with the last pending marker
    if (currentParagraphRuns.length > 0) {
        if (pendingMarker) {
            applyMarkerToRuns(pendingMarker, currentParagraphRuns);
        } else {
            // No marker available - apply defaults
            const defaultPStyle = { align: 'left' };
            for (const run of currentParagraphRuns) {
                run.paragraphStyle = defaultPStyle;
                run.bullet = null;
            }
        }
    }

    // Cleanup trailing newlines
    if (runs.length > 0) {
        const lastRun = runs[runs.length - 1];
        if (lastRun.text.endsWith('\n')) {
            lastRun.text = lastRun.text.substring(0, lastRun.text.length - 1);
            if (lastRun.text.length === 0 && runs.length > 1) {
                runs.pop();
            }
        }
    }

    if (plainText.endsWith('\n')) {
        plainText = plainText.substring(0, plainText.length - 1);
    }

    return { plainText, runs, firstRunStyle };
}

function extractPlainTextAdvanced(textElements) {
    let text = '';
    for (const elem of textElements) {
        if (elem.textRun?.content) {
            text += elem.textRun.content;
        }
    }
    return text;
}

function extractParagraphStyleAdvanced(textElements) {
    for (const elem of textElements) {
        if (elem.paragraphMarker?.style) {
            const ps = elem.paragraphMarker.style;

            // Helper to get Points value
            const getPt = (prop) => {
                if (!prop || prop.magnitude === undefined) return 0;
                // If explicitly PT, return raw magnitude. Otherwise assume EMU and divide.
                return (prop.unit === 'PT') ? prop.magnitude : (prop.magnitude / ADVANCED_EMU_PER_PT);
            };

            return {
                align: mapAlignmentAdvanced(ps.alignment),
                indentStart: getPt(ps.indentStart),
                indentFirstLine: getPt(ps.indentFirstLine),
                spaceAbove: getPt(ps.spaceAbove),
                spaceBelow: getPt(ps.spaceBelow),
                lineSpacing: ps.lineSpacing || 100
            };
        }
    }
    return null;
}

function mapAlignmentAdvanced(alignment) {
    const map = {
        'START': 'left',
        'CENTER': 'center',
        'END': 'right',
        'JUSTIFIED': 'justify'
    };
    return map[alignment] || 'left';
}

// ============================================================================
// IMAGE EXTRACTION
// ============================================================================

function extractImageAdvanced(element, base) {
    const image = element.image;
    return {
        type: 'image',
        id: element.objectId,
        left: base.x,
        top: base.y,
        width: base.w,
        height: base.h,
        rotation: base.rotation,
        url: image.contentUrl || image.sourceUrl || '',
        sourceUrl: image.sourceUrl || null,
        originalWidth: base.w,
        originalHeight: base.h
    };
}

// ============================================================================
// TABLE EXTRACTION
// ============================================================================

function extractTableAdvanced(element, base) {
    const table = element.table;
    const data = [];

    const rows = table.tableRows || [];
    for (const row of rows) {
        const rowData = [];
        const cells = row.tableCells || [];

        for (const cell of cells) {
            let cellData = { text: '' };

            if (cell.text?.textElements) {
                const textData = extractTextAdvanced(cell.text.textElements);
                const firstStyle = textData.firstRunStyle || {};

                cellData = {
                    text: textData.plainText.trim(),
                    bold: firstStyle.bold || false,
                    italic: firstStyle.italic || false,
                    color: resolveThemeColor(firstStyle.color) || '#000000',
                    fontSize: firstStyle.fontSize || 12,
                    fontFamily: firstStyle.fontFamily || 'Arial',
                    align: 'center'
                };

                const paraStyle = extractParagraphStyleAdvanced(cell.text.textElements);
                if (paraStyle) {
                    cellData.align = paraStyle.align;
                }

                if (textData.runs.length > 1) {
                    cellData.textRuns = textData.runs.map(run => ({
                        ...run,
                        color: resolveThemeColor(run.color)
                    }));
                }
            }

            if (cell.tableCellProperties?.tableCellBackgroundFill) {
                cellData.fillColor = resolveThemeColor(
                    extractFillAdvanced(cell.tableCellProperties.tableCellBackgroundFill)
                );
            } else {
                cellData.fillColor = 'transparent';
            }

            rowData.push(cellData);
        }

        data.push(rowData);
    }

    return {
        type: 'table',
        ...base,
        data: data
    };
}

// ============================================================================
// LINE EXTRACTION
// ============================================================================

function extractLineAdvanced(element, base) {
    const line = element.line;
    let lineColor = '#000000';
    if (line.lineProperties?.lineFill?.solidFill?.color?.rgbColor) {
        lineColor = rgbToHexAdvanced(line.lineProperties.lineFill.solidFill.color.rgbColor);
    } else if (line.lineProperties?.lineFill?.solidFill?.color?.themeColor) {
        lineColor = resolveThemeColor('theme:' + line.lineProperties.lineFill.solidFill.color.themeColor);
    }

    return {
        type: 'line',
        ...base,
        color: lineColor,
        startArrow: line.lineProperties?.startArrow || 'NONE',
        endArrow: line.lineProperties?.endArrow || 'NONE',
        weight: line.lineProperties?.weight?.magnitude
            ? line.lineProperties.weight.magnitude / ADVANCED_EMU_PER_PT
            : 1,
        dashStyle: line.lineProperties?.dashStyle || 'SOLID',
        startConnect: null,
        endConnect: null
    };
}

// ============================================================================
// CHART EXTRACTION
// ============================================================================

function extractChartAdvanced(element, base) {
    const chart = element.sheetsChart;

    // Check if the chart has a content URL (image representation)
    // This is available for embedded/non-linked charts
    const contentUrl = chart.contentUrl;

    // Try to determine if we can access the source spreadsheet
    const spreadsheetId = chart.spreadsheetId || '';
    const chartId = chart.chartId || 0;

    // If we have a content URL but no spreadsheet access, extract as image
    if (contentUrl && !spreadsheetId) {
        Logger.log('Chart has contentUrl but no spreadsheetId - extracting as image');
        return {
            type: 'image',
            ...base,
            url: contentUrl,
            sourceType: 'chart',
            originalChartId: chartId
        };
    }

    // If we have a spreadsheet ID, try to return as sheetsChart
    // The generation side will handle access errors gracefully
    if (spreadsheetId) {
        return {
            type: 'sheetsChart',
            ...base,
            spreadsheetId: spreadsheetId,
            chartId: chartId,
            embedType: 'IMAGE',
            // Include contentUrl as fallback for generation
            contentUrl: contentUrl || null
        };
    }

    // Fallback: If we have contentUrl, use it as image
    if (contentUrl) {
        Logger.log('Chart extraction fallback to image: ' + contentUrl);
        return {
            type: 'image',
            ...base,
            url: contentUrl,
            sourceType: 'chart'
        };
    }

    // Last resort: return placeholder
    Logger.log('Chart has no extractable content - returning placeholder');
    return {
        type: 'shape',
        ...base,
        shape: 'RECTANGLE',
        fillColor: '#f0f0f0',
        text: 'Chart (Not Extractable)',
        fontSize: 12,
        color: '#888888'
    };
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function extractFillAdvanced(fill) {
    if (!fill) return 'transparent';
    if (fill.propertyState === 'NOT_RENDERED') return 'transparent';

    if (fill.solidFill?.color?.rgbColor) {
        return rgbToHexAdvanced(fill.solidFill.color.rgbColor);
    }

    if (fill.solidFill?.color?.themeColor) {
        // Visual Accuracy: Resolve theme color to hex immediately
        return _themeColorMap[fill.solidFill.color.themeColor] || 'transparent';
    }

    return 'transparent';
}

function extractOutlineColorAdvanced(outline) {
    if (!outline) return 'none';
    if (outline.propertyState === 'NOT_RENDERED') return 'none';
    if (!outline.outlineFill?.solidFill) return 'none';

    const solidFill = outline.outlineFill.solidFill;
    if (solidFill.color?.rgbColor) {
        return rgbToHexAdvanced(solidFill.color.rgbColor);
    }

    if (solidFill.color?.themeColor) {
        // Visual Accuracy: Resolve theme color to hex immediately
        return _themeColorMap[solidFill.color.themeColor] || 'none';
    }

    return 'none';
}

function rgbToHexAdvanced(rgb) {
    if (!rgb) return '#000000';
    const r = Math.round((rgb.red || 0) * 255);
    const g = Math.round((rgb.green || 0) * 255);
    const b = Math.round((rgb.blue || 0) * 255);
    return '#' + [r, g, b].map(x => x.toString(16).padStart(2, '0')).join('');
}

function emuToPt(emu) {
    return Math.round(emu / ADVANCED_EMU_PER_PT);
}
