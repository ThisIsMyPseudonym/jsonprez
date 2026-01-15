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

// Known Google Slides built-in theme color schemes
// These are predefined themes where colorScheme isn't returned by the API
const KNOWN_THEME_COLORS = {
    'Simple Light': {
        'DARK1': '#000000',
        'DARK2': '#434343',
        'LIGHT1': '#ffffff',
        'LIGHT2': '#f3f3f3',
        'ACCENT1': '#4285f4',
        'ACCENT2': '#ea4335',
        'ACCENT3': '#fbbc04',
        'ACCENT4': '#34a853',
        'ACCENT5': '#ff6d01',
        'ACCENT6': '#46bdc6',
        'HYPERLINK': '#1155cc',
        'FOLLOWED_HYPERLINK': '#1155cc'
    },
    'Parallax': {
        'DARK1': '#212121',
        'DARK2': '#616161',
        'LIGHT1': '#ffffff',
        'LIGHT2': '#f5f5f5',
        'ACCENT1': '#009688',
        'ACCENT2': '#ff5722',
        'ACCENT3': '#ffc107',
        'ACCENT4': '#8bc34a',
        'ACCENT5': '#03a9f4',
        'ACCENT6': '#e91e63',
        'HYPERLINK': '#009688',
        'FOLLOWED_HYPERLINK': '#009688'
    },
    // Google's Material Design / Blue themes
    'Material': {
        'DARK1': '#2196f3',  // Blue - common in Google themes
        'DARK2': '#1976d2',
        'LIGHT1': '#ffffff',
        'LIGHT2': '#e3f2fd',
        'ACCENT1': '#2196f3',
        'ACCENT2': '#f44336',
        'ACCENT3': '#ffeb3b',
        'ACCENT4': '#4caf50',
        'ACCENT5': '#ff9800',
        'ACCENT6': '#9c27b0',
        'HYPERLINK': '#1976d2',
        'FOLLOWED_HYPERLINK': '#1976d2'
    },
    'Swiss': {
        'DARK1': '#2196f3',  // Blue header theme
        'DARK2': '#424242',
        'LIGHT1': '#ffffff',
        'LIGHT2': '#f5f5f5',
        'ACCENT1': '#db4437',
        'ACCENT2': '#4285f4',
        'ACCENT3': '#f4b400',
        'ACCENT4': '#0f9d58',
        'ACCENT5': '#ab47bc',
        'ACCENT6': '#00acc1',
        'HYPERLINK': '#4285f4',
        'FOLLOWED_HYPERLINK': '#4285f4'
    }
};

// Default theme colors for when colorScheme isn't available and theme is unknown
const DEFAULT_THEME_COLORS = {
    'DARK1': '#000000',      // Typically black
    'DARK2': '#595959',      // Typically dark gray
    'LIGHT1': '#ffffff',     // Typically white
    'LIGHT2': '#eeeeee',     // Typically light gray
    'ACCENT1': '#4285f4',    // Blue
    'ACCENT2': '#34a853',    // Green
    'ACCENT3': '#fbbc05',    // Yellow
    'ACCENT4': '#ea4335',    // Red
    'ACCENT5': '#46bdc6',    // Cyan
    'ACCENT6': '#7baaf7',    // Light blue
    'HYPERLINK': '#1a73e8',
    'FOLLOWED_HYPERLINK': '#681da8'
};

// Active theme colors - set during extraction based on presentation's theme
let _activeThemeColors = DEFAULT_THEME_COLORS;

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
            const themeKey = slideBg.solidFill.color.themeColor;
            const color = _themeColorMap[themeKey] || _activeThemeColors[themeKey] || DEFAULT_THEME_COLORS[themeKey];
            if (color) {
                Logger.log('Slide BG from solidFill theme: ' + themeKey + ' -> ' + color);
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

    // Extract fill color from the shape properties
    let fillColor = undefined;
    const fill = shape?.shapeProperties?.shapeBackgroundFill;
    if (fill && fill.propertyState !== 'NOT_RENDERED' && fill.propertyState !== 'INHERIT') {
        fillColor = extractFillAdvanced(fill);
    }

    return {
        fontSize: firstRun.fontSize?.magnitude,
        fontFamily: firstRun.fontFamily,
        color: firstRun.foregroundColor ? extractFillAdvanced(firstRun.foregroundColor) : undefined,
        bold: firstRun.bold,
        italic: firstRun.italic,
        fillColor: fillColor  // Add fill color inheritance
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
// RESOLVED COLOR CACHE (using SlidesApp for actual rendered colors)
// ============================================================================

// Cache of resolved fill colors: objectId -> hex color
let _resolvedFillColors = {};

// Cache of resolved theme colors per slide (keyed by slide index)
// This allows each slide to have its own theme color resolution based on its master
let _resolvedThemeColorMapsPerSlide = {};

/**
 * Get the resolved theme color map for a specific slide
 * @param {number} slideIndex - The slide index (0-based)
 * @returns {Object} The theme color map for that slide, or empty object if not found
 */
function getResolvedThemeColorMap(slideIndex) {
    return _resolvedThemeColorMapsPerSlide[slideIndex] || {};
}

/**
 * Build a cache of resolved fill colors using SlidesApp.
 * SlidesApp returns actual rendered colors, resolving theme references.
 * This runs BEFORE extraction to populate _resolvedFillColors.
 */
function buildResolvedColorCache(presentationId) {
    _resolvedFillColors = {};
    _resolvedThemeColorMapsPerSlide = {};

    try {
        const presentation = SlidesApp.openById(presentationId);
        const slides = presentation.getSlides();

        const themeColorTypes = [
            SlidesApp.ThemeColorType.DARK1,
            SlidesApp.ThemeColorType.DARK2,
            SlidesApp.ThemeColorType.LIGHT1,
            SlidesApp.ThemeColorType.LIGHT2,
            SlidesApp.ThemeColorType.ACCENT1,
            SlidesApp.ThemeColorType.ACCENT2,
            SlidesApp.ThemeColorType.ACCENT3,
            SlidesApp.ThemeColorType.ACCENT4,
            SlidesApp.ThemeColorType.ACCENT5,
            SlidesApp.ThemeColorType.ACCENT6,
            SlidesApp.ThemeColorType.HYPERLINK,
            SlidesApp.ThemeColorType.FOLLOWED_HYPERLINK
        ];

        Logger.log('[RESOLVE] Building resolved color cache for ' + slides.length + ' slides');

        for (let slideIndex = 0; slideIndex < slides.length; slideIndex++) {
            const slide = slides[slideIndex];

            // Build theme color map for THIS slide's color scheme
            // Each slide may use a different master with different theme colors
            // Store in per-slide dictionary so extraction can use the correct colors
            const slideThemeColorMap = {};
            try {
                const colorScheme = slide.getColorScheme();
                for (const themeType of themeColorTypes) {
                    try {
                        const concreteColor = colorScheme.getConcreteColor(themeType);
                        if (concreteColor) {
                            const hex = concreteColor.asRgbColor().asHexString();
                            slideThemeColorMap[themeType.toString()] = hex;
                        }
                    } catch (e) {
                        // Skip
                    }
                }
                // Store this slide's color map
                _resolvedThemeColorMapsPerSlide[slideIndex] = slideThemeColorMap;
            } catch (e) {
                Logger.log('[RESOLVE] Could not get color scheme for slide ' + slideIndex + ': ' + e.message);
                _resolvedThemeColorMapsPerSlide[slideIndex] = {}; // Empty map on error
            }

            // Get ALL page elements, not just shapes
            const pageElements = slide.getPageElements();
            Logger.log('[RESOLVE] Slide ' + slideIndex + ' has ' + pageElements.length + ' page elements');

            for (const element of pageElements) {
                try {
                    const elementType = element.getPageElementType();
                    const objectId = element.getObjectId();

                    if (elementType === SlidesApp.PageElementType.SHAPE) {
                        const shape = element.asShape();
                        extractShapeColor(shape, objectId, slideThemeColorMap);
                    } else if (elementType === SlidesApp.PageElementType.GROUP) {
                        const group = element.asGroup();
                        processGroupForColors(group, slideThemeColorMap);
                    }
                    // Note: Images, tables, etc. don't have fill colors we need to resolve
                } catch (e) {
                    Logger.log('[RESOLVE] Error processing element: ' + e.message);
                }
            }
        }

        Logger.log('[RESOLVE] Cached ' + Object.keys(_resolvedFillColors).length + ' resolved colors');
    } catch (e) {
        Logger.log('[RESOLVE] Error building color cache: ' + e.message);
    }
}

/**
 * Extract fill color from a shape using SlidesApp
 * @param {Shape} shape - The SlidesApp shape object
 * @param {string} objectId - The shape's object ID
 * @param {Object} themeColorMap - The theme color map for this slide
 */
function extractShapeColor(shape, objectId, themeColorMap) {
    try {
        const fill = shape.getFill();
        if (!fill) {
            return;
        }

        const fillType = fill.getType();

        if (fillType === SlidesApp.FillType.SOLID) {
            const solidFill = fill.getSolidFill();
            if (solidFill) {
                const color = solidFill.getColor();
                if (color) {
                    const colorType = color.getColorType();

                    if (colorType === SlidesApp.ColorType.RGB) {
                        const rgbColor = color.asRgbColor();
                        if (rgbColor) {
                            const hex = rgbColor.asHexString();
                            _resolvedFillColors[objectId] = hex;
                            Logger.log('[RESOLVE] Shape ' + objectId + ' RGB fill: ' + hex);
                        }
                    } else if (colorType === SlidesApp.ColorType.THEME) {
                        // Resolve theme color using the slide's theme color map
                        const themeColor = color.asThemeColor();
                        const themeColorType = themeColor.getThemeColorType();

                        // Look up the resolved color from this slide's theme color map
                        const resolvedHex = themeColorMap[themeColorType.toString()];
                        if (resolvedHex) {
                            _resolvedFillColors[objectId] = resolvedHex;
                            Logger.log('[RESOLVE] Shape ' + objectId + ' theme color ' + themeColorType + ' resolved to: ' + resolvedHex);
                        } else {
                            Logger.log('[RESOLVE] Shape ' + objectId + ' theme color ' + themeColorType + ' not found in map');
                        }
                    }
                }
            }
        }
    } catch (e) {
        Logger.log('[RESOLVE] Error extracting shape color for ' + objectId + ': ' + e.message);
    }
}

/**
 * Process a group recursively to extract resolved colors
 * @param {Group} group - The SlidesApp group object
 * @param {Object} themeColorMap - The theme color map for this slide
 */
function processGroupForColors(group, themeColorMap) {
    try {
        const children = group.getChildren();
        for (const child of children) {
            const elementType = child.getPageElementType();
            if (elementType === SlidesApp.PageElementType.SHAPE) {
                const shape = child.asShape();
                const objectId = child.getObjectId();
                extractShapeColor(shape, objectId, themeColorMap);
            } else if (elementType === SlidesApp.PageElementType.GROUP) {
                processGroupForColors(child.asGroup(), themeColorMap);
            }
        }
    } catch (e) {
        Logger.log('[RESOLVE] Error processing group: ' + e.message);
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

    // Build resolved color cache using SlidesApp (gets actual rendered colors)
    buildResolvedColorCache(presentationId);

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
    _activeThemeColors = DEFAULT_THEME_COLORS; // Reset before extraction
    _themeColorMap = extractThemeColorMap(presentation);
    buildBackgroundIndex(presentation);
    log('Theme colors extracted: ' + Object.keys(_themeColorMap).join(', '));
    log('Active theme colors: ' + Object.keys(_activeThemeColors).join(', '));

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
        return extractSlideAdvanced(slide, idx);
    });

    const json = {
        config: {
            title: presentation.title || 'input ' + presentationId,
            rawMode: rawMode,
            sourcePresentationId: presentationId, // For Phase 2 copy operations
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
    let themeName = null;

    try {
        const masters = presentation.masters;
        Logger.log('[THEME] Masters count: ' + (masters ? masters.length : 0));
        if (!masters || masters.length === 0) return colorMap;

        // Check all masters for color scheme and theme name
        for (let i = 0; i < masters.length; i++) {
            const master = masters[i];

            // Get theme name from first master
            if (!themeName && master.masterProperties?.displayName) {
                themeName = master.masterProperties.displayName;
                Logger.log('[THEME] Theme name: ' + themeName);
            }

            // Try standard colorScheme location first
            let colorScheme = master.masterProperties?.colorScheme?.colors;
            if (colorScheme && colorScheme.length > 0) {
                Logger.log('[THEME] Found explicit colorScheme in master ' + i + ': ' + colorScheme.length + ' colors');
                colorScheme.forEach(c => {
                    if (c.type && c.color?.rgbColor) {
                        colorMap[c.type] = rgbToHexAdvanced(c.color.rgbColor);
                    }
                });
            }
        }

        // If no explicit colorScheme, check if we know this theme
        if (Object.keys(colorMap).length === 0 && themeName && KNOWN_THEME_COLORS[themeName]) {
            Logger.log('[THEME] Using known theme colors for: ' + themeName);
            _activeThemeColors = { ...KNOWN_THEME_COLORS[themeName] };
        }

        // IMPORTANT: Override DARK1 if we detect blue layouts
        // This handles "Keep original styles" cases where content from a blue theme
        // was pasted into a black-DARK1 theme presentation
        if (presentation.layouts) {
            const layoutColors = detectProminentLayoutColors(presentation.layouts);
            if (layoutColors.prominentBlue && _activeThemeColors['DARK1'] === '#000000') {
                Logger.log('[THEME] Detected blue layouts - overriding DARK1 to: ' + layoutColors.prominentBlue);
                _activeThemeColors['DARK1'] = layoutColors.prominentBlue;
            }
        }

        // If we extracted colors, merge them with active
        if (Object.keys(colorMap).length > 0) {
            _activeThemeColors = { ...DEFAULT_THEME_COLORS, ..._activeThemeColors, ...colorMap };
        }

        Logger.log('[THEME] Final colorMap: ' + JSON.stringify(colorMap));
        Logger.log('[THEME] Active theme colors DARK1=' + _activeThemeColors['DARK1']);
    } catch (e) {
        Logger.log('[THEME] Error extracting theme colors: ' + e.message);
    }
    return colorMap;
}

/**
 * Extract RGB colors from page elements to build theme color map.
 * Looks for shapes with explicit RGB fills that might represent theme colors.
 */
function extractColorsFromPageElements(elements, colorMap) {
    if (!elements) return;

    for (const element of elements) {
        // Check shape fills
        if (element.shape?.shapeProperties?.shapeBackgroundFill?.solidFill) {
            const fill = element.shape.shapeProperties.shapeBackgroundFill.solidFill;
            if (fill.color?.rgbColor && !fill.color?.themeColor) {
                const hex = rgbToHexAdvanced(fill.color.rgbColor);
                // Try to identify what theme color this might be based on luminance
                const rgb = fill.color.rgbColor;
                const luminance = (rgb.red || 0) * 0.299 + (rgb.green || 0) * 0.587 + (rgb.blue || 0) * 0.114;

                if (luminance < 0.2 && !colorMap['DARK1']) {
                    colorMap['DARK1'] = hex;
                    Logger.log('[THEME] Inferred DARK1 from shape: ' + hex);
                } else if (luminance > 0.8 && !colorMap['LIGHT1']) {
                    colorMap['LIGHT1'] = hex;
                    Logger.log('[THEME] Inferred LIGHT1 from shape: ' + hex);
                }
            }
        }

        // Check text colors for additional theme color hints
        if (element.shape?.text?.textElements) {
            for (const textEl of element.shape.text.textElements) {
                if (textEl.textRun?.style?.foregroundColor?.opaqueColor?.rgbColor) {
                    const rgb = textEl.textRun.style.foregroundColor.opaqueColor.rgbColor;
                    const hex = rgbToHexAdvanced(rgb);
                    const luminance = (rgb.red || 0) * 0.299 + (rgb.green || 0) * 0.587 + (rgb.blue || 0) * 0.114;

                    if (luminance < 0.2 && !colorMap['DARK1']) {
                        colorMap['DARK1'] = hex;
                    }
                }
            }
        }
    }
}

function resolveThemeColor(color) {
    if (!color) return '#000000';
    if (typeof color === 'string' && color.startsWith('#')) return color;
    if (typeof color === 'string' && color.startsWith('theme:')) {
        const themeKey = color.substring(6);
        return _themeColorMap[themeKey] || _activeThemeColors[themeKey] || DEFAULT_THEME_COLORS[themeKey] || '#000000';
    }
    return color;
}

/**
 * Detect prominent colors used in layouts to help infer theme colors.
 * This helps handle "Keep original styles" cases where pasted content
 * uses different theme colors than the destination presentation.
 */
function detectProminentLayoutColors(layouts) {
    const result = { prominentBlue: null };
    const blueColors = [];

    for (const layout of layouts) {
        // Check layout background
        const bg = layout.pageProperties?.pageBackgroundFill;
        if (bg?.solidFill?.color?.rgbColor) {
            const rgb = bg.solidFill.color.rgbColor;
            const hex = rgbToHexAdvanced(rgb);

            // Check if this is a blue color (high blue component, lower red/green)
            if ((rgb.blue || 0) > 0.6 && (rgb.blue || 0) > (rgb.red || 0) && (rgb.blue || 0) > (rgb.green || 0)) {
                blueColors.push(hex);
            }
        }

        // Also check shapes on layouts for blue fills
        if (layout.pageElements) {
            for (const el of layout.pageElements) {
                if (el.shape?.shapeProperties?.shapeBackgroundFill?.solidFill?.color?.rgbColor) {
                    const rgb = el.shape.shapeProperties.shapeBackgroundFill.solidFill.color.rgbColor;
                    const hex = rgbToHexAdvanced(rgb);

                    if ((rgb.blue || 0) > 0.6 && (rgb.blue || 0) > (rgb.red || 0) && (rgb.blue || 0) > (rgb.green || 0)) {
                        blueColors.push(hex);
                    }
                }
            }
        }
    }

    // If we found blue colors, use the most common one (or first)
    if (blueColors.length > 0) {
        // Count occurrences
        const counts = {};
        blueColors.forEach(c => { counts[c] = (counts[c] || 0) + 1; });

        // Find most common
        let maxCount = 0;
        for (const [color, count] of Object.entries(counts)) {
            if (count > maxCount) {
                maxCount = count;
                result.prominentBlue = color;
            }
        }
        Logger.log('[THEME] Detected blue colors in layouts: ' + JSON.stringify(counts));
    }

    return result;
}

// ============================================================================
// SLIDE EXTRACTION
// ============================================================================

function extractSlideAdvanced(slide, slideIndex) {
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

    // Track z-index for proper layer ordering during rebuild
    let zIndexCounter = 0;

    elements.forEach((element, idx) => {
        // Pass layoutId and slideIndex down to elements
        const extracted = extractElementAdvanced(element, null, layoutId, slideIndex);
        if (extracted) {
            if (Array.isArray(extracted)) {
                log('    [' + idx + '] Group with ' + extracted.length + ' children');
                // DEBUG: Log group details for slide 14
                if (slideIndex === 13) {
                    Logger.log('[SLIDE14_GROUP] Element ' + idx + ' is a group with ' + extracted.length + ' children');
                    extracted.forEach((child, childIdx) => {
                        Logger.log('[SLIDE14_GROUP]   Child ' + childIdx + ': type=' + child.type +
                            ' x=' + Math.round(child.x) + ' y=' + Math.round(child.y) +
                            ' w=' + Math.round(child.w) + ' h=' + Math.round(child.h) +
                            (child.shapeType ? ' shapeType=' + child.shapeType : '') +
                            (child.color ? ' color=' + child.color : ''));
                    });
                }
                // Assign incrementing zIndex to each child to preserve order
                extracted.forEach(child => {
                    child.zIndex = zIndexCounter++;
                });
                slideData.elements.push(...extracted);
            } else {
                log('    [' + idx + '] ' + extracted.type + ' "' + (extracted.text || extracted.objectId || '').substring(0, 30) + '"');
                // DEBUG: Log line details for slide 14
                if (slideIndex === 13 && extracted.type === 'line') {
                    Logger.log('[SLIDE14_LINE] Line: x=' + Math.round(extracted.x) + ' y=' + Math.round(extracted.y) +
                        ' w=' + Math.round(extracted.w) + ' h=' + Math.round(extracted.h) +
                        ' color=' + extracted.color + ' weight=' + extracted.weight);
                }
                extracted.zIndex = zIndexCounter++;
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

    // Parent transform matrix components
    const pScaleX = parent.scaleX !== undefined ? parent.scaleX : 1;
    const pScaleY = parent.scaleY !== undefined ? parent.scaleY : 1;
    const pShearX = parent.shearX || 0;
    const pShearY = parent.shearY || 0;
    const pTranslateX = parent.translateX || 0;
    const pTranslateY = parent.translateY || 0;

    // Child transform matrix components
    const cScaleX = child.scaleX !== undefined ? child.scaleX : 1;
    const cScaleY = child.scaleY !== undefined ? child.scaleY : 1;
    const cShearX = child.shearX || 0;
    const cShearY = child.shearY || 0;
    const cTranslateX = child.translateX || 0;
    const cTranslateY = child.translateY || 0;

    // Proper 2x2 matrix multiplication for the linear part:
    // [ pScaleX  pShearX ]   [ cScaleX  cShearX ]
    // [ pShearY  pScaleY ] x [ cShearY  cScaleY ]
    const newScaleX = pScaleX * cScaleX + pShearX * cShearY;
    const newShearX = pScaleX * cShearX + pShearX * cScaleY;
    const newShearY = pShearY * cScaleX + pScaleY * cShearY;
    const newScaleY = pShearY * cShearX + pScaleY * cScaleY;

    // Translation: parent translation + parent matrix * child translation
    const newTranslateX = pTranslateX + pScaleX * cTranslateX + pShearX * cTranslateY;
    const newTranslateY = pTranslateY + pShearY * cTranslateX + pScaleY * cTranslateY;

    return {
        scaleX: newScaleX,
        scaleY: newScaleY,
        shearX: newShearX,
        shearY: newShearY,
        translateX: newTranslateX,
        translateY: newTranslateY,
        rotation: (parent.rotation || 0) + (child.rotation || 0)
    };
}

function extractElementAdvanced(element, parentTransform, layoutId, slideIndex) {
    const elementTransform = element.transform || {};
    const transform = parentTransform
        ? composeTransforms(parentTransform, elementTransform)
        : elementTransform;

    // DEBUG: Log transform and size for investigation
    if (element.image) {
        console.log('DEBUG_IMAGE_EXTRACT: ' + element.objectId);
        console.log('DEBUG_SIZE: ' + JSON.stringify(element.size));
        console.log('DEBUG_TRANSFORM: ' + JSON.stringify(element.transform));

        // Also log image properties to check for recolor/crop
        if (element.image && element.image.imageProperties) {
            console.log('DEBUG_IMAGE_PROPS: ' + JSON.stringify(element.image.imageProperties));
        }
    }

    // TARGETED DEBUG for bar chart shapes (rectangles with shear)
    const hasShear = (transform.shearX && Math.abs(transform.shearX) > 0.01) ||
                     (transform.shearY && Math.abs(transform.shearY) > 0.01);
    if (hasShear && element.shape) {
        console.log('[[SHEAR_DEBUG]] ' + element.objectId);
        console.log('  Raw element.transform: ' + JSON.stringify(element.transform));
        console.log('  Composed transform: ' + JSON.stringify(transform));
        console.log('  Element size: ' + JSON.stringify(element.size));
        console.log('  Parent transform was: ' + (parentTransform ? JSON.stringify(parentTransform) : 'none'));
    }

    // IMPORTANT: Google Slides API omits transform properties that are 0.
    // For rotated shapes (e.g., 270°), scaleX and scaleY are 0 (cos(270°)=0)
    // and are omitted from the API response. We must default to 0, not 1.
    // The API includes explicit scaleX/scaleY=1 for non-rotated shapes.
    // Only default to 1 if the transform object is completely empty (identity).
    const hasAnyTransformValue = transform.scaleX !== undefined ||
                                  transform.scaleY !== undefined ||
                                  transform.shearX !== undefined ||
                                  transform.shearY !== undefined;
    const defaultScale = hasAnyTransformValue ? 0 : 1;

    const scaleX = transform.scaleX !== undefined ? transform.scaleX : defaultScale;
    const scaleY = transform.scaleY !== undefined ? transform.scaleY : defaultScale;
    const shearX = transform.shearX || 0;
    const shearY = transform.shearY || 0;
    const translateX = transform.translateX || 0;
    const translateY = transform.translateY || 0;

    const baseWidth = element.size?.width?.magnitude || 0;
    const baseHeight = element.size?.height?.magnitude || 0;

    // For rotated shapes, the transform matrix encodes both rotation AND scale.
    // Matrix form: [scaleX, shearX] = [Sw*cos, -Sh*sin]
    //              [shearY, scaleY]   [Sw*sin,  Sh*cos]
    // Where Sw = width scale factor, Sh = height scale factor
    //
    // To extract actual scale factors from a rotation+scale matrix:
    // Sw = sqrt(scaleX² + shearY²)
    // Sh = sqrt(shearX² + scaleY²)
    //
    // This correctly handles rotated shapes (e.g., 270° where cos=0, sin=-1)
    // where the scale is encoded in shear components rather than scale components.

    const scaleW = Math.sqrt(scaleX * scaleX + shearY * shearY);
    const scaleH = Math.sqrt(shearX * shearX + scaleY * scaleY);

    // Calculate rotation from matrix components.
    // For a rotation+scale matrix: [Sw*cos, -Sh*sin; Sw*sin, Sh*cos]
    // rotation = atan2(shearY, scaleX) = atan2(Sw*sin, Sw*cos) = theta
    // The scale factors cancel out in atan2.
    // Convert from radians to degrees.
    let rotationRad = Math.atan2(shearY, scaleX);
    let rotationDeg = rotationRad * (180 / Math.PI);
    // Normalize to 0-360 range
    if (rotationDeg < 0) {
        rotationDeg += 360;
    }
    // Round to avoid floating point issues (e.g., 269.9999 -> 270)
    rotationDeg = Math.round(rotationDeg * 100) / 100;

    // Determine flip by checking determinant sign
    // det = scaleX * scaleY - shearX * shearY
    // Negative determinant means there's a flip
    const det = scaleX * scaleY - shearX * shearY;
    const hasFlip = det < 0;

    const actualWidth = baseWidth * scaleW;
    const actualHeight = baseHeight * scaleH;

    // Pre-compute values for position conversion (used in both logging and final calculation)
    const cos = Math.cos(rotationRad);
    const sin = Math.sin(rotationRad);
    const halfW = actualWidth / 2;
    const halfH = actualHeight / 2;

    // Verify extracted dimensions
    if (VERBOSE_LOGGING || true) { // FORCE LOG for debugging
        const phType = element.shape?.placeholder?.type || 'NONE';
        console.log('[ExtDims] ID:' + element.objectId +
            ' BaseW:' + baseWidth + ' BaseH:' + baseHeight +
            ' ScaleX:' + scaleX.toFixed(4) + ' ScaleY:' + scaleY.toFixed(4) +
            ' ShearX:' + shearX.toFixed(4) + ' ShearY:' + shearY.toFixed(4) +
            ' ScaleW:' + scaleW.toFixed(4) + ' ScaleH:' + scaleH.toFixed(4) +
            ' ActualW:' + actualWidth.toFixed(0) + ' ActualH:' + actualHeight.toFixed(0) +
            ' RotDeg:' + rotationDeg +
            ' PH:' + phType);
        // Additional log for position conversion
        console.log('  TranslateX:' + translateX.toFixed(0) + ' TranslateY:' + translateY.toFixed(0) +
            ' -> TopLeftX:' + (translateX - halfW * (1 - cos) - halfH * sin).toFixed(0) +
            ' TopLeftY:' + (translateY - halfH * (1 - cos) + halfW * sin).toFixed(0));
    }

    // For flip detection with rotation, we use the determinant.
    // det < 0 means there's exactly one flip (either H or V, not both).
    // To distinguish flipH vs flipV, we check which axis was reflected
    // by comparing the rotation extracted from atan2 with the stored rotation.
    // For simplicity, if det < 0, we assume flipH (most common case).
    // Note: Both flipH AND flipV together = 180° rotation (det > 0).
    let flipH = false;
    let flipV = false;
    if (hasFlip) {
        // Single flip detected - default to flipH
        flipH = true;
    }

    // Convert from matrix translation (translateX/Y) to top-left corner position.
    // buildTransform expects top-left corner and computes matrix translation as:
    //   tx = x + (w/2)(1-cos) + (h/2)sin
    //   ty = y + (h/2)(1-cos) - (w/2)sin
    // So we reverse this to get x, y from tx, ty:
    //   x = tx - (w/2)(1-cos) - (h/2)sin
    //   y = ty - (h/2)(1-cos) + (w/2)sin
    // Note: cos, sin, halfW, halfH are already computed above
    const topLeftX = translateX - halfW * (1 - cos) - halfH * sin;
    const topLeftY = translateY - halfH * (1 - cos) + halfW * sin;

    const base = {
        objectId: element.objectId,
        // Computed values for fallback/compatibility
        x: emuToPt(topLeftX),
        y: emuToPt(topLeftY),
        w: emuToPt(actualWidth),
        h: emuToPt(actualHeight),
        rotation: rotationDeg,
        flipH: flipH,
        flipV: flipV,
        // RAW VALUES for direct passthrough (most accurate)
        // baseSize is the original size BEFORE transform is applied
        // composedTransform is the exact transform matrix from the API
        baseSize: {
            width: baseWidth,   // In EMU
            height: baseHeight  // In EMU
        },
        composedTransform: transform // Raw matrix for high fidelity
    };

    try {
        if (element.shape) {
            return extractShapeAdvanced(element, base, layoutId, slideIndex);
        }
        if (element.image) {
            return extractImageAdvanced(element, base);
        }
        if (element.table) {
            return extractTableAdvanced(element, base, slideIndex);
        }
        if (element.line) {
            return extractLineAdvanced(element, base);
        }
        if (element.sheetsChart) {
            return extractChartAdvanced(element, base);
        }
        if (element.video) {
            return extractVideoAdvanced(element, base);
        }
        if (element.elementGroup) {
            const children = element.elementGroup.children || [];

            // Check if this group contains curved/freeform lines that can't be accurately reproduced
            const hasCurvedLines = children.some(child => {
                if (child.line) {
                    const lineType = child.line.lineType || '';
                    const lineCategory = child.line.lineCategory || '';
                    // Freeform/scribble lines have no lineType, or are CURVED connectors
                    // These can't be accurately reproduced as the API doesn't expose curve control points
                    const isCurved = lineCategory === 'CURVED' ||
                                     lineType.includes('CURVED') ||
                                     (lineType === '' && lineCategory === ''); // Freeform lines
                    if (isCurved) {
                        Logger.log('[GROUP_DETECT] Found curved/freeform line in group ' + element.objectId + ': lineType=' + lineType + ' lineCategory=' + lineCategory);
                    }
                    return isCurved;
                }
                return false;
            });

            if (hasCurvedLines) {
                // Extract as a copyGroup reference instead of flattening
                Logger.log('[GROUP_DETECT] Group ' + element.objectId + ' contains curved lines - marking for copy from source');

                // Groups don't have their own dimensions - calculate bounding box from children
                let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;

                // Get the GROUP's transform - needed to compose with children's local transforms
                const groupTransform = element.transform || {};
                const grpScaleX = groupTransform.scaleX !== undefined ? groupTransform.scaleX : 1;
                const grpScaleY = groupTransform.scaleY !== undefined ? groupTransform.scaleY : 1;
                const grpShearX = groupTransform.shearX || 0;
                const grpShearY = groupTransform.shearY || 0;
                const grpTranslateX = groupTransform.translateX || 0;
                const grpTranslateY = groupTransform.translateY || 0;

                Logger.log('[GROUP_TRANSFORM] scaleX=' + grpScaleX.toFixed(4) + ' scaleY=' + grpScaleY.toFixed(4) +
                           ' translateX=' + emuToPt(grpTranslateX).toFixed(1) + ' translateY=' + emuToPt(grpTranslateY).toFixed(1));

                children.forEach((child, idx) => {
                    const childTransform = child.transform || {};
                    const childSize = child.size || {};

                    // Get child's base size
                    const childBaseW = childSize.width ? childSize.width.magnitude : 0;
                    const childBaseH = childSize.height ? childSize.height.magnitude : 0;

                    // Skip "template connector" elements (3000000 EMU = 236pt base size)
                    // These are placeholder-sized elements that distort bounding box calculations
                    const TEMPLATE_SIZE = 3000000; // EMU
                    if (childBaseW === TEMPLATE_SIZE || childBaseH === TEMPLATE_SIZE) {
                        Logger.log('[GROUP_CHILD_' + idx + '] SKIPPED - template connector size');
                        return; // Skip this child
                    }

                    // Get child's LOCAL transform components (relative to group)
                    const childScaleX = childTransform.scaleX !== undefined ? childTransform.scaleX : 1;
                    const childScaleY = childTransform.scaleY !== undefined ? childTransform.scaleY : 1;
                    const childLocalX = childTransform.translateX || 0;
                    const childLocalY = childTransform.translateY || 0;

                    // COMPOSE with group transform to get WORLD coordinates
                    // World_X = grpScaleX * childLocalX + grpShearX * childLocalY + grpTranslateX
                    // World_Y = grpShearY * childLocalX + grpScaleY * childLocalY + grpTranslateY
                    const childWorldX = grpScaleX * childLocalX + grpShearX * childLocalY + grpTranslateX;
                    const childWorldY = grpShearY * childLocalX + grpScaleY * childLocalY + grpTranslateY;

                    // Calculate child's actual dimensions (absolute values, accounting for group scale)
                    const childActualW = Math.abs(childBaseW * childScaleX * grpScaleX);
                    const childActualH = Math.abs(childBaseH * childScaleY * grpScaleY);

                    // For lines, translate is one corner - determine which based on scale signs
                    // Positive scale: line extends right/down from translate
                    // Negative scale: line extends left/up from translate
                    // Use composed scale to determine direction
                    const composedScaleX = childScaleX * grpScaleX;
                    const composedScaleY = childScaleY * grpScaleY;

                    let childMinX = childWorldX;
                    let childMinY = childWorldY;
                    let childMaxX = childWorldX;
                    let childMaxY = childWorldY;

                    if (composedScaleX >= 0) {
                        childMaxX = childWorldX + childActualW;
                    } else {
                        childMinX = childWorldX - childActualW;
                    }

                    if (composedScaleY >= 0) {
                        childMaxY = childWorldY + childActualH;
                    } else {
                        childMinY = childWorldY - childActualH;
                    }

                    // Debug logging
                    Logger.log('[GROUP_CHILD_' + idx + '] baseW=' + emuToPt(childBaseW).toFixed(0) +
                               ' baseH=' + emuToPt(childBaseH).toFixed(0) +
                               ' localX=' + emuToPt(childLocalX).toFixed(1) + ' localY=' + emuToPt(childLocalY).toFixed(1) +
                               ' -> world: (' + emuToPt(childMinX).toFixed(1) + ',' + emuToPt(childMinY).toFixed(1) +
                               ') to (' + emuToPt(childMaxX).toFixed(1) + ',' + emuToPt(childMaxY).toFixed(1) + ')');

                    // Update bounding box
                    if (childActualW > 0 && childActualH > 0) {
                        minX = Math.min(minX, childMinX);
                        minY = Math.min(minY, childMinY);
                        maxX = Math.max(maxX, childMaxX);
                        maxY = Math.max(maxY, childMaxY);
                    }
                });

                // If we couldn't calculate bounds, use defaults
                if (minX === Infinity) {
                    minX = 0; minY = 0; maxX = 1270000; maxY = 1270000;
                }

                // Add padding to account for line stroke width (lines extend beyond geometric bounds)
                // 4pt padding on each side should cover most stroke widths
                const STROKE_PADDING = 4 * 12700; // 4pt in EMU
                minX -= STROKE_PADDING;
                minY -= STROKE_PADDING;
                maxX += STROKE_PADDING;
                maxY += STROKE_PADDING;

                const groupX = minX;
                const groupY = minY;
                const groupW = maxX - minX;
                const groupH = maxY - minY;

                Logger.log('[GROUP_DETECT] Calculated bounds: x=' + emuToPt(groupX).toFixed(1) +
                           ' y=' + emuToPt(groupY).toFixed(1) +
                           ' w=' + emuToPt(groupW).toFixed(1) +
                           ' h=' + emuToPt(groupH).toFixed(1) + ' pt');

                return {
                    type: 'copyGroup',
                    objectId: element.objectId,
                    x: emuToPt(groupX),
                    y: emuToPt(groupY),
                    w: emuToPt(groupW),
                    h: emuToPt(groupH),
                    sourceObjectId: element.objectId,
                    sourceSlideIndex: slideIndex,
                    reason: 'Contains curved/freeform lines that cannot be reproduced via API'
                };
            }

            // Normal group - flatten and extract children
            const extracted = children
                .map(c => extractElementAdvanced(c, transform, layoutId, slideIndex)) // Pass layoutId and slideIndex recursively
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

function extractShapeAdvanced(element, base, layoutId, slideIndex) {
    const shape = element.shape;
    const shapeType = shape.shapeType || 'RECTANGLE';
    const isTextBox = shapeType === 'TEXT_BOX';

    // Extract text content - pass slideIndex for proper theme color resolution
    let textContent = '';
    let textRuns = [];
    let textStyle = null;

    if (shape.text && shape.text.textElements) {
        const textData = extractTextAdvanced(shape.text.textElements, slideIndex);
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

        // Extract fill color for text boxes (same logic as shapes)
        if (fill && fill.propertyState !== 'NOT_RENDERED' && fill.propertyState !== 'INHERIT') {
            result.fillColor = resolveThemeColor(extractFillAdvanced(fill, element.objectId));
            result._originalFillColor = result.fillColor;
        } else if (fill && fill.propertyState === 'NOT_RENDERED') {
            result.fillColor = 'transparent';
            result._originalFillColor = 'transparent';
        }

        // Extract border/outline for text boxes (same logic as shapes)
        if (outline) {
            const extractedBorderColor = extractOutlineColorAdvanced(outline);
            const extractedBorderWidth = outline.weight?.magnitude ? outline.weight.magnitude / ADVANCED_EMU_PER_PT : 0;
            const extractedBorderDash = extractOutlineDashStyle(outline);
            if (extractedBorderColor && extractedBorderColor !== 'none') {
                result.borderColor = resolveThemeColor(extractedBorderColor);
                result.borderWidth = extractedBorderWidth;
                result.borderDash = extractedBorderDash;
            }
        }

        const autofit = shape.shapeProperties?.autofit;
        if (autofit) {
            result.paddingTop = (autofit.topOffset || 0) / ADVANCED_EMU_PER_PT;
            result.paddingBottom = (autofit.bottomOffset || 0) / ADVANCED_EMU_PER_PT;
            result.paddingLeft = (autofit.leftOffset || 0) / ADVANCED_EMU_PER_PT;
            result.paddingRight = (autofit.rightOffset || 0) / ADVANCED_EMU_PER_PT;
        }

        // Vertical alignment (contentAlignment)
        const contentAlignment = shape.shapeProperties?.contentAlignment;
        if (contentAlignment) {
            result.verticalAlign = mapContentAlignmentAdvanced(contentAlignment);
        }

        return result;
    } else {
        // Resolve fill color - check for inheritance from placeholder
        let resolvedFillColor = 'transparent';

        // DEBUG: Log fill data for shapes
        if (fill) {
            console.log('[FILL_DEBUG] Shape ' + element.objectId + ' fill: ' + JSON.stringify(fill).substring(0, 200));
        } else {
            console.log('[FILL_DEBUG] Shape ' + element.objectId + ' has no fill property');
        }

        if (fill && fill.propertyState !== 'NOT_RENDERED' && fill.propertyState !== 'INHERIT') {
            // Explicit fill on this shape - pass objectId for SlidesApp resolved colors
            resolvedFillColor = extractFillAdvanced(fill, element.objectId);
            console.log('[FILL_DEBUG] Extracted fill: ' + resolvedFillColor);
        } else if (fill && fill.propertyState === 'INHERIT') {
            // INHERIT means use the fill from the layout/master
            // First try to get it from placeholder inheritance
            const inheritedFill = resolveInheritedProperty(element, 'fillColor', layoutId);
            if (inheritedFill) {
                resolvedFillColor = inheritedFill;
                console.log('[FILL_DEBUG] Inherited fill from placeholder: ' + resolvedFillColor);
            } else {
                // If no placeholder inheritance, check if fill has solidFill data despite INHERIT state
                // (Some API responses include both propertyState and solidFill)
                if (fill.solidFill) {
                    resolvedFillColor = extractFillAdvanced(fill, element.objectId);
                    console.log('[FILL_DEBUG] Got fill from solidFill despite INHERIT: ' + resolvedFillColor);
                }
            }
        } else {
            // Try to inherit fill from placeholder (if shape has a placeholder type)
            const inheritedFill = resolveInheritedProperty(element, 'fillColor', layoutId);
            if (inheritedFill) {
                resolvedFillColor = inheritedFill;
                console.log('[FILL_DEBUG] Inherited fill: ' + resolvedFillColor);
            }
        }

        // DEBUG: Log outline data for shapes
        if (outline) {
            Logger.log('[OUTLINE_DEBUG] Shape ' + element.objectId + ' outline: ' +
                JSON.stringify(outline).substring(0, 300));
        }

        const extractedBorderColor = outline ? extractOutlineColorAdvanced(outline) : 'none';
        const extractedBorderWidth = outline?.weight?.magnitude ? outline.weight.magnitude / ADVANCED_EMU_PER_PT : 0;
        const extractedBorderDash = extractOutlineDashStyle(outline);

        Logger.log('[OUTLINE_DEBUG] Shape ' + element.objectId + ' borderColor=' + extractedBorderColor +
            ' borderWidth=' + extractedBorderWidth + ' borderDash=' + extractedBorderDash);

        const result = {
            type: 'shape',
            ...base,
            shape: shapeType,
            fillColor: resolveThemeColor(resolvedFillColor),
            borderColor: resolveThemeColor(extractedBorderColor),
            borderWidth: extractedBorderWidth,
            borderDash: extractedBorderDash,
            flipH: base.flipH,
            flipV: base.flipV
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

        // Vertical alignment (contentAlignment)
        const shapeContentAlignment = shape.shapeProperties?.contentAlignment;
        if (shapeContentAlignment) {
            result.verticalAlign = mapContentAlignmentAdvanced(shapeContentAlignment);
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
function extractTextAdvanced(textElements, slideIndex) {
    const runs = [];
    let plainText = '';
    let firstRunStyle = null;

    // Get the resolved theme color map for this specific slide
    const slideThemeColors = getResolvedThemeColorMap(slideIndex);

    // CORRECTED APPROACH (marker-first):
    // 1. When we see a paragraphMarker, save it as "pending" for upcoming text runs
    // 2. Collect textRuns into currentParagraphRuns
    // 3. When we see the next paragraphMarker, apply pending marker to collected runs
    // 4. At the end, apply any pending marker to remaining runs

    // Runs in the current paragraph (will receive the pending marker)
    let currentParagraphRuns = [];

    // Pending paragraph marker to apply to the next text runs
    let pendingMarker = null;


    // Helper to convert units to points
    const getPt = (prop) => {
        if (!prop || prop.magnitude === undefined) return undefined;
        return (prop.unit === 'PT') ? prop.magnitude : (prop.magnitude / ADVANCED_EMU_PER_PT);
    };

    // Helper to build paragraph style and bullet data from a marker
    const buildParagraphData = (marker) => {
        const ps = marker.style || {};
        const bullet = marker.bullet;

        // Use explicit values only - don't inherit from previous paragraphs
        // undefined means "use default (0pt)", not "inherit from previous"
        const indentStart = getPt(ps.indentStart);
        const indentFirstLine = getPt(ps.indentFirstLine);

        const pStyle = {
            align: mapAlignmentAdvanced(ps.alignment),
            direction: ps.direction,
            spacingMode: ps.spacingMode,
            spaceAbove: getPt(ps.spaceAbove),
            spaceBelow: getPt(ps.spaceBelow),
            lineSpacing: ps.lineSpacing || 100,
            indentStart: indentStart,
            indentFirstLine: indentFirstLine
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
                const themeKey = style.foregroundColor.opaqueColor.themeColor;
                // Use this slide's resolved theme colors first (most accurate), then fall back to static maps
                color = slideThemeColors[themeKey] || _themeColorMap[themeKey] || _activeThemeColors[themeKey] || DEFAULT_THEME_COLORS[themeKey] || '#000000';
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

    // Inherit missing fontSize/fontFamily from adjacent runs
    // This ensures whitespace and other runs without explicit styles
    // get the formatting from nearby content runs
    inheritMissingRunStyles(runs);

    return { plainText, runs, firstRunStyle };
}

/**
 * Inherit missing fontSize/fontFamily from adjacent runs.
 * When the API doesn't return style info for certain runs (especially whitespace),
 * those runs would get default formatting during generation.
 * This function ensures they inherit from adjacent content runs instead.
 */
function inheritMissingRunStyles(runs) {
    if (runs.length === 0) return;

    // First pass: forward inheritance (from previous runs)
    for (let i = 1; i < runs.length; i++) {
        const run = runs[i];
        const prevRun = runs[i - 1];

        if (run.fontSize === undefined && prevRun.fontSize !== undefined) {
            run.fontSize = prevRun.fontSize;
        }
        if (run.fontFamily === undefined && prevRun.fontFamily !== undefined) {
            run.fontFamily = prevRun.fontFamily;
        }
    }

    // Second pass: backward inheritance (for leading runs that are still missing styles)
    for (let i = runs.length - 2; i >= 0; i--) {
        const run = runs[i];
        const nextRun = runs[i + 1];

        if (run.fontSize === undefined && nextRun.fontSize !== undefined) {
            run.fontSize = nextRun.fontSize;
        }
        if (run.fontFamily === undefined && nextRun.fontFamily !== undefined) {
            run.fontFamily = nextRun.fontFamily;
        }
    }
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

function mapContentAlignmentAdvanced(contentAlignment) {
    const map = {
        'TOP': 'top',
        'MIDDLE': 'middle',
        'BOTTOM': 'bottom'
    };
    return map[contentAlignment] || 'top';
}

// ============================================================================
// IMAGE EXTRACTION
// ============================================================================

function extractImageAdvanced(element, base) {
    const image = element.image;
    const result = {
        type: 'image',
        id: element.objectId,
        x: base.x,
        y: base.y,
        w: base.w,
        h: base.h,
        left: base.x, // Keep legacy alias
        top: base.y,  // Keep legacy alias
        width: base.w, // Keep legacy alias
        height: base.h, // Keep legacy alias
        rotation: base.rotation,
        flipH: base.flipH,
        flipV: base.flipV,
        // RAW FIDELITY PASSTHROUGH
        rawSize: element.size,
        rawTransform: base.composedTransform, // We need to expose this from base or element
        url: image.contentUrl || image.sourceUrl, // CRITICAL FIX: Use contentUrl as primary source
        sourceUrl: image.sourceUrl || null,
        originalWidth: base.w,
        originalHeight: base.h
    };

    // Extract crop properties if present
    const crop = image.imageProperties?.cropProperties;
    if (crop) {
        result.crop = {
            left: crop.leftOffset || 0,
            right: crop.rightOffset || 0,
            top: crop.topOffset || 0,
            bottom: crop.bottomOffset || 0,
            angle: crop.angle || 0
        };
    }

    // Extract recolor
    const recolor = image.imageProperties?.recolor;
    if (recolor && recolor.recolorStops) {
        result.recolor = recolor;
    }

    // Extract corrections (brightness, contrast, transparency)
    const props = image.imageProperties || {};
    if (props.brightness !== undefined) result.brightness = props.brightness;
    if (props.contrast !== undefined) result.contrast = props.contrast;
    if (props.transparency !== undefined) result.transparency = props.transparency;

    // Extract border/outline for images
    const outline = props.outline;
    if (outline && outline.propertyState !== 'NOT_RENDERED') {
        const extractedBorderColor = extractOutlineColorAdvanced(outline);
        const extractedBorderWidth = outline.weight?.magnitude ? outline.weight.magnitude / ADVANCED_EMU_PER_PT : 0;
        if (extractedBorderColor && extractedBorderColor !== 'none' && extractedBorderWidth > 0) {
            result.borderColor = resolveThemeColor(extractedBorderColor);
            result.borderWidth = extractedBorderWidth;
        }
    }

    return result;
}

// ============================================================================
// TABLE EXTRACTION
// ============================================================================

/**
 * Extract border properties from a table cell border
 * @param {Object} border - The border object (borderTop, borderBottom, etc.)
 * @returns {Object|null} - Border properties or null if not rendered
 */
function extractTableBorder(border) {
    if (!border) return null;

    // Check if border is rendered
    if (border.dashStyle === 'INVISIBLE' || !border.weight) return null;

    const result = {
        weight: border.weight?.magnitude ? border.weight.magnitude / ADVANCED_EMU_PER_PT : 1,
        dashStyle: border.dashStyle || 'SOLID'
    };

    // Extract border color
    if (border.tableBorderFill?.solidFill?.color) {
        const colorObj = border.tableBorderFill.solidFill.color;
        if (colorObj.rgbColor) {
            result.color = rgbToHexAdvanced(colorObj.rgbColor);
        } else if (colorObj.themeColor) {
            result.color = resolveThemeColor('theme:' + colorObj.themeColor);
        }
    }

    return result;
}

function extractTableAdvanced(element, base, slideIndex) {
    const table = element.table;
    const data = [];
    const rowHeights = [];
    const columnWidths = [];

    // Extract column widths from tableColumns
    const tableColumns = table.tableColumns || [];
    Logger.log('[TABLE_DEBUG] tableColumns count: ' + tableColumns.length);
    tableColumns.forEach((col, i) => {
        if (col.columnWidth?.magnitude) {
            const widthPt = col.columnWidth.magnitude / ADVANCED_EMU_PER_PT;
            columnWidths.push(widthPt);
            if (i === 0) {
                Logger.log('[TABLE_DEBUG] First column width: ' + widthPt.toFixed(2) + 'pt');
            }
        } else {
            columnWidths.push(null);
        }
    });
    Logger.log('[TABLE_DEBUG] Column widths: ' + JSON.stringify(columnWidths.map(w => w ? w.toFixed(2) : null)));

    // Table borders are stored at the TABLE level, not cell level
    // horizontalBorderRows[row][col] = border BELOW row (or above for row 0)
    // verticalBorderRows[row][col] = border to the RIGHT of col (or left for col 0)
    const horizontalBorders = table.horizontalBorderRows || [];
    const verticalBorders = table.verticalBorderRows || [];

    // Debug: log border structure
    if (horizontalBorders.length > 0) {
        Logger.log('[TABLE_BORDER_DEBUG] horizontalBorderRows count: ' + horizontalBorders.length);
        if (horizontalBorders[0]?.tableBorderCells?.[0]) {
            Logger.log('[TABLE_BORDER_DEBUG] First horizontal border: ' + JSON.stringify(horizontalBorders[0].tableBorderCells[0]));
        }
    }
    if (horizontalBorders.length > 1 && horizontalBorders[1]?.tableBorderCells?.[0]) {
        Logger.log('[TABLE_BORDER_DEBUG] Second horizontal border (below header): ' + JSON.stringify(horizontalBorders[1].tableBorderCells[0]));
    }

    const rows = table.tableRows || [];
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        const row = rows[rowIndex];
        const rowData = [];
        const cells = row.tableCells || [];

        // Extract row height
        if (row.tableRowProperties?.minRowHeight?.magnitude) {
            rowHeights.push(row.tableRowProperties.minRowHeight.magnitude / ADVANCED_EMU_PER_PT);
        } else {
            rowHeights.push(null); // Will use default
        }

        for (const cell of cells) {
            let cellData = { text: '' };

            if (cell.text?.textElements) {
                const textData = extractTextAdvanced(cell.text.textElements, slideIndex);
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

                // Debug: log first cell's font size
                if (rowIndex === 0 && rowData.length === 0) {
                    Logger.log('[CELL_DEBUG] First cell fontSize: ' + cellData.fontSize + ', fontFamily: ' + cellData.fontFamily);
                }

                const paraStyle = extractParagraphStyleAdvanced(cell.text.textElements);
                if (paraStyle) {
                    cellData.align = paraStyle.align;
                    // Extract line spacing - crucial for compact tables
                    if (paraStyle.lineSpacing) {
                        cellData.lineSpacing = paraStyle.lineSpacing;
                    }
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

            // Extract cell padding - crucial for compact tables
            const tcp = cell.tableCellProperties;
            if (tcp) {
                const getPadding = (prop) => {
                    if (!prop || prop.magnitude === undefined) return null;
                    return prop.unit === 'PT' ? prop.magnitude : prop.magnitude / ADVANCED_EMU_PER_PT;
                };
                const padTop = getPadding(tcp.paddingTop);
                const padBottom = getPadding(tcp.paddingBottom);
                const padLeft = getPadding(tcp.paddingLeft);
                const padRight = getPadding(tcp.paddingRight);

                // Only include if any padding is explicitly set
                if (padTop !== null || padBottom !== null || padLeft !== null || padRight !== null) {
                    cellData.padding = {
                        top: padTop !== null ? padTop : 0,
                        bottom: padBottom !== null ? padBottom : 0,
                        left: padLeft !== null ? padLeft : 0,
                        right: padRight !== null ? padRight : 0
                    };
                }
            }

            // Debug: Log cell properties for first cell to check padding
            if (rowIndex === 0 && rowData.length === 0) {
                Logger.log('[CELL_DEBUG] First cell tableCellProperties: ' + JSON.stringify(cell.tableCellProperties));
                Logger.log('[CELL_DEBUG] First cell padding extracted: ' + JSON.stringify(cellData.padding));
                // Also log text elements for line spacing info
                if (cell.text?.textElements) {
                    const paraEl = cell.text.textElements.find(e => e.paragraphMarker);
                    if (paraEl) {
                        Logger.log('[CELL_DEBUG] First cell paragraphStyle: ' + JSON.stringify(paraEl.paragraphMarker?.style));
                    }
                }
            }

            // Extract cell borders from table-level border arrays
            const borders = {};
            const colIndex = rowData.length; // Current column index

            // Top border: horizontalBorders[rowIndex] (border above this row)
            if (horizontalBorders[rowIndex]?.tableBorderCells?.[colIndex]) {
                const borderCell = horizontalBorders[rowIndex].tableBorderCells[colIndex];
                const topBorder = extractTableBorder(borderCell.tableBorderProperties);
                if (topBorder) borders.top = topBorder;
            }

            // Bottom border: horizontalBorders[rowIndex + 1] (border below this row)
            if (horizontalBorders[rowIndex + 1]?.tableBorderCells?.[colIndex]) {
                const borderCell = horizontalBorders[rowIndex + 1].tableBorderCells[colIndex];
                const bottomBorder = extractTableBorder(borderCell.tableBorderProperties);
                if (bottomBorder) borders.bottom = bottomBorder;
            }

            // Left border: verticalBorders[rowIndex][colIndex]
            if (verticalBorders[rowIndex]?.tableBorderCells?.[colIndex]) {
                const borderCell = verticalBorders[rowIndex].tableBorderCells[colIndex];
                const leftBorder = extractTableBorder(borderCell.tableBorderProperties);
                if (leftBorder) borders.left = leftBorder;
            }

            // Right border: verticalBorders[rowIndex][colIndex + 1]
            if (verticalBorders[rowIndex]?.tableBorderCells?.[colIndex + 1]) {
                const borderCell = verticalBorders[rowIndex].tableBorderCells[colIndex + 1];
                const rightBorder = extractTableBorder(borderCell.tableBorderProperties);
                if (rightBorder) borders.right = rightBorder;
            }

            if (Object.keys(borders).length > 0) {
                cellData.borders = borders;
            }

            rowData.push(cellData);
        }

        data.push(rowData);
    }

    // Use API minRowHeight values directly - these are the actual stored row heights
    const tableHeight = base.h || 0;
    const numRows = data.length;
    Logger.log('[TABLE_DEBUG] Table height: ' + tableHeight + 'pt, rows: ' + numRows);
    Logger.log('[TABLE_DEBUG] API minRowHeights (using directly): ' + JSON.stringify(rowHeights.map(h => h ? h.toFixed(2) : null)));

    // Always include rowHeights if we have any (don't filter)
    const hasRowHeights = rowHeights.length > 0 && rowHeights.some(h => h !== null);

    // Debug: log first cell borders
    if (data.length > 0 && data[0].length > 0 && data[0][0].borders) {
        Logger.log('[TABLE_BORDER_DEBUG] First cell borders: ' + JSON.stringify(data[0][0].borders));
    }
    // Debug: count cells with borders
    let borderCount = 0;
    data.forEach(row => row.forEach(cell => { if (cell.borders) borderCount++; }));
    Logger.log('[TABLE_BORDER_DEBUG] Cells with borders: ' + borderCount + ' / ' + (data.length * (data[0]?.length || 0)));

    // Include columnWidths if we have any
    const hasColumnWidths = columnWidths.length > 0 && columnWidths.some(w => w !== null);

    return {
        type: 'table',
        ...base,
        data: data,
        ...(hasRowHeights && { rowHeights: rowHeights }),
        ...(hasColumnWidths && { columnWidths: columnWidths })
    };
}

// ============================================================================
// LINE EXTRACTION
// ============================================================================

function extractLineAdvanced(element, base) {
    const line = element.line;
    const lineType = line.lineType || 'none';
    const lineCategory = line.lineCategory || 'none';

    // DEBUG: Log line type and category
    Logger.log('[LINE_DEBUG] Line ' + element.objectId +
        ' lineType=' + lineType +
        ' lineCategory=' + lineCategory +
        ' w=' + Math.round(base.w) + ' h=' + Math.round(base.h));

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
// VIDEO EXTRACTION
// ============================================================================

function extractVideoAdvanced(element, base) {
    const video = element.video;
    const result = {
        type: 'video',
        ...base,
        source: video.source || 'YOUTUBE',
        videoId: video.id || null,
        url: video.url || null
    };

    // Extract border/outline for videos
    const outline = video.videoProperties?.outline;
    if (outline && outline.propertyState !== 'NOT_RENDERED') {
        const extractedBorderColor = extractOutlineColorAdvanced(outline);
        const extractedBorderWidth = outline.weight?.magnitude ? outline.weight.magnitude / ADVANCED_EMU_PER_PT : 0;
        if (extractedBorderColor && extractedBorderColor !== 'none' && extractedBorderWidth > 0) {
            result.borderColor = resolveThemeColor(extractedBorderColor);
            result.borderWidth = extractedBorderWidth;
        }
    }

    return result;
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

function extractFillAdvanced(fill, objectId) {
    if (!fill) return 'transparent';
    if (fill.propertyState === 'NOT_RENDERED') return 'transparent';

    // If we have an explicit RGB color, use it directly (most common case)
    if (fill.solidFill?.color?.rgbColor) {
        return rgbToHexAdvanced(fill.solidFill.color.rgbColor);
    }

    // ONLY for theme color references: use SlidesApp resolved color if available
    // This is the specific case where API returns "DARK1" but we need the actual rendered color
    if (fill.solidFill?.color?.themeColor) {
        const themeKey = fill.solidFill.color.themeColor;

        // First try SlidesApp resolved color (actual rendered color)
        if (objectId && _resolvedFillColors[objectId]) {
            const resolvedColor = _resolvedFillColors[objectId];
            Logger.log('[FILL_DEBUG] Theme color ' + themeKey + ' resolved via SlidesApp: ' + resolvedColor);
            return resolvedColor;
        }

        // Fallback: Try theme color maps
        const color = _themeColorMap[themeKey] || _activeThemeColors[themeKey] || DEFAULT_THEME_COLORS[themeKey];
        if (color) {
            Logger.log('[FILL_DEBUG] Theme color ' + themeKey + ' resolved via map: ' + color);
            return color;
        }
        Logger.log('[FILL_DEBUG] Unknown theme color: ' + themeKey);
        return '#000000';
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
        const themeKey = solidFill.color.themeColor;
        // Try: 1) extracted colorMap, 2) active theme colors, 3) defaults
        return _themeColorMap[themeKey] || _activeThemeColors[themeKey] || DEFAULT_THEME_COLORS[themeKey] || 'none';
    }

    return 'none';
}

function extractOutlineDashStyle(outline) {
    if (!outline) return 'solid';
    if (outline.propertyState === 'NOT_RENDERED') return 'solid';
    // Map API dashStyle values to our JSON schema values
    // API values: SOLID, DOT, DASH, DASH_DOT, LONG_DASH, LONG_DASH_DOT
    const dashStyle = outline.dashStyle;
    if (!dashStyle || dashStyle === 'SOLID') return 'solid';

    const dashMap = {
        'DOT': 'dot',
        'DASH': 'dash',
        'DASH_DOT': 'dashDot',
        'LONG_DASH': 'longDash',
        'LONG_DASH_DOT': 'longDashDot'
    };
    return dashMap[dashStyle] || 'solid';
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
