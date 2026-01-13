/**
 * @fileoverview Validation logic for the Slides Engine.
 */

// ============================================================================
// UTILITY HELPERS
// ============================================================================

/**
 * Clamp value between min and max
 * @param {number} value
 * @param {number} min
 * @param {number} max
 * @returns {number}
 */
function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

/**
 * Normalize property names to handle common mistakes and aliases
 * @param {Object} obj
 * @returns {Object}
 */
function normalizeProperties(obj) {
  if (!obj || typeof obj !== 'object') return obj;

  const normalized = {};
  for (let key in obj) {
    if (!obj.hasOwnProperty(key)) continue;
    const normalizedKey = CONFIG.PROPERTY_ALIASES[key] || key;
    normalized[normalizedKey] = obj[key];
  }
  return normalized;
}

/**
 * Normalize hex color codes.
 * @param {string} color
 * @returns {string|null} Normalised hex string or null
 */
function normalizeColor(color) {
  if (!color || color === 'none' || color === 'transparent') return null;

  // Ensure # prefix
  if (!color.startsWith('#')) {
    color = '#' + color;
  }

  // Expand shorthand (#fff -> #ffffff)
  if (color.length === 4) {
    color = '#' + color[1] + color[1] + color[2] + color[2] + color[3] + color[3];
  }

  return color.toLowerCase();
}

// ============================================================================
// VALIDATION LOGIC
// ============================================================================

/**
 * Check if Advanced Slides Service is enabled
 * @throws {Error} if service is not available
 */
function validateAdvancedService() {
  if (typeof Slides === 'undefined') {
    throw new Error(
      'Advanced Slides Service not enabled. ' +
      'Go to Apps Script Editor -> Services -> Add -> Google Slides API'
    );
  }
}

/**
 * Validate and sanitize element data.
 * Modifies the element object in place for efficient processing.
 * @param {Object} element - The element specification
 * @returns {Object} Validated and normalized element
 */
function validateElement(element) {
  if (!element || !element.type) {
    throw new Error('Element must have a type property');
  }

  element = normalizeProperties(element);

  // Validate dimensions
  if (element.w !== undefined) {
    element.w = clamp(element.w, CONFIG.LIMITS.MIN_DIMENSION, CONFIG.LIMITS.MAX_DIMENSION);
  }
  if (element.h !== undefined) {
    element.h = clamp(element.h, CONFIG.LIMITS.MIN_DIMENSION, CONFIG.LIMITS.MAX_DIMENSION);
  }

  // Validate font size
  if (element.fontSize) {
    element.fontSize = clamp(element.fontSize, CONFIG.LIMITS.MIN_FONT_SIZE, CONFIG.LIMITS.MAX_FONT_SIZE);
  }

  // Preserve original fillColor for transparent detection
  if (element.fillColor) {
    element._originalFillColor = element.fillColor;
  }

  // NOTE: Theme resolution is now handled by ThemeService, but we keep the hook here
  // We'll trust that ThemeService.resolveAllColors(element) is called before or we call it here if we import it.
  // For now, we'll assume the Coordinator handles the order: Normalize -> Theme Resolve -> Validate.
  // However, the original code mixed them. Let's make this pure validation + normalization
  
  // Validate text length
  if (element.text && element.text.length > CONFIG.LIMITS.MAX_TEXT_LENGTH) {
    element.text = element.text.substring(0, CONFIG.LIMITS.MAX_TEXT_LENGTH) + '...';
  }

  return element;
}

/**
 * Validate entire JSON specification
 * @param {Object} json - The full presentation object
 * @returns {Object} The validated json
 */
function validateJSON(json) {
  if (!json || !json.slides || !Array.isArray(json.slides)) {
    throw new Error('JSON must have a slides array');
  }

  if (json.slides.length === 0) {
    throw new Error('Presentation must have at least one slide');
  }

  if (json.slides.length > CONFIG.LIMITS.MAX_SLIDES) {
    throw new Error(`Maximum ${CONFIG.LIMITS.MAX_SLIDES} slides allowed`);
  }

  json.slides.forEach((slide, index) => {
    if (slide.elements && slide.elements.length > CONFIG.LIMITS.MAX_ELEMENTS_PER_SLIDE) {
      throw new Error(`Slide ${index + 1} exceeds ${CONFIG.LIMITS.MAX_ELEMENTS_PER_SLIDE} elements`);
    }
  });

  return json;
}
