/**
 * @fileoverview Utility functions for the Slides Engine.
 */

/**
 * Generate unique object ID for Slides API
 * Must be 5-50 chars, start with alphanumeric or underscore
 * @returns {string}
 */
function generateObjectId() {
  const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let result = 'obj_';
  for (let i = 0; i < 16; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

/**
 * Generate deterministic object ID for predictable references (used in groups)
 * Format: obj_s{slideIndex}_e{elementIndex}[_{suffix}]
 * @param {number} slideIndex
 * @param {number|string} elementIndex
 * @param {string} suffix - Optional suffix for nested elements
 * @returns {string}
 */
function generateDeterministicId(slideIndex, elementIndex, suffix) {
  const base = 'obj_s' + slideIndex + '_e' + elementIndex;
  return suffix ? base + '_' + suffix : base;
}

/**
 * Transform virtual canvas coordinates to Google Slides points
 * @param {number} x
 * @param {number} y
 * @param {number} w
 * @param {number} h
 * @returns {Object} {x, y, w, h}
 */
function transformCoordinates(x, y, w, h) {
  return {
    x: (x || 0) * SCALE,
    y: (y || 0) * SCALE,
    w: (w || 0) * SCALE,
    h: (h || 0) * SCALE
  };
}

/**
 * Build element transform for Slides API with optional rotation
 * Rotation is applied around the element's center.
 * @returns {Object} Transform object for Slides API
 */
function buildTransform(x, y, rotation, w, h) {
  // Convert to points
  const px = (x || 0) * SCALE;
  const py = (y || 0) * SCALE;
  const pw = (w || 100) * SCALE;
  const ph = (h || 100) * SCALE;

  // No rotation - simple position transform
  if (!rotation || rotation === 0) {
    return {
      scaleX: 1,
      scaleY: 1,
      shearX: 0,
      shearY: 0,
      translateX: px,
      translateY: py,
      unit: 'PT'
    };
  }

  // Convert degrees to radians (positive = clockwise)
  const radians = (rotation * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);

  // Rotate around center logic
  // new_center = center
  // tx = px + pw/2 - (pw/2*cos - ph/2*sin)
  // ty = py + ph/2 - (pw/2*sin + ph/2*cos)
  const tx = px + (pw / 2) * (1 - cos) + (ph / 2) * sin;
  const ty = py + (ph / 2) * (1 - cos) - (pw / 2) * sin;

  return {
    scaleX: cos,
    scaleY: cos,
    shearX: -sin,
    shearY: sin,
    translateX: tx,
    translateY: ty,
    unit: 'PT'
  };
}

/**
 * Build size specification
 * @returns {Object}
 */
function buildSize(w, h) {
  const coords = transformCoordinates(0, 0, w, h);
  return {
    width: { magnitude: coords.w, unit: 'PT' },
    height: { magnitude: coords.h, unit: 'PT' }
  };
}

/**
 * Build a hyperlink request for text or shapes
 * @param {string} objectId 
 * @param {string|Object} link 
 * @param {string} elementType 
 * @returns {Object|null}
 */
function buildLinkRequest(objectId, link, elementType) {
  if (!link) return null;

  // Build the link object
  let linkObj = {};

  if (typeof link === 'string') {
    // Simple URL string
    linkObj.url = link;
  } else if (typeof link === 'object') {
    if (link.url) {
      linkObj.url = link.url;
    } else if (link.slideIndex !== undefined) {
      // Link to specific slide (zero-based index)
      linkObj.slideIndex = link.slideIndex;
    } else if (link.relativeLink) {
      // NEXT_SLIDE, PREVIOUS_SLIDE, FIRST_SLIDE, LAST_SLIDE
      linkObj.relativeLink = link.relativeLink.toUpperCase().replace(/-/g, '_');
    } else if (link.slide !== undefined) {
      // Alias
      linkObj.slideIndex = link.slide;
    }
  }

  // Return appropriate request type
  if (elementType === 'text') {
    return {
      updateTextStyle: {
        objectId: objectId,
        style: {
          link: linkObj
        },
        textRange: { type: 'ALL' },
        fields: 'link'
      }
    };
  } else if (elementType === 'shape') {
    return {
      updateShapeProperties: {
        objectId: objectId,
        shapeProperties: {
          link: linkObj
        },
        fields: 'link'
      }
    };
  }

  return null;
}
