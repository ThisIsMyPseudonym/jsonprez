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
function buildTransform(x, y, rotation, w, h, flipH, flipV) {
  // Convert to points
  const px = (x || 0) * SCALE;
  const py = (y || 0) * SCALE;
  const pw = (w || 100) * SCALE;
  const ph = (h || 100) * SCALE;

  // Rotation setup
  const radians = ((rotation || 0) * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);

  // Base scale factors
  let sx = 1;
  let sy = 1;

  if (flipH) sx = -1;
  if (flipV) sy = -1;

  // IMPORTANT: 
  // If we have rotation AND flip, the order matters.
  // Google Slides API Transform matrix is:
  // [ scaleX  shearX  translateX ]
  // [ shearY  scaleY  translateY ]
  //
  // Our rotation logic below calculates the components for a rotation regarding the center.
  // If we flip, we simply invert the scale components relative to the rotated frame?
  //
  // Let's assume the input (x,y) is the translation component of the matrix (because that's what we extract).
  // AND we extract rotation.
  //
  // IF we are reconstructing from Raw Extraction (where x, y = translateX/Y directly):
  // We want to recreate the matrix [ sx*cos, -sy*sin, x ] ... mixed with rotation.
  //
  // However, existing buildTransform includes logic `px + (pw/2)...` which implies it expects x,y to be TOP-LEFT of unrotated box.
  // BUT AdvancedExtractor emits `x` = `translateX` (matrix component).
  //
  // If `x` comes from `translateX` (matrix), then we should NOT be doing the `(pw/2)` compensation math IF that math was intended to convert "visual top-left" to "matrix translate".
  //
  // HYPOTHESIS: `AdvancedExtractor` emits RAW matrix translation. `buildTransform` applies "Rotation Compensation".
  // This essentially "double compensates" IF the input was already matrix translation.
  //
  // IF `x` is matrix translation, then `translateX` should just be `px`.
  //
  // However, removing that logic might break other things if they rely on it.
  //
  // Let's stick to the "Flip Fix" for now:
  // We want `scaleX` to be negative if flipped.
  // The matrix construction for Rotation + Scale(flip):
  // R = [ cos -sin ]
  //     [ sin  cos ]
  // S = [ sx   0   ]
  //     [ 0    sy  ]
  // R * S = [ sx*cos  -sy*sin ]
  //         [ sx*sin   sy*cos ]
  //
  // `buildTransform` currently returns:
  // scaleX: cos, scaleY: cos, shearX: -sin, shearY: sin
  // This matches Rotation matrix [ cos  -sin ] (shearX is -sin)
  //                              [ sin   cos ] (shearY is sin)
  // Wait, API `shearX` corresponds to row 0, col 1?
  // Slides API:
  // Matrix = [ scaleX shearX translateX ]
  //          [ shearY scaleY translateY ]
  //
  // Rotation theta:
  // [ cos(t)  -sin(t)  0 ]
  // [ sin(t)   cos(t)  0 ]
  //
  // So scaleX=cos, shearX=-sin, shearY=sin, scaleY=cos.
  // This matches current code.
  //
  // If we Flip H (scale x by -1) BEFORE Rotation (Local Flip):
  // R * S_flip = [ cos -sin ] * [ -1  0 ] = [ -cos  -sin ]
  //              [ sin  cos ]   [  0  1 ]   [ -sin   cos ]
  // scaleX = -cos, shearX = -sin
  // shearY = -sin, scaleY = cos
  //
  // If we Flip H AFTER Rotation (Global Flip?):
  // S_flip * R = [ -1  0 ] * [ cos -sin ] = [ -cos  sin ]
  //              [  0  1 ]   [ sin  cos ]   [ sin   cos ]
  //
  // Usually "Flip Horizontal" in editors is a local flip (flip the object, then rotate it).
  //
  // So we apply SX, SY to the rotation components.

  const finalScaleX = sx * cos;
  const finalScaleY = sy * cos;
  const finalShearX = -sy * sin; // If sy=-1 (flipV), this term flips sign? Check math: R*S -> [sx*c, -sy*s; sx*s, sy*c].
  // Actually:
  // [ c -s ] [ sx 0 ]   [ c*sx  -s*sy ]
  // [ s  c ] [ 0 sy ] = [ s*sx   c*sy ]
  //
  // scaleX = c*sx
  // shearX = -s*sy
  // shearY = s*sx
  // scaleY = c*sy

  const finalShearY = sx * sin;

  // Transform Origin Adjustment
  // If we simply pass `px` through as `translateX`, we assume `px` IS the matrix translation.
  // If we keep the `tx` calculation, we assume `px` is top-left of untransformed box.
  //
  // Given `AdvancedExtractor.js`: `x: emuToPt(translateX)`.
  // Using matrix translation directly is safest for fidelity.
  //
  // Let's TRY to use `px` directly as `tx` for the FLIP case or generally?
  // If we change it generally, we might break normal shapes if their `x` was derived differently.
  // But `AdvancedExtractor` is the source.
  //
  // For this specific task (Fix Flip), we will keep the `tx` logic but apply the flipped scales.
  // However, if `flipH` is true, the `tx` logic might need adjustment because `pw` is width.
  //
  // Let's just apply the matrix scaling factors first and see.

  const tx = px + (pw / 2) * (1 - cos) + (ph / 2) * sin;
  const ty = py + (ph / 2) * (1 - cos) - (pw / 2) * sin;

  // NOTE: If we really trust the Extractor's X/Y to be Matrix Translate X/Y, we should just return `px/py`.
  // But let's stick to modifying scale for now to fix the "mirror" issue.

  return {
    scaleX: finalScaleX,
    scaleY: finalScaleY,
    shearX: finalShearX,
    shearY: finalShearY,
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
