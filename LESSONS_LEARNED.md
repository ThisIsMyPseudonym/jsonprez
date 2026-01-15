# Google Slides API: Lessons Learned & Best Practices

> **Purpose**: Comprehensive guide for developers working with Google Slides API extraction and generation. Covers API quirks, proven patterns, and critical gotchas discovered during development.

---

## Table of Contents
1. [Architecture Overview](#architecture-overview)
2. [API Fundamentals](#api-fundamentals)
3. [Extraction Best Practices](#extraction-best-practices)
4. [Generation Best Practices](#generation-best-practices)
5. [Critical Gotchas](#critical-gotchas)
6. [Text Handling](#text-handling)
7. [Colors & Themes](#colors--themes)
8. [Transforms, Rotation & Dimensions](#transforms-rotation--dimensions)
9. [Shapes & Elements](#shapes--elements)
10. [Tables](#tables)
11. [Phase 2 Operations](#phase-2-operations-slidesapp)
12. [Debugging Guide](#debugging-guide)

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────┐
│                        Controller.gs                         │
│  Entry points: generatePresentation() / importPresentation() │
└──────────────────────────┬──────────────────────────────────┘
                           │
       ┌───────────────────┼───────────────────┐
       ▼                   ▼                   ▼
┌──────────────┐   ┌───────────────┐   ┌─────────────────┐
│ Extractor    │   │ SlideBuilders │   │ SlidesApiAdapter│
│ (SlidesApp)  │   │ (Build reqs)  │   │ (Execute API)   │
└──────────────┘   └───────────────┘   └─────────────────┘
                           │
                   ┌───────┴───────┐
                   ▼               ▼
           ┌──────────────┐ ┌───────────────┐
           │ ThemeService │ │ Config/ENUMS  │
           └──────────────┘ └───────────────┘
```

### Two-Phase Execution
| Phase | API Used | Capabilities |
|-------|----------|--------------|
| **Phase 1** | REST `batchUpdate` | Create slides, shapes, text, images, lines, tables |
| **Phase 2** | SlidesApp | Charts, speaker notes, grouping, complex operations |

> **Why Phase 2?** Some operations (embedded charts, speaker notes) are only available through SlidesApp, not the REST API.

---

## API Fundamentals

### Two APIs, Different Purposes

```javascript
// REST API (via Advanced Service) - Fast, batch operations
Slides.Presentations.batchUpdate(resource, presentationId);

// SlidesApp - Object-oriented, some exclusive features
SlidesApp.openById(presentationId).getSlides()[0].getPageElements();
```

### Scale Factor
The API uses **points** (72 DPI), while our JSON uses a virtual 1000-unit canvas.

```javascript
const SCALE = 720 / 1000;  // = 0.72

// Extraction: API → JSON
const jsonSize = apiPoints / SCALE;

// Generation: JSON → API
const apiPoints = jsonSize * SCALE;
```

---

## Extraction Best Practices

### Element Type Detection
```javascript
const type = element.getPageElementType();

switch(type) {
  case SlidesApp.PageElementType.SHAPE:
    // Could be TEXT_BOX or actual shape - check getShapeType()
    const shape = element.asShape();
    if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
      // It's a text box
    } else {
      // It's a shape with possible text
    }
    break;
  case SlidesApp.PageElementType.IMAGE:
    // ...
}
```

### Text Style Extraction - Critical Pattern
> ⚠️ **Key Learning**: `textRange.getTextStyle()` on full range returns `null` for properties that vary.

```javascript
// ❌ WRONG - Returns null for fontSize if text has mixed sizes
const style = textRange.getTextStyle();
const fontSize = style.getFontSize();  // null!

// ✅ CORRECT - Get style from first character
const firstCharRange = textRange.getRange(0, 1);
const style = firstCharRange.getTextStyle();
const fontSize = style.getFontSize();  // 32 (actual value)
```

### Text Runs for Mixed Formatting
When text has multiple formats (bold title + regular body), extract each run:

```javascript
function extractTextRuns(shape) {
  const textRange = shape.getText();
  const runs = [];
  const textRuns = textRange.getRuns();
  
  for (let i = 0; i < textRuns.length; i++) {
    const run = textRuns[i];
    const style = run.getTextStyle();
    
    runs.push({
      text: run.asString(),
      fontSize: style.getFontSize() / SCALE,
      bold: style.isBold() === true,
      // ... other properties
    });
  }
  return runs;
}
```

### Stripping Invisible Characters
Google Slides text often contains invisible Unicode characters that cause issues.

```javascript
// Strip invisible chars but PRESERVE line breaks
const cleanText = rawText.replace(
  /[\x00-\x09\x0B\x0C\x0E-\x1F\x7F\u200B-\u200D\uFEFF\uFFFC\uFFFD]/g, 
  ''
).trim();

// Note: \x0A (newline) and \x0D (carriage return) are preserved
```

---

## Generation Best Practices

### Request Building Pattern
All element creation follows this pattern:

```javascript
function buildElementRequests(element, slideId) {
  const requests = [];
  const shapeId = element.objectId || Utilities.getUuid();
  
  // 1. Create the element
  requests.push({
    createShape: {
      objectId: shapeId,
      shapeType: 'TEXT_BOX',
      elementProperties: {
        pageObjectId: slideId,
        size: { width: { magnitude: w, unit: 'PT' }, height: { magnitude: h, unit: 'PT' }},
        transform: { scaleX: 1, scaleY: 1, translateX: x, translateY: y, unit: 'PT' }
      }
    }
  });
  
  // 2. Insert text
  requests.push({
    insertText: { objectId: shapeId, text: element.text, insertionIndex: 0 }
  });
  
  // 3. Style the text
  requests.push({
    updateTextStyle: {
      objectId: shapeId,
      style: { fontSize: { magnitude: 14, unit: 'PT' }, bold: true },
      textRange: { type: 'ALL' },
      fields: 'fontSize,bold'
    }
  });
  
  return requests;
}
```

### Text Runs - Index Alignment Critical
> ⚠️ **Key Learning**: When using `textRuns`, the inserted text MUST match the cumulative textRuns length exactly.

```javascript
// ❌ WRONG - textContent may differ from textRuns after stripping
requests.push({ insertText: { objectId: shapeId, text: textContent } });
// Then apply textRuns styles with FIXED_RANGE → INDEX MISMATCH ERROR

// ✅ CORRECT - Use concatenated textRuns text
const runsText = element.textRuns.map(r => r.text).join('');
requests.push({ insertText: { objectId: shapeId, text: runsText } });

// Now indices match
let currentIndex = 0;
for (const run of element.textRuns) {
  requests.push({
    updateTextStyle: {
      objectId: shapeId,
      style: { fontSize: { magnitude: run.fontSize * SCALE, unit: 'PT' }},
      textRange: { 
        type: 'FIXED_RANGE',
        startIndex: currentIndex,
        endIndex: currentIndex + run.text.length
      },
      fields: 'fontSize'
    }
  });
  currentIndex += run.text.length;
}
```

---

## Critical Gotchas

### 1. Empty Text Styling Error
```
Error: The object (...) has no text.
```
**Cause**: Trying to style text that contains only invisible characters.  
**Solution**: Validate text has visible content before styling:
```javascript
const hasTextContent = textContent && textContent.length > 0 && /[\w\d\p{L}\p{P}]/u.test(textContent);
if (hasTextContent) {
  // Safe to insert and style
}
```

### 2. Index Out of Bounds Error
```
Error: The end index (126) should not be greater than the existing text length (124).
```
**Cause**: textRuns length doesn't match inserted text.  
**Solution**: Use concatenated textRuns text for insertion (see above).

### 3. null from getTextStyle()
**Cause**: Mixed styles in text range.  
**Solution**: Get style from first character or iterate through runs.

### 4. Theme Colors Not Resolved
```json
"fillColor": "theme:ACCENT1"  // Not an actual color!
```
**Cause**: `asRgbColor()` fails for some theme colors during extraction.  
**Solution**: Handle in generation via ThemeService mapping:
```javascript
if (colorValue.startsWith('theme:')) {
  const mapping = {
    'ACCENT1': '#4285f4',
    'ACCENT2': '#34a853',
    // ...
  };
  return mapping[colorValue.substring(6)];
}
```

### 5. Cache/Propagation Issues
After `batchUpdate`, changes may not immediately appear in SlidesApp.
**Solution**: Save, close, and re-open:
```javascript
presentation.saveAndClose();
Utilities.sleep(1000);
presentation = SlidesApp.openById(presentationId);
```

### 6. TEXT_BOX Default Newline Styling
> ⚠️ **Key Learning**: When you create a TEXT_BOX shape, Google Slides automatically creates it with a default newline character that has default styling (typically Arial 18pt). This character persists even when you insert new text.

**Problem**: When you insert text at index 0, the default newline gets pushed to the END of your text content. If you style only your inserted text, this trailing newline keeps its default Arial 18pt styling, which can affect text positioning and appearance.

```
Before insertion: "\n" (default, styled Arial 18)
After inserting "\nBing\n" at index 0:
  Index 0: \n (your text - styled correctly)
  Index 1-4: Bing (your text - styled correctly)
  Index 5: \n (your text - styled correctly)
  Index 6: \n (default newline - STILL Arial 18!)
```

**Symptom**: Text appears vertically misaligned because the trailing newline has a different (larger) font size.

**Solution**: After styling all your text runs, add one more `updateTextStyle` request to style the trailing default newline:

```javascript
// After styling all runs, style the trailing default newline
const runsText = element.textRuns.map(r => r.text).join('');
const lastRun = element.textRuns[element.textRuns.length - 1];

requests.push({
  updateTextStyle: {
    objectId: shapeId,
    style: {
      fontSize: { magnitude: lastRun.fontSize || defaultFontSize, unit: 'PT' },
      fontFamily: lastRun.fontFamily || defaultFontFamily
    },
    textRange: {
      type: 'FIXED_RANGE',
      startIndex: runsText.length,      // Position of trailing newline
      endIndex: runsText.length + 1     // Just that one character
    },
    fields: 'fontSize,fontFamily'
  }
});
```

> **Note**: This only applies when using `FIXED_RANGE` for individual runs. If using `textRange: { type: 'ALL' }` for styling, the trailing newline is automatically included.

---

## Text Handling

### Text Box vs Shape Text
| Property | Text Box | Shape with Text |
|----------|----------|-----------------|
| Element Type | `SHAPE` | `SHAPE` |
| Shape Type | `TEXT_BOX` | `RECTANGLE`, etc. |
| Text Access | `shape.getText()` | `shape.getText()` |
| JSON Type | `"type": "text"` | `"type": "shape"` with `"text"` property |

### Alignment Mapping
```javascript
const ALIGNMENT_MAP = {
  'left': 'START',
  'center': 'CENTER', 
  'right': 'END',
  'justify': 'JUSTIFIED'
};
```

---

## Colors & Themes

### Color Extraction Flow
```javascript
function extractColor(colorObj) {
  const colorType = colorObj.getColorType();
  
  if (colorType === SlidesApp.ColorType.RGB) {
    return rgbToHex(colorObj.asRgbColor());
  } 
  else if (colorType === SlidesApp.ColorType.THEME) {
    // Try to resolve to RGB
    try {
      return rgbToHex(colorObj.asRgbColor());
    } catch (e) {
      // Fallback to theme reference
      return 'theme:' + colorObj.getThemeColorType();
    }
  }
  return '#000000';  // Fallback
}
```

### Theme Color Resolution (Generation)
ThemeService maps theme references to actual colors:
```javascript
const themeColorMapping = {
  'DARK1': '#1e293b',
  'LIGHT1': '#ffffff',
  'ACCENT1': '#4285f4',  // Google Blue
  'ACCENT2': '#34a853',  // Google Green
  'ACCENT3': '#fbbc04',  // Google Yellow
  'ACCENT4': '#ea4335',  // Google Red
  // ...
};
```

### Theme Color Resolution (Extraction) - Critical Pattern

> ⚠️ **Key Learning**: The Advanced Slides API returns theme color references (e.g., `themeColor: "DARK1"`) without the actual RGB values. The color scheme that defines what these references mean varies per slide based on which master it uses.

**Problem**: When slides are pasted with "Keep original styles" from another presentation, they visually preserve colors but the API still returns theme color references. The rendered color depends on which master/color scheme the slide uses.

**Solution**: Use SlidesApp to resolve theme colors via the slide's own color scheme:

```javascript
// ❌ WRONG - Using first master's color scheme for all slides
const colorScheme = masters[0].getColorScheme();
const darkColor = colorScheme.getConcreteColor(SlidesApp.ThemeColorType.DARK1);

// ✅ CORRECT - Use each slide's specific color scheme
const colorScheme = slide.getColorScheme();
const darkColor = colorScheme.getConcreteColor(SlidesApp.ThemeColorType.DARK1);
const hex = darkColor.asRgbColor().asHexString();
```

> ⚠️ **Key Learning**: You cannot call `color.asRgbColor()` directly on a theme color - it throws "Object is not of type RgbColor". You must use `colorScheme.getConcreteColor()` to resolve theme colors.

```javascript
// ❌ WRONG - Throws error for theme colors
const color = solidFill.getColor();
if (color.getColorType() === SlidesApp.ColorType.THEME) {
  const rgb = color.asRgbColor();  // ERROR: Object is not of type RgbColor
}

// ✅ CORRECT - Look up via color scheme
if (color.getColorType() === SlidesApp.ColorType.THEME) {
  const themeColor = color.asThemeColor();
  const themeType = themeColor.getThemeColorType();
  const colorScheme = slide.getColorScheme();
  const resolved = colorScheme.getConcreteColor(themeType);
  const hex = resolved.asRgbColor().asHexString();
}
```

### Hybrid Extraction Pattern (SlidesApp + Advanced API)

For accurate color extraction, use a two-pass approach:

1. **Pre-pass with SlidesApp**: Build a cache of resolved fill colors before extraction
2. **Main extraction with Advanced API**: Use the cache for theme color resolution

```javascript
// Cache of resolved fill colors: objectId -> hex color
let _resolvedFillColors = {};

function buildResolvedColorCache(presentationId) {
  const presentation = SlidesApp.openById(presentationId);
  const slides = presentation.getSlides();

  for (const slide of slides) {
    const colorScheme = slide.getColorScheme();
    // Build theme color map for THIS slide
    const themeColorMap = {};
    for (const themeType of themeColorTypes) {
      const concrete = colorScheme.getConcreteColor(themeType);
      themeColorMap[themeType.toString()] = concrete.asRgbColor().asHexString();
    }

    // Extract colors for all shapes on this slide
    for (const element of slide.getPageElements()) {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = element.asShape();
        const color = shape.getFill().getSolidFill().getColor();

        if (color.getColorType() === SlidesApp.ColorType.RGB) {
          _resolvedFillColors[element.getObjectId()] = color.asRgbColor().asHexString();
        } else if (color.getColorType() === SlidesApp.ColorType.THEME) {
          const themeType = color.asThemeColor().getThemeColorType();
          _resolvedFillColors[element.getObjectId()] = themeColorMap[themeType.toString()];
        }
      }
    }
  }
}

// During extraction, check cache first
function extractFill(fill, objectId) {
  if (fill.solidFill?.color?.themeColor && _resolvedFillColors[objectId]) {
    return _resolvedFillColors[objectId];  // Use pre-resolved color
  }
  // ... fallback to direct extraction
}
```

> **Why this matters**: A presentation can have multiple masters with different color schemes. A shape on slide 1 using `DARK1` might be blue (from master A), while a shape on slide 2 using `DARK1` might be black (from master B). Using `slide.getColorScheme()` ensures each slide's colors are resolved correctly.

### Theme Color Resolution Timing - Critical Architecture Issue (SOLVED)

> ⚠️ **Key Learning**: Theme colors must be resolved per-slide because different slides can use different masters with different color schemes.

**Original Problem**:
The original `_resolvedThemeColorMap` was a single global map that got overwritten for each slide during `buildResolvedColorCache()`. By the time extraction ran, only the LAST slide's theme colors remained in memory.

```
buildResolvedColorCache() runs:
  - Slide 0: _resolvedThemeColorMap = { LIGHT2: '#1a5c30' }  // dark green
  - Slide 1: _resolvedThemeColorMap = { LIGHT2: '#f0f0f0' }  // overwrites!
  - Slide 2: _resolvedThemeColorMap = { LIGHT2: '#e0e0e0' }  // overwrites!

extractSlides() runs AFTER cache loop completes:
  - All slides use Slide 2's colors → Slide 0's dark green text becomes grey
```

**Failed Fix Attempt**: Using only slide 0's colors for all slides caused theme colors to apply incorrectly to slides using different masters.

**Implemented Solution**: Per-slide color map storage with slide index threading:

```javascript
// Store per-slide color maps (keyed by slide index)
let _resolvedThemeColorMapsPerSlide = {};

// Helper to retrieve per-slide colors
function getResolvedThemeColorMap(slideIndex) {
    return _resolvedThemeColorMapsPerSlide[slideIndex] || {};
}

// During cache building - store each slide's colors separately:
for (let slideIndex = 0; slideIndex < slides.length; slideIndex++) {
    const slideThemeColorMap = {};
    const colorScheme = slide.getColorScheme();
    for (const themeType of themeColorTypes) {
        const concreteColor = colorScheme.getConcreteColor(themeType);
        slideThemeColorMap[themeType.toString()] = concreteColor.asRgbColor().asHexString();
    }
    _resolvedThemeColorMapsPerSlide[slideIndex] = slideThemeColorMap;
}

// During extraction - pass slideIndex through the entire pipeline:
function extractSlideAdvanced(slide, slideIndex) { ... }
function extractElementAdvanced(element, parentTransform, layoutId, slideIndex) { ... }
function extractShapeAdvanced(element, base, layoutId, slideIndex) { ... }
function extractTextAdvanced(textElements, slideIndex) {
    const slideThemeColors = getResolvedThemeColorMap(slideIndex);
    // Use slideThemeColors[themeKey] for lookups
}
function extractTableAdvanced(element, base, slideIndex) { ... }
```

**Key Implementation Details**:
1. `buildResolvedColorCache()` stores colors in `_resolvedThemeColorMapsPerSlide[slideIndex]`
2. `slideIndex` is threaded through: `extractSlideAdvanced` → `extractElementAdvanced` → `extractShapeAdvanced`/`extractTableAdvanced` → `extractTextAdvanced`
3. Color lookup priority: `slideThemeColors[themeKey]` → `_themeColorMap[themeKey]` → `_activeThemeColors[themeKey]` → `DEFAULT_THEME_COLORS[themeKey]`

This ensures each slide's text colors are resolved using that slide's master's color scheme.

---

## Transforms, Rotation & Dimensions

### Understanding the Transform Matrix

Google Slides API represents position, scale, and rotation as a 2D affine transformation matrix:

```
Transform = [ scaleX   shearX   translateX ]
            [ shearY   scaleY   translateY ]
            [   0        0          1      ]
```

The API returns these as individual properties: `scaleX`, `scaleY`, `shearX`, `shearY`, `translateX`, `translateY`.

### Rotation is Encoded in the Matrix - Critical Pattern

> ⚠️ **Key Learning**: Rotation is NOT stored as a separate property. It's encoded in the scaleX/scaleY/shearX/shearY values using trigonometry.

For a rotated + scaled shape, the matrix components are:
```
[ scaleX, shearX ] = [ Sw × cos(θ),  -Sh × sin(θ) ]
[ shearY, scaleY ]   [ Sw × sin(θ),   Sh × cos(θ) ]
```
Where:
- `Sw` = width scale factor
- `Sh` = height scale factor
- `θ` = rotation angle in radians

**Extracting actual dimensions from a rotated shape:**
```javascript
// Extract scale factors (actual width/height multipliers)
const scaleW = Math.sqrt(scaleX * scaleX + shearY * shearY);
const scaleH = Math.sqrt(shearX * shearX + scaleY * scaleY);

// Extract rotation angle
const rotationRad = Math.atan2(shearY, scaleX);
const rotationDeg = rotationRad * (180 / Math.PI);

// Actual dimensions
const actualWidth = baseWidth * scaleW;
const actualHeight = baseHeight * scaleH;
```

### API Omits Zero Values - Critical Gotcha

> ⚠️ **Key Learning**: The API omits transform properties that are 0. For a 270° rotation, `cos(270°) = 0`, so `scaleX` and `scaleY` are omitted. You must NOT default these to 1.

```javascript
// ❌ WRONG - Assumes scaleX=1 when omitted
const scaleX = transform.scaleX || 1;  // Breaks for 270° rotation!

// ✅ CORRECT - Default to 0 if other transform values are present
const hasAnyTransformValue = transform.scaleX !== undefined ||
                              transform.scaleY !== undefined ||
                              transform.shearX !== undefined ||
                              transform.shearY !== undefined;
const defaultScale = hasAnyTransformValue ? 0 : 1;

const scaleX = transform.scaleX !== undefined ? transform.scaleX : defaultScale;
const scaleY = transform.scaleY !== undefined ? transform.scaleY : defaultScale;
```

### Translation vs Top-Left Position

The API's `translateX`/`translateY` represent the transformed position (related to the shape's anchor/center after rotation), NOT the top-left corner. For our JSON schema, we need the top-left corner.

**Conversion formulas:**
```javascript
// Given: translateX, translateY (from API)
// Given: actualWidth, actualHeight, rotationRad (computed above)
const cos = Math.cos(rotationRad);
const sin = Math.sin(rotationRad);
const halfW = actualWidth / 2;
const halfH = actualHeight / 2;

// Convert to top-left corner
const topLeftX = translateX - halfW * (1 - cos) - halfH * sin;
const topLeftY = translateY - halfH * (1 - cos) + halfW * sin;
```

**Reverse (for generation - top-left to translation):**
```javascript
// Given: x, y (top-left corner from JSON)
const translateX = x + halfW * (1 - cos) + halfH * sin;
const translateY = y + halfH * (1 - cos) - halfW * sin;
```

### Composing Transforms for Groups - Critical Pattern

> ⚠️ **Key Learning**: Children's transforms are LOCAL to the group coordinate space, NOT world coordinates. The API documentation confirms: "The transforms of a group's children are relative to the group's transform."

When elements are inside groups, their transform is relative to the group. To get the absolute transform, multiply (compose) parent × child matrices:

```javascript
function composeTransforms(parent, child) {
  if (!parent) return child;
  if (!child) return parent;

  // Parent matrix components (default to identity if missing)
  const pScaleX = parent.scaleX !== undefined ? parent.scaleX : 1;
  const pScaleY = parent.scaleY !== undefined ? parent.scaleY : 1;
  const pShearX = parent.shearX || 0;
  const pShearY = parent.shearY || 0;
  const pTranslateX = parent.translateX || 0;
  const pTranslateY = parent.translateY || 0;

  // Child matrix components
  const cScaleX = child.scaleX !== undefined ? child.scaleX : 1;
  const cScaleY = child.scaleY !== undefined ? child.scaleY : 1;
  const cShearX = child.shearX || 0;
  const cShearY = child.shearY || 0;
  const cTranslateX = child.translateX || 0;
  const cTranslateY = child.translateY || 0;

  // 2x2 matrix multiplication for the linear part
  return {
    scaleX: pScaleX * cScaleX + pShearX * cShearY,
    shearX: pScaleX * cShearX + pShearX * cScaleY,
    shearY: pShearY * cScaleX + pScaleY * cShearY,
    scaleY: pShearY * cShearX + pScaleY * cScaleY,
    // Translation: parent translation + parent matrix × child translation
    translateX: pTranslateX + pScaleX * cTranslateX + pShearX * cTranslateY,
    translateY: pTranslateY + pShearY * cTranslateX + pScaleY * cTranslateY
  };
}
```

### Flip Detection

Flips (horizontal or vertical mirroring) are detected via the matrix determinant:

```javascript
const det = scaleX * scaleY - shearX * shearY;
const hasFlip = det < 0;  // Negative determinant = one flip

// Note: Both flipH AND flipV together = 180° rotation (det > 0)
// To distinguish flipH vs flipV, additional context is needed.
// Default to flipH (most common case).
```

### Complete Dimension Extraction Example

```javascript
function extractDimensions(element, parentTransform) {
  // Compose with parent if inside a group
  const transform = parentTransform
    ? composeTransforms(parentTransform, element.transform || {})
    : (element.transform || {});

  // Handle API omitting zero values
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

  // Extract true scale and rotation
  const scaleW = Math.sqrt(scaleX * scaleX + shearY * shearY);
  const scaleH = Math.sqrt(shearX * shearX + scaleY * scaleY);
  const rotationRad = Math.atan2(shearY, scaleX);
  let rotationDeg = rotationRad * (180 / Math.PI);
  if (rotationDeg < 0) rotationDeg += 360;

  // Actual dimensions
  const actualWidth = baseWidth * scaleW;
  const actualHeight = baseHeight * scaleH;

  // Convert translation to top-left
  const cos = Math.cos(rotationRad);
  const sin = Math.sin(rotationRad);
  const halfW = actualWidth / 2;
  const halfH = actualHeight / 2;
  const topLeftX = translateX - halfW * (1 - cos) - halfH * sin;
  const topLeftY = translateY - halfH * (1 - cos) + halfW * sin;

  // Flip detection
  const det = scaleX * scaleY - shearX * shearY;
  const flipH = det < 0;

  return {
    x: emuToPt(topLeftX),
    y: emuToPt(topLeftY),
    w: emuToPt(actualWidth),
    h: emuToPt(actualHeight),
    rotation: rotationDeg,
    flipH: flipH,
    flipV: false
  };
}
```

---

## Shapes & Elements

### Shape Type Mapping
```javascript
// User-friendly → API constant
const SHAPE_TYPE_MAP = {
  'rectangle': 'RECTANGLE',
  'roundRect': 'ROUND_RECTANGLE',
  'ellipse': 'ELLIPSE',
  'circle': 'ELLIPSE',
  'triangle': 'TRIANGLE',
  // ... many more
};
```

### Shape Adjustments (Corner Radius) - API Limitation

> ⚠️ **Key Learning**: The Google Slides API does NOT expose shape adjustment values (like corner radius for rounded rectangles).

In the Google Slides UI, you can drag the yellow diamond handles to adjust properties like:
- Corner radius on `ROUND_RECTANGLE` shapes
- Arrow head sizes on arrows
- Callout pointer positions

**However, the API returns NO adjustment data:**
```javascript
// What we receive from the API for a ROUND_RECTANGLE:
{
  shapeType: "ROUND_RECTANGLE",
  shapeProperties: {
    contentAlignment: "...",
    shapeBackgroundFill: {...},
    outline: {...},
    shadow: {...},
    autofit: {...}
    // NO "adjustments" property!
  }
}
```

**Impact**: When extracting and regenerating rounded rectangles:
- The shape type (`ROUND_RECTANGLE`) is preserved
- The corner radius resets to the default value
- Shapes with "maxed out" corner curves will appear less curved after regeneration

**Workaround**: None available. Neither the REST API nor SlidesApp expose shape adjustment values. The Shape class in SlidesApp has ~60 methods but none for corner radius or geometry adjustments. This is a fundamental limitation of both APIs.

---

### Unsupported Elements
Some elements can't be extracted/recreated:
- **WordArt** - Rendered as `unsupported`
- **SmartArt** - Rendered as `unsupported`
- **Linked Charts** - Recreated via Phase 2 with spreadsheet
- **Custom Shape Adjustments** - Corner radius, callout positions (API doesn't expose)

### Outline Dash Styles

Shape outlines can have dash styles (solid, dotted, dashed, etc.). The API provides these values:
- `SOLID`, `DOT`, `DASH`, `DASH_DOT`, `LONG_DASH`, `LONG_DASH_DOT`

**Extraction**: Check `outline.dashStyle` and map to JSON-friendly values:
```javascript
function extractOutlineDashStyle(outline) {
    if (!outline || outline.propertyState === 'NOT_RENDERED') return 'solid';
    const dashStyle = outline.dashStyle;
    if (!dashStyle || dashStyle === 'SOLID') return 'solid';
    const dashMap = {
        'DOT': 'dot', 'DASH': 'dash', 'DASH_DOT': 'dashDot',
        'LONG_DASH': 'longDash', 'LONG_DASH_DOT': 'longDashDot'
    };
    return dashMap[dashStyle] || 'solid';
}
```

**Generation**: Use `DASH_STYLE_MAP` in Config.js to convert back to API values.

### Phase 2 Image Z-Order - Critical Pattern

> **Key Learning**: Images routed through Phase 2 (SlidesApp) are inserted LAST, putting them on top of all shapes created in Phase 1.

**Problem**: Google-internal image URLs (googleusercontent.com) can't be fetched by the REST API. They must be fetched via SlidesApp using `UrlFetchApp` with OAuth token, then inserted via `slide.insertImage()`. This inserts them after all other elements, breaking z-order.

**Symptom**: Callout boxes, overlays, and shapes that should be in front of images appear behind them.

**Solution - Use `sendBackward()` with Correct Count**:

> ⚠️ **Key Learning**: Do NOT use `sendToBack()` for z-order restoration. It moves the element behind ALL other elements. Instead, use `sendBackward()` the correct number of times based on the element's original position.

```javascript
// Track original element index during building
phase2Service.addProactiveImage(slideIndex, {
  objectId: objectId,
  slideId: slideId,
  element: element,
  elementIndex: elementIndex,     // Original position in z-order
  totalElements: totalElements    // Total elements on slide
});

// In Phase 2, calculate correct positioning
const elementsInFront = totalElements - elementIndex - 1;

if (elementsInFront > 0) {
  // Move backward the exact number of positions needed
  for (let j = 0; j < elementsInFront; j++) {
    image.sendBackward();
  }
}
```

**Example**: An image originally at index 26/30 should have 3 elements in front of it (indices 27, 28, 29). Call `sendBackward()` 3 times, NOT `sendToBack()` which would put it behind all 29 other elements.

### CopyGroup - Handling Curved/Freeform Lines

> ⚠️ **Key Learning**: Groups containing curved or freeform lines CANNOT be accurately reproduced via the Slides API because it doesn't expose Bézier curve control points.

**Solution**: Capture these groups as cropped slide thumbnails:

1. **Detection**: During extraction, detect groups with curved lines:
```javascript
const hasCurvedLines = children.some(child => {
  if (child.line) {
    const lineType = child.line.lineType || '';
    const lineCategory = child.line.lineCategory || '';
    // Freeform lines have no lineType, or are CURVED connectors
    return lineCategory === 'CURVED' ||
           lineType.includes('CURVED') ||
           (lineType === '' && lineCategory === '');
  }
  return false;
});
```

2. **Bounding Box Calculation**: Calculate group bounds from children, composing transforms:
```javascript
// Get GROUP's transform
const grpScaleX = groupTransform.scaleX || 1;
const grpTranslateX = groupTransform.translateX || 0;
// ... etc

// For each child, COMPOSE with group transform to get world coordinates
const childWorldX = grpScaleX * childLocalX + grpShearX * childLocalY + grpTranslateX;
const childWorldY = grpShearY * childLocalX + grpScaleY * childLocalY + grpTranslateY;
```

3. **Skip Template Connectors**: Elements with base size 3000000 EMU (236pt) are template placeholders that distort bounding box calculations:
```javascript
const TEMPLATE_SIZE = 3000000; // EMU
if (childBaseW === TEMPLATE_SIZE || childBaseH === TEMPLATE_SIZE) {
  return; // Skip this child
}
```

4. **Bounds Validation**: Skip groups entirely outside slide bounds:
```javascript
if (groupXEmu >= slideWidth || groupYEmu >= slideHeight ||
    groupRight <= 0 || groupBottom <= 0) {
  Logger.log('CopyGroup: SKIPPED - group entirely outside slide bounds');
  return;
}
```

5. **Phase 2 Processing**: Fetch slide thumbnail, crop to group bounds, apply transform:
```javascript
// Get thumbnail of source slide
const thumbnailResponse = Slides.Presentations.Pages.getThumbnail(
  sourcePresentationId, sourceSlide.objectId,
  { 'thumbnailProperties.mimeType': 'PNG', 'thumbnailProperties.thumbnailSize': 'LARGE' }
);

// Calculate crop fractions
const cropLeft = groupXEmu / slideWidth;
const cropTop = groupYEmu / slideHeight;
const cropRight = 1 - ((groupXEmu + groupWEmu) / slideWidth);
const cropBottom = 1 - ((groupYEmu + groupHEmu) / slideHeight);

// Scale = targetSize / sourceSize (simple ratio)
const scaleX = groupWEmu / srcW;
const scaleY = groupHEmu / srcH;
```

**Key Insight**: The crop determines what portion of the source fills the element. The visible portion is STRETCHED to fill the element rectangle. Scale = targetSize / sourceSize (no visible fraction math needed).

6. **Stroke Padding**: Lines extend beyond their geometric bounds due to stroke width. Add padding to the bounding box:
```javascript
const STROKE_PADDING = 4 * 12700; // 4pt in EMU
minX -= STROKE_PADDING;
minY -= STROKE_PADDING;
maxX += STROKE_PADDING;
maxY += STROKE_PADDING;
```

---

## Tables

### Table Cell Extraction
```javascript
for (let r = 0; r < table.getNumRows(); r++) {
  for (let c = 0; c < table.getNumColumns(); c++) {
    const cell = table.getCell(r, c);
    const text = cell.getText().asString().trim();
    const textStyle = cell.getText().getTextStyle();
    
    // Get cell fill
    const fill = cell.getFill();
    let fillColor = 'transparent';
    if (fill.getType() === SlidesApp.FillType.SOLID) {
      fillColor = extractColor(fill.getSolidFill().getColor());
    }
  }
}
```

### Table Generation
Tables are created row-by-row:
```javascript
requests.push({
  createTable: {
    objectId: tableId,
    rows: data.length,
    columns: data[0].length,
    elementProperties: { /* position */ }
  }
});

// Then populate cells with insertText and updateTableCellProperties
```

---

## Phase 2 Operations (SlidesApp)

### When to Use Phase 2
- **Speaker Notes** - Only via SlidesApp
- **Embedded Charts** - Requires SpreadsheetApp + SlidesApp
- **Element Grouping** - Only via SlidesApp
- **Complex Animations** - Not available in REST API

### Synchronization Pattern
```javascript
// Wait for slides to propagate after batchUpdate
for (let i = 0; i < maxRetries; i++) {
  if (slides.length >= expectedSlideCount) break;
  
  presentation.saveAndClose();
  Utilities.sleep(1000);
  presentation = SlidesApp.openById(presentationId);
  slides = presentation.getSlides();
}
```

---

## Debugging Guide

### Enable Logging
```javascript
Logger.log('Processing element: ' + JSON.stringify(element));
```

### Common Error Messages

| Error | Likely Cause | Solution |
|-------|-------------|----------|
| `has no text` | Styling empty/invisible text | Validate visible content |
| `end index greater than text length` | textRuns index mismatch | Use concatenated textRuns text |
| `Invalid enum value` | Bad shape/arrow type | Check ENUMS mapping |
| `Resource not found` | Element doesn't exist yet | Add synchronization delay |

### Viewing Logs
1. Open Apps Script editor
2. Click **Executions** in left sidebar
3. Click on execution to see logs

### Request Validation
```javascript
Logger.log('Executing batchUpdate. Req count: ' + requests.length);

// Log problematic requests
requests.forEach((req, i) => {
  if (req.updateTextStyle) {
    Logger.log(`Request ${i}: updateTextStyle on ${req.updateTextStyle.objectId}`);
  }
});
```

---

## Quick Reference

### JSON Schema (Simplified)
```json
{
  "config": {
    "title": "Presentation Title",
    "theme": {
      "colors": { "primary": "#4285f4" },
      "fonts": { "heading": "Google Sans" }
    }
  },
  "slides": [{
    "background": "#ffffff",
    "elements": [
      {
        "type": "text",
        "x": 100, "y": 100, "w": 400, "h": 50,
        "text": "Hello World",
        "fontSize": 24,
        "bold": true,
        "textRuns": [
          { "text": "Hello ", "bold": true, "fontSize": 24 },
          { "text": "World", "bold": false, "fontSize": 18 }
        ]
      }
    ]
  }]
}
```

### File Responsibilities
| File | Purpose |
|------|---------|
| `Controller.gs` | Entry points, orchestration |
| `PresentationExtractor.gs` | Read slides → JSON |
| `SlideBuilders.gs` | JSON → API requests |
| `SlidesApiAdapter.gs` | Execute API calls |
| `ThemeService.gs` | Color/font resolution |
| `Config.gs` | Constants, enums, mappings |
| `Validation.gs` | Input validation |
| `Phase2Service.gs` | Queue for SlidesApp operations |

---

*Last updated: January 15, 2026*
