# JsonPrez Production Readiness Plan

> **Version**: Orion v7.12 → v8.0 Production Release
> **Date**: January 15, 2026
> **Objective**: Stabilize, harden, and finalize JsonPrez as a production-ready Google Slides JSON compiler

---

## Executive Summary

JsonPrez is a mature Google Slides API wrapper that enables JSON-to-Slides and Slides-to-JSON workflows. After extensive development, the codebase has strong feature coverage but requires hardening in error handling, code quality, and testing before production deployment.

### Current State Assessment

| Category | Status | Risk Level |
|----------|--------|------------|
| **Feature Completeness** | 85% | Low |
| **Error Handling** | 60% | **High** |
| **Code Quality** | 65% | Medium |
| **Testing Coverage** | 20% | **High** |
| **Documentation** | 80% | Low |
| **Deployment Pipeline** | 90% | Low |

---

## Phase 1: Critical Bug Fixes & Stability (Week 1)

### 1.1 Silent Error Swallowing (CRITICAL)

**Problem**: Multiple locations swallow errors without user notification.

| Location | Issue | Fix |
|----------|-------|-----|
| `Client/Client.html:76-78` | JSON parse errors hidden | Add toast notification on parse failure |
| `Client/Client.html:1000` | formatJSON() fails silently | Show error in editor status |
| `Controller.js:252` | Logging errors swallowed | Remove empty catch or log to error channel |
| `SlidesApiAdapter.js:765` | CopyGroup errors logged only | Surface to response object |

**Implementation**:
```javascript
// Client.html - Replace silent catch
try {
  presentationData = JSON.parse(val);
} catch (e) {
  showToast('Invalid JSON: ' + e.message, 'error');
  return;
}
```

### 1.2 API Quota/Rate Limit Handling (CRITICAL)

**Problem**: No handling for 429 (Too Many Requests) or 403 (Quota Exceeded) responses.

**Solution**: Implement exponential backoff in `Controller.js`:

```javascript
// Add to Config.js
RETRY: {
  MAX_ATTEMPTS: 5,
  INITIAL_DELAY_MS: 1000,
  MAX_DELAY_MS: 32000,
  BACKOFF_MULTIPLIER: 2,
  QUOTA_ERROR_CODES: [429, 403]
}

// Implement in SlidesApiAdapter.js
function executeWithBackoff(operation, maxAttempts = 5) {
  let delay = CONFIG.RETRY.INITIAL_DELAY_MS;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return operation();
    } catch (e) {
      if (isQuotaError(e) && attempt < maxAttempts) {
        Logger.log('Quota limit hit, backing off: ' + delay + 'ms');
        Utilities.sleep(delay);
        delay = Math.min(delay * CONFIG.RETRY.BACKOFF_MULTIPLIER, CONFIG.RETRY.MAX_DELAY_MS);
      } else {
        throw e;
      }
    }
  }
}
```

### 1.3 Array Bounds Validation

**Problem**: Phase 2 operations assume slide indices are valid without checking.

**Locations**: `SlidesApiAdapter.js` lines 197, 529, 579

**Fix**:
```javascript
// Before accessing slides[index]
if (index < 0 || index >= slides.length) {
  Logger.log('Invalid slide index: ' + index + ' (total: ' + slides.length + ')');
  return;
}
```

---

## Phase 2: Code Quality Improvements (Week 2)

### 2.1 Eliminate Code Duplication

**Priority 1: Consolidate Retry Logic**
- Merge `executeBatchWithRetry` and `executeBatchWithRetryV2` in Controller.js
- Remove unused `executeBatchWithRetry` (dead code)

**Priority 2: Extract Color Helper**
```javascript
// New file: Server/ColorHelper.js
const ColorHelper = {
  toRgbApi(color, themeService) {
    const resolved = themeService.resolveThemeColor(color);
    return hexToRgbApi(resolved);
  },

  isTransparent(color) {
    return color === 'transparent' || color === 'none' || !color;
  }
};
```

**Priority 3: Centralize Alignment Maps**
- Use existing `ENUMS.VERTICAL_ALIGNMENT_MAP` from Config.js
- Remove duplicate `contentAlignmentMap` definitions in SlideBuilders.js

### 2.2 Extract Magic Numbers to Config

Add to `Config.js`:
```javascript
RUNTIME: {
  LOCK_TIMEOUT_MS: 30000,
  LOG_TRUNCATE_LENGTH: 50,
  SHEAR_TOLERANCE: 0.001,
  MIN_COLUMN_WIDTH_PT: 32
},

RETRY: {
  MAX_ATTEMPTS: 5,
  ERROR_PATTERN: /Invalid requests\[(\d+)\]/
}
```

### 2.3 Remove Dead Code

| Item | Location | Action |
|------|----------|--------|
| `executeBatchWithRetry` | Controller.js:286-325 | Delete |
| `testImport` | Controller.js:196-219 | Move to test file or delete |
| Commented brightness/contrast | SlideBuilders.js:981-996 | Delete |
| Self-referential LIST_PRESETS | Config.js:107-120 | Consolidate |

### 2.4 Split Long Functions

**Target**: `buildTextContentRequests` (305 lines)

Split into:
- `buildTextInsertRequest()` - Text insertion
- `buildTextStyleRequests()` - Per-run styling
- `buildParagraphStyleRequests()` - Bullet/indent handling
- `buildTrailingNewlineStyle()` - Default newline fix

---

## Phase 3: Error Handling Standardization (Week 2-3)

### 3.1 Centralized Error Response Format

```javascript
// Server/ErrorHandler.js
const ErrorHandler = {
  // User-safe error messages
  USER_MESSAGES: {
    QUOTA_EXCEEDED: 'API limit reached. Please wait a moment and try again.',
    INVALID_JSON: 'The JSON data is invalid. Please check the format.',
    PRESENTATION_NOT_FOUND: 'Could not access the presentation. Check the ID and permissions.',
    UNKNOWN: 'An unexpected error occurred. Please try again.'
  },

  toResponse(error, context) {
    const userMessage = this.classifyError(error);
    Logger.log('[ERROR] ' + context + ': ' + error.message + '\n' + error.stack);

    return {
      status: 'error',
      message: userMessage,
      code: error.code || 'UNKNOWN',
      // Never expose stack traces to client
    };
  },

  classifyError(error) {
    if (error.message.includes('quota')) return this.USER_MESSAGES.QUOTA_EXCEEDED;
    if (error.message.includes('JSON')) return this.USER_MESSAGES.INVALID_JSON;
    // ... more classifications
    return this.USER_MESSAGES.UNKNOWN;
  }
};
```

### 3.2 Unified Logging Configuration

```javascript
// Config.js
LOGGING: {
  LEVEL: 'INFO',  // DEBUG, INFO, WARN, ERROR
  INCLUDE_TIMESTAMP: true,
  PREFIX_FORMAT: '[{level}:{module}]'
}

// Replace all Logger.log calls with:
function log(level, module, message) {
  if (shouldLog(level)) {
    Logger.log(formatLogMessage(level, module, message));
  }
}
```

---

## Phase 4: Input Validation Hardening (Week 3)

### 4.1 Schema Validation

Add element type validation in `Validation.js`:

```javascript
const ELEMENT_SCHEMAS = {
  text: { required: ['x', 'y'], optional: ['text', 'textRuns', 'fontSize'] },
  shape: { required: ['x', 'y', 'shape'], optional: ['text', 'fillColor'] },
  image: { required: ['x', 'y', 'url'], optional: ['crop', 'border'] },
  // ... etc
};

function validateElementSchema(element) {
  const schema = ELEMENT_SCHEMAS[element.type];
  if (!schema) {
    throw new Error('Unknown element type: ' + element.type);
  }

  for (const field of schema.required) {
    if (element[field] === undefined) {
      throw new Error(element.type + ' requires field: ' + field);
    }
  }
}
```

### 4.2 URL Validation for Images

```javascript
function validateImageUrl(url) {
  if (!url || typeof url !== 'string') {
    return { valid: false, reason: 'URL is required' };
  }

  // Must start with http:// or https://
  if (!url.match(/^https?:\/\//i)) {
    return { valid: false, reason: 'URL must use http or https protocol' };
  }

  // Check for common image extensions or Google domains
  const isGoogleHosted = url.includes('googleusercontent.com') ||
                          url.includes('drive.google.com');
  const hasImageExt = url.match(/\.(png|jpg|jpeg|gif|webp|svg)(\?|$)/i);

  if (!isGoogleHosted && !hasImageExt) {
    return { valid: true, warning: 'URL may not be an image' };
  }

  return { valid: true };
}
```

### 4.3 Color Validation

```javascript
function validateColor(color) {
  if (!color) return null;
  if (color === 'transparent' || color === 'none') return color;
  if (color.startsWith('theme:')) return color;

  // Validate hex format
  const hexMatch = color.match(/^#?([0-9a-f]{3}|[0-9a-f]{6})$/i);
  if (!hexMatch) {
    Logger.log('Invalid color format: ' + color + ', using default');
    return CONFIG.DEFAULTS.FILL_COLOR;
  }

  return color.startsWith('#') ? color : '#' + color;
}
```

---

## Phase 5: Testing Infrastructure (Week 3-4)

### 5.1 Test Framework Setup

Create `Server/Tests.js`:

```javascript
/**
 * Test runner for JsonPrez
 * Run via Apps Script editor: testAll()
 */

function testAll() {
  const results = [];

  results.push(testColorHelper());
  results.push(testValidation());
  results.push(testTransformMath());
  results.push(testTextBuilding());

  const passed = results.filter(r => r.passed).length;
  const failed = results.filter(r => !r.passed).length;

  Logger.log('=== TEST RESULTS ===');
  Logger.log('Passed: ' + passed + '/' + results.length);
  Logger.log('Failed: ' + failed);

  results.filter(r => !r.passed).forEach(r => {
    Logger.log('FAILED: ' + r.name + ' - ' + r.error);
  });

  return { passed, failed, results };
}

function testColorHelper() {
  try {
    const rgb = hexToRgbApi('#FF0000');
    assert(rgb.red === 1, 'Red should be 1');
    assert(rgb.green === 0, 'Green should be 0');
    assert(rgb.blue === 0, 'Blue should be 0');
    return { name: 'ColorHelper', passed: true };
  } catch (e) {
    return { name: 'ColorHelper', passed: false, error: e.message };
  }
}

function assert(condition, message) {
  if (!condition) throw new Error('Assertion failed: ' + message);
}
```

### 5.2 Integration Test Cases

Create test presentations covering:

| Test Case | JSON File | Validates |
|-----------|-----------|-----------|
| Basic shapes | test_shapes.json | All shape types render correctly |
| Text formatting | test_text.json | Fonts, sizes, colors, bullets |
| Images | test_images.json | External URLs, Google URLs, cropping |
| Tables | test_tables.json | Cell formatting, borders, alignment |
| Groups | test_groups.json | Nesting, transforms, copyGroup |
| Round-trip | test_roundtrip.json | Extract → Generate → Compare |

### 5.3 Automated Validation Script

```javascript
function validateRoundTrip(presentationId) {
  // 1. Extract original
  const originalJson = importPresentation(presentationId, false);

  // 2. Generate new presentation
  const newPresId = generatePresentation(originalJson);

  // 3. Extract generated
  const generatedJson = importPresentation(newPresId, false);

  // 4. Compare key metrics
  const comparison = {
    slideCount: originalJson.slides.length === generatedJson.slides.length,
    elementCounts: compareElementCounts(originalJson, generatedJson),
    colorAccuracy: compareColors(originalJson, generatedJson),
    textAccuracy: compareText(originalJson, generatedJson)
  };

  return comparison;
}
```

---

## Phase 6: Documentation & Deployment (Week 4)

### 6.1 API Documentation

Create `API.md` with:
- JSON schema reference with all element types
- Required vs optional fields per element
- Example payloads for each element type
- Error codes and meanings
- Rate limits and quotas

### 6.2 Deployment Checklist

```markdown
## Pre-Deployment Checklist

### Code Quality
- [ ] All tests pass (`testAll()`)
- [ ] No console errors in Dev UI
- [ ] LESSONS_LEARNED.md is current

### Manual Testing
- [ ] testpresentation.json generates correctly
- [ ] test_comprehensive.json generates correctly
- [ ] Round-trip test passes for sample presentation
- [ ] Dev UI import/export workflow works

### Git
- [ ] All changes committed
- [ ] Commit message follows convention
- [ ] No sensitive data in commits

### Deployment
- [ ] `clasp push` succeeds
- [ ] Web app accessible at deployed URL
- [ ] Test with fresh browser session
- [ ] Verify OAuth scopes are minimal needed
```

### 6.3 Version Bump

Update version references:
- Controller.js UI version string → "v8.0"
- Create git tag: `v8.0.0-production`
- Update enhancementbrief.md status

---

## Feature Gaps to Address (Post-Production)

### Not Blocking Production

| Feature | Current Status | Priority |
|---------|----------------|----------|
| Gradient fills | Unsupported | Low |
| Shape adjustments (corner radius) | API limitation | N/A |
| Animations | API limitation | N/A |
| Merged table cells | Unsupported | Medium |
| Image brightness/contrast | Unsupported | Low |

### Planned Enhancements (from enhancementbrief.md)

1. **Synchronous Batch Grouping** - Eliminate Phase 2 polling for groups (partially done)
2. **Deep Style Inheritance** - Resolve Master→Layout→Element style chain
3. **Client Preview Padding** - Match slide padding in preview CSS

---

## Risk Mitigation

### High-Risk Areas

| Risk | Mitigation |
|------|------------|
| API quota exhaustion | Exponential backoff + chunked batches for large decks |
| Google-hosted image failures | Retry with backoff + fallback to placeholder |
| CopyGroup position errors | Bounds validation + skip off-slide groups |
| Theme color drift | Per-slide color resolution with SlidesApp |

### Rollback Plan

1. Keep previous clasp deployment as backup
2. Document current Script ID and deployment version
3. If critical issues found: `clasp versions` → `clasp deploy -V <previous>`

---

## Success Criteria

### Production Ready When:

1. **Zero Critical Bugs**: No silent failures, proper error messages
2. **Test Coverage**: All test JSON files generate without errors
3. **Round-Trip Fidelity**: Extract→Generate produces visually similar output
4. **Performance**: 50-slide presentation generates in <60 seconds
5. **Documentation**: LESSONS_LEARNED.md and API.md are complete

---

## Timeline Summary

| Week | Phase | Deliverables |
|------|-------|--------------|
| 1 | Critical Fixes | Error handling, quota limits, bounds checking |
| 2 | Code Quality | Duplication removal, magic numbers, dead code |
| 2-3 | Error Standardization | Unified logging, user-friendly messages |
| 3 | Input Validation | Schema validation, URL/color validation |
| 3-4 | Testing | Test framework, integration tests |
| 4 | Documentation & Deploy | API docs, deployment, version bump |

---

*Plan created: January 15, 2026*
*Target production release: v8.0*
