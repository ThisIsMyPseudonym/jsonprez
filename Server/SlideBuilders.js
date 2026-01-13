/**
 * @fileoverview Builders for Google Slides API requests.
 */

// ============================================================================
// LOGGING
// ============================================================================

const BUILDER_VERBOSE_LOGGING = true;

function builderLog(message, level) {
  level = level || 'INFO';
  if (BUILDER_VERBOSE_LOGGING || level === 'ERROR' || level === 'WARN') {
    Logger.log('[BUILDER:' + level + '] ' + message);
  }
}

// ============================================================================
// HELPER BUILDERS
// ============================================================================

/**
 * Build requests for a fake shadow (semi-transparent shape behind main element)
 * @param {Object} element
 * @param {string} slideId
 * @param {string} shapeType
 * @returns {Array}
 */
function buildFakeShadowRequests(element, slideId, shapeType) {
  if (!element.shadow) return [];

  const requests = [];

  // Get shadow config
  let shadowConfig;
  if (element.shadow === true) {
    shadowConfig = { ...CONFIG.SHADOW_PRESETS.medium };
  } else if (typeof element.shadow === 'string') {
    shadowConfig = { ...(CONFIG.SHADOW_PRESETS[element.shadow] || CONFIG.SHADOW_PRESETS.medium) };
  } else if (typeof element.shadow === 'object') {
    shadowConfig = { ...CONFIG.SHADOW_PRESETS.medium, ...element.shadow };
  } else {
    return [];
  }

  // Calculate offset
  let offsetX, offsetY;
  if (shadowConfig.angle !== undefined && shadowConfig.distance !== undefined) {
    const angleRad = (shadowConfig.angle * Math.PI) / 180;
    offsetX = shadowConfig.distance * Math.cos(angleRad);
    offsetY = shadowConfig.distance * Math.sin(angleRad);
  } else {
    offsetX = shadowConfig.offsetX !== undefined ? shadowConfig.offsetX : 5;
    offsetY = shadowConfig.offsetY !== undefined ? shadowConfig.offsetY : 5;
  }

  const spread = shadowConfig.spread !== undefined ? shadowConfig.spread : 4;
  const opacity = shadowConfig.opacity !== undefined ? shadowConfig.opacity : 0.15;

  let shadowColor = shadowConfig.color || '#000000';
  if (shadowColor === 'inherit') {
    shadowColor = element.fillColor || element.fill || element.background || '#3b82f6';
  }
  shadowColor = themeService.resolveThemeColor(shadowColor);
  const rgb = themeService.hexToRgbApi(normalizeColor(shadowColor));

  const x = (element.x || 0) * SCALE;
  const y = (element.y || 0) * SCALE;
  const w = (element.w || 100) * SCALE;
  const h = (element.h || 100) * SCALE;
  const scaledOffsetX = offsetX * SCALE;
  const scaledOffsetY = offsetY * SCALE;
  const scaledSpread = spread * SCALE;

  const shadowId = generateObjectId();
  const shadowW = w + scaledSpread;
  const shadowH = h + scaledSpread;
  const shadowX = x + scaledOffsetX - scaledSpread * 0.5;
  const shadowY = y + scaledOffsetY - scaledSpread * 0.5;

  // Create shape
  requests.push({
    createShape: {
      objectId: shadowId,
      shapeType: shapeType || 'RECTANGLE',
      elementProperties: {
        pageObjectId: slideId,
        size: {
          width: { magnitude: shadowW, unit: 'PT' },
          height: { magnitude: shadowH, unit: 'PT' }
        },
        transform: {
          scaleX: 1, scaleY: 1, translateX: shadowX, translateY: shadowY, unit: 'PT'
        }
      }
    }
  });

  // Style shadow
  if (rgb) {
    requests.push({
      updateShapeProperties: {
        objectId: shadowId,
        shapeProperties: {
          shapeBackgroundFill: {
            solidFill: { color: { rgbColor: rgb }, alpha: opacity }
          },
          outline: { propertyState: 'NOT_RENDERED' }
        },
        fields: 'shapeBackgroundFill,outline'
      }
    });
  }

  return requests;
}

// ============================================================================
// ELEMENT BUILDERS
// ============================================================================

/**
 * Shared helper to build text content requests (insert text, style runs, paragraphs).
 * Used by both buildTextRequests (TextBox) and buildShape (Shape with text).
 */
function buildTextContentRequests(element, shapeId) {
  const requests = [];
  let textContent = element.text || '';
  if (element.items && Array.isArray(element.items)) {
    textContent = element.items.map(item => {
      if (typeof item === 'object' && item.text) {
        const indent = '\t'.repeat(item.indent || 0);
        return indent + item.text;
      }
      return String(item);
    }).join('\n');
  }

  // Clean text but preserve formatting chars
  textContent = textContent.replace(/[\x00-\x09\x0B\x0C\x0E-\x1F\x7F\u200B-\u200D\uFEFF\uFFFC\uFFFD]/g, '');

  const hasTextContent = textContent && textContent.length > 0;

  if (!hasTextContent) {
    return requests;
  }

  // Check if we have per-run styling available (mixed formatting)
  if (element.textRuns && element.textRuns.length > 1) {
    const runsText = element.textRuns.map(r => r.text).join('');
    requests.push({
      insertText: { objectId: shapeId, text: runsText, insertionIndex: 0 }
    });
    // Clean slate: Remove any default bullets from the theme/shape
    // Use FIXED_RANGE covering the whole text to be safer than 'ALL'
    requests.push({
      deleteParagraphBullets: {
        objectId: shapeId,
        textRange: {
          type: 'FIXED_RANGE',
          startIndex: 0,
          endIndex: runsText.length
        }
      }
    });

    // CRITICAL: Collect bullet requests separately to ensure proper ordering
    // deleteParagraphBullets MUST come AFTER createParagraphBullets to prevent list bleed
    const bulletCreateRequests = [];
    const bulletDeleteRequests = [];
    const paragraphStyleRequests = [];

    let currentIndex = 0;
    for (let i = 0; i < element.textRuns.length; i++) {
      const run = element.textRuns[i];
      const runLength = run.text.length;

      if (runLength === 0) continue;

      const textStyle = {};
      const fields = [];

      // Font Size (No SCALE for PT values coming from AdvancedExtractor)
      const fontSize = (run.fontSize || element.fontSize || CONFIG.DEFAULTS.FONT_SIZE);
      textStyle.fontSize = { magnitude: fontSize, unit: 'PT' };
      fields.push('fontSize');

      textStyle.fontFamily = themeService.resolveThemeFont(run.fontFamily || element.fontFamily);
      fields.push('fontFamily');

      const textColor = themeService.resolveThemeColor(run.color || element.color || CONFIG.DEFAULTS.TEXT_COLOR);
      const rgb = themeService.hexToRgbApi(textColor);
      if (rgb) {
        textStyle.foregroundColor = { opaqueColor: { rgbColor: rgb } };
        fields.push('foregroundColor');
      }

      if (run.bold !== undefined) { textStyle.bold = run.bold; fields.push('bold'); }
      else if (element.bold) { textStyle.bold = true; fields.push('bold'); }

      if (run.italic !== undefined) { textStyle.italic = run.italic; fields.push('italic'); }
      if (run.underline !== undefined) { textStyle.underline = run.underline; fields.push('underline'); }
      if (run.strikethrough !== undefined) { textStyle.strikethrough = run.strikethrough; fields.push('strikethrough'); }
      if (run.smallCaps !== undefined) { textStyle.smallCaps = run.smallCaps; fields.push('smallCaps'); }

      if (run.baselineOffset && run.baselineOffset !== 'NONE') {
        textStyle.baselineOffset = run.baselineOffset;
        fields.push('baselineOffset');
      }

      if (run.link && run.link.url) {
        const url = String(run.link.url).trim();
        if (url.length > 0) {
          const safeUrl = (url.indexOf('://') === -1 && !url.startsWith('mailto:')) ? 'https://' + url : url;
          if (safeUrl.indexOf('.') !== -1 || safeUrl.startsWith('mailto:')) {
            textStyle.link = { url: safeUrl };
            fields.push('link');
          }
        }
      }

      if (fields.length > 0) {
        requests.push({
          updateTextStyle: {
            objectId: shapeId,
            style: textStyle,
            textRange: {
              type: 'FIXED_RANGE',
              startIndex: currentIndex,
              endIndex: currentIndex + runLength
            },
            fields: fields.join(',')
          }
        });
      }

      const isParagraphStart = (i === 0) || (element.textRuns[i - 1].text.endsWith('\n'));

      // DEBUG: Trace logic for crucial paragraph formatting
      Logger.log('[BUILDER:TRACE] Run ' + i + ': ' + runLength + ' chars. Text: "' +
        (run.text.substring(0, 10).replace(/\n/g, '\\n')) + '...". IsParaStart: ' + isParagraphStart +
        '. HasBullet: ' + !!(run.bullet && (run.bullet.listId || run.bullet.glyph)));

      if (isParagraphStart) {
        // STRICT CHECK: Ensure bullet object is valid and not just empty
        if (run.bullet && (run.bullet.listId || run.bullet.glyph)) {
          bulletCreateRequests.push({
            createParagraphBullets: {
              objectId: shapeId,
              bulletPreset: 'BULLET_DISC_CIRCLE_SQUARE',
              textRange: {
                type: 'FIXED_RANGE',
                startIndex: currentIndex,
                endIndex: currentIndex + runLength
              }
            }
          });
        } else {
          bulletDeleteRequests.push({
            deleteParagraphBullets: {
              objectId: shapeId,
              textRange: {
                type: 'FIXED_RANGE',
                startIndex: currentIndex,
                endIndex: currentIndex + runLength
              }
            }
          });
        }

        if (run.paragraphStyle) {
          const ps = run.paragraphStyle;
          const pStyle = {};
          const pFields = [];

          if (ps.align) {
            const alignMap = { 'left': 'START', 'center': 'CENTER', 'right': 'END', 'justify': 'JUSTIFIED' };
            pStyle.alignment = alignMap[ps.align] || 'START';
            pFields.push('alignment');
          }
          if (ps.indentStart !== undefined) {
            pStyle.indentStart = { magnitude: ps.indentStart, unit: 'PT' };
            pFields.push('indentStart');
          }
          if (ps.indentFirstLine !== undefined) {
            pStyle.indentFirstLine = { magnitude: ps.indentFirstLine, unit: 'PT' };
            pFields.push('indentFirstLine');
          }

          Logger.log('[BUILDER:TRACE] Collecting Indent: ' + JSON.stringify(pStyle) +
            ' to range ' + currentIndex + '-' + (currentIndex + runLength));

          if (ps.spaceAbove !== undefined) {
            pStyle.spaceAbove = { magnitude: ps.spaceAbove * SCALE, unit: 'PT' };
            pFields.push('spaceAbove');
          }
          if (ps.spaceBelow !== undefined) {
            pStyle.spaceBelow = { magnitude: ps.spaceBelow * SCALE, unit: 'PT' };
            pFields.push('spaceBelow');
          }
          if (ps.lineSpacing !== undefined) {
            pStyle.lineSpacing = ps.lineSpacing;
            pFields.push('lineSpacing');
          }

          if (pFields.length > 0) {
            // CRITICAL: Collect now, push later (after bullet handling)
            paragraphStyleRequests.push({
              updateParagraphStyle: {
                objectId: shapeId,
                style: pStyle,
                textRange: {
                  type: 'FIXED_RANGE',
                  startIndex: currentIndex,
                  endIndex: currentIndex + runLength
                },
                fields: pFields.join(',')
              }
            });
          }
        }
      }
      currentIndex += runLength;
    }

    // CRITICAL: Push requests in the correct order
    // 1. First, add all createParagraphBullets (these can cause list bleed)
    // 2. Then, add all deleteParagraphBullets (these clean up the bleed AND can clear indentation)
    // 3. Finally, add all updateParagraphStyle (to restore indentation after bullet removal)
    Logger.log('[BUILDER:TRACE] Request ordering: ' + bulletCreateRequests.length + ' bullet creates, ' +
      bulletDeleteRequests.length + ' bullet deletes, ' + paragraphStyleRequests.length + ' paragraph styles');
    requests.push(...bulletCreateRequests);
    requests.push(...bulletDeleteRequests);
    requests.push(...paragraphStyleRequests);

  } else {
    // Legacy/Simple Text Path
    requests.push({
      insertText: { objectId: shapeId, text: textContent, insertionIndex: 0 }
    });

    const textStyle = {};
    const fields = [];

    const fontSize = (element.fontSize || CONFIG.DEFAULTS.FONT_SIZE) * SCALE;
    textStyle.fontSize = { magnitude: fontSize, unit: 'PT' };
    fields.push('fontSize');

    textStyle.fontFamily = themeService.resolveThemeFont(element.fontFamily);
    fields.push('fontFamily');

    const textColor = themeService.resolveThemeColor(element.color || CONFIG.DEFAULTS.TEXT_COLOR);
    const rgb = themeService.hexToRgbApi(textColor);
    if (rgb) {
      textStyle.foregroundColor = { opaqueColor: { rgbColor: rgb } };
      fields.push('foregroundColor');
    }

    if (element.bold) { textStyle.bold = true; fields.push('bold'); }
    if (element.italic) { textStyle.italic = true; fields.push('italic'); }
    if (element.underline) { textStyle.underline = true; fields.push('underline'); }

    if (fields.length > 0) {
      requests.push({
        updateTextStyle: {
          objectId: shapeId,
          style: textStyle,
          textRange: { type: 'ALL' },
          fields: fields.join(',')
        }
      });
    }

    if (element.align) {
      requests.push({
        updateParagraphStyle: {
          objectId: shapeId,
          style: { alignment: ENUMS.ALIGNMENT_MAP[element.align] || 'START' },
          textRange: { type: 'ALL' },
          fields: 'alignment'
        }
      });
    }
  }

  // Link handling (global/simple)
  if (element.link && (!element.textRuns || element.textRuns.length <= 1)) {
    // ... (Simple link handling omitted for brevity, handled by caller or assumed less critical for now)
  }

  return requests;
}

/**
 * Build requests for TEXT element (New Wrapper)
 */
function buildTextRequests(element, slideId) {
  element = validateElement(element);
  const requests = [];
  const shapeId = element.objectId || generateObjectId();

  if (element.shadow && element.background) {
    requests.push(...buildFakeShadowRequests(element, slideId, 'RECTANGLE'));
  }

  requests.push({
    createShape: {
      objectId: shapeId,
      shapeType: 'TEXT_BOX',
      elementProperties: {
        pageObjectId: slideId,
        size: buildSize(element.w || 200, element.h || 50),
        transform: buildTransform(element.x || 0, element.y || 0, element.rotation, element.w || 200, element.h || 50)
      }
    }
  });

  requests.push(...buildTextContentRequests(element, shapeId));

  return requests;
}


/**
 * Build requests for SHAPE element
 */
function buildShapeRequests(element, slideId) {
  element = validateElement(element);
  const requests = [];
  const shapeId = element.objectId || generateObjectId();
  const shapeType = ENUMS.SHAPE_TYPE_MAP[element.shape] || 'RECTANGLE';

  if (element.shadow) {
    requests.push(...buildFakeShadowRequests(element, slideId, shapeType));
  }

  requests.push({
    createShape: {
      objectId: shapeId,
      shapeType: shapeType,
      elementProperties: {
        pageObjectId: slideId,
        size: buildSize(element.w || 100, element.h || 100),
        transform: buildTransform(element.x || 0, element.y || 0, element.rotation, element.w || 100, element.h || 100)
      }
    }
  });

  const shapeProperties = {};
  const fields = [];

  const originalFillColor = element._originalFillColor || element.fillColor;
  const isTransparent = originalFillColor === 'transparent' || originalFillColor === 'none';

  if (isTransparent) {
    shapeProperties.shapeBackgroundFill = { propertyState: 'NOT_RENDERED' };
    fields.push('shapeBackgroundFill');
    builderLog('  Shape ' + shapeId + ': Fill is transparent/none');
  } else if (element.fillColor) {
    const fillColor = themeService.resolveThemeColor(element.fillColor);
    const rgb = themeService.hexToRgbApi(fillColor);
    builderLog('  Shape ' + shapeId + ': Fill ' + element.fillColor + ' -> ' + fillColor + ' -> ' + (rgb ? 'RGB OK' : 'RGB FAIL'));

    if (rgb) {
      shapeProperties.shapeBackgroundFill = {
        solidFill: {
          color: { rgbColor: rgb },
          alpha: element.alpha !== undefined ? element.alpha : (element.fillAlpha !== undefined ? element.fillAlpha : 1)
        }
      };
      fields.push('shapeBackgroundFill');
    }
  }

  // Border
  if (element.borderColor && element.borderColor !== 'none') {
    const borderColor = themeService.resolveThemeColor(element.borderColor);
    const borderRgb = themeService.hexToRgbApi(borderColor);
    if (borderRgb) {
      const outline = {
        outlineFill: {
          solidFill: { color: { rgbColor: borderRgb }, alpha: element.borderAlpha !== undefined ? element.borderAlpha : 1 }
        },
        weight: { magnitude: (element.borderWidth || CONFIG.DEFAULTS.LINE_WEIGHT) * SCALE, unit: 'PT' },
        propertyState: 'RENDERED'
      };
      if (element.borderDash) {
        outline.dashStyle = ENUMS.DASH_STYLE_MAP[element.borderDash] || 'SOLID';
      }
      shapeProperties.outline = outline;
      fields.push('outline');
    }
  } else {
    shapeProperties.outline = { propertyState: 'NOT_RENDERED' };
    fields.push('outline');
  }

  if (fields.length > 0) {
    requests.push({
      updateShapeProperties: {
        objectId: shapeId,
        shapeProperties: shapeProperties,
        fields: fields.join(',')
      }
    });
  }

  requests.push(...buildTextContentRequests(element, shapeId));

  if (element.link) {
    const linkRequest = buildLinkRequest(shapeId, element.link, 'shape');
    if (linkRequest) requests.push(linkRequest);
  }

  return requests;
}

/**
 * Build requests for ICON element
 */
function buildIconRequests(element, slideId) {
  element = validateElement(element);
  const requests = [];
  const iconId = element.objectId || generateObjectId();
  const shapeType = 'ELLIPSE';

  if (element.shadow) {
    requests.push(...buildFakeShadowRequests(element, slideId, shapeType));
  }

  const width = element.w || element.size || 50;
  const height = element.h || element.size || 50;

  requests.push({
    createShape: {
      objectId: iconId,
      shapeType: shapeType,
      elementProperties: {
        pageObjectId: slideId,
        size: buildSize(width, height),
        transform: buildTransform(element.x || 0, element.y || 0, element.rotation, width, height)
      }
    }
  });

  const color = themeService.resolveThemeColor(element.color || CONFIG.DEFAULTS.ICON_COLOR);
  const rgb = themeService.hexToRgbApi(color);
  const bgOpacity = element.bgOpacity !== undefined ? element.bgOpacity : CONFIG.DEFAULTS.ICON_BG_OPACITY;

  if (rgb) {
    requests.push({
      updateShapeProperties: {
        objectId: iconId,
        shapeProperties: {
          shapeBackgroundFill: { solidFill: { color: { rgbColor: rgb } }, alpha: bgOpacity },
          outline: { propertyState: 'NOT_RENDERED' }
        },
        fields: 'shapeBackgroundFill,outline'
      }
    });
  }

  const iconText = element.text || element.icon || '★';

  requests.push({
    insertText: { objectId: iconId, text: iconText, insertionIndex: 0 }
  });

  requests.push({
    updateTextStyle: {
      objectId: iconId,
      style: {
        fontSize: { magnitude: (element.fontSize || height * 0.6) * SCALE, unit: 'PT' },
        foregroundColor: { opaqueColor: { rgbColor: rgb } },
        bold: true
      },
      fields: 'fontSize,foregroundColor,bold'
    }
  });

  requests.push({
    updateParagraphStyle: {
      objectId: iconId,
      style: { alignment: 'CENTER' },
      fields: 'alignment'
    }
  });

  if (element.link) {
    const linkRequest = buildLinkRequest(iconId, element.link, 'shape');
    if (linkRequest) requests.push(linkRequest);
  }

  return requests;
}

/**
 * Build requests for VIDEO element
 */
function buildVideoRequests(element, slideId) {
  const requests = [];
  const videoId = element.objectId || generateObjectId();
  const source = element.source || 'YOUTUBE';
  const id = element.videoId || element.id;

  if (!id) {
    return buildShapeRequests({
      ...element,
      type: 'shape',
      shape: 'RECTANGLE',
      text: 'VIDEO (Missing ID)',
      fillColor: '#000000',
      color: '#ffffff'
    }, slideId);
  }

  requests.push({
    createVideo: {
      objectId: videoId,
      source: source,
      id: id,
      elementProperties: {
        pageObjectId: slideId,
        size: buildSize(element.w || 300, element.h || 200),
        transform: buildTransform(element.x || 0, element.y || 0, element.rotation, element.w || 300, element.h || 200)
      }
    }
  });

  if (element.borderColor) {
    const c = themeService.resolveThemeColor(element.borderColor);
    const rgb = themeService.hexToRgbApi(c);
    if (rgb) {
      requests.push({
        updateVideoProperties: {
          objectId: videoId,
          outline: {
            outlineFill: { solidFill: { color: { rgbColor: rgb } } },
            weight: { magnitude: (element.borderWidth || 1) * SCALE, unit: 'PT' },
            propertyState: 'RENDERED'
          },
          fields: 'outline'
        }
      });
    }
  }

  return requests;
}

/**
 * Build requests for IMAGE element
 */
function buildImageRequests(element, slideId) {
  element = validateElement(element);
  const requests = [];
  const imageId = element.objectId || generateObjectId();

  if (element.shadow) {
    requests.push(...buildFakeShadowRequests(element, slideId, 'RECTANGLE'));
  }

  requests.push({
    createImage: {
      objectId: imageId,
      url: element.url || 'https://via.placeholder.com/800x600?text=Image+Not+Found',
      elementProperties: {
        pageObjectId: slideId,
        size: buildSize(element.w || 200, element.h || 200),
        transform: buildTransform(element.x || 0, element.y || 0, element.rotation, element.w || 200, element.h || 200)
      }
    }
  });

  if (element.borderColor || element.borderWidth) {
    const outline = {
      weight: { magnitude: (element.borderWidth || 1) * SCALE, unit: 'PT' },
      propertyState: 'RENDERED'
    };
    if (element.borderColor) {
      const c = themeService.resolveThemeColor(element.borderColor);
      const rgb = themeService.hexToRgbApi(c);
      if (rgb) outline.outlineFill = { solidFill: { color: { rgbColor: rgb } } };
    }

    requests.push({
      updateImageProperties: {
        objectId: imageId,
        outline: outline,
        fields: 'outline'
      }
    });
  }

  if (element.link) {
    const linkRequest = buildLinkRequest(imageId, element.link, 'image');
    if (linkRequest) requests.push(linkRequest);
  }

  return requests;
}

/**
 * Build requests for TABLE element
 */
function buildTableRequests(element, slideId) {
  const requests = [];
  const tableId = element.objectId || generateObjectId();
  const data = element._data || element.data;

  if (!data || !Array.isArray(data) || data.length === 0) return [];

  const rows = data.length;
  const cols = data[0].length;

  requests.push({
    createTable: {
      objectId: tableId,
      rows: rows,
      columns: cols,
      elementProperties: {
        pageObjectId: slideId,
        transform: buildTransform(element.x || 0, element.y || 0, 0, element.w || 400, element.h || 300)
      }
    }
  });

  data.forEach((row, r) => {
    row.forEach((cellValue, c) => {
      let text = '';
      let style = {};
      let cellStyle = {};

      if (cellValue === null || cellValue === undefined) {
        text = '';
      } else if (typeof cellValue === 'object') {
        text = cellValue.text || '';
        style = {
          bold: cellValue.bold,
          italic: cellValue.italic,
          fontSize: cellValue.fontSize ? { magnitude: cellValue.fontSize, unit: 'PT' } : undefined,
          fontFamily: cellValue.fontFamily,
          backgroundColor: cellValue.backgroundColor
        };
        if (cellValue.color) {
          const hex = themeService.resolveThemeColor(cellValue.color);
          const rgb = themeService.hexToRgbApi(hex);
          if (rgb) style.foregroundColor = { opaqueColor: { rgbColor: rgb } };
        }
        if (cellValue.fillColor) {
          const hex = themeService.resolveThemeColor(cellValue.fillColor);
          const rgb = themeService.hexToRgbApi(hex);
          if (rgb && hex !== 'transparent') cellStyle.tableCellBackgroundFill = { solidFill: { color: { rgbColor: rgb } } };
        }
        if (cellValue.align) {
          cellStyle.contentAlignment = (cellValue.align === 'center') ? 'MIDDLE' : (cellValue.align === 'bottom' ? 'BOTTOM' : 'TOP');
        }
      } else {
        text = String(cellValue);
      }

      if (typeof cellValue === 'object' && cellValue.textRuns && cellValue.textRuns.length > 1) {
        const joinedText = cellValue.textRuns.map(r => r.text).join('');
        if (joinedText.length > 0) {
          text = joinedText;
        }
      }

      if (text && /\S/.test(text)) {
        requests.push({
          insertText: {
            objectId: tableId,
            cellLocation: { rowIndex: r, columnIndex: c },
            text: text,
            insertionIndex: 0
          }
        });

        if (cellValue.textRuns && cellValue.textRuns.length > 1) {
          let currentIndex = 0;
          cellValue.textRuns.forEach(run => {
            const runText = run.text;
            if (!runText || runText.length === 0) return;

            const runLength = runText.length;
            const runStyle = {
              fontSize: run.fontSize ? { magnitude: run.fontSize, unit: 'PT' } : undefined,
              fontFamily: run.fontFamily,
              bold: run.bold,
              italic: run.italic,
              underline: run.underline,
              strikethrough: run.strikethrough
            };

            if (run.color) {
              const hex = themeService.resolveThemeColor(run.color);
              const rgb = themeService.hexToRgbApi(hex);
              if (rgb) runStyle.foregroundColor = { opaqueColor: { rgbColor: rgb } };
            }

            const fields = Object.keys(runStyle).filter(k => runStyle[k] !== undefined).join(',');

            if (fields.length > 0) {
              requests.push({
                updateTextStyle: {
                  objectId: tableId,
                  cellLocation: { rowIndex: r, columnIndex: c },
                  textRange: {
                    type: 'FIXED_RANGE',
                    startIndex: currentIndex,
                    endIndex: currentIndex + runLength
                  },
                  style: runStyle,
                  fields: fields
                }
              });
            }
            currentIndex += runLength;
          });

        } else if (Object.keys(style).length > 0) {
          const fields = Object.keys(style).map(k => k).join(',');
          if (fields.length > 0) {
            requests.push({
              updateTextStyle: {
                objectId: tableId,
                cellLocation: { rowIndex: r, columnIndex: c },
                style: style,
                fields: fields
              }
            });
          }
        }

        if (typeof cellValue === 'object' && cellValue.align) {
          let alignType = 'START';
          if (cellValue.align === 'center') alignType = 'CENTER';
          else if (cellValue.align === 'right') alignType = 'END';
          else if (cellValue.align === 'justify') alignType = 'JUSTIFIED';

          requests.push({
            updateParagraphStyle: {
              objectId: tableId,
              cellLocation: { rowIndex: r, columnIndex: c },
              style: { alignment: alignType },
              fields: 'alignment'
            }
          });
        }
      }

      if (Object.keys(cellStyle).length > 0) {
        const fields = Object.keys(cellStyle).map(k => k).join(',');
        requests.push({
          updateTableCellProperties: {
            objectId: tableId,
            tableRange: { location: { rowIndex: r, columnIndex: c }, rowSpan: 1, columnSpan: 1 },
            tableCellProperties: cellStyle,
            fields: fields
          }
        });
      }

      if (r === 0 && element.header && typeof cellValue !== 'object') {
        requests.push({
          updateTextStyle: {
            objectId: tableId,
            cellLocation: { rowIndex: r, columnIndex: c },
            style: { bold: true },
            fields: 'bold'
          }
        });

        const headerBg = themeService.resolveThemeColor(element.headerBg || CONFIG.DEFAULTS.TABLE_HEADER_BG || '#f1f5f9');
        const rgb = themeService.hexToRgbApi(headerBg);
        if (rgb) {
          requests.push({
            updateTableCellProperties: {
              objectId: tableId,
              tableRange: { location: { rowIndex: r, columnIndex: c }, rowSpan: 1, columnSpan: 1 },
              tableCellProperties: {
                tableCellBackgroundFill: { solidFill: { color: { rgbColor: rgb } } }
              },
              fields: 'tableCellBackgroundFill'
            }
          });
        }
      }
    });
  });

  return requests;
}

// Helper for elbow connectors
function buildElbowAsSegments(element, slideId) {
  const requests = [];
  const x1 = element.x1 || 0;
  const y1 = element.y1 || 0;
  const x2 = element.x2 || 100;
  const y2 = element.y2 || 100;
  const bendDirection = element.bendDirection || 'horizontal-first';

  let midX, midY;
  if (bendDirection === 'vertical-first') {
    midX = x1;
    midY = y2;
  } else {
    midX = x2;
    midY = y1;
  }

  const segment1 = {
    ...element,
    x1, y1, x2: midX, y2: midY,
    connector: null,
    startArrow: element.startArrow,
    endArrow: 'NONE',
    startConnect: element.startConnect,
    endConnect: null
  };

  const segment2 = {
    ...element,
    x1: midX, y1: midY, x2, y2,
    connector: null,
    startArrow: 'NONE',
    endArrow: element.endArrow,
    startConnect: null,
    endConnect: element.endConnect
  };

  const seg1HasLength = (x1 !== midX || y1 !== midY);
  const seg2HasLength = (midX !== x2 || midY !== y2);

  if (seg1HasLength) requests.push(...buildLineRequests(segment1, slideId).requests);
  if (seg2HasLength) requests.push(...buildLineRequests(segment2, slideId).requests);
  if (!seg1HasLength && !seg2HasLength) requests.push(...buildLineRequests(element, slideId).requests);

  return requests;
}

/**
 * Build requests for LINE element
 */
function buildLineRequests(element, slideId) {
  const requests = [];
  const deferredConnections = [];
  const lineId = element.objectId || generateObjectId();

  if (element.connector === 'elbow' || element.connector === 'bent') {
    const segments = buildElbowAsSegments(element, slideId);
    return { requests: segments, deferredConnections: [], objectId: lineId };
  }

  const x = parseFloat(element.x || 0);
  const y = parseFloat(element.y || 0);
  const w = parseFloat(element.w !== undefined ? element.w : 100);
  const h = parseFloat(element.h !== undefined ? element.h : 0);
  const ex1 = parseFloat(element.x1);
  const ey1 = parseFloat(element.y1);
  const ex2 = parseFloat(element.x2);
  const ey2 = parseFloat(element.y2);

  const x1 = (!isNaN(ex1) ? ex1 : x) * SCALE;
  const y1 = (!isNaN(ey1) ? ey1 : y) * SCALE;
  const x2 = (!isNaN(ex2) ? ex2 : (x + w)) * SCALE;
  const y2 = (!isNaN(ey2) ? ey2 : (y + h)) * SCALE;

  let width = Math.abs(x2 - x1);
  let height = Math.abs(y2 - y1);

  if (isNaN(width)) width = 100;
  if (isNaN(height)) height = 0;

  if (width < 0.1 && height < 0.1) {
    width = 1;
    height = 0;
  }

  const safeW = width;
  const safeH = height;

  const tx = Math.min(x1, x2);
  const ty = Math.min(y1, y2);

  const isAntiDiagonal = (x2 > x1 && y2 < y1) || (x2 < x1 && y2 > y1);

  let finalScaleY = 1;
  let finalTranslateY = ty;

  if (isAntiDiagonal) {
    finalScaleY = -1;
    finalTranslateY = ty + safeH;
  }

  requests.push({
    createLine: {
      objectId: lineId,
      lineCategory: 'STRAIGHT',
      elementProperties: {
        pageObjectId: slideId,
        size: { width: { magnitude: safeW, unit: 'PT' }, height: { magnitude: safeH, unit: 'PT' } },
        transform: { scaleX: 1, scaleY: finalScaleY, translateX: tx, translateY: finalTranslateY, unit: 'PT' }
      }
    }
  });

  const style = {};
  const fields = [];

  style.weight = { magnitude: (element.weight || CONFIG.DEFAULTS.LINE_WEIGHT) * SCALE, unit: 'PT' };
  fields.push('weight');

  const color = themeService.resolveThemeColor(element.color || '#000000');
  const rgb = themeService.hexToRgbApi(color);
  if (rgb) {
    style.lineFill = { solidFill: { color: { rgbColor: rgb } } };
    fields.push('lineFill');
  }

  if (element.dashStyle) {
    style.dashStyle = ENUMS.DASH_STYLE_MAP[element.dashStyle] || 'SOLID';
    fields.push('dashStyle');
  }

  if (element.startArrow) {
    style.startArrow = ENUMS.ARROW_TYPE_MAP[element.startArrow] || 'NONE';
    fields.push('startArrow');
  }
  if (element.endArrow) {
    style.endArrow = ENUMS.ARROW_TYPE_MAP[element.endArrow] || 'NONE';
    fields.push('endArrow');
  }

  if (fields.length > 0) {
    requests.push({
      updateLineProperties: {
        objectId: lineId,
        lineProperties: style,
        fields: fields.join(',')
      }
    });
  }

  const hasStartConnect = element.startConnect && element.startConnect.objectId;
  const hasEndConnect = element.endConnect && element.endConnect.objectId;

  if (hasStartConnect || hasEndConnect) {
    const connectionProperties = {};
    const connectionFields = [];

    if (hasStartConnect) {
      connectionProperties.startConnection = {
        connectedObjectId: element.startConnect.objectId,
        connectionSiteIndex: element.startConnect.site || 0
      };
      connectionFields.push('startConnection');
    }

    if (hasEndConnect) {
      connectionProperties.endConnection = {
        connectedObjectId: element.endConnect.objectId,
        connectionSiteIndex: element.endConnect.site || 0
      };
      connectionFields.push('endConnection');
    }

    deferredConnections.push({
      updateLineProperties: {
        objectId: lineId,
        lineProperties: connectionProperties,
        fields: connectionFields.join(',')
      }
    });
  }

  return { requests, deferredConnections, objectId: lineId };
}

function buildChartRequests(element, slideId, slideIndex) {
  phase2Service.addChart(slideIndex, element);
  return { requests: [], spreadsheetIds: [], objectId: null, deferredConnections: [] };
}

/**
 * Build requests for LINKED SHEETS CHART
 */
function buildSheetsChartRequests(element, slideId) {
  const requests = [];
  const chartId = element.objectId || generateObjectId();

  if (!element.spreadsheetId || !element.chartId) {
    console.warn('Skipping sheetsChart due to missing spreadsheetId or chartId', element);
    return { requests: [], deferredConnections: [], objectId: null };
  }

  const x = (element.x || 0) * SCALE;
  const y = (element.y || 0) * SCALE;
  const width = (element.w || 400) * SCALE;
  const height = (element.h || 300) * SCALE;

  try {
    const ss = SpreadsheetApp.openById(element.spreadsheetId);
  } catch (e) {
    // If we have a contentUrl from extraction, use it as image fallback
    if (element.contentUrl) {
      Logger.log('Cannot access spreadsheet ' + element.spreadsheetId + '. Using contentUrl as image fallback.');

      requests.push({
        createImage: {
          objectId: chartId,
          url: element.contentUrl,
          elementProperties: {
            pageObjectId: slideId,
            size: { width: { magnitude: width, unit: 'PT' }, height: { magnitude: height, unit: 'PT' } },
            transform: { scaleX: 1, scaleY: 1, translateX: x, translateY: y, unit: 'PT' }
          }
        }
      });

      return { requests, deferredConnections: [], objectId: chartId };
    }

    // No contentUrl available - create placeholder
    Logger.log('Cannot access spreadsheet ' + element.spreadsheetId + ' and no contentUrl. Creating placeholder.');

    requests.push({
      createShape: {
        objectId: chartId,
        shapeType: 'RECTANGLE',
        elementProperties: {
          pageObjectId: slideId,
          size: { width: { magnitude: width, unit: 'PT' }, height: { magnitude: height, unit: 'PT' } },
          transform: { scaleX: 1, scaleY: 1, translateX: x, translateY: y, unit: 'PT' }
        }
      }
    });

    requests.push({
      updateShapeProperties: {
        objectId: chartId,
        shapeProperties: {
          shapeBackgroundFill: { solidFill: { color: { rgbColor: { red: 0.9, green: 0.9, blue: 0.9 } } } },
          outline: { outlineFill: { solidFill: { color: { rgbColor: { red: 0.6, green: 0.6, blue: 0.6 } } } }, weight: { magnitude: 1, unit: 'PT' } }
        },
        fields: 'shapeBackgroundFill,outline'
      }
    });

    requests.push({
      insertText: {
        objectId: chartId,
        text: 'Chart Unavailable\nNo access to source spreadsheet'
      }
    });

    requests.push({
      updateTextStyle: {
        objectId: chartId,
        style: {
          foregroundColor: { opaqueColor: { rgbColor: { red: 0.4, green: 0.4, blue: 0.4 } } },
          fontSize: { magnitude: 12, unit: 'PT' },
          bold: true
        },
        textRange: { type: 'ALL' },
        fields: 'foregroundColor,fontSize,bold'
      }
    });

    return { requests, deferredConnections: [], objectId: chartId };
  }

  requests.push({
    createSheetsChart: {
      objectId: chartId,
      spreadsheetId: element.spreadsheetId,
      chartId: element.chartId,
      linkingMode: (element.embedType === 'IMAGE') ? 'NOT_LINKED_IMAGE' : 'LINKED',
      elementProperties: {
        pageObjectId: slideId,
        size: { width: { magnitude: width, unit: 'PT' }, height: { magnitude: height, unit: 'PT' } },
        transform: { scaleX: 1, scaleY: 1, translateX: x, translateY: y, unit: 'PT' }
      }
    }
  });

  return { requests, deferredConnections: [], objectId: chartId };
}


/**
 * Main element request dispatcher
 */
function buildElementRequests(element, slideId, slideIndex, elementIndex) {
  try {
    let requests = [];
    let spreadsheetIds = [];
    let objectId = null;
    let deferredConnections = [];

    builderLog('Building ' + element.type + ' at (' + (element.x || 0) + ',' + (element.y || 0) + ') size ' + (element.w || 0) + 'x' + (element.h || 0));

    switch (element.type) {
      case 'text':
        requests = buildTextRequests(element, slideId);
        if (requests.length > 0 && requests[0].createShape) objectId = requests[0].createShape.objectId;
        break;
      case 'shape':
        requests = buildShapeRequests(element, slideId);
        if (requests.length > 0 && requests[0].createShape) objectId = requests[0].createShape.objectId;
        break;
      case 'chart':
        buildChartRequests(element, slideId, slideIndex);
        break;
      case 'image':
        requests = buildImageRequests(element, slideId);
        if (requests.length > 0 && requests[0].createImage) objectId = requests[0].createImage.objectId;
        break;
      case 'table':
        requests = buildTableRequests(element, slideId);
        if (requests.length > 0 && requests[0].createTable) objectId = requests[0].createTable.objectId;
        break;
      case 'line':
        const lineResult = buildLineRequests(element, slideId);
        requests = lineResult.requests;
        deferredConnections = lineResult.deferredConnections;
        objectId = lineResult.objectId;
        break;
      case 'icon':
        requests = buildIconRequests(element, slideId);
        if (requests.length > 0 && requests[0].createShape) objectId = requests[0].createShape.objectId;
        break;
      case 'video':
        requests = buildVideoRequests(element, slideId);
        if (requests.length > 0 && (requests[0].createVideo || requests[0].createShape)) objectId = (requests[0].createVideo || requests[0].createShape).objectId;
        break;
      case 'sheetsChart':
        const chartResult = buildSheetsChartRequests(element, slideId);
        requests = chartResult.requests;
        objectId = chartResult.objectId;
        break;
      case 'wordArt':
        requests = buildTextRequests({
          ...element,
          type: 'text',
          fontSize: element.fontSize || 48,
          bold: true,
          align: 'center'
        }, slideId);
        if (requests.length > 0 && requests[0].createShape) objectId = requests[0].createShape.objectId;
        break;
      case 'group':
        // ATOMIC/RECURSIVE GROUPING
        // 1. Build all children recursively (depth-first)
        const groupChildrenIds = [];
        const groupId = element.objectId || generateObjectId();

        if (element.elements && element.elements.length > 0) {
          element.elements.forEach((child, idx) => {
            const childResult = buildElementRequests(child, slideId, slideIndex, elementIndex + '_' + idx);
            requests.push(...childResult.requests); // Add child creation logic to batch
            deferredConnections.push(...(childResult.deferredConnections || []));

            if (childResult.objectId) {
              groupChildrenIds.push(childResult.objectId);
            }
          });
        }

        // 2. Add the createGroup request to the END of the batch (after children are created)
        if (groupChildrenIds.length >= 1) {
          // Note: Single-element groups are valid in API but weird. API requires at least 1 child?
          // Actually createGroup requires at least 1 child.
          requests.push({
            createGroup: {
              objectId: groupId,
              childrenObjectIds: groupChildrenIds
            }
          });
          objectId = groupId;
        }
        break;
      case 'unsupported':
        Logger.log('Skipping unsupported element: ' + (element.objectId || 'unknown'));
        break;
      default:
        Logger.log('Unknown or unimplemented element type: ' + element.type);
    }

    if (objectId) {
      phase2Service.recordElementId(slideIndex, elementIndex || 0, objectId);
      element._objectId = objectId;
    }

    return { requests, spreadsheetIds, objectId, deferredConnections };
  } catch (e) {
    Logger.log('Error building element requests: ' + e.message);
    return { requests: [], spreadsheetIds: [], objectId: null, deferredConnections: [] };
  }
}

/**
 * Build ALL requests for the entire presentation
 * @param {Object} json
 * @param {string} firstSlideId
 * @param {string} presentationId
 * @returns {Object} { requests, spreadsheetIds }
 */
function buildAllRequests(json, firstSlideId, presentationId) {
  builderLog('=== GENERATION START [DEBUG CANARY MARKER-FIRST-FIX] ===');
  builderLog('Total slides: ' + (json.slides ? json.slides.length : 0));
  builderLog('Presentation ID: ' + presentationId);

  const requests = [];
  const spreadsheetIds = [];
  const deferredConnections = [];

  if (!json.slides || json.slides.length === 0) {
    builderLog('No slides to generate', 'WARN');
    return { requests, spreadsheetIds };
  }

  json.slides.forEach((slide, slideIndex) => {
    const slideId = slideIndex === 0 ? firstSlideId : generateObjectId();
    builderLog('--- Processing Slide ' + (slideIndex + 1) + ' (ID: ' + slideId + ') ---');
    builderLog('  Background: ' + (slide.backgroundImage || slide.background || 'default'));
    builderLog('  Elements: ' + (slide.elements ? slide.elements.length : 0));

    if (slideIndex > 0) {
      requests.push({
        createSlide: {
          objectId: slideId,
          insertionIndex: slideIndex,
          slideLayoutReference: { predefinedLayout: 'BLANK' }
        }
      });
    }

    if (slide.backgroundImage) {
      requests.push({
        updatePageProperties: {
          objectId: slideId,
          pageProperties: {
            pageBackgroundFill: {
              stretchedPictureFill: { contentUrl: slide.backgroundImage }
            }
          },
          fields: 'pageBackgroundFill.stretchedPictureFill'
        }
      });
    } else if (slide.background) {
      const bgColor = themeService.resolveThemeColor(slide.background);
      const rgb = themeService.hexToRgbApi(bgColor);
      if (rgb) {
        requests.push({
          updatePageProperties: {
            objectId: slideId,
            pageProperties: {
              pageBackgroundFill: {
                solidFill: { color: { rgbColor: rgb } }
              }
            },
            fields: 'pageBackgroundFill.solidFill.color'
          }
        });
      }
    }

    if (slide.elements) {
      slide.elements.forEach((el, idx) => { el._originalIndex = idx; });
      slide.elements.sort((a, b) => (a.zIndex || 0) - (b.zIndex || 0));

      slide.elements.forEach((element, idx) => {
        const result = buildElementRequests(element, slideId, slideIndex, element._originalIndex);

        if (result.requests) requests.push(...result.requests);
        if (result.spreadsheetIds) spreadsheetIds.push(...result.spreadsheetIds);
        if (result.deferredConnections) deferredConnections.push(...result.deferredConnections);
      });
    }

    // Phase 2 groups loop removed - now handled recursively in Phase 1
    // Keep Phase 2 for notes only
    if (slide.speakerNotes || slide.notes) {
      phase2Service.addSpeakerNotes(slideIndex, slide.speakerNotes || slide.notes);
    }
  });

  builderLog('=== GENERATION COMPLETE ===');
  builderLog('Total API requests: ' + requests.length);
  builderLog('Deferred connections: ' + deferredConnections.length);
  builderLog('Spreadsheet IDs: ' + spreadsheetIds.length);

  return { requests, connectionRequests: deferredConnections, spreadsheetIds };
}
