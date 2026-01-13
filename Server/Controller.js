/**
 * @fileoverview Main Controller. Entry point for the web app and execution.
 */

// ============================================================================
// WEB APP ENTRY POINTS
// ============================================================================

/**
 * Serve the web app
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Client/Index');
  const html = template.evaluate();

  html.setTitle('Slides Engine v7.12 Orion')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  return html;
}

/**
 * Include helper for HTML templates
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================================
// SERVER SIDE API
// ============================================================================

/**
 * Generate Presentation from JSON
 * @param {string} jsonString
 * @returns {Object} result
 */
function generatePresentation(jsonString) {
  try {
    if (!jsonString) {
      throw new Error('No JSON data provided. If running manually, this is expected. Please use the Web App.');
    }
    const json = JSON.parse(jsonString);

    // 1. Validation
    validateJSON(json);

    // 2. Theme Setup
    themeService.setTheme(json.config ? json.config.theme : null);

    // 3. Create Presentation
    const title = (json.config && json.config.title) || 'New Presentation';
    const presentationId = slidesApi.createPresentation(title);

    // 4. Build Requests (Phase 1)
    // We need to get the first slide ID. Since we just created it, it has one slide.
    // We can get it via SlidesApp or just assume we insert others. 
    // To allow modifying the first slide, we need its ID.
    const presentation = SlidesApp.openById(presentationId);
    const firstSlide = presentation.getSlides()[0];
    const firstSlideId = firstSlide.getObjectId();

    // Clean up first slide artifacts (remove default placeholders)
    // We remove all default elements (Title/Subtitle placeholders) to make it clean.
    firstSlide.getPageElements().forEach(element => element.remove());

    // CRITICAL: Save and close to flush changes and release the file lock.
    // This allows the subsequent API batchUpdate to work on a fresh state
    // and prevents caching issues in Phase 2.
    presentation.saveAndClose();

    // Phase 2 queue reset
    phase2Service.reset();

    const buildResult = buildAllRequests(json, firstSlideId, presentationId);

    // 5. Execute Phase 1 (Batch Update)
    if (buildResult.requests.length > 0) {
      slidesApi.batchUpdate(presentationId, buildResult.requests);
    }

    // 5.5. Execute Phase 1.5 (Connections)
    // Must be done after elements are created but before Phase 2
    if (buildResult.connectionRequests && buildResult.connectionRequests.length > 0) {
      Logger.log('Executing Phase 1.5: ' + buildResult.connectionRequests.length + ' connections');
      slidesApi.batchUpdate(presentationId, buildResult.connectionRequests);
    }


    // 6. Execute Phase 2 (Charts, Shadows, etc via SlidesApp)
    // Pass slide count for robust synchronization
    slidesApi.executePhase2(presentationId, json.slides.length);

    return {
      status: 'success',
      presentationId: presentationId,
      url: 'https://docs.google.com/presentation/d/' + presentationId + '/edit',
      slideCount: json.slides.length
    };

  } catch (e) {
    Logger.log('ERROR: ' + e.message + '\n' + e.stack);
    return {
      status: 'error',
      message: e.message
    };
  }
}

/**
 * Get template JSONs for the UI
 */
function getTemplate(id) {
  // Return different templates based on ID
  // For now return a basic one
  return JSON.stringify({
    config: { title: "New Deck", theme: DEFAULT_THEME },
    slides: [{ background: "#ffffff", elements: [{ type: "text", text: "Hello World", x: 100, y: 100 }] }]
  }, null, 2);
}

/**
 * Import a presentation ID to JSON
 * @param {string} presentationId
 * @param {boolean} rawMode - If true, skip master/theme inheritance and copy exact styles
 */
function importPresentation(presentationId, rawMode) {
  try {
    const options = { rawMode: rawMode || false };
    const json = extractPresentationAdvanced(presentationId, options);
    return {
      status: 'success',
      json: JSON.stringify(json, null, 2)
    };
  } catch (e) {
    Logger.log('Import error: ' + e.message + '\n' + e.stack);
    return {
      status: 'error',
      message: e.message
    };
  }
}
