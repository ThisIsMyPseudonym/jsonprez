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
  if (e.parameter && e.parameter.page === 'dev') {
    return HtmlService.createHtmlOutputFromFile('Client/Dev')
      .setTitle('SlidesEngine Dev Runner')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const template = HtmlService.createTemplateFromFile('Client/Index');
  const html = template.evaluate();

  html.setTitle('Slides Engine v7.12 Orion')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  return html;
}

/**
 * Handle POST requests (API access)
 */
function doPost(e) {
  // Use lock to prevent concurrency issues if multiple requests come in
  const lock = LockService.getScriptLock();
  lock.tryLock(30000); // Wait up to 30s

  try {
    if (!e.postData || !e.postData.contents) {
      throw new Error('No Data');
    }

    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    let response = {};

    if (action === 'import') {
      response = importPresentation(request.presentationId, request.rawMode);
    } else if (action === 'generate') {
      // support both json string and object
      const jsonString = typeof request.json === 'string' ? request.json : JSON.stringify(request.json);
      response = generatePresentation(jsonString);
    } else {
      response = { status: 'error', message: 'Unknown action' };
    }

    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.message,
      stack: error.stack
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
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

    // 5. Execute Phase 1 (Batch Update) with Retry Logic
    if (buildResult.requests.length > 0) {
      executeBatchWithRetryV2(presentationId, buildResult.requests);
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

/**
 * Test Import for Debugging
 */
function testImport() {
  const ID = '1zv6VClAWFd0m6AWsagAQaMlhC1wWScD_sd45qXHSM90';
  const result = importPresentation(ID, false);
  const logs = [];
  if (result.status === 'success') {
    const json = JSON.parse(result.json);
    // Find all images
    let imageCount = 0;
    json.slides.forEach((slide, sIdx) => {
      slide.elements.forEach((el, eIdx) => {
        if (el.type === 'image') {
          imageCount++;
          const msg = `[DEBUG_IMG] Slide ${sIdx} El ${eIdx}: URL=${el.url ? el.url.substring(0, 50) + '...' : 'NULL'} Source=${el.sourceUrl}`;
          Logger.log(msg);
          logs.push(msg);
        }
      });
    });
    logs.push(`Found ${imageCount} images.`);
  } else {
    logs.push('Import Failed: ' + result.message);
  }
  return logs.join('\n');
}

/**
 * Execute batch requests with retry logic for failing images
 */
function executeBatchWithRetryV2(presentationId, requests) {
  let attempts = 0;
  const maxAttempts = 5;

  while (attempts < maxAttempts) {
    try {
      if (requests.length === 0) break;
      slidesApi.batchUpdate(presentationId, requests);
      break;
    } catch (e) {
      attempts++;
      Logger.log('DEBUG_RETRY: Catch block entered. Attempts: ' + attempts);
      Logger.log('DEBUG_RETRY: Error Message: ' + e.message);

      // Try multiple matches. Sometimes error message is prefixed.
      const match = e.message.match(/Invalid requests\[(\d+)\]/);

      if (match) {
        const index = parseInt(match[1], 10);
        Logger.log('DEBUG_RETRY: Matched Index: ' + index);

        if (index >= requests.length) {
          Logger.log('DEBUG_RETRY: Index OOB. Len: ' + requests.length);
          throw e; // Can't recover
        }

        const failReq = requests[index];
        // Safely log keys
        try { Logger.log('DEBUG_RETRY: FailReq Keys: ' + JSON.stringify(Object.keys(failReq))); } catch (ex) { }

        // If it's an image creation error, defer it to Phase 2
        if (failReq && (failReq.createImage || failReq.replaceImage)) { // Handle replaceImage too just in case
          Logger.log(`API Error on Request #${index} (Image). Deferring to Phase 2 fallback.`);

          // Add to deferred queue
          // We need pageObjectId to know where to put it. 
          // For replaceImage, we need logic. For createImage, we have elementProperties.pageObjectId.

          if (failReq.createImage) {
            const pageId = failReq.createImage.elementProperties.pageObjectId;
            phase2Service.addDeferredImage(pageId, failReq.createImage);
          } else {
            // replaceImage doesn't create a new image, it replaces an existing one.
            // We can't really "defer" it easily as insertImage creates NEW image.
            // Just Log and Skip for now to avoid crash.
            Logger.log('Skipping failed ReplaceImage request.');
          }

          // Remove the bad request and continue (retry loop will submit the rest)
          requests.splice(index, 1);
          continue;
        } else {
          Logger.log(`API Error on Request #${index} (NOT Image). Logic: ${JSON.stringify(failReq)}`);
          throw e;
        }
      } else {
        Logger.log('DEBUG_RETRY: No Regex Match for Invalid requests.');
      }
      throw e; // Non-indexable error
    }
  }
}
function executeBatchWithRetry(presentationId, requests) {
  let attempts = 0;
  const maxAttempts = 5; // Avoid infinite loops if multiple images fail

  while (attempts < maxAttempts) {
    try {
      if (requests.length === 0) break;
      slidesApi.batchUpdate(presentationId, requests);
      break; // Success
    } catch (e) {
      attempts++;
      const match = e.message.match(/Invalid requests\[(\d+)\]/);
      if (match) {
        const index = parseInt(match[1], 10);
        const failReq = requests[index];

        // If it's an image creation error, defer it to Phase 2
        if (failReq && failReq.createImage) {
          Logger.log(`API Error on Request #${index}. Deferring Image to Phase 2.`);

          // Add to deferred queue
          // We need pageObjectId to know where to put it. 
          const pageId = failReq.createImage.elementProperties.pageObjectId;

          // Capture the spec for Phase 2
          phase2Service.addDeferredImage(pageId, failReq.createImage);

          // Remove the bad request and continue (retry loop will submit the rest)
          requests.splice(index, 1);
          continue;
        } else {
          // If it's NOT an image error (e.g. text/shape), we can't auto-fix it.
          Logger.log(`API Error on Request #${index} (NOT Image). Logic: ${JSON.stringify(failReq)}`);
          throw e;
        }
      }
      throw e; // Non-indexable error
    }
  }
}
