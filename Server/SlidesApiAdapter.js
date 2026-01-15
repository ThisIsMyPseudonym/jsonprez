/**
 * @fileoverview Adapter for Google Slides API.
 */

class SlidesApiAdapter {
  constructor() {
    validateAdvancedService();
  }

  /**
   * Create a new presentation
   * @param {string} title
   * @returns {string} Presentation ID
   */
  createPresentation(title) {
    const presentation = Slides.Presentations.create({
      title: title || 'New Presentation'
    });
    return presentation.presentationId;
  }

  /**
   * Execute batch update
   * @param {string} presentationId
   * @param {Array} requests
   */
  batchUpdate(presentationId, requests) {
    if (!requests || requests.length === 0) return;

    Logger.log('Executing batchUpdate. Req count: ' + requests.length + '. Pres ID type: ' + typeof presentationId);

    // Validate request structure to avoid confusing API errors
    // requests must be an array of objects
    if (!Array.isArray(requests)) {
      throw new Error('batchUpdate expects an array of requests.');
    }

    try {
      const resource = { requests: requests };
      // Log first request (usually createTable or createShape)
      if (requests.length > 0) {
        Logger.log('DEBUG_REQ_0: ' + JSON.stringify(requests[0]));
      }
      if (requests.length > 17) {
        Logger.log('DEBUG_REQ_17: ' + JSON.stringify(requests[17]));
      }
      const response = Slides.Presentations.batchUpdate(resource, presentationId);
      if (response && response.replies) {
        Logger.log('batchUpdate response: ' + response.replies.length + ' replies');
        // Log first reply to check for errors
        if (response.replies.length > 0) {
          Logger.log('First reply: ' + JSON.stringify(response.replies[0]));
        }
      }
    } catch (e) {
      Logger.log('Advanced Service Error: ' + e.message + '. Attempting REST API fallback.');

      // Fallback: Direct REST API call
      // Sometimes Advanced Services wrapper is finicky with types or environments.
      // We use UrlFetchApp to bypass the wrapper.
      try {
        const url = `https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`;
        const options = {
          method: 'post',
          contentType: 'application/json',
          headers: {
            Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
          },
          payload: JSON.stringify({ requests: requests })
        };
        UrlFetchApp.fetch(url, options);
        Logger.log('Fallback REST Call Successful');
      } catch (e2) {
        Logger.log('Fallback REST Error: ' + e2.message);
        throw e; // Throw original error or new error? Best to throw original for clarity unless fallback worked.
        // Actually if fallback fails, we want that error too.
        throw new Error('Slides API Failed (Advanced & REST): ' + e.message + ' | ' + e2.message);
      }
    }
  }

  /**
   * Execute Phase 2 Operations (SlidesApp)
   * @param {string} presentationId
   */
  executePhase2(presentationId, expectedSlideCount = 0) {
    if (!CONFIG.PHASE2.ENABLED) return;

    Logger.log('=== PHASE 2 EXECUTION ===');

    // Check what operations are actually needed
    const charts = phase2Service.getCharts();
    const notes = phase2Service.getSpeakerNotes();
    const groups = phase2Service.getGroups();
    const proactiveImages = phase2Service.getProactiveImages();
    const copyGroups = phase2Service.getCopyGroups();

    // Groups are now handled atomically in Phase 1 via createGroup API
    // Charts, speaker notes, proactive images, and copyGroups need Phase 2 (SlidesApp operations)
    const needsPhase2 = charts.length > 0 || notes.length > 0 || proactiveImages.length > 0 || copyGroups.length > 0;

    if (!needsPhase2) {
      Logger.log('Phase 2: No deferred operations needed. Groups handled atomically in Phase 1.');
      if (groups.length > 0) {
        Logger.log('Note: ' + groups.length + ' groups were processed in Phase 1 batch.');
      }
      return;
    }

    Logger.log('Phase 2: Processing ' + charts.length + ' charts, ' + notes.length + ' speaker notes, ' + proactiveImages.length + ' proactive images, ' + copyGroups.length + ' copyGroups');

    // SYNC: Poll for up to 30s until slides appear (only when needed)
    const maxRetries = 30;
    let presentation = SlidesApp.openById(presentationId);
    let slides = presentation.getSlides();

    // If expected count provided, wait until we see them
    if (expectedSlideCount > 0) {
      Logger.log('Waiting for ' + expectedSlideCount + ' slides to propagate...');
      for (let i = 0; i < maxRetries; i++) {
        if (slides.length >= expectedSlideCount) {
          Logger.log('Slides propagated after ' + i + ' seconds.');
          break;
        }
        // Close the instance to flush cache before retrying
        presentation.saveAndClose();
        Utilities.sleep(1000);
        presentation = SlidesApp.openById(presentationId); // Re-open to clear cache
        slides = presentation.getSlides(); // Re-fetch
      }

      if (slides.length < expectedSlideCount) {
        Logger.log('WARNING: Timeout waiting for slides. Saw ' + slides.length + ', expected ' + expectedSlideCount + '. Phase 2 may fail.');
      }
    } else {
      // Just original logic if no count (fallback)
      Utilities.sleep(2000);
      presentation = SlidesApp.openById(presentationId);
      slides = presentation.getSlides();
    }

    // =========================================================================
    // 0a. PROACTIVE IMAGE ROUTING (Hybrid SlidesApp + API)
    // =========================================================================
    //
    // PROBLEM:
    // Images extracted from Google Slides have `contentUrl` pointing to
    // googleusercontent.com domains. These URLs require OAuth authentication
    // and CANNOT be used with the Slides API `createImage` request, which
    // expects publicly accessible URLs. The API returns "problem retrieving
    // image" errors.
    //
    // SOLUTION:
    // Route these images through SlidesApp which CAN fetch authenticated URLs
    // via UrlFetchApp (using the script's OAuth token). Then apply transforms
    // and crop via the Slides API.
    //
    // CRITICAL MATH INSIGHT (prevents aspect ratio distortion):
    // --------------------------------------------------------
    // When an image has crop properties, the visible portion of the source
    // image is STRETCHED to fill the entire element rectangle.
    //
    //   Element size = Target visible size (NOT target/visibleFraction!)
    //
    // The scale factors should be:
    //   scaleX = targetWidth / sourceWidth
    //   scaleY = targetHeight / sourceHeight
    //
    // This produces NON-UNIFORM scale when crop is asymmetric, which is CORRECT.
    // For example, an image cropped horizontally (38.5% visible width) at
    // 68x38pt target needs scaleX≈46 and scaleY≈121 (very different values).
    //
    // WRONG approach (causes stretching):
    //   elementWidth = targetWidth / visibleFraction  // DO NOT DO THIS
    //
    // The crop determines WHAT portion of source fills the element, not the
    // element size itself.
    //
    // IMPLEMENTATION STEPS:
    // 1. Insert image via SlidesApp.insertImage(blob) - gets it into the preso
    // 2. saveAndClose() to flush changes to Google's servers
    // 3. Read element via Slides.Presentations.get() to get source size in EMU
    // 4. Calculate scale factors: target size / source size
    // 5. Apply ABSOLUTE transform via API (includes position + scale)
    // 6. Apply crop properties via API updateImageProperties
    //
    // See also: SlideBuilders.js needsSlidesAppRouting() for detection logic
    // =========================================================================
    if (proactiveImages && proactiveImages.length > 0) {
      Logger.log('Phase 2: Processing ' + proactiveImages.length + ' proactively routed images');

      // Store inserted image data for later API processing
      const insertedImages = [];

      proactiveImages.forEach(item => {
        try {
          const slide = slides[item.slideIndex];
          const element = item.element;

          if (!slide) {
            Logger.log('Proactive Image: Slide not found at index ' + item.slideIndex);
            return;
          }

          // Fetch blob with OAuth (this is why we use SlidesApp - it can auth)
          const response = UrlFetchApp.fetch(element.url, { muteHttpExceptions: true });
          if (response.getResponseCode() !== 200) {
            Logger.log('Proactive Image: Failed to fetch URL (code ' + response.getResponseCode() + '): ' + element.url);
            return;
          }
          const blob = response.getBlob();

          // Insert via SlidesApp - just get it into the presentation
          const image = slide.insertImage(blob);
          const objectId = image.getObjectId();
          const slideId = slide.getObjectId();

          Logger.log('Proactive Image inserted: ' + objectId + ' on slide ' + slideId);

          // Store for Phase 2 API processing (include z-order info from original item)
          insertedImages.push({
            objectId: objectId,
            slideId: slideId,
            element: element,
            elementIndex: item.elementIndex,
            totalElements: item.totalElements
          });

        } catch (e) {
          Logger.log('Proactive Image Insert Error: ' + e.message);
        }
      });

      // CRITICAL: Save and close to flush SlidesApp changes before reading via API
      if (insertedImages.length > 0) {
        Logger.log('Phase 2: Saving presentation to flush SlidesApp inserts...');
        presentation.saveAndClose();
        Utilities.sleep(500);

        // Now read the presentation via API to get the actual transforms
        const preso = Slides.Presentations.get(presentationId);
        const propertyRequests = [];

        insertedImages.forEach(item => {
          try {
            const element = item.element;

            // Find the page and element in the API response
            const page = preso.slides.find(s => s.objectId === item.slideId);
            if (!page) {
              Logger.log('Proactive Image: Could not find slide ' + item.slideId + ' in API response');
              return;
            }

            const pageElement = page.pageElements.find(e => e.objectId === item.objectId);
            if (!pageElement || !pageElement.image) {
              Logger.log('Proactive Image: Could not find element ' + item.objectId + ' in API response');
              return;
            }

            // Get the current transform from API
            const currentTransform = pageElement.transform || {};
            const curScaleX = currentTransform.scaleX || 1;
            const curScaleY = currentTransform.scaleY || 1;
            const curTranslateX = currentTransform.translateX || 0;
            const curTranslateY = currentTransform.translateY || 0;

            // Get the source size from API (this is the reference size for transforms)
            const sourceSize = pageElement.size || {};
            const sourceWidth = sourceSize.width ? sourceSize.width.magnitude : 0;
            const sourceHeight = sourceSize.height ? sourceSize.height.magnitude : 0;

            // Current displayed size = source * scale
            const currentDisplayWidth = sourceWidth * curScaleX;
            const currentDisplayHeight = sourceHeight * curScaleY;

            Logger.log('API Element: source=' + sourceWidth.toFixed(1) + 'x' + sourceHeight.toFixed(1) +
                       ' (EMU), curScale=' + curScaleX.toFixed(4) + 'x' + curScaleY.toFixed(4) +
                       ', displayed=' + currentDisplayWidth.toFixed(1) + 'x' + currentDisplayHeight.toFixed(1) + ' EMU');

            // Target display size (in EMU - API uses EMU)
            const EMU_PER_PT = 12700;
            const targetWidth = (element.w || 200) * SCALE * EMU_PER_PT;
            const targetHeight = (element.h || 200) * SCALE * EMU_PER_PT;
            const targetX = (element.x || 0) * SCALE * EMU_PER_PT;
            const targetY = (element.y || 0) * SCALE * EMU_PER_PT;

            // Calculate crop visible fractions
            let visibleFractionX = 1;
            let visibleFractionY = 1;
            if (element.crop) {
              visibleFractionX = 1 - (element.crop.left || 0) - (element.crop.right || 0);
              visibleFractionY = 1 - (element.crop.top || 0) - (element.crop.bottom || 0);
              if (visibleFractionX <= 0) visibleFractionX = 1;
              if (visibleFractionY <= 0) visibleFractionY = 1;
            }

            Logger.log('Target: ' + (targetWidth/EMU_PER_PT).toFixed(1) + 'x' + (targetHeight/EMU_PER_PT).toFixed(1) +
                       ' pt, crop visible: ' + visibleFractionX.toFixed(3) + 'x' + visibleFractionY.toFixed(3));

            // Element size = target visible size (NOT divided by visible fraction!)
            // The crop determines what portion of source fills the element rectangle
            // The visible portion of source is stretched to fill the entire element
            const elementWidth = targetWidth;
            const elementHeight = targetHeight;

            // New scale factors relative to source size
            // This produces non-uniform scale when crop is asymmetric (which is correct!)
            const newScaleX = elementWidth / sourceWidth;
            const newScaleY = elementHeight / sourceHeight;

            Logger.log('New transform: scaleX=' + newScaleX.toFixed(4) + ' scaleY=' + newScaleY.toFixed(4) +
                       ' translateX=' + targetX.toFixed(0) + ' translateY=' + targetY.toFixed(0));

            // Build new transform (ABSOLUTE replaces the whole transform)
            let transformObj = {
              scaleX: newScaleX,
              scaleY: newScaleY,
              shearX: 0,
              shearY: 0,
              translateX: targetX,
              translateY: targetY,
              unit: 'EMU'
            };

            // Handle rotation
            if (element.rotation) {
              const rad = (element.rotation * Math.PI) / 180;
              const cos = Math.cos(rad);
              const sin = Math.sin(rad);
              transformObj.scaleX = newScaleX * cos;
              transformObj.shearX = -newScaleX * sin;
              transformObj.shearY = newScaleY * sin;
              transformObj.scaleY = newScaleY * cos;
            }

            propertyRequests.push({
              updatePageElementTransform: {
                objectId: item.objectId,
                transform: transformObj,
                applyMode: 'ABSOLUTE'
              }
            });

            // Apply crop via API
            if (element.crop) {
              Logger.log('Queuing crop: L=' + (element.crop.left || 0).toFixed(3) +
                         ' R=' + (element.crop.right || 0).toFixed(3) +
                         ' T=' + (element.crop.top || 0).toFixed(3) +
                         ' B=' + (element.crop.bottom || 0).toFixed(3));
              propertyRequests.push({
                updateImageProperties: {
                  objectId: item.objectId,
                  imageProperties: {
                    cropProperties: {
                      leftOffset: element.crop.left || 0,
                      rightOffset: element.crop.right || 0,
                      topOffset: element.crop.top || 0,
                      bottomOffset: element.crop.bottom || 0,
                      angle: element.crop.angle || 0
                    }
                  },
                  fields: 'cropProperties'
                }
              });
            }

            // Recolor
            if (element.recolor) {
              propertyRequests.push({
                updateImageProperties: {
                  objectId: item.objectId,
                  imageProperties: { recolor: element.recolor },
                  fields: 'recolor'
                }
              });
            }

            // Border
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
              propertyRequests.push({
                updateImageProperties: {
                  objectId: item.objectId,
                  imageProperties: { outline: outline },
                  fields: 'outline'
                }
              });
            }

            // Link
            if (element.link) {
              const linkReq = buildLinkRequest(item.objectId, element.link, 'image');
              if (linkReq) propertyRequests.push(linkReq);
            }

          } catch (e) {
            Logger.log('Proactive Image Property Error: ' + e.message);
          }
        });

        // Execute all property updates
        if (propertyRequests.length > 0) {
          Logger.log('Phase 2: Applying ' + propertyRequests.length + ' image properties via API');
          try {
            this.batchUpdate(presentationId, propertyRequests);
          } catch (e) {
            Logger.log('Phase 2 Property Batch Error: ' + e.message);
          }
        }

        // Reopen presentation for any subsequent Phase 2 operations
        presentation = SlidesApp.openById(presentationId);
        slides = presentation.getSlides();

        // Z-ORDER FIX: Move images to correct position AFTER all transforms are applied
        // Images inserted via Phase 2 end up at the TOP (front) of the z-order.
        // We need to move each image BACK to its correct relative position.
        // For an image originally at index i out of n elements:
        //   - There are (n - i - 1) elements that should be IN FRONT of it
        //   - So we need to call sendBackward() that many times
        Logger.log('Phase 2: Fixing z-order for ' + insertedImages.length + ' images');
        insertedImages.forEach(item => {
          try {
            // Calculate how many elements should be in front of this image
            const elementsInFront = item.totalElements !== undefined && item.elementIndex !== undefined
              ? item.totalElements - item.elementIndex - 1
              : 0;

            if (elementsInFront === 0) {
              Logger.log('Z-order: ' + item.objectId + ' was at front (idx ' + item.elementIndex + '/' + item.totalElements + '), keeping at front');
              return; // Image should stay at front
            }

            const slide = presentation.getSlideById(item.slideId);
            if (slide) {
              const elements = slide.getPageElements();
              for (let i = 0; i < elements.length; i++) {
                if (elements[i].getObjectId() === item.objectId) {
                  // Move backward the correct number of positions
                  for (let j = 0; j < elementsInFront; j++) {
                    elements[i].sendBackward();
                  }
                  Logger.log('Z-order fixed: ' + item.objectId + ' moved back ' + elementsInFront + ' positions (was idx ' + item.elementIndex + '/' + item.totalElements + ')');
                  break;
                }
              }
            }
          } catch (e) {
            Logger.log('Z-order fix error: ' + e.message);
          }
        });
      }
    }

    // 0b. Deferred Image Processing (Fallback for API failures)
    const images = phase2Service.getImages();
    if (images && images.length > 0) {
      Logger.log('Phase 2: Processing ' + images.length + ' deferred images (API Fallback)');
      images.forEach(item => {
        try {
          const slideId = item.slideIndex; // In Controller.js retry logic, we store pageObjectId here
          const spec = item.imageSpec; // The original createImage request object

          let slide = presentation.getSlideById(slideId);
          if (!slide) {
            // Fallback: try to find by iterating if ID lookup fails (rare but possible with mixed ID types)
            slides.forEach(s => { if (s.getObjectId() === slideId) slide = s; });
          }

          if (slide) {
            // Fetch blob (this uses script's auth, bypassing the public URL requirement)
            const response = UrlFetchApp.fetch(spec.url);
            const blob = response.getBlob();

            const image = slide.insertImage(blob);

            // Apply Transforms (EMU to Points)
            const EMU_PER_PT = 12700;

            if (spec.elementProperties) {
              const props = spec.elementProperties;
              const size = props.size;
              const transform = props.transform;

              if (size) {
                if (size.width && size.width.magnitude) image.setWidth(size.width.magnitude / EMU_PER_PT);
                if (size.height && size.height.magnitude) image.setHeight(size.height.magnitude / EMU_PER_PT);
              }

              if (transform) {
                // simple scale/translate map
                const scaleX = transform.scaleX || 1;
                const scaleY = transform.scaleY || 1;
                const tx = transform.translateX || 0;
                const ty = transform.translateY || 0;

                image.setLeft(tx / EMU_PER_PT);
                image.setTop(ty / EMU_PER_PT);
                // SlidesApp doesn't support setting scaleX/Y directly easily without affecting size?
                // Actually setWidth/Height sets the visual size. 
                // We trust the size calculation above.
                // If there's rotation/shear, it's harder in Phase 2.
                // Assuming mostly standard images for now.
              }
            }
          } else {
            Logger.log('Phase 2 Image Error: Could not find slide with ID ' + slideId);
          }
        } catch (e) {
          Logger.log('Phase 2 Image Insert Failed: ' + e.message);
        }
      });
    }

    // 1. Chart Processing
    if (charts.length > 0) Logger.log('Processing ' + charts.length + ' charts');

    charts.forEach(item => {
      try {
        const slide = slides[item.slideIndex];
        const element = item.chartSpec;

        // Setup data
        const ss = SpreadsheetApp.create(`Temp Chart Data - ${Date.now()}`);
        const sheet = ss.getSheets()[0];
        const data = element.data || [['Category', 'Value'], ['Initial', 1]];
        if (data.length > 0) sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

        // Build chart
        const range = sheet.getRange(1, 1, data.length, data[0].length);
        const chartBuilder = sheet.newChart();
        const typeMap = { 'BAR': Charts.ChartType.BAR, 'COLUMN': Charts.ChartType.COLUMN, 'LINE': Charts.ChartType.LINE, 'AREA': Charts.ChartType.AREA, 'PIE': Charts.ChartType.PIE, 'SCATTER': Charts.ChartType.SCATTER };
        const chartType = typeMap[(element.chartType || 'COLUMN').toUpperCase()] || Charts.ChartType.COLUMN;

        chartBuilder.setChartType(chartType).addRange(range).setPosition(1, 1, 0, 0);
        if (element.title) chartBuilder.setOption('title', element.title);
        if (element.colors) chartBuilder.setOption('colors', element.colors);
        if (element.isStacked) chartBuilder.setOption('isStacked', 'absolute');

        const spreadSheetChart = chartBuilder.build();
        sheet.insertChart(spreadSheetChart);

        // Insert
        try {
          const sourceChart = sheet.getCharts()[0];
          const slideChart = slide.insertSheetsChart(sourceChart);
          slideChart.setLeft((element.x || 0) * SCALE);
          slideChart.setTop((element.y || 0) * SCALE);
          slideChart.setWidth((element.w || 400) * SCALE);
          slideChart.setHeight((element.h || 300) * SCALE);
        } catch (insertError) {
          Logger.log('Chart insert failed: ' + insertError.message + '. Trying as Image.');
          const sourceChart = sheet.getCharts()[0];
          const slideChart = slide.insertSheetsChartAsImage(sourceChart);
          slideChart.setLeft((element.x || 0) * SCALE);
          slideChart.setTop((element.y || 0) * SCALE);
          slideChart.setWidth((element.w || 400) * SCALE);
          slideChart.setHeight((element.h || 300) * SCALE);
        }
      } catch (e) {
        Logger.log('Phase 2 Chart Error: ' + e.message);
      }
    });

    // 2. Speaker Notes
    if (notes && notes.length > 0) Logger.log('Processing ' + notes.length + ' speaker notes');

    notes.forEach(item => {
      try {
        const slide = slides[item.slideIndex];
        const notesPage = slide.getNotesPage();
        const shape = notesPage.getSpeakerNotesShape();
        if (shape && shape.getText()) {
          shape.getText().setText(item.notes);
        }
      } catch (e) {
        Logger.log('Phase 2 Speaker Notes Error: ' + e.message);
      }
    });

    // 3. CopyGroups Processing (complex groups with curved lines)
    // =========================================================================
    // These are groups containing curved/freeform lines that cannot be
    // recreated via the Slides API (no control point data exposed).
    // We capture them as images from the source presentation's slide thumbnail.
    // =========================================================================
    if (copyGroups && copyGroups.length > 0) {
      Logger.log('Processing ' + copyGroups.length + ' copyGroups (curved line groups)');

      copyGroups.forEach(item => {
        try {
          const targetSlide = slides[item.slideIndex];
          if (!targetSlide) {
            Logger.log('CopyGroup: Target slide not found at index ' + item.slideIndex);
            return;
          }

          // Get source presentation info
          const sourcePres = Slides.Presentations.get(item.sourcePresentationId);
          if (!sourcePres || !sourcePres.slides) {
            Logger.log('CopyGroup: Source presentation not found: ' + item.sourcePresentationId);
            return;
          }

          // Find source slide
          const sourceSlide = sourcePres.slides[item.sourceSlideIndex];
          if (!sourceSlide) {
            Logger.log('CopyGroup: Source slide not found at index ' + item.sourceSlideIndex);
            return;
          }

          // Get slide dimensions (for calculating crop fractions)
          const pageSize = sourcePres.pageSize || {};
          const slideWidth = pageSize.width ? pageSize.width.magnitude : 9144000; // Default 720pt in EMU
          const slideHeight = pageSize.height ? pageSize.height.magnitude : 5143500; // Default 405pt in EMU
          const EMU_PER_PT = 12700;

          // Calculate group bounds in EMU early for bounds validation
          const groupXEmu = (item.x || 0) * SCALE * EMU_PER_PT;
          const groupYEmu = (item.y || 0) * SCALE * EMU_PER_PT;
          const groupWEmu = (item.w || 100) * SCALE * EMU_PER_PT;
          const groupHEmu = (item.h || 100) * SCALE * EMU_PER_PT;

          // BOUNDS VALIDATION: Skip groups that are entirely outside the slide
          const groupRight = groupXEmu + groupWEmu;
          const groupBottom = groupYEmu + groupHEmu;

          if (groupXEmu >= slideWidth || groupYEmu >= slideHeight ||
              groupRight <= 0 || groupBottom <= 0) {
            Logger.log('CopyGroup: SKIPPED - group entirely outside slide bounds');
            Logger.log('  Group: x=' + groupXEmu + ' y=' + groupYEmu + ' right=' + groupRight + ' bottom=' + groupBottom);
            Logger.log('  Slide: width=' + slideWidth + ' height=' + slideHeight);
            return;
          }

          // Check if visible area is too small (less than 1% in both dimensions)
          const visibleLeft = Math.max(0, groupXEmu);
          const visibleTop = Math.max(0, groupYEmu);
          const visibleRight = Math.min(slideWidth, groupRight);
          const visibleBottom = Math.min(slideHeight, groupBottom);
          const visibleW = visibleRight - visibleLeft;
          const visibleH = visibleBottom - visibleTop;

          if (visibleW < slideWidth * 0.01 && visibleH < slideHeight * 0.01) {
            Logger.log('CopyGroup: SKIPPED - visible area too small');
            Logger.log('  Visible: w=' + visibleW + ' h=' + visibleH);
            return;
          }

          // Get thumbnail of source slide
          Logger.log('CopyGroup: Getting thumbnail for slide ' + sourceSlide.objectId);
          const thumbnailResponse = Slides.Presentations.Pages.getThumbnail(
            item.sourcePresentationId,
            sourceSlide.objectId,
            { 'thumbnailProperties.mimeType': 'PNG', 'thumbnailProperties.thumbnailSize': 'LARGE' }
          );

          if (!thumbnailResponse || !thumbnailResponse.contentUrl) {
            Logger.log('CopyGroup: Failed to get thumbnail');
            return;
          }

          // Fetch the thumbnail image
          const thumbnailBlob = UrlFetchApp.fetch(thumbnailResponse.contentUrl).getBlob();

          // Insert the image into target slide
          const insertedImage = targetSlide.insertImage(thumbnailBlob);
          const objectId = insertedImage.getObjectId();

          Logger.log('CopyGroup: Inserted thumbnail image ' + objectId);

          // Get the inserted image's source size via API
          presentation.saveAndClose();
          Utilities.sleep(300);
          const targetPres = Slides.Presentations.get(presentationId);
          const targetPage = targetPres.slides.find(s => s.objectId === targetSlide.getObjectId());
          const imgElement = targetPage ? targetPage.pageElements.find(e => e.objectId === objectId) : null;

          if (!imgElement || !imgElement.size) {
            Logger.log('CopyGroup: Could not find inserted image element');
            presentation = SlidesApp.openById(presentationId);
            slides = presentation.getSlides();
            return;
          }

          const srcW = imgElement.size.width ? imgElement.size.width.magnitude : 1;
          const srcH = imgElement.size.height ? imgElement.size.height.magnitude : 1;

          Logger.log('CopyGroup: Thumbnail source size: ' + srcW + ' x ' + srcH + ' EMU');
          Logger.log('CopyGroup: Slide size: ' + slideWidth + ' x ' + slideHeight + ' EMU');

          // Group bounds already calculated above for validation (groupXEmu, groupYEmu, groupWEmu, groupHEmu)
          Logger.log('CopyGroup: Group bounds in EMU - x:' + groupXEmu + ' y:' + groupYEmu + ' w:' + groupWEmu + ' h:' + groupHEmu);

          // Crop fractions (relative to full slide)
          const cropLeft = Math.max(0, Math.min(0.99, groupXEmu / slideWidth));
          const cropTop = Math.max(0, Math.min(0.99, groupYEmu / slideHeight));
          const cropRight = Math.max(0, Math.min(0.99, 1 - ((groupXEmu + groupWEmu) / slideWidth)));
          const cropBottom = Math.max(0, Math.min(0.99, 1 - ((groupYEmu + groupHEmu) / slideHeight)));

          Logger.log('CopyGroup: Crop fractions - L:' + cropLeft.toFixed(3) + ' T:' + cropTop.toFixed(3) +
                     ' R:' + cropRight.toFixed(3) + ' B:' + cropBottom.toFixed(3));

          // CORRECTED APPROACH (from proactive images pattern):
          // Element size = target visible size (NOT divided by visible fraction!)
          // The crop determines what portion of source fills the element rectangle
          // The visible portion of source is STRETCHED to fill the entire element
          //
          // Scale = targetSize / sourceSize (simple ratio, no crop math)
          // Translate = target position
          // Crop is applied separately and stretches to fill element
          const scaleX = groupWEmu / srcW;
          const scaleY = groupHEmu / srcH;

          Logger.log('CopyGroup: Scale factors - scaleX:' + scaleX.toFixed(4) + ' scaleY:' + scaleY.toFixed(4));
          Logger.log('CopyGroup: Position - translateX:' + groupXEmu + ' translateY:' + groupYEmu + ' EMU');

          const updateRequests = [];

          // Transform: scale to slide proportions, position at group location
          updateRequests.push({
            updatePageElementTransform: {
              objectId: objectId,
              transform: {
                scaleX: scaleX,
                scaleY: scaleY,
                shearX: 0,
                shearY: 0,
                translateX: groupXEmu,
                translateY: groupYEmu,
                unit: 'EMU'
              },
              applyMode: 'ABSOLUTE'
            }
          });

          // Apply crop to show only the group region
          updateRequests.push({
            updateImageProperties: {
              objectId: objectId,
              imageProperties: {
                cropProperties: {
                  leftOffset: cropLeft,
                  rightOffset: cropRight,
                  topOffset: cropTop,
                  bottomOffset: cropBottom
                }
              },
              fields: 'cropProperties'
            }
          });

          try {
            this.batchUpdate(presentationId, updateRequests);
            Logger.log('CopyGroup: Applied transform and crop for ' + objectId);
          } catch (e) {
            Logger.log('CopyGroup: batchUpdate error - ' + e.message);
          }

          // Reopen presentation for subsequent operations
          presentation = SlidesApp.openById(presentationId);
          slides = presentation.getSlides();

          // Z-ORDER FIX: Move copyGroup image to correct position
          // Images inserted via Phase 2 end up at the TOP (front) of the z-order.
          // We need to move each image BACK to its correct relative position.
          try {
            // Calculate how many elements should be in front of this copyGroup
            const elementsInFront = item.totalElements !== undefined && item.elementIndex !== undefined
              ? item.totalElements - item.elementIndex - 1
              : 0;

            const slide = slides[item.slideIndex];
            if (slide) {
              const elements = slide.getPageElements();
              for (let i = 0; i < elements.length; i++) {
                if (elements[i].getObjectId() === objectId) {
                  if (elementsInFront === 0) {
                    Logger.log('CopyGroup: Z-order - ' + objectId + ' at front (idx ' + item.elementIndex + '/' + item.totalElements + '), keeping at front');
                  } else {
                    // Move backward the correct number of positions
                    for (let j = 0; j < elementsInFront; j++) {
                      elements[i].sendBackward();
                    }
                    Logger.log('CopyGroup: Z-order fixed - ' + objectId + ' moved back ' + elementsInFront + ' positions (was idx ' + item.elementIndex + '/' + item.totalElements + ')');
                  }
                  break;
                }
              }
            }
          } catch (zErr) {
            Logger.log('CopyGroup: Z-order fix error - ' + zErr.message);
          }

        } catch (e) {
          Logger.log('CopyGroup Error: ' + e.message);
        }
      });
    }

    // 4. Grouping (legacy fallback - groups now handled in Phase 1)
    if (groups && groups.length > 0) Logger.log('Processing ' + groups.length + ' groups (fallback)');

    groups.forEach(item => {
      try {
        const slide = slides[item.slideIndex];
        const pageElements = slide.getPageElements();
        const elementsToGroup = [];

        // Find elements by tracked objectId
        item.elementIds.forEach(id => {
          // We must traverse existing elements to find match
          // SlidesApp doesn't have getElementById easily for generic elements without iteration
          for (let i = 0; i < pageElements.length; i++) {
            if (pageElements[i].getObjectId() === id) {
              elementsToGroup.push(pageElements[i]);
              break;
            }
          }
        });

        if (elementsToGroup.length > 1) {
          slide.group(elementsToGroup);
        }
      } catch (e) {
        Logger.log('Phase 2 Grouping Error: ' + e.message);
      }
    });
  }
}

const slidesApi = new SlidesApiAdapter();
