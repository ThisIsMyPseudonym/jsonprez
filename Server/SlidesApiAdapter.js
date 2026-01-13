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
      Slides.Presentations.batchUpdate(resource, presentationId);
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

    // Groups are now handled atomically in Phase 1 via createGroup API
    // Only charts and speaker notes need Phase 2 (SlidesApp operations)
    const needsPhase2 = charts.length > 0 || notes.length > 0;

    if (!needsPhase2) {
      Logger.log('Phase 2: No deferred operations needed. Groups handled atomically in Phase 1.');
      if (groups.length > 0) {
        Logger.log('Note: ' + groups.length + ' groups were processed in Phase 1 batch.');
      }
      return;
    }

    Logger.log('Phase 2: Processing ' + charts.length + ' charts, ' + notes.length + ' speaker notes');

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

    // 3. Grouping (legacy fallback - groups now handled in Phase 1)
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
