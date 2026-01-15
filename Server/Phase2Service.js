/**
 * @fileoverview Phase 2 Service for tracking deferred operations.
 *
 * Phase 2 handles operations that cannot be done via the Slides API batchUpdate:
 *
 * QUEUES:
 * - charts:          Sheets-linked charts (require SpreadsheetApp)
 * - speakerNotes:    Speaker notes (require SlidesApp)
 * - groups:          Element grouping (fallback, now mostly handled in Phase 1)
 * - images:          REACTIVE fallback for API createImage failures
 * - proactiveImages: PROACTIVE routing for Google-internal URLs
 *
 * PROACTIVE vs REACTIVE IMAGE HANDLING:
 * - Proactive: Detected BEFORE API call, routed directly to SlidesApp
 *   (preserves crop, transform, recolor - see needsSlidesAppRouting)
 * - Reactive: API createImage failed, deferred to Phase 2 as fallback
 *   (limited property preservation)
 *
 * See ARCHITECTURE.md for full documentation on image handling.
 */

// Queue storage
const _phase2Queue = {
  charts: [],           // { slideIndex, chartSpec }
  speakerNotes: [],     // { slideIndex, notes }
  groups: [],           // { slideIndex, elementIds: [] }
  images: [],           // { slideIndex, imageSpec } - Reactive fallback for API failures
  proactiveImages: [],  // { slideIndex, objectId, slideId, element } - Proactive SlidesApp routing
  copyGroups: [],       // { slideIndex, sourcePresId, sourceObjectId, x, y, w, h } - Copy from source
  elementIds: {},       // { 'slide_0_element_5': 'obj_abc123' }

  reset: function () {
    this.charts = [];
    this.speakerNotes = [];
    this.groups = [];
    this.images = [];
    this.proactiveImages = [];
    this.copyGroups = [];
    this.elementIds = {};
  },

  // Store element ID for later lookup
  recordElementId: function (slideIndex, elementIndex, objectId) {
    const key = `slide_${slideIndex}_element_${elementIndex}`;
    this.elementIds[key] = objectId;
  },

  // Get element ID
  getElementId: function (slideIndex, elementIndex) {
    const key = `slide_${slideIndex}_element_${elementIndex}`;
    return this.elementIds[key];
  }
};

/**
 * Phase 2 Service Class
 */
class Phase2Service {
  constructor() {
    this.queue = _phase2Queue;
  }

  reset() { this.queue.reset(); }

  addChart(slideIndex, chartSpec) {
    this.queue.charts.push({ slideIndex, chartSpec });
  }

  addSpeakerNotes(slideIndex, notes) {
    this.queue.speakerNotes.push({ slideIndex, notes });
  }

  addDeferredImage(slideIndex, imageSpec) {
    this.queue.images.push({ slideIndex, imageSpec });
  }

  /**
   * Add image for proactive SlidesApp routing with full element data
   * @param {number} slideIndex - Index of the slide
   * @param {Object} imageData - { objectId, slideId, element }
   */
  addProactiveImage(slideIndex, imageData) {
    this.queue.proactiveImages.push({ slideIndex, ...imageData });
  }

  addGroup(slideIndex, elementIds) {
    this.queue.groups.push({ slideIndex, elementIds });
  }

  /**
   * Add a group to be copied from source presentation (for complex/curved line groups)
   * @param {number} slideIndex - Target slide index
   * @param {Object} copySpec - { sourcePresId, sourceSlideIndex, sourceObjectId, x, y, w, h }
   */
  addCopyGroup(slideIndex, copySpec) {
    this.queue.copyGroups.push({ slideIndex, ...copySpec });
  }

  getCopyGroups() { return this.queue.copyGroups; }

  recordElementId(slideIndex, elementIndex, objectId) {
    this.queue.recordElementId(slideIndex, elementIndex, objectId);
  }

  getElementId(slideIndex, elementIndex) {
    return this.queue.getElementId(slideIndex, elementIndex);
  }

  getCharts() { return this.queue.charts; }
  getSpeakerNotes() { return this.queue.speakerNotes; }
  getGroups() { return this.queue.groups; }
  getImages() { return this.queue.images; }
  getProactiveImages() { return this.queue.proactiveImages; }
}

const phase2Service = new Phase2Service();
