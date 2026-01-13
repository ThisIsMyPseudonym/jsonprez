/**
 * @fileoverview Phase 2 Service for tracking deferred operations.
 */

// Queue storage
const _phase2Queue = {
  charts: [],       // { slideIndex, chartSpec }
  speakerNotes: [], // { slideIndex, notes }
  groups: [],       // { slideIndex, elementIds: [] }
  elementIds: {},   // { 'slide_0_element_5': 'obj_abc123' }
  
  reset: function() {
    this.charts = [];
    this.speakerNotes = [];
    this.groups = [];
    this.elementIds = {};
  },

  // Store element ID for later lookup
  recordElementId: function(slideIndex, elementIndex, objectId) {
    const key = `slide_${slideIndex}_element_${elementIndex}`;
    this.elementIds[key] = objectId;
  },
  
  // Get element ID
  getElementId: function(slideIndex, elementIndex) {
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

  addGroup(slideIndex, elementIds) {
    this.queue.groups.push({ slideIndex, elementIds });
  }

  recordElementId(slideIndex, elementIndex, objectId) {
    this.queue.recordElementId(slideIndex, elementIndex, objectId);
  }

  getElementId(slideIndex, elementIndex) {
    return this.queue.getElementId(slideIndex, elementIndex);
  }

  getCharts() { return this.queue.charts; }
  getSpeakerNotes() { return this.queue.speakerNotes; }
  getGroups() { return this.queue.groups; }
}

const phase2Service = new Phase2Service();
