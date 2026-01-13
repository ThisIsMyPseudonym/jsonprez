/**
 * @fileoverview Configuration constants and global settings for the Slides Engine.
 */

/**
 * Global Configuration Object
 * @constant
 */
const CONFIG = {
  VERSION: "7.12.0",
  CODENAME: "Orion",

  // Phase 2 (SlidesApp) features - can be disabled if causing issues
  PHASE2: {
    ENABLED: true,
    SPEAKER_NOTES: true,
    GROUPS: true,
    LINKED_CHARTS: true
  },

  // Canvas now uses raw points (same as Google Slides)
  CANVAS: {
    WIDTH: 720,
    HEIGHT: 405,
    ASPECT_RATIO: 16 / 9
  },

  // Google Slides actual dimensions (points)
  SLIDES: {
    WIDTH: 720,
    HEIGHT: 405
  },

  // Property name aliases for user convenience
  PROPERTY_ALIASES: {
    'width': 'w',
    'height': 'h',
    'fontColor': 'color',
    'backgroundColor': 'background'
  },

  // Default values (aligned with modern design / DEFAULT_THEME)
  DEFAULTS: {
    FONT_SIZE: 16,
    FONT_FAMILY: 'Roboto',
    TEXT_COLOR: '#1e293b',      // Softer than pure black, matches theme.text
    SHAPE_FILL: '#3b82f6',      // Theme primary blue
    ICON_COLOR: '#3b82f6',      // Theme primary for icons
    ICON_BG_OPACITY: 0.12,      // Subtle background circle
    LINE_WEIGHT: 2,
    TABLE_HEADER_BG: '#f1f5f9', // Light gray, matches theme.surface variant
    TABLE_FONT_SIZE: 14,        // Slightly smaller than body text
    // Padding defaults (in points) for text boxes and shapes
    PADDING_TOP: 5.0,
    PADDING_BOTTOM: 5.0,
    PADDING_LEFT: 7.2,
    PADDING_RIGHT: 7.2
  },

  // Validation limits
  LIMITS: {
    MAX_SLIDES: 100,
    MAX_ELEMENTS_PER_SLIDE: 100,
    MAX_TEXT_LENGTH: 10000,
    MIN_DIMENSION: 10,
    MAX_DIMENSION: 1000,
    MIN_FONT_SIZE: 6,
    MAX_FONT_SIZE: 200
  },

  // Fake shadow presets (workaround for read-only shadow API)
  // Simple single-layer shadows - clean and modern
  // Properties: angle (degrees), distance, spread (size increase), opacity, color
  SHADOW_PRESETS: {
    // Standard drop shadows (bottom-right, 135°)
    subtle: { angle: 135, distance: 2, spread: 0, opacity: 0.08, color: '#000000' },
    medium: { angle: 135, distance: 4, spread: 2, opacity: 0.12, color: '#000000' },
    large: { angle: 135, distance: 6, spread: 4, opacity: 0.18, color: '#000000' },
    dramatic: { angle: 135, distance: 10, spread: 6, opacity: 0.25, color: '#000000' },
    // Directional shadows
    bottom: { angle: 90, distance: 5, spread: 2, opacity: 0.15, color: '#000000' },
    top: { angle: 270, distance: 5, spread: 2, opacity: 0.15, color: '#000000' },
    left: { angle: 180, distance: 5, spread: 2, opacity: 0.15, color: '#000000' },
    right: { angle: 0, distance: 5, spread: 2, opacity: 0.15, color: '#000000' },
    // Special effects
    glow: { angle: 0, distance: 0, spread: 12, opacity: 0.25, color: 'inherit' },
    soft: { angle: 135, distance: 3, spread: 8, opacity: 0.10, color: '#000000' },
    hard: { angle: 135, distance: 3, spread: 0, opacity: 0.20, color: '#000000' },
    floating: { angle: 90, distance: 12, spread: 8, opacity: 0.20, color: '#000000' }
  },

  // List/bullet presets for createParagraphBullets
  LIST_PRESETS: {
    // Bulleted lists
    'bullet': 'BULLET_DISC_CIRCLE_SQUARE',
    'disc': 'BULLET_DISC_CIRCLE_SQUARE',
    'arrow': 'BULLET_ARROW_DIAMOND_DISC',
    'star': 'BULLET_STAR_CIRCLE_SQUARE',
    'checkbox': 'BULLET_CHECKBOX',
    'diamond': 'BULLET_DIAMOND_CIRCLE_SQUARE',
    // Numbered lists
    'number': 'NUMBERED_DIGIT_ALPHA_ROMAN',
    'numbered': 'NUMBERED_DIGIT_ALPHA_ROMAN',
    'alpha': 'NUMBERED_UPPERALPHA_ALPHA_ROMAN',
    'roman': 'NUMBERED_UPPERROMAN_UPPERALPHA_DIGIT',
    // Full preset names also accepted
    'BULLET_DISC_CIRCLE_SQUARE': 'BULLET_DISC_CIRCLE_SQUARE',
    'BULLET_ARROW_DIAMOND_DISC': 'BULLET_ARROW_DIAMOND_DISC',
    'BULLET_STAR_CIRCLE_SQUARE': 'BULLET_STAR_CIRCLE_SQUARE',
    'BULLET_CHECKBOX': 'BULLET_CHECKBOX',
    'BULLET_DIAMOND_CIRCLE_SQUARE': 'BULLET_DIAMOND_CIRCLE_SQUARE',
    'BULLET_DIAMONDX_ARROW3D_SQUARE': 'BULLET_DIAMONDX_ARROW3D_SQUARE',
    'BULLET_ARROW3D_CIRCLE_SQUARE': 'BULLET_ARROW3D_CIRCLE_SQUARE',
    'BULLET_LEFTTRIANGLE_DIAMOND_DISC': 'BULLET_LEFTTRIANGLE_DIAMOND_DISC',
    'NUMBERED_DIGIT_ALPHA_ROMAN': 'NUMBERED_DIGIT_ALPHA_ROMAN',
    'NUMBERED_DIGIT_ALPHA_ROMAN_PARENS': 'NUMBERED_DIGIT_ALPHA_ROMAN_PARENS',
    'NUMBERED_DIGIT_NESTED': 'NUMBERED_DIGIT_NESTED',
    'NUMBERED_UPPERALPHA_ALPHA_ROMAN': 'NUMBERED_UPPERALPHA_ALPHA_ROMAN',
    'NUMBERED_UPPERROMAN_UPPERALPHA_DIGIT': 'NUMBERED_UPPERROMAN_UPPERALPHA_DIGIT',
    'NUMBERED_ZERODIGIT_ALPHA_ROMAN': 'NUMBERED_ZERODIGIT_ALPHA_ROMAN'
  }
};

/**
 * Calculate scale factor
 * @constant
 */
const SCALE = CONFIG.SLIDES.WIDTH / CONFIG.CANVAS.WIDTH;

/**
 * Mappings for Shapes, Lines, Alignments, etc.
 * @constant
 */
const ENUMS = {
  SHAPE_TYPE_MAP: {
    // Basic shapes (uppercase API names)
    'RECTANGLE': 'RECTANGLE',
    'ROUNDED_RECTANGLE': 'ROUND_RECTANGLE',
    'ROUND_RECTANGLE': 'ROUND_RECTANGLE',
    'ELLIPSE': 'ELLIPSE',
    'CIRCLE': 'ELLIPSE',
    'TRIANGLE': 'TRIANGLE',
    'DIAMOND': 'DIAMOND',
    'PARALLELOGRAM': 'PARALLELOGRAM',
    'TRAPEZOID': 'TRAPEZOID',
    'PENTAGON': 'PENTAGON',
    'HEXAGON': 'HEXAGON',
    'HEPTAGON': 'HEPTAGON',
    'OCTAGON': 'OCTAGON',
    'DECAGON': 'DECAGON',
    'DODECAGON': 'DODECAGON',
    // Basic shapes (lowercase user-friendly aliases)
    'rectangle': 'RECTANGLE',
    'roundRectangle': 'ROUND_RECTANGLE',
    'roundedRectangle': 'ROUND_RECTANGLE',
    'ellipse': 'ELLIPSE',
    'circle': 'ELLIPSE',
    'triangle': 'TRIANGLE',
    'diamond': 'DIAMOND',
    'parallelogram': 'PARALLELOGRAM',
    'trapezoid': 'TRAPEZOID',
    'pentagon': 'PENTAGON',
    'hexagon': 'HEXAGON',
    'heptagon': 'HEPTAGON',
    'octagon': 'OCTAGON',
    'chord': 'CHORD',
    'CHORD': 'CHORD',
    // Stars and banners
    'STAR_4': 'STAR_4', 'STAR_5': 'STAR_5', 'STAR_6': 'STAR_6',
    'STAR_7': 'STAR_7', 'STAR_8': 'STAR_8', 'STAR_10': 'STAR_10',
    'STAR_12': 'STAR_12', 'STAR_16': 'STAR_16', 'STAR_24': 'STAR_24',
    'STAR_32': 'STAR_32', 'RIBBON': 'RIBBON', 'RIBBON_2': 'RIBBON_2',
    'star': 'STAR_5', 'star4': 'STAR_4', 'star5': 'STAR_5',
    'star6': 'STAR_6', 'star7': 'STAR_7', 'star8': 'STAR_8',
    'star10': 'STAR_10', 'star12': 'STAR_12', 'ribbon': 'RIBBON',
    // Arrows
    'ARROW_RIGHT': 'RIGHT_ARROW', 'ARROW_LEFT': 'LEFT_ARROW',
    'ARROW_UP': 'UP_ARROW', 'ARROW_DOWN': 'DOWN_ARROW',
    'ARROW_LEFT_RIGHT': 'LEFT_RIGHT_ARROW', 'ARROW_UP_DOWN': 'UP_DOWN_ARROW',
    'RIGHT_ARROW': 'RIGHT_ARROW', 'LEFT_ARROW': 'LEFT_ARROW',
    'UP_ARROW': 'UP_ARROW', 'DOWN_ARROW': 'DOWN_ARROW',
    'CHEVRON': 'CHEVRON', 'NOTCHED_RIGHT_ARROW': 'NOTCHED_RIGHT_ARROW',
    'HOME_PLATE': 'HOME_PLATE',
    'arrowRight': 'RIGHT_ARROW', 'arrowLeft': 'LEFT_ARROW',
    'arrowUp': 'UP_ARROW', 'arrowDown': 'DOWN_ARROW', 'chevron': 'CHEVRON',
    // Callouts
    'CLOUD': 'CLOUD', 'CLOUD_CALLOUT': 'CLOUD',
    'WEDGE_RECTANGLE_CALLOUT': 'WEDGE_RECTANGLE_CALLOUT',
    'WEDGE_ROUND_RECTANGLE_CALLOUT': 'WEDGE_RECTANGLE_CALLOUT',
    'WEDGE_ELLIPSE_CALLOUT': 'WEDGE_ELLIPSE_CALLOUT',
    'cloud': 'CLOUD', 'callout': 'WEDGE_RECTANGLE_CALLOUT',
    'speechBubble': 'WEDGE_RECTANGLE_CALLOUT',
    // Math and special
    'PLUS': 'PLUS', 'PIE': 'PIE', 'ARC': 'ARC', 'DONUT': 'DONUT',
    'NO_SMOKING': 'NO_SMOKING', 'BLOCK_ARC': 'BLOCK_ARC',
    'FOLDED_CORNER': 'FOLDED_CORNER', 'FRAME': 'FRAME',
    'HALF_FRAME': 'HALF_FRAME', 'CORNER': 'CORNER', 'CUBE': 'CUBE',
    'CAN': 'CAN', 'LIGHTNING_BOLT': 'LIGHTNING_BOLT', 'HEART': 'HEART',
    'SUN': 'SUN', 'MOON': 'MOON', 'SMILEY_FACE': 'SMILEY_FACE',
    'IRREGULAR_SEAL_1': 'IRREGULAR_SEAL_1', 'IRREGULAR_SEAL_2': 'IRREGULAR_SEAL_2',
    'plus': 'PLUS', 'pie': 'PIE', 'arc': 'ARC', 'donut': 'DONUT',
    'cube': 'CUBE', 'can': 'CAN', 'cylinder': 'CAN',
    'lightning': 'LIGHTNING_BOLT', 'lightningBolt': 'LIGHTNING_BOLT',
    'bolt': 'LIGHTNING_BOLT', 'heart': 'HEART', 'sun': 'SUN',
    'moon': 'MOON', 'smiley': 'SMILEY_FACE', 'smileyFace': 'SMILEY_FACE',
    'face': 'SMILEY_FACE', 'burst': 'IRREGULAR_SEAL_1',
    'explosion': 'IRREGULAR_SEAL_2', 'frame': 'FRAME',
    'foldedCorner': 'FOLDED_CORNER',
    // Flowchart
    'FLOWCHART_PROCESS': 'RECTANGLE', 'FLOWCHART_DECISION': 'DIAMOND',
    'FLOWCHART_ALTERNATE_PROCESS': 'ROUND_RECTANGLE',
    'FLOWCHART_DATA': 'PARALLELOGRAM', 'FLOWCHART_TERMINATOR': 'ROUND_RECTANGLE',
    'FLOWCHART_DOCUMENT': 'SNIP_ROUND_RECTANGLE', 'FLOWCHART_CONNECTOR': 'ELLIPSE',
    'FLOWCHART_PREPARATION': 'HEXAGON', 'FLOWCHART_MANUAL_INPUT': 'TRAPEZOID',
    'FLOWCHART_MERGE': 'TRIANGLE', 'FLOWCHART_DELAY': 'HALF_FRAME',
    // Braces
    'LEFT_BRACE': 'LEFT_BRACE', 'RIGHT_BRACE': 'RIGHT_BRACE',
    'LEFT_BRACKET': 'LEFT_BRACKET', 'RIGHT_BRACKET': 'RIGHT_BRACKET',
    'leftBrace': 'LEFT_BRACE', 'rightBrace': 'RIGHT_BRACE',
    'leftBracket': 'LEFT_BRACKET', 'rightBracket': 'RIGHT_BRACKET',
    // Text
    'TEXT_BOX': 'TEXT_BOX', 'textBox': 'TEXT_BOX'
  },

  ARROW_TYPE_MAP: {
    // User-friendly aliases -> API values
    'NONE': 'NONE', 'ARROW': 'STEALTH_ARROW', 'OPEN_ARROW': 'OPEN_ARROW',
    'CIRCLE': 'FILL_CIRCLE', 'OPEN_CIRCLE': 'OPEN_CIRCLE',
    'DIAMOND': 'FILL_DIAMOND', 'OPEN_DIAMOND': 'OPEN_DIAMOND',
    'SQUARE': 'FILL_SQUARE', 'OPEN_SQUARE': 'OPEN_SQUARE',
    // Passthrough for raw API values (from extraction)
    'STEALTH_ARROW': 'STEALTH_ARROW', 'FILL_ARROW': 'FILL_ARROW',
    'FILL_CIRCLE': 'FILL_CIRCLE', 'FILL_DIAMOND': 'FILL_DIAMOND', 'FILL_SQUARE': 'FILL_SQUARE'
  },

  DASH_STYLE_MAP: {
    'SOLID': 'SOLID', 'DOT': 'DOT', 'DASH': 'DASH',
    'DASH_DOT': 'DASH_DOT', 'LONG_DASH': 'LONG_DASH',
    'LONG_DASH_DOT': 'LONG_DASH_DOT',
    'solid': 'SOLID', 'dot': 'DOT', 'dotted': 'DOT',
    'dash': 'DASH', 'dashed': 'DASH',
    'dash-dot': 'DASH_DOT', 'dashDot': 'DASH_DOT',
    'long-dash': 'LONG_DASH', 'longDash': 'LONG_DASH',
    'long-dash-dot': 'LONG_DASH_DOT', 'longDashDot': 'LONG_DASH_DOT'
  },

  ALIGNMENT_MAP: {
    'left': 'START', 'center': 'CENTER',
    'right': 'END', 'justify': 'JUSTIFIED'
  },

  VERTICAL_ALIGNMENT_MAP: {
    'top': 'TOP', 'middle': 'MIDDLE', 'bottom': 'BOTTOM'
  },

  LAYOUT_MAP: {
    'blank': 'BLANK', 'title': 'TITLE', 'titleSlide': 'TITLE',
    'title-slide': 'TITLE', 'titleAndBody': 'TITLE_AND_BODY',
    'title-and-body': 'TITLE_AND_BODY', 'titleBody': 'TITLE_AND_BODY',
    'titleAndTwoColumns': 'TITLE_AND_TWO_COLUMNS',
    'title-and-two-columns': 'TITLE_AND_TWO_COLUMNS',
    'twoColumns': 'TITLE_AND_TWO_COLUMNS', 'titleOnly': 'TITLE_ONLY',
    'title-only': 'TITLE_ONLY', 'sectionHeader': 'SECTION_HEADER',
    'section-header': 'SECTION_HEADER', 'section': 'SECTION_HEADER',
    'sectionTitle': 'SECTION_TITLE_AND_DESCRIPTION',
    'section-title': 'SECTION_TITLE_AND_DESCRIPTION',
    'oneColumn': 'ONE_COLUMN_TEXT', 'one-column': 'ONE_COLUMN_TEXT',
    'mainPoint': 'MAIN_POINT', 'main-point': 'MAIN_POINT',
    'bigNumber': 'BIG_NUMBER', 'big-number': 'BIG_NUMBER',
    'captionOnly': 'CAPTION_ONLY', 'caption-only': 'CAPTION_ONLY',
    'caption': 'CAPTION_ONLY',
    'BLANK': 'BLANK', 'TITLE': 'TITLE', 'TITLE_AND_BODY': 'TITLE_AND_BODY',
    'TITLE_AND_TWO_COLUMNS': 'TITLE_AND_TWO_COLUMNS',
    'TITLE_ONLY': 'TITLE_ONLY', 'SECTION_HEADER': 'SECTION_HEADER',
    'SECTION_TITLE_AND_DESCRIPTION': 'SECTION_TITLE_AND_DESCRIPTION',
    'ONE_COLUMN_TEXT': 'ONE_COLUMN_TEXT', 'MAIN_POINT': 'MAIN_POINT',
    'BIG_NUMBER': 'BIG_NUMBER', 'CAPTION_ONLY': 'CAPTION_ONLY'
  }
};
