/**
 * @fileoverview Theme and Color Resolution Service.
 */

/**
 * Default theme configuration
 * @const
 */
const DEFAULT_THEME = {
  colors: {
    primary: '#3b82f6',
    secondary: '#10b981',
    accent: '#f59e0b',
    background: '#ffffff',
    surface: '#f8fafc',
    text: '#1e293b',
    textLight: '#64748b',
    success: '#22c55e',
    warning: '#eab308',
    error: '#ef4444',
    info: '#3b82f6'
  },
  fonts: {
    heading: 'Roboto',
    body: 'Roboto',
    mono: 'Roboto Mono'
  }
};

/**
 * Theme Manager Class
 */
class ThemeService {
  constructor() {
    this.currentTheme = null;
  }

  /**
   * Set the current theme from JSON config
   * @param {Object} themeConfig - Theme configuration from JSON
   */
  setTheme(themeConfig) {
    if (!themeConfig) {
      this.currentTheme = DEFAULT_THEME;
      return;
    }

    this.currentTheme = {
      colors: { ...DEFAULT_THEME.colors, ...(themeConfig.colors || {}) },
      fonts: { ...DEFAULT_THEME.fonts, ...(themeConfig.fonts || {}) }
    };
  }

  /**
   * Resolve a theme color name to its hex value
   * @param {string} colorValue - Color value (hex or theme name like "primary")
   * @returns {string|null} - Resolved hex color or null
   */
  resolveThemeColor(colorValue) {
    if (!colorValue) return null;

    // If it's already a hex color, return as-is
    if (colorValue.startsWith('#')) return colorValue;

    // Handle theme:ACCENT1 style references (from extraction)
    if (colorValue.startsWith('theme:')) {
      const themeColorType = colorValue.substring(6); // Remove 'theme:' prefix
      // Map Google Slides theme color types to our theme colors
      const themeColorMapping = {
        'DARK1': this.currentTheme?.colors?.text || '#1e293b',
        'DARK2': this.currentTheme?.colors?.textLight || '#64748b',
        'LIGHT1': this.currentTheme?.colors?.background || '#ffffff',
        'LIGHT2': this.currentTheme?.colors?.surface || '#f8fafc',
        'ACCENT1': this.currentTheme?.colors?.primary || '#4285f4',
        'ACCENT2': this.currentTheme?.colors?.secondary || '#34a853',
        'ACCENT3': this.currentTheme?.colors?.accent || '#fbbc04',
        'ACCENT4': this.currentTheme?.colors?.error || '#ea4335',
        'ACCENT5': this.currentTheme?.colors?.accent5 || '#46bdc6',
        'ACCENT6': this.currentTheme?.colors?.accent6 || '#7baaf7',
        'HYPERLINK': '#1a73e8',
        'FOLLOWED_HYPERLINK': '#660099'
      };
      return themeColorMapping[themeColorType] || this.currentTheme?.colors?.primary || '#4285f4';
    }

    // Check if it's a theme color name
    const theme = this.currentTheme || DEFAULT_THEME;
    if (theme.colors && theme.colors[colorValue]) {
      return theme.colors[colorValue];
    }

    // Not a theme color - return as-is (might be a color name like 'none')
    return colorValue;
  }

  /**
   * Resolve a theme font name to its font family
   * @param {string} fontValue - Font value (font family or theme name like "heading")
   * @returns {string} - Resolved font family
   */
  resolveThemeFont(fontValue) {
    if (!fontValue) return CONFIG.DEFAULTS.FONT_FAMILY;

    // Check if it's a theme font name
    const theme = this.currentTheme || DEFAULT_THEME;
    if (theme.fonts && theme.fonts[fontValue]) {
      return theme.fonts[fontValue];
    }

    // Not a theme font - return as-is
    return fontValue;
  }

  /**
   * Convert hex color to Slides API RGB format (0-1 range)
   * @param {string} hex 
   * @returns {Object|null} {red, green, blue} or null
   */
  hexToRgbApi(hex) {
    hex = normalizeColor(hex); // Using global validator Helper or local one? 
    // Ideally Validation.gs exposes normalizeColor. In GAS they share scope.
    if (!hex) return null;

    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
      red: parseInt(result[1], 16) / 255,
      green: parseInt(result[2], 16) / 255,
      blue: parseInt(result[3], 16) / 255
    } : null;
  }
}

// Export singleton instance
const themeService = new ThemeService();
