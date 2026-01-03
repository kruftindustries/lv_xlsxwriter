/*
 * libxlsxwriter_LV.h - LabVIEW-compatible header for libxlsxwriter
 *
 * This header is specifically designed for LabVIEW's Import Shared Library wizard.
 * All preprocessor conditionals have been removed for compatibility.
 * All pointer types use unsigned long (32-bit handle).
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2025, John McNamara, jmcnamara@cpan.org.
 */

#ifndef __LIBXLSXWRITER_LV_H__
#define __LIBXLSXWRITER_LV_H__

/* ============================================================================
 * Basic Type Definitions (no preprocessor conditionals)
 * ============================================================================ */

typedef signed char        int8_t;
typedef signed short       int16_t;
typedef signed int         int32_t;
typedef unsigned char      uint8_t;
typedef unsigned short     uint16_t;
typedef unsigned int       uint32_t;
typedef unsigned int       size_t;

/* Pointer-sized unsigned integer for array of string pointers.
 * For 32-bit DLL: use U32 array in LabVIEW
 * For 64-bit DLL: use U64 array in LabVIEW
 * The DLL is compiled for the target platform, so this typedef
 * will be correct when the DLL is built. */
typedef unsigned long      uintptr_t;

/* ============================================================================
 * Library-Specific Type Definitions
 * ============================================================================ */

typedef uint32_t lxw_row_t;
typedef uint16_t lxw_col_t;
typedef uint32_t lxw_color_t;

/* ============================================================================
 * Opaque Handle Types (use unsigned long for 32-bit pointers)
 *
 * In LabVIEW, map these to "Unsigned 32-bit Integer" (U32)
 * ============================================================================ */

typedef unsigned long lxw_workbook;
typedef unsigned long lxw_worksheet;
typedef unsigned long lxw_chartsheet;
typedef unsigned long lxw_chart;
typedef unsigned long lxw_chart_series;
typedef unsigned long lxw_chart_axis;
typedef unsigned long lxw_format;
typedef unsigned long lxw_series_error_bars;
typedef unsigned long lxw_styles;
typedef unsigned long lxw_relationships;
typedef unsigned long lxw_drawing;
typedef unsigned long lxw_file_handle;

/* ============================================================================
 * Error Codes
 * ============================================================================ */

typedef enum lxw_error {
    LXW_NO_ERROR = 0,
    LXW_ERROR_MEMORY_MALLOC_FAILED = 1,
    LXW_ERROR_CREATING_XLSX_FILE = 2,
    LXW_ERROR_CREATING_TMPFILE = 3,
    LXW_ERROR_READING_TMPFILE = 4,
    LXW_ERROR_ZIP_FILE_OPERATION = 5,
    LXW_ERROR_ZIP_PARAMETER_ERROR = 6,
    LXW_ERROR_ZIP_BAD_ZIP_FILE = 7,
    LXW_ERROR_ZIP_INTERNAL_ERROR = 8,
    LXW_ERROR_ZIP_FILE_ADD = 9,
    LXW_ERROR_ZIP_CLOSE = 10,
    LXW_ERROR_FEATURE_NOT_SUPPORTED = 11,
    LXW_ERROR_NULL_PARAMETER_IGNORED = 12,
    LXW_ERROR_PARAMETER_VALIDATION = 13,
    LXW_ERROR_PARAMETER_IS_EMPTY = 14,
    LXW_ERROR_SHEETNAME_LENGTH_EXCEEDED = 15,
    LXW_ERROR_INVALID_SHEETNAME_CHARACTER = 16,
    LXW_ERROR_SHEETNAME_START_END_APOSTROPHE = 17,
    LXW_ERROR_SHEETNAME_ALREADY_USED = 18,
    LXW_ERROR_32_STRING_LENGTH_EXCEEDED = 19,
    LXW_ERROR_128_STRING_LENGTH_EXCEEDED = 20,
    LXW_ERROR_255_STRING_LENGTH_EXCEEDED = 21,
    LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED = 22,
    LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND = 23,
    LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE = 24,
    LXW_ERROR_WORKSHEET_MAX_URL_LENGTH_EXCEEDED = 25,
    LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED = 26,
    LXW_ERROR_IMAGE_DIMENSIONS = 27
} lxw_error;

/* ============================================================================
 * Simple Structures (LabVIEW can create clusters for these)
 * ============================================================================ */

typedef struct lxw_datetime {
    int32_t year;
    int32_t month;
    int32_t day;
    int32_t hour;
    int32_t min;
    double sec;
} lxw_datetime;

typedef struct lxw_chart_line {
    lxw_color_t color;
    uint8_t none;
    float width;
    uint8_t dash_type;
    uint8_t transparency;
} lxw_chart_line;

typedef struct lxw_chart_fill {
    lxw_color_t color;
    uint8_t none;
    uint8_t transparency;
} lxw_chart_fill;

typedef struct lxw_chart_pattern {
    lxw_color_t fg_color;
    lxw_color_t bg_color;
    uint8_t type;
} lxw_chart_pattern;

typedef struct lxw_chart_gradient_fill {
    uint8_t type;
    lxw_color_t colors[4];
    uint8_t num_colors;
    double angle;
} lxw_chart_gradient_fill;

typedef struct lxw_chart_layout {
    double x;
    double y;
    double width;
    double height;
} lxw_chart_layout;

typedef struct lxw_chart_font {
    unsigned long name;
    double size;
    uint8_t bold;
    uint8_t italic;
    uint8_t underline;
    int32_t rotation;
    lxw_color_t color;
    uint8_t pitch_family;
    uint8_t charset;
    int8_t baseline;
} lxw_chart_font;

typedef struct lxw_chart_point {
    unsigned long line;
    unsigned long fill;
    unsigned long pattern;
} lxw_chart_point;

typedef struct lxw_image_options {
    int32_t x_offset;
    int32_t y_offset;
    double x_scale;
    double y_scale;
    uint32_t row;
    uint16_t col;
    unsigned long url;
    unsigned long tip;
    uint8_t object_position;
    unsigned long description;
    unsigned long decorative;
} lxw_image_options;

typedef struct lxw_chart_options {
    int32_t x_offset;
    int32_t y_offset;
    double x_scale;
    double y_scale;
    uint8_t object_position;
    unsigned long description;  /* const char* - set to 0 for NULL */
    uint8_t decorative;
} lxw_chart_options;

/* Structure for custom data labels on chart series.
 * For simple labels, only set 'value' (string) and/or 'hide' (uint8_t).
 * The pointer fields (font, line, fill, pattern) should be set to 0 (NULL)
 * unless you need custom formatting for individual labels.
 */
typedef struct lxw_chart_data_label {
    unsigned long value;        /* const char* - label string or formula like "=Sheet1!$C$2" */
    uint8_t hide;               /* Set to 1 to hide this data label */
    unsigned long font;         /* lxw_chart_font* - set to 0 for default */
    unsigned long line;         /* lxw_chart_line* - set to 0 for default */
    unsigned long fill;         /* lxw_chart_fill* - set to 0 for default */
    unsigned long pattern;      /* lxw_chart_pattern* - set to 0 for default */
} lxw_chart_data_label;

typedef struct lxw_row_col_options {
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
} lxw_row_col_options;

typedef struct lxw_protection {
    uint8_t no_select_locked_cells;
    uint8_t no_select_unlocked_cells;
    uint8_t format_cells;
    uint8_t format_columns;
    uint8_t format_rows;
    uint8_t insert_columns;
    uint8_t insert_rows;
    uint8_t insert_hyperlinks;
    uint8_t delete_columns;
    uint8_t delete_rows;
    uint8_t sort;
    uint8_t autofilter;
    uint8_t pivot_tables;
    uint8_t scenarios;
    uint8_t objects;
    uint8_t no_content;
    uint8_t no_objects;
} lxw_protection;

typedef struct lxw_header_footer_options {
    double margin;
    unsigned long image_left;
    unsigned long image_center;
    unsigned long image_right;
} lxw_header_footer_options;

typedef struct lxw_textbox_options {
    uint32_t width;
    uint32_t height;
    int32_t x_offset;
    int32_t y_offset;
    double x_scale;
    double y_scale;
    uint8_t object_position;
    unsigned long description;
    uint8_t decorative;
} lxw_textbox_options;

typedef struct lxw_button_options {
    unsigned long caption;
    unsigned long macro;
    unsigned long description;
    uint32_t width;
    uint32_t height;
    int32_t x_offset;
    int32_t y_offset;
    double x_scale;
    double y_scale;
} lxw_button_options;

typedef struct lxw_workbook_options {
    uint8_t constant_memory;
    unsigned long tmpdir;
    uint8_t use_zip64;
    unsigned long output_buffer;
    unsigned long output_buffer_size;
} lxw_workbook_options;

typedef struct lxw_doc_properties {
    unsigned long title;
    unsigned long subject;
    unsigned long author;
    unsigned long manager;
    unsigned long company;
    unsigned long category;
    unsigned long keywords;
    unsigned long comments;
    unsigned long status;
    unsigned long hyperlink_base;
    int64_t created;
} lxw_doc_properties;

/* ============================================================================
 * Enumeration Types (for LabVIEW custom control generation)
 * ============================================================================ */

typedef enum lxw_chart_type {
    LXW_CHART_NONE = 0,
    LXW_CHART_AREA = 1,
    LXW_CHART_AREA_STACKED = 2,
    LXW_CHART_AREA_STACKED_PERCENT = 3,
    LXW_CHART_BAR = 4,
    LXW_CHART_BAR_STACKED = 5,
    LXW_CHART_BAR_STACKED_PERCENT = 6,
    LXW_CHART_COLUMN = 7,
    LXW_CHART_COLUMN_STACKED = 8,
    LXW_CHART_COLUMN_STACKED_PERCENT = 9,
    LXW_CHART_DOUGHNUT = 10,
    LXW_CHART_LINE = 11,
    LXW_CHART_LINE_STACKED = 12,
    LXW_CHART_LINE_STACKED_PERCENT = 13,
    LXW_CHART_PIE = 14,
    LXW_CHART_SCATTER = 15,
    LXW_CHART_SCATTER_STRAIGHT = 16,
    LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS = 17,
    LXW_CHART_SCATTER_SMOOTH = 18,
    LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS = 19,
    LXW_CHART_RADAR = 20,
    LXW_CHART_RADAR_WITH_MARKERS = 21,
    LXW_CHART_RADAR_FILLED = 22,
    LXW_CHART_STOCK = 23
} lxw_chart_type;

typedef enum lxw_chart_legend_position {
    LXW_CHART_LEGEND_NONE = 0,
    LXW_CHART_LEGEND_RIGHT = 1,
    LXW_CHART_LEGEND_LEFT = 2,
    LXW_CHART_LEGEND_TOP = 3,
    LXW_CHART_LEGEND_BOTTOM = 4,
    LXW_CHART_LEGEND_TOP_RIGHT = 5,
    LXW_CHART_LEGEND_OVERLAY_RIGHT = 6,
    LXW_CHART_LEGEND_OVERLAY_LEFT = 7,
    LXW_CHART_LEGEND_OVERLAY_TOP_RIGHT = 8
} lxw_chart_legend_position;

typedef enum lxw_chart_marker_type {
    LXW_CHART_MARKER_AUTOMATIC = 0,
    LXW_CHART_MARKER_NONE = 1,
    LXW_CHART_MARKER_SQUARE = 2,
    LXW_CHART_MARKER_DIAMOND = 3,
    LXW_CHART_MARKER_TRIANGLE = 4,
    LXW_CHART_MARKER_X = 5,
    LXW_CHART_MARKER_STAR = 6,
    LXW_CHART_MARKER_SHORT_DASH = 7,
    LXW_CHART_MARKER_LONG_DASH = 8,
    LXW_CHART_MARKER_CIRCLE = 9,
    LXW_CHART_MARKER_PLUS = 10,
    LXW_CHART_MARKER_DOT = 11
} lxw_chart_marker_type;

typedef enum lxw_chart_axis_label_position {
    LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO = 0,
    LXW_CHART_AXIS_LABEL_POSITION_HIGH = 1,
    LXW_CHART_AXIS_LABEL_POSITION_LOW = 2,
    LXW_CHART_AXIS_LABEL_POSITION_NONE = 3
} lxw_chart_axis_label_position;

typedef enum lxw_alignment {
    LXW_ALIGN_NONE = 0,
    LXW_ALIGN_LEFT = 1,
    LXW_ALIGN_CENTER = 2,
    LXW_ALIGN_RIGHT = 3,
    LXW_ALIGN_FILL = 4,
    LXW_ALIGN_JUSTIFY = 5,
    LXW_ALIGN_CENTER_ACROSS = 6,
    LXW_ALIGN_DISTRIBUTED = 7,
    LXW_ALIGN_VERTICAL_TOP = 8,
    LXW_ALIGN_VERTICAL_BOTTOM = 9,
    LXW_ALIGN_VERTICAL_CENTER = 10,
    LXW_ALIGN_VERTICAL_JUSTIFY = 11,
    LXW_ALIGN_VERTICAL_DISTRIBUTED = 12
} lxw_alignment;

typedef enum lxw_border_style {
    LXW_BORDER_NONE = 0,
    LXW_BORDER_THIN = 1,
    LXW_BORDER_MEDIUM = 2,
    LXW_BORDER_DASHED = 3,
    LXW_BORDER_DOTTED = 4,
    LXW_BORDER_THICK = 5,
    LXW_BORDER_DOUBLE = 6,
    LXW_BORDER_HAIR = 7,
    LXW_BORDER_MEDIUM_DASHED = 8,
    LXW_BORDER_DASH_DOT = 9,
    LXW_BORDER_MEDIUM_DASH_DOT = 10,
    LXW_BORDER_DASH_DOT_DOT = 11,
    LXW_BORDER_MEDIUM_DASH_DOT_DOT = 12,
    LXW_BORDER_SLANT_DASH_DOT = 13
} lxw_border_style;

typedef enum lxw_diagonal_border_type {
    LXW_DIAGONAL_BORDER_UP = 1,
    LXW_DIAGONAL_BORDER_DOWN = 2,
    LXW_DIAGONAL_BORDER_UP_DOWN = 3
} lxw_diagonal_border_type;

typedef enum lxw_pattern_type {
    LXW_PATTERN_NONE = 0,
    LXW_PATTERN_SOLID = 1,
    LXW_PATTERN_MEDIUM_GRAY = 2,
    LXW_PATTERN_DARK_GRAY = 3,
    LXW_PATTERN_LIGHT_GRAY = 4,
    LXW_PATTERN_DARK_HORIZONTAL = 5,
    LXW_PATTERN_DARK_VERTICAL = 6,
    LXW_PATTERN_DARK_DOWN = 7,
    LXW_PATTERN_DARK_UP = 8,
    LXW_PATTERN_DARK_GRID = 9,
    LXW_PATTERN_DARK_TRELLIS = 10,
    LXW_PATTERN_LIGHT_HORIZONTAL = 11,
    LXW_PATTERN_LIGHT_VERTICAL = 12,
    LXW_PATTERN_LIGHT_DOWN = 13,
    LXW_PATTERN_LIGHT_UP = 14,
    LXW_PATTERN_LIGHT_GRID = 15,
    LXW_PATTERN_LIGHT_TRELLIS = 16,
    LXW_PATTERN_GRAY_125 = 17,
    LXW_PATTERN_GRAY_0625 = 18
} lxw_pattern_type;

typedef enum lxw_chart_gradient_fill_type {
    LXW_CHART_GRADIENT_FILL_LINEAR = 1,
    LXW_CHART_GRADIENT_FILL_RADIAL = 2,
    LXW_CHART_GRADIENT_FILL_RECTANGULAR = 3,
    LXW_CHART_GRADIENT_FILL_PATH = 4
} lxw_chart_gradient_fill_type;

typedef enum lxw_defined_color {
    LXW_COLOR_BLACK = 0x000000,
    LXW_COLOR_NAVY = 0x000080,
    LXW_COLOR_BLUE = 0x0000FF,
    LXW_COLOR_GREEN = 0x008000,
    LXW_COLOR_CYAN = 0x00FFFF,
    LXW_COLOR_LIME = 0x00FF00,
    LXW_COLOR_ORANGE = 0xFF6600,
    LXW_COLOR_BROWN = 0x800000,
    LXW_COLOR_PURPLE = 0x800080,
    LXW_COLOR_GRAY = 0x808080,
    LXW_COLOR_SILVER = 0xC0C0C0,
    LXW_COLOR_RED = 0xFF0000,
    LXW_COLOR_MAGENTA = 0xFF00FF,
    LXW_COLOR_PINK = 0xFF00FF,
    LXW_COLOR_YELLOW = 0xFFFF00,
    LXW_COLOR_WHITE = 0xFFFFFF
} lxw_defined_color;

typedef enum lxw_format_underlines {
    LXW_UNDERLINE_NONE = 0,
    LXW_UNDERLINE_SINGLE = 1,
    LXW_UNDERLINE_DOUBLE = 2,
    LXW_UNDERLINE_SINGLE_ACCOUNTING = 3,
    LXW_UNDERLINE_DOUBLE_ACCOUNTING = 4
} lxw_format_underlines;

typedef enum lxw_format_scripts {
    LXW_FONT_SUPERSCRIPT = 1,
    LXW_FONT_SUBSCRIPT = 2
} lxw_format_scripts;

typedef enum lxw_filter_criteria {
    LXW_FILTER_CRITERIA_NONE = 0,
    LXW_FILTER_CRITERIA_EQUAL_TO = 1,
    LXW_FILTER_CRITERIA_NOT_EQUAL_TO = 2,
    LXW_FILTER_CRITERIA_GREATER_THAN = 3,
    LXW_FILTER_CRITERIA_LESS_THAN = 4,
    LXW_FILTER_CRITERIA_GREATER_THAN_OR_EQUAL_TO = 5,
    LXW_FILTER_CRITERIA_LESS_THAN_OR_EQUAL_TO = 6,
    LXW_FILTER_CRITERIA_BLANKS = 7,
    LXW_FILTER_CRITERIA_NON_BLANKS = 8
} lxw_filter_criteria;

typedef enum lxw_filter_operator {
    LXW_FILTER_AND = 0,
    LXW_FILTER_OR = 1
} lxw_filter_operator;

typedef struct lxw_filter_rule {
    uint8_t criteria;
    const char* value_string;
    double value;
} lxw_filter_rule;

/* ============================================================================
 * Workbook Functions
 * ============================================================================ */

lxw_workbook workbook_new(const char *filename);
lxw_workbook workbook_new_opt(const char *filename, unsigned long options);
lxw_format workbook_add_format(lxw_workbook workbook);
lxw_chart workbook_add_chart(lxw_workbook workbook, uint8_t chart_type);
lxw_error workbook_close(lxw_workbook workbook);
lxw_error workbook_set_properties(lxw_workbook workbook, unsigned long properties);
lxw_error workbook_set_custom_property_number(lxw_workbook workbook, const char *name, double value);
lxw_error workbook_set_custom_property_integer(lxw_workbook workbook, const char *name, int32_t value);
lxw_error workbook_set_custom_property_boolean(lxw_workbook workbook, const char *name, uint8_t value);
lxw_error workbook_set_custom_property_datetime(lxw_workbook workbook, const char *name, lxw_datetime *datetime);
lxw_format workbook_get_default_url_format(lxw_workbook workbook);
lxw_error workbook_add_vba_project(lxw_workbook workbook, const char *filename);
lxw_error workbook_add_signed_vba_project(lxw_workbook workbook, const char *vba_project, const char *signature);
lxw_error workbook_set_vba_name(lxw_workbook workbook, const char *name);
void workbook_read_only_recommended(lxw_workbook workbook);
void workbook_use_1904_epoch(lxw_workbook workbook);
void workbook_set_size(lxw_workbook workbook, uint16_t width, uint16_t height);

/* ============================================================================
 * Worksheet Functions
 * ============================================================================ */

lxw_error worksheet_write_number(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, double number, lxw_format format);
lxw_error worksheet_write_datetime(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, lxw_datetime *datetime, lxw_format format);
lxw_error worksheet_write_unixtime(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, int64_t unixtime, lxw_format format);
lxw_error worksheet_write_boolean(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, int value, lxw_format format);
lxw_error worksheet_write_blank(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, lxw_format format);
lxw_error worksheet_write_rich_string(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, unsigned long rich_string, lxw_format format);

lxw_error worksheet_set_row(lxw_worksheet worksheet, lxw_row_t row, double height, lxw_format format);
lxw_error worksheet_set_row_opt(lxw_worksheet worksheet, lxw_row_t row, double height, lxw_format format, lxw_row_col_options *options);
lxw_error worksheet_set_row_pixels(lxw_worksheet worksheet, lxw_row_t row, uint32_t pixels, lxw_format format);
lxw_error worksheet_set_row_pixels_opt(lxw_worksheet worksheet, lxw_row_t row, uint32_t pixels, lxw_format format, lxw_row_col_options *options);
lxw_error worksheet_set_column(lxw_worksheet worksheet, lxw_col_t first_col, lxw_col_t last_col, double width, lxw_format format);
lxw_error worksheet_set_column_opt(lxw_worksheet worksheet, lxw_col_t first_col, lxw_col_t last_col, double width, lxw_format format, lxw_row_col_options *options);
lxw_error worksheet_set_column_pixels(lxw_worksheet worksheet, lxw_col_t first_col, lxw_col_t last_col, uint32_t pixels, lxw_format format);
lxw_error worksheet_set_column_pixels_opt(lxw_worksheet worksheet, lxw_col_t first_col, lxw_col_t last_col, uint32_t pixels, lxw_format format, lxw_row_col_options *options);

lxw_error worksheet_insert_image(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *filename);
lxw_error worksheet_insert_image_opt(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *filename, lxw_image_options *options);
lxw_error worksheet_insert_image_buffer(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const unsigned char *image_buffer, size_t image_size);
lxw_error worksheet_insert_image_buffer_opt(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const unsigned char *image_buffer, size_t image_size, lxw_image_options *options);
lxw_error worksheet_insert_chart(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, lxw_chart chart);
lxw_error worksheet_insert_chart_opt(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, lxw_chart chart, lxw_chart_options *options);
lxw_error worksheet_insert_checkbox(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, uint8_t checked);
lxw_error worksheet_insert_button(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, lxw_button_options *options);
lxw_error worksheet_embed_image(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *filename);
lxw_error worksheet_embed_image_opt(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *filename, lxw_image_options *options);
lxw_error worksheet_embed_image_buffer(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const unsigned char *image_buffer, size_t image_size);
lxw_error worksheet_embed_image_buffer_opt(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const unsigned char *image_buffer, size_t image_size, lxw_image_options *options);
lxw_error worksheet_add_table(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, unsigned long options);

lxw_error worksheet_autofilter(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);

lxw_error worksheet_data_validation_cell(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, unsigned long validation);
lxw_error worksheet_data_validation_range(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, unsigned long validation);
lxw_error worksheet_conditional_format_cell(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, unsigned long conditional_format);
lxw_error worksheet_conditional_format_range(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, unsigned long conditional_format);

void worksheet_activate(lxw_worksheet worksheet);
void worksheet_select(lxw_worksheet worksheet);
void worksheet_hide(lxw_worksheet worksheet);
void worksheet_set_first_sheet(lxw_worksheet worksheet);
void worksheet_freeze_panes(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col);
void worksheet_split_panes(lxw_worksheet worksheet, double vertical, double horizontal);
void worksheet_freeze_panes_opt(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t top_row, lxw_col_t left_col, uint8_t type);
void worksheet_split_panes_opt(lxw_worksheet worksheet, double vertical, double horizontal, lxw_row_t top_row, lxw_col_t left_col);

lxw_error worksheet_filter_column_lv(lxw_worksheet worksheet, lxw_col_t col, uint8_t criteria, const char *value_string, double value);
lxw_error worksheet_filter_column2_lv(lxw_worksheet worksheet, lxw_col_t col, uint8_t criteria1, const char *value_string1, double value1, uint8_t criteria2, const char *value_string2, double value2, uint8_t and_or);


lxw_error worksheet_set_selection(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void worksheet_set_top_left_cell(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col);
void worksheet_set_landscape(lxw_worksheet worksheet);
void worksheet_set_portrait(lxw_worksheet worksheet);
void worksheet_set_page_view(lxw_worksheet worksheet);
void worksheet_set_paper(lxw_worksheet worksheet, uint8_t paper_type);
void worksheet_set_margins(lxw_worksheet worksheet, double left, double right, double top, double bottom);
lxw_error worksheet_set_h_pagebreaks(lxw_worksheet worksheet, unsigned long breaks);
lxw_error worksheet_set_v_pagebreaks(lxw_worksheet worksheet, unsigned long breaks);
void worksheet_print_across(lxw_worksheet worksheet);
void worksheet_set_zoom(lxw_worksheet worksheet, uint16_t scale);
void worksheet_gridlines(lxw_worksheet worksheet, uint8_t option);
void worksheet_center_horizontally(lxw_worksheet worksheet);
void worksheet_center_vertically(lxw_worksheet worksheet);
void worksheet_print_row_col_headers(lxw_worksheet worksheet);
lxw_error worksheet_repeat_rows(lxw_worksheet worksheet, lxw_row_t first_row, lxw_row_t last_row);
lxw_error worksheet_repeat_columns(lxw_worksheet worksheet, lxw_col_t first_col, lxw_col_t last_col);
lxw_error worksheet_print_area(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void worksheet_fit_to_pages(lxw_worksheet worksheet, uint16_t width, uint16_t height);
void worksheet_set_start_page(lxw_worksheet worksheet, uint16_t start_page);
void worksheet_set_print_scale(lxw_worksheet worksheet, uint16_t scale);
void worksheet_right_to_left(lxw_worksheet worksheet);
void worksheet_hide_zero(lxw_worksheet worksheet);
void worksheet_set_tab_color(lxw_worksheet worksheet, lxw_color_t color);
void worksheet_protect(lxw_worksheet worksheet, const char *password, lxw_protection *options);
void worksheet_outline_settings(lxw_worksheet worksheet, uint8_t visible, uint8_t symbols_below, uint8_t symbols_right, uint8_t auto_style);
void worksheet_set_default_row(lxw_worksheet worksheet, double height, uint8_t hide_unused_rows);
lxw_error worksheet_set_vba_name(lxw_worksheet worksheet, const char *name);
void worksheet_show_comments(lxw_worksheet worksheet);
lxw_error worksheet_ignore_errors(lxw_worksheet worksheet, uint8_t type, const char *range);
lxw_error worksheet_set_background(lxw_worksheet worksheet, const char *filename);
lxw_error worksheet_set_background_buffer(lxw_worksheet worksheet, const unsigned char *image_buffer, size_t image_size);
void worksheet_print_black_and_white(lxw_worksheet worksheet);
lxw_error worksheet_set_header_opt(lxw_worksheet worksheet, const char *string, lxw_header_footer_options *options);
lxw_error worksheet_set_footer_opt(lxw_worksheet worksheet, const char *string, lxw_header_footer_options *options);
lxw_error worksheet_write_array_formula(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *formula, lxw_format format);
lxw_error worksheet_write_dynamic_array_formula(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *formula, lxw_format format);
lxw_error worksheet_write_dynamic_formula(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format format);
lxw_error worksheet_write_formula_num(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format format, double result);
lxw_error worksheet_write_formula_str(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format format, const char *result);
lxw_error worksheet_write_url_opt(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *url, lxw_format format, const char *string, const char *tooltip);
lxw_error worksheet_write_comment_opt(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *string, unsigned long options);
lxw_error worksheet_write_array_formula_num(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *formula, lxw_format format, double result);
lxw_error worksheet_write_dynamic_array_formula_num(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *formula, lxw_format format, double result);
lxw_error worksheet_write_dynamic_formula_num(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format format, double result);
void worksheet_set_error_cell(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col);

/* ============================================================================
 * Chartsheet Functions
 * ============================================================================ */

lxw_error chartsheet_set_chart(lxw_chartsheet chartsheet, lxw_chart chart);
lxw_error chartsheet_set_chart_opt(lxw_chartsheet chartsheet, lxw_chart chart, lxw_chart_options *options);
void chartsheet_activate(lxw_chartsheet chartsheet);
void chartsheet_select(lxw_chartsheet chartsheet);
void chartsheet_hide(lxw_chartsheet chartsheet);
void chartsheet_set_first_sheet(lxw_chartsheet chartsheet);
void chartsheet_set_tab_color(lxw_chartsheet chartsheet, lxw_color_t color);
void chartsheet_protect(lxw_chartsheet chartsheet, const char *password, lxw_protection *options);
void chartsheet_set_zoom(lxw_chartsheet chartsheet, uint16_t scale);
void chartsheet_set_landscape(lxw_chartsheet chartsheet);
void chartsheet_set_portrait(lxw_chartsheet chartsheet);
void chartsheet_set_paper(lxw_chartsheet chartsheet, uint8_t paper_type);
void chartsheet_set_margins(lxw_chartsheet chartsheet, double left, double right, double top, double bottom);
lxw_error chartsheet_set_header_opt(lxw_chartsheet chartsheet, const char *string, lxw_header_footer_options *options);
lxw_error chartsheet_set_footer_opt(lxw_chartsheet chartsheet, const char *string, lxw_header_footer_options *options);

/* ============================================================================
 * Chart Functions
 * ============================================================================ */

lxw_chart_series chart_add_series_impl(lxw_chart chart, const char *categories, const char *values, uint8_t y2_axis);
void chart_series_set_line(lxw_chart_series series, lxw_chart_line *line);
void chart_series_set_fill(lxw_chart_series series, lxw_chart_fill *fill);
void chart_series_set_invert_if_negative(lxw_chart_series series);
void chart_series_set_pattern(lxw_chart_series series, lxw_chart_pattern *pattern);
void chart_series_set_gradient(lxw_chart_series series, lxw_chart_gradient_fill *gradient);
void chart_series_set_marker_type(lxw_chart_series series, uint8_t type);
void chart_series_set_marker_size(lxw_chart_series series, uint8_t size);
void chart_series_set_marker_line(lxw_chart_series series, lxw_chart_line *line);
void chart_series_set_marker_fill(lxw_chart_series series, lxw_chart_fill *fill);
void chart_series_set_marker_pattern(lxw_chart_series series, lxw_chart_pattern *pattern);
lxw_error chart_series_set_points(lxw_chart_series series, unsigned long points);
void chart_series_set_smooth(lxw_chart_series series, uint8_t smooth);
void chart_series_set_labels(lxw_chart_series series);
void chart_series_set_labels_options(lxw_chart_series series, uint8_t show_name, uint8_t show_category, uint8_t show_value);
lxw_error chart_series_set_labels_custom(lxw_chart_series series, unsigned long data_labels);
void chart_series_set_labels_separator(lxw_chart_series series, uint8_t separator);
void chart_series_set_labels_position(lxw_chart_series series, uint8_t position);
void chart_series_set_labels_leader_line(lxw_chart_series series);
void chart_series_set_labels_legend(lxw_chart_series series);
void chart_series_set_labels_percentage(lxw_chart_series series);
void chart_series_set_labels_font(lxw_chart_series series, lxw_chart_font *font);
void chart_series_set_labels_line(lxw_chart_series series, lxw_chart_line *line);
void chart_series_set_labels_fill(lxw_chart_series series, lxw_chart_fill *fill);
void chart_series_set_labels_pattern(lxw_chart_series series, lxw_chart_pattern *pattern);
void chart_series_set_trendline(lxw_chart_series series, uint8_t type, uint8_t value);
void chart_series_set_trendline_forecast(lxw_chart_series series, double forward, double backward);
void chart_series_set_trendline_equation(lxw_chart_series series);
void chart_series_set_trendline_r_squared(lxw_chart_series series);
void chart_series_set_trendline_intercept(lxw_chart_series series, double intercept);
void chart_series_set_trendline_line(lxw_chart_series series, lxw_chart_line *line);
lxw_series_error_bars chart_series_get_error_bars(lxw_chart_series series, uint8_t axis_type);
void chart_series_set_error_bars(lxw_series_error_bars error_bars, uint8_t type, double value);
void chart_series_set_error_bars_direction(lxw_series_error_bars error_bars, uint8_t direction);
void chart_series_set_error_bars_endcap(lxw_series_error_bars error_bars, uint8_t endcap);
void chart_series_set_error_bars_line(lxw_series_error_bars error_bars, lxw_chart_line *line);

lxw_chart_axis chart_axis_get(lxw_chart chart, uint8_t axis_type);
void chart_axis_set_name_layout(lxw_chart_axis axis, lxw_chart_layout *layout);
void chart_axis_set_name_font(lxw_chart_axis axis, lxw_chart_font *font);
void chart_axis_set_num_font(lxw_chart_axis axis, lxw_chart_font *font);
void chart_axis_set_line(lxw_chart_axis axis, lxw_chart_line *line);
void chart_axis_set_fill(lxw_chart_axis axis, lxw_chart_fill *fill);
void chart_axis_set_pattern(lxw_chart_axis axis, lxw_chart_pattern *pattern);
void chart_axis_set_reverse(lxw_chart_axis axis);
void chart_axis_set_crossing(lxw_chart_axis axis, double value);
void chart_axis_set_crossing_max(lxw_chart_axis axis);
void chart_axis_set_crossing_min(lxw_chart_axis axis);
void chart_axis_off(lxw_chart_axis axis);
void chart_axis_set_position(lxw_chart_axis axis, uint8_t position);
void chart_axis_set_label_position(lxw_chart_axis axis, uint8_t position);
void chart_axis_set_label_align(lxw_chart_axis axis, uint8_t align);
void chart_axis_set_min(lxw_chart_axis axis, double min);
void chart_axis_set_max(lxw_chart_axis axis, double max);
void chart_axis_set_log_base(lxw_chart_axis axis, uint16_t log_base);
void chart_axis_set_major_tick_mark(lxw_chart_axis axis, uint8_t type);
void chart_axis_set_minor_tick_mark(lxw_chart_axis axis, uint8_t type);
void chart_axis_set_interval_unit(lxw_chart_axis axis, uint16_t unit);
void chart_axis_set_interval_tick(lxw_chart_axis axis, uint16_t unit);
void chart_axis_set_major_unit(lxw_chart_axis axis, double unit);
void chart_axis_set_minor_unit(lxw_chart_axis axis, double unit);
void chart_axis_set_display_units(lxw_chart_axis axis, uint8_t units);
void chart_axis_set_display_units_visible(lxw_chart_axis axis, uint8_t visible);
void chart_axis_major_gridlines_set_visible(lxw_chart_axis axis, uint8_t visible);
void chart_axis_minor_gridlines_set_visible(lxw_chart_axis axis, uint8_t visible);
void chart_axis_major_gridlines_set_line(lxw_chart_axis axis, lxw_chart_line *line);
void chart_axis_minor_gridlines_set_line(lxw_chart_axis axis, lxw_chart_line *line);

void chart_title_set_name_font(lxw_chart chart, lxw_chart_font *font);
void chart_title_off(lxw_chart chart);
void chart_legend_set_position(lxw_chart chart, uint8_t position);
void chart_legend_set_font(lxw_chart chart, lxw_chart_font *font);
lxw_error chart_legend_delete_series(lxw_chart chart, unsigned long delete_series);
void chart_chartarea_set_line(lxw_chart chart, lxw_chart_line *line);
void chart_chartarea_set_fill(lxw_chart chart, lxw_chart_fill *fill);
void chart_chartarea_set_pattern(lxw_chart chart, lxw_chart_pattern *pattern);
void chart_chartarea_set_gradient(lxw_chart chart, lxw_chart_gradient_fill *gradient);
void chart_plotarea_set_line(lxw_chart chart, lxw_chart_line *line);
void chart_plotarea_set_fill(lxw_chart chart, lxw_chart_fill *fill);
void chart_plotarea_set_pattern(lxw_chart chart, lxw_chart_pattern *pattern);
void chart_plotarea_set_gradient(lxw_chart chart, lxw_chart_gradient_fill *gradient);
void chart_plotarea_set_layout(lxw_chart chart, lxw_chart_layout *layout);
void chart_combine(lxw_chart chart, lxw_chart combined_chart);
void chart_title_set_layout(lxw_chart chart, lxw_chart_layout *layout);
void chart_title_set_overlay(lxw_chart chart, uint8_t overlay);
void chart_legend_set_layout(lxw_chart chart, lxw_chart_layout *layout);
void chart_set_style(lxw_chart chart, uint8_t style_id);
void chart_set_table(lxw_chart chart);
void chart_set_table_grid(lxw_chart chart, uint8_t horizontal, uint8_t vertical, uint8_t outline, uint8_t legend_keys);
void chart_set_table_font(lxw_chart chart, lxw_chart_font *font);
void chart_set_up_down_bars(lxw_chart chart);
void chart_set_up_down_bars_format(lxw_chart chart, lxw_chart_line *up_bar_line, lxw_chart_fill *up_bar_fill, lxw_chart_line *down_bar_line, lxw_chart_fill *down_bar_fill);
void chart_set_drop_lines(lxw_chart chart, lxw_chart_line *line);
void chart_set_high_low_lines(lxw_chart chart, lxw_chart_line *line);
void chart_set_series_overlap(lxw_chart chart, int8_t overlap);
void chart_set_series_gap(lxw_chart chart, uint16_t gap);
void chart_set_series_overlap_y2(lxw_chart chart, int8_t overlap);
void chart_set_series_gap_y2(lxw_chart chart, uint16_t gap);
void chart_show_blanks_as(lxw_chart chart, uint8_t option);
void chart_show_hidden_data(lxw_chart chart);
void chart_set_rotation(lxw_chart chart, uint16_t rotation);
void chart_set_hole_size(lxw_chart chart, uint8_t size);

/* Axis accessor functions for LabVIEW */
lxw_chart_axis chart_get_x_axis(lxw_chart chart);
lxw_chart_axis chart_get_y_axis(lxw_chart chart);
lxw_chart_axis chart_get_y2_axis(lxw_chart chart);

/* ============================================================================
 * Format Functions
 * ============================================================================ */

void format_set_font_size(lxw_format format, double size);
void format_set_font_color(lxw_format format, lxw_color_t color);
void format_set_bold(lxw_format format);
void format_set_italic(lxw_format format);
void format_set_underline(lxw_format format, uint8_t style);
void format_set_font_strikeout(lxw_format format);
void format_set_font_script(lxw_format format, uint8_t style);
void format_set_num_format_index(lxw_format format, uint8_t index);
void format_set_unlocked(lxw_format format);
void format_set_hidden(lxw_format format);
void format_set_align(lxw_format format, uint8_t alignment);
void format_set_text_wrap(lxw_format format);
void format_set_rotation(lxw_format format, int16_t angle);
void format_set_indent(lxw_format format, uint8_t level);
void format_set_shrink(lxw_format format);
void format_set_pattern(lxw_format format, uint8_t pattern);
void format_set_bg_color(lxw_format format, lxw_color_t color);
void format_set_fg_color(lxw_format format, lxw_color_t color);
void format_set_border(lxw_format format, uint8_t style);
void format_set_bottom(lxw_format format, uint8_t style);
void format_set_top(lxw_format format, uint8_t style);
void format_set_left(lxw_format format, uint8_t style);
void format_set_right(lxw_format format, uint8_t style);
void format_set_border_color(lxw_format format, lxw_color_t color);
void format_set_bottom_color(lxw_format format, lxw_color_t color);
void format_set_top_color(lxw_format format, lxw_color_t color);
void format_set_left_color(lxw_format format, lxw_color_t color);
void format_set_right_color(lxw_format format, lxw_color_t color);
void format_set_diag_type(lxw_format format, uint8_t value);
void format_set_diag_border(lxw_format format, uint8_t value);
void format_set_diag_color(lxw_format format, lxw_color_t color);
void format_set_font_outline(lxw_format format);
void format_set_font_shadow(lxw_format format);
void format_set_font_family(lxw_format format, uint8_t value);
void format_set_font_charset(lxw_format format, uint8_t value);
void format_set_font_scheme(lxw_format format, const char *font_scheme);
void format_set_font_condense(lxw_format format);
void format_set_font_extend(lxw_format format);
void format_set_reading_order(lxw_format format, uint8_t value);
void format_set_theme(lxw_format format, uint8_t value);
void format_set_hyperlink(lxw_format format);
void format_set_color_indexed(lxw_format format, uint8_t value);
void format_set_font_only(lxw_format format);
void format_set_quote_prefix(lxw_format format);
void format_set_checkbox(lxw_format format);

/* ============================================================================
 * Cell/Range Reference Structures (for LabVIEW-compatible functions)
 * ============================================================================ */

typedef struct lxw_cell_ref {
    lxw_row_t row;
    lxw_col_t col;
} lxw_cell_ref;

typedef struct lxw_col_range {
    lxw_col_t first_col;
    lxw_col_t last_col;
} lxw_col_range;

typedef struct lxw_range_ref {
    lxw_row_t first_row;
    lxw_col_t first_col;
    lxw_row_t last_row;
    lxw_col_t last_col;
} lxw_range_ref;

/* ============================================================================
 * Cell/Range Reference Functions (LabVIEW-compatible alternatives to macros)
 * ============================================================================ */

/**
 * Convert an Excel `A1` cell string into row and col values.
 * Example: lxw_parse_cell("A1", &row, &col);
 */
void lxw_parse_cell(const char *cell_str, lxw_row_t *row, lxw_col_t *col);

/**
 * Convert an Excel `A:B` column range into first_col and last_col values.
 * Example: lxw_parse_cols("B:D", &first_col, &last_col);
 */
void lxw_parse_cols(const char *cols_str, lxw_col_t *first_col, lxw_col_t *last_col);

/**
 * Convert an Excel `A1:B2` range into row/col values.
 * Example: lxw_parse_range("A1:K42", &first_row, &first_col, &last_row, &last_col);
 */
void lxw_parse_range(const char *range_str, lxw_row_t *first_row, lxw_col_t *first_col, lxw_row_t *last_row, lxw_col_t *last_col);

/* ============================================================================
 * Utility Functions
 * ============================================================================ */

uint32_t lxw_name_to_row(const char *row_str);
uint16_t lxw_name_to_col(const char *col_str);
uint32_t lxw_name_to_row_2(const char *row_str);
uint16_t lxw_name_to_col_2(const char *col_str);

const char* lxw_version(void);
uint16_t lxw_version_id(void);
char* lxw_strerror(lxw_error error_num);
double lxw_datetime_to_excel_datetime(lxw_datetime *datetime);
int32_t lxw_unixtime_to_excel_date(int64_t unixtime);
double lxw_unixtime_to_excel_date_epoch(int64_t unixtime, uint8_t is_date_1904);

/* ============================================================================
 * LabVIEW Wrapper Functions with ANSI to UTF-8 conversion
 *
 * These functions automatically convert ANSI-encoded strings to UTF-8.
 * Use these instead of the standard functions when passing strings from LabVIEW.
 * File name functions are NOT included - those should use the standard versions.
 * ============================================================================ */

/* Worksheet write functions (with ANSI to UTF-8 conversion) */
lxw_error worksheet_write_string_lv(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *string, lxw_format format);
lxw_error worksheet_write_formula_lv(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format format);
lxw_error worksheet_write_url_lv(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *url, lxw_format format);
lxw_error worksheet_write_comment_lv(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *string);
lxw_error worksheet_set_header_lv(lxw_worksheet worksheet, const char *header);
lxw_error worksheet_set_footer_lv(lxw_worksheet worksheet, const char *footer);
lxw_error worksheet_merge_range_lv(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *string, lxw_format format);

/* Worksheet write functions (no conversion - use when string is already UTF-8) */
lxw_error worksheet_write_string(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *string, lxw_format format);
lxw_error worksheet_write_formula(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *formula, lxw_format format);
lxw_error worksheet_write_rich_string(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, unsigned long rich_strings, lxw_format format);
lxw_error worksheet_write_url(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *url, lxw_format format);
lxw_error worksheet_write_comment(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *string);
lxw_error worksheet_set_header(lxw_worksheet worksheet, const char *header);
lxw_error worksheet_set_footer(lxw_worksheet worksheet, const char *footer);
lxw_error worksheet_merge_range(lxw_worksheet worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, const char *string, lxw_format format);

/* Chart functions */
lxw_chart_series chart_add_series_lv(lxw_chart chart, const char *categories, const char *values, uint8_t y2_axis);
void chart_series_set_name_lv(lxw_chart_series series, const char *name);
void chart_axis_set_name_lv(lxw_chart_axis axis, const char *name);
void chart_title_set_name_lv(lxw_chart chart, const char *name);

/* Format functions */
void format_set_font_name_lv(lxw_format format, const char *font_name);
void format_set_num_format_lv(lxw_format format, const char *num_format);

/* Workbook functions (except filenames) */
lxw_worksheet workbook_add_worksheet_lv(lxw_workbook workbook, const char *sheetname);
lxw_chartsheet workbook_add_chartsheet_lv(lxw_workbook workbook, const char *sheetname);
lxw_error workbook_define_name_lv(lxw_workbook workbook, const char *name, const char *formula);
lxw_worksheet workbook_get_worksheet_by_name_lv(lxw_workbook workbook, const char *name);
lxw_chartsheet workbook_get_chartsheet_by_name_lv(lxw_workbook workbook, const char *name);
lxw_error workbook_validate_sheet_name_lv(lxw_workbook workbook, const char *sheetname);
lxw_error workbook_set_custom_property_string_lv(lxw_workbook workbook, const char *name, const char *value);

/* Worksheet functions */
void worksheet_set_comments_author_lv(lxw_worksheet worksheet, const char *author);
lxw_error worksheet_insert_textbox_lv(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *text);
lxw_error worksheet_insert_textbox_opt_lv(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *text, lxw_textbox_options *options);

/* Chartsheet functions */
lxw_error chartsheet_set_header_lv(lxw_chartsheet chartsheet, const char *header);
lxw_error chartsheet_set_footer_lv(lxw_chartsheet chartsheet, const char *footer);

/* Additional chart functions */
void chart_series_set_trendline_name_lv(lxw_chart_series series, const char *name);
void chart_axis_set_num_format_lv(lxw_chart_axis axis, const char *num_format);
void chart_series_set_labels_num_format_lv(lxw_chart_series series, const char *num_format);

/* Custom data labels wrapper - simplified version for LabVIEW.
 * Takes separate arrays for values and hide flags, plus a count.
 *
 * Parameters:
 *   series      - The chart series from chart_add_series
 *   values      - Array of pointer-sized integers containing string addresses
 *                 (from LabVIEW MoveBlock). Use 0 for default label.
 *   hide_flags  - Array of uint8_t (1 = hide label, 0 = show). Can be NULL.
 *   count       - Number of elements in the arrays
 *
 * In LabVIEW:
 *   1. For each string, use MoveBlock to get the string pointer
 *   2. Build an array of uintptr_t (U32 on 32-bit, U64 on 64-bit) with these pointers
 *   3. Pass this array as 'values'
 *   4. Build a U8 array for hide_flags (or pass NULL/0 if not hiding any)
 */
lxw_error chart_series_set_labels_custom_lv(lxw_chart_series series,
                                            uintptr_t *values,
                                            uint8_t *hide_flags,
                                            uint16_t count);

/* Chart functions with sheetname parameters */
void chart_series_set_categories_lv(lxw_chart_series series, const char *sheetname, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void chart_series_set_values_lv(lxw_chart_series series, const char *sheetname, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
void chart_series_set_name_range_lv(lxw_chart_series series, const char *sheetname, lxw_row_t row, lxw_col_t col);
void chart_axis_set_name_range_lv(lxw_chart_axis axis, const char *sheetname, lxw_row_t row, lxw_col_t col);
void chart_title_set_name_range_lv(lxw_chart chart, const char *sheetname, lxw_row_t row, lxw_col_t col);

/* File path functions (ANSI to UTF-8 conversion for file operations) */
lxw_workbook workbook_new_lv(const char *filename);
lxw_workbook workbook_new_opt_lv(const char *filename, lxw_workbook_options *options);
lxw_error worksheet_insert_image_lv(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *filename);
lxw_error worksheet_insert_image_opt_lv(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *filename, lxw_image_options *options);
lxw_error worksheet_embed_image_lv(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *filename);
lxw_error worksheet_embed_image_opt_lv(lxw_worksheet worksheet, lxw_row_t row, lxw_col_t col, const char *filename, lxw_image_options *options);
lxw_error worksheet_set_background_lv(lxw_worksheet worksheet, const char *filename);
lxw_error workbook_add_vba_project_lv(lxw_workbook workbook, const char *filename);
lxw_error workbook_add_signed_vba_project_lv(lxw_workbook workbook, const char *vba_project, const char *signature);

#endif /* __LIBXLSXWRITER_LV_H__ */
