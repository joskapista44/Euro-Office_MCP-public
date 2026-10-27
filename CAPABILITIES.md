# Capabilities -- Euro-Office_MCP 1.1.0

The accepted read operations and write intents of this release. `tools/list` returns the same sets in each tool's
input schema; a test (`test/static/public-capabilities-doc.test.cjs`) fails if this file and the source lists differ.
Anything not listed is refused before any editor call.

Every write runs as the calling requester, is persisted with a confirmed save and verified by reading back; a step
that cannot be proven is refused with a named reason, not guessed.

## `xlsx_read` -- 16 read operations

- `workbook_overview`
- `read_range`
- `read_validations`
- `read_charts`
- `read_pivots`
- `read_drawings`
- `read_conditional_formats`
- `read_layout`
- `read_freeze_panes`
- `read_sheet_state`
- `read_page_setup`
- `read_filter`
- `read_comments`
- `read_print_area`
- `read_defined_names`
- `read_runs`

Not offered: `read_print_rendering` (NOT_MEASURED).

## `docx_read` -- 12 read operations

- `read_text`
- `read_formatting`
- `read_tables`
- `read_headers_footers`
- `read_page_setup`
- `read_bookmarks`
- `read_numbering`
- `read_styles`
- `read_notes`
- `read_images`
- `read_comments`
- `read_section_flags`

Not offered: `read_table_grid` (PARTIAL).

## `xlsx_execute_task` -- 62 write intents

- `write_range`
- `insert_rows`
- `delete_rows`
- `rename_sheet`
- `unmerge_range`
- `set_chart_title`
- `format_range`
- `merge_range`
- `sort_range`
- `add_conditional_format`
- `set_print_area`
- `layout_range`
- `freeze_panes`
- `unfreeze_panes`
- `set_validation`
- `clear_validation`
- `delete_conditional_format`
- `clear_filter`
- `clear_print_area`
- `create_sheet`
- `rename_chart`
- `set_chart_axis_title`
- `clear_range`
- `set_chart_style`
- `insert_columns`
- `delete_columns`
- `delete_sheet`
- `move_sheet`
- `set_chart`
- `set_chart_legend`
- `copy_range`
- `copy_chart`
- `copy_sheet`
- `move_range`
- `delete_chart`
- `filter_range`
- `set_defined_name`
- `delete_defined_name`
- `delete_cells`
- `set_print_titles`
- `clear_print_titles`
- `set_page_layout`
- `move_chart`
- `insert_shape`
- `insert_image`
- `format_run`
- `set_print_fit`
- `set_header_footer`
- `set_page_break`
- `set_first_page_number`
- `set_print_centering`
- `set_print_scale`
- `set_paper_size`
- `set_header_footer_flag`
- `refresh_pivot`
- `create_pivot`
- `write_array_formula`
- `add_comment`
- `set_shape_text`
- `set_chart_point_data_label`
- `set_shape_rotation`
- `set_sheet_visibility`

## `docx_execute_task` -- 38 write intents

- `rewrite_paragraph`
- `insert_paragraph_before`
- `insert_paragraph_after`
- `delete_paragraph`
- `move_paragraph_before`
- `move_paragraph_after`
- `apply_list`
- `set_list_level`
- `indent_list_item`
- `outdent_list_item`
- `remove_list`
- `replace_table_cell`
- `format_text`
- `insert_table_row_before`
- `insert_table_row_after`
- `delete_table_row`
- `insert_table_column_before`
- `insert_table_column_after`
- `delete_table_column`
- `merge_table_cells`
- `split_table_cell`
- `format_table`
- `set_page_setup`
- `edit_header`
- `edit_footer`
- `define_style`
- `modify_style`
- `set_numbering`
- `set_list_type`
- `add_footnote`
- `add_endnote`
- `add_link`
- `add_bookmark`
- `resolve_comment`
- `reply_comment`
- `add_comment_on_match`
- `set_image_size`
- `set_image_wrap`
