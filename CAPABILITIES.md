# Capabilities -- Euro-Office_MCP 1.1.0

The accepted read operations and write intents of this release. `tools/list` returns the same sets in each tool's
input schema; a test (`test/static/public-capabilities-doc.test.cjs`) fails if this file and the source lists differ.
Anything not listed is refused before any editor call.

Every write runs as the calling requester, is persisted with a confirmed save and verified by reading back; a step
that cannot be proven is refused with a named reason, not guessed.

## `xlsx_read` -- 22 read operations

- `workbook_overview`
- `read_range`
- `read_validations`
- `read_charts`
- `read_pivots`
- `read_drawings`
- `read_conditional_formats`
- `read_effective_fill`
- `read_layout`
- `read_freeze_panes`
- `read_sheet_state`
- `read_page_setup`
- `read_filter`
- `read_comments`
- `read_tables`
- `read_hyperlinks`
- `read_protection`
- `read_print_area`
- `read_defined_names`
- `read_runs`
- `read_array_formulas`
- `read_print_rendering`

`read_charts`: a chart's `dataLabels` is the chart-type group setting only (a series' own labels override it); `dataLabelsEffective` is the state the points actually show, read per point (one four-flag state, `mixed`, or a named `UNOBSERVABLE:<reason>`).

`read_array_formulas` reports each array formula's anchor, ref and formula for a range; an array that crosses the requested range is reported as such, without a ref, and two adjacent arrays with the same formula are not told apart.

`read_filter` reports a colour filter's kind (cell or font colour) and, on the Document Server Bridge `/21`, its colour; a colour filter cannot be written. An empty value list is reported as ambiguous (nothing visible, blanks only or dates only read the same).

`read_protection` reports whether a sheet is protected and its 15 allowed / forbidden operations, and with a range the locked / hidden state of the range's non-empty cells. It never reports a password.

`read_effective_fill` reports, for up to 200 named cells in a range, the conditional-format-evaluated fill colour actually shown (not the rule) -- `read_conditional_formats` reports the rule, never the evaluated per-cell result. Only the fill is read; every other evaluated style field is not reported.

`read_print_rendering` reports how many pages the editor lays one sheet out on for printing under its stored print settings, on the Document Server Bridge `/23`; a hidden sheet and a sheet over the 1500-page limit are reported as named gaps, never as a number, and reading never changes the document.

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

## `xlsx_execute_task` -- 87 write intents

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
- `edit_comment`
- `delete_comment`
- `reply_to_comment`
- `set_comment_resolved`
- `set_shape_text`
- `set_chart_point_data_label`
- `set_shape_rotation`
- `set_sheet_visibility`
- `set_hyperlink`
- `remove_hyperlink`
- `create_table`
- `insert_table_rows`
- `delete_table_rows`
- `insert_table_columns`
- `delete_table_columns`
- `set_table_style`
- `set_table_flag`
- `set_table_totals_row`
- `rename_table`
- `resize_table`
- `set_header_footer_margin`
- `set_page_order`
- `protect_sheet`
- `unprotect_sheet`
- `set_cell_protection`
- `protect_workbook`
- `unprotect_workbook`
- `set_autofilter_range`
- `remove_autofilter`

`protect_sheet`, `unprotect_sheet` and `set_cell_protection` are passwordless: a sheet protected with a password is refused, never changed. `set_cell_protection` is refused on a protected sheet (unprotect it first). They need the Document Server Bridge `/22`.

`protect_workbook` and `unprotect_workbook` lock or unlock the whole workbook's structure (passwordless), through Bridge `/24`. `unprotect_workbook` on a currently locked workbook is always refused, never changed, until a future release can tell a passwordless lock apart from a password-protected one.

`set_hyperlink` and `remove_hyperlink` act on one cell (A1). A multi-cell range is refused before anything is written, with the reason `xlsx-hyperlink-range-unsupported` (the engine does not apply a hyperlink to a range).

`set_autofilter_range` turns on a range's AutoFilter buttons with no active criteria yet. If the sheet already has the buttons on the same range with no active criteria, it is a no-op; with active criteria it is refused (`xlsx-autofilter-already-present`), never silently cleared. `remove_autofilter` takes the AutoFilter off the sheet entirely -- stronger than clearing an active filter's criteria -- and is a no-op (not a refusal) when there is none to remove. Both are refused (`xlsx-filter-range-conflict`) when the sheet's existing AutoFilter is on a different range.

## `docx_execute_task` -- 46 write intents

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
- `insert_table`
- `insert_section_break`
- `insert_image`
- `insert_paragraphs`
- `insert_shape`
- `insert_chart`
- `delete_image`
- `replace_image`

## `document_create` -- 3 formats

- `docx`
- `xlsx`
- `pptx`

A new, empty document of the format, produced by DocBuilder and written to `target`: `local` (an absolute path on the
server host) or `nextcloud` (a path in the requester's own files). An existing file is never replaced. A `nextcloud`
target answers the new `fileId`; every further edit goes through the co-editing tools with that `fileId`.

## Notes

- `insert_table`: if `afterParagraph` names the document's last body paragraph, the engine cannot
  end the body on a table and writes one empty paragraph after the new table (measured; not an error).
- `insert_section_break`: sets only the new section's break type (`nextPage` or `continuous`); use
  `set_page_setup` afterward, addressing the new section by its index, to set its columns or other
  page-setup fields.
