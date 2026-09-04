from constant import ABS_DIR
import os,copy,re
from win32com.client import constants
import win32com.client as win32
from tool.file_trans import FileConverter
from tool.basetool import ContextToolsConfig,BaseTool

class TableReader():
    def __init__(self, pyconfig=ContextToolsConfig(config_path="config/Tools/reader/table_reader_config.yaml")):
        self.config = pyconfig.config

    def pt_to_convert(self, value, unit):
        if value is None or value in [99999, 9999999]:
            return 99999
        value = float(value)
        # Speed up: read the converted value directly
        cm_unit = 28.346456692913385
        inches_unit = 72.0
        """Convert spacing values from various units to points (pt)."""
        if unit == "pt" or unit == "point":
            return value
        elif unit == "cm":
            return round(value / cm_unit, 2)
        elif unit == "mm":
            return round(10 * value / cm_unit, 2)
        elif unit == "inches":
            return round(value / inches_unit, 2)
        else:
            raise ValueError(f"不支持的单位: {unit}")

    def pt_to_percent(self, value, PageSetup):
        page_width = PageSetup.PageWidth - PageSetup.LeftMargin - PageSetup.RightMargin
        percent = round(value / page_width * 100, 2)
        return percent

    def get_table(self, doc, table_index):
        try:
            # Check if document object is valid
            if doc is None:
                raise ValueError("Document object cannot be None")

            # Check if document object has Tables attribute
            if not hasattr(doc, 'Tables'):
                raise ValueError("The provided object is not a valid Word document object")

            # Validate table index
            if not isinstance(table_index, int) or table_index < 1:
                raise ValueError(f"Table index must be a positive integer, got: {table_index}")

            # Get tables collection from document
            tables = doc.Tables

            # Check if any tables exist
            if tables.Count == 0:
                raise Exception("No tables found in the document")

            # Check if table index is within range
            if table_index > tables.Count:
                raise ValueError(
                    f"Table index {table_index} is out of range. Document contains {tables.Count} table(s)")

            # Get specified table (Word table indexing starts from 1)
            table = tables(table_index)

            # Validate table object
            if table is None:
                raise Exception(f"Retrieved table object is None for index: {table_index}")

            return table

        except ValueError as ve:
            # Re-raise ValueError with original type
            raise ve
        except Exception as e:
            # Wrap other exceptions with clearer error message
            error_msg = f"Failed to get table (index:{table_index}): {str(e)}"
            raise Exception(error_msg) from e

    def __get_column_width(self, table, col_index: int = 1):
        """Get width info for a table column (or all columns): pt value + rule.

        :param table: Word table object
        :param col_index: Column number (1 = first column, 0 = average of all columns)
        :return: dict with width (pt) and rule (exactly / auto / mixed)"""
        if table.Columns.Count == 0:
            raise ValueError("Table has no columns")
        if col_index < 0 or col_index > table.Columns.Count:
            raise ValueError(f"Invalid col_index: {col_index}, table has {table.Columns.Count} columns")
        # Width rule mapping table
        width_type_map = {
            constants.wdPreferredWidthPoints: "exactly",
            constants.wdPreferredWidthAuto: "auto"
        }
        # ---- Get average across all columns ----
        if col_index == 0:
            total_width = 0.0
            rule_set = set()
            for col in table.Columns:
                total_width += col.Width
                rule_set.add(col.PreferredWidthType)
            avg_width = total_width / table.Columns.Count if table.Columns.Count > 0 else 0
            if len(rule_set) == 1:
                rule = width_type_map.get(next(iter(rule_set)), "unknown")
            else:
                rule = "mixed"
            width_pt = avg_width
        # ---- Get the specified column ----
        else:
            col = table.Columns(col_index)
            width_pt = col.Width
            rule = width_type_map.get(col.PreferredWidthType, "unknown")
        return {
            "width": width_pt,
            "rule": rule
        }

    def __get_row_height(self, table, row_index: int = 1):
        """Get height info for a table row: pt value + rule.

        :param table: Word table object
        :param row_index: Row number (1 = first row, 0 = average of all rows)
        :return: dict with height (pt or None) and rule (exactly/at_least/auto/mixed)"""
        if table.Rows.Count == 0:
            raise ValueError("Table has no rows")
        if row_index < 0 or row_index > table.Rows.Count:
            raise ValueError(f"Invalid row_index: {row_index}, table has {table.Rows.Count} rows")

        # Mapping table
        rule_map = {
            constants.wdRowHeightExactly: "exactly",
            constants.wdRowHeightAtLeast: "at_least",
            constants.wdRowHeightAuto: "auto",
        }

        # ---- Handle whole-table average case ----
        if row_index == 0:
            total_height = 0.0
            valid_count = 0
            rule_set = set()

            for row in table.Rows:
                rule_code = row.HeightRule
                rule_str = rule_map.get(rule_code, "unknown")
                rule_set.add(rule_str)

                # Auto height or invalid values (99999) are excluded from the average
                if rule_code == constants.wdRowHeightAuto or row.Height >= 99999:
                    continue

                total_height += row.Height
                valid_count += 1

            avg_height = total_height / valid_count if valid_count > 0 else None
            rule = "mixed" if len(rule_set) > 1 else next(iter(rule_set), "unknown")

            return {"height": avg_height, "rule": rule}

        # ---- Get the specified row ----
        else:
            row = table.Rows(row_index)
            rule_code = row.HeightRule
            rule = rule_map.get(rule_code, "unknown")

            # If auto height or 99999, height is meaningless
            if rule_code == constants.wdRowHeightAuto or row.Height >= 99999:
                height_pt = None
            else:
                height_pt = row.Height

            return {"height": height_pt, "rule": rule}

    def __get_table_width(self, table):
        return {
            "width": table.PreferredWidth,
        }

    def __get_table_height(self, table):
        """Get the total height of the whole table (pt).
        :param table: Word table object
        :return: dict containing:
            {
                "height": float | None,   # Total table height (pt); None if all rows are auto
                "rule": str               # overall rule: exactly / at_least / auto / mixed
            }"""
        if table.Rows.Count == 0:
            raise ValueError("Table has no rows")
        total_height = 0.0
        valid_count = 0
        rule_set = set()
        # Iterate all rows and accumulate row heights
        for i in range(1, table.Rows.Count + 1):
            row_info = self.__get_row_height(table, i)
            rule_set.add(row_info["rule"])

            # Accumulate valid heights
            if row_info["height"] is not None:
                total_height += row_info["height"]
                valid_count += 1

        # Determine the overall rule
        if len(rule_set) == 1:
            overall_rule = next(iter(rule_set))
        else:
            overall_rule = "mixed"

        # If all rows use auto height, total height cannot be computed
        if valid_count == 0:
            total_height_pt = None
        else:
            total_height_pt = total_height
        return {
            "height": total_height_pt,
            "rule": overall_rule
        }

    def __get_text_wrapping(self, table):
        return {
            "wrapping_style": table.Rows.WrapAroundText
        }

    def __get_pagination(self, table):
        return {
            "allow_break_across_pages": table.Rows.AllowBreakAcrossPages,
            "repeat_header": table.Rows(1).HeadingFormat if table.Rows.Count > 0 else 0,
            "keep_with_next": table.Range.ParagraphFormat.KeepWithNext,
            "page_break_before": table.Range.ParagraphFormat.PageBreakBefore
        }

    def __get_table_alignment(self, table):
        # ==== Get overall table horizontal alignment ====
        horizontal_map = {
            0: "left",  # wdAlignParagraphLeft
            1: "center",  # wdAlignParagraphCenter
            2: "right"  # wdAlignParagraphRight
        }
        horizontal_value = table.Rows.Alignment
        horizontal_align = horizontal_map.get(horizontal_value, "unknown")
        # ==== Get overall table vertical alignment ====
        vertical_map = {
            0: "top",  # wdCellAlignVerticalTop
            1: "center",  # wdCellAlignVerticalCenter
            3: "bottom"  # wdCellAlignVerticalBottom
        }
        align_set = set()
        for row in table.Rows:
            for cell in row.Cells:
                vertical_value = cell.VerticalAlignment
                vertical_align = vertical_map.get(vertical_value, "unknown")
                align_set.add(vertical_align)
        # Determine the result
        if len(align_set) == 1:
            final_vertical_align = align_set.pop()  # Only one alignment style
        else:
            final_vertical_align = "mix"  # Multiple alignment styles present
        # print("Overall table vertical alignment:", final_vertical_align)

        return {
            "horizontal_align": horizontal_align,
            "vertical_align": final_vertical_align
        }

    def __get_left_indent(self, table):
        return {
            "indent": table.Rows.LeftIndent
        }

    def __get_cell_vertical_alignment(self, table, row_index=1, col_index=1):
        # ==== Get vertical alignment of the specified cell ====
        vertical_map = {
            0: "top",  # wdCellAlignVerticalTop
            1: "center",  # wdCellAlignVerticalCenter
            3: "bottom"  # wdCellAlignVerticalBottom
        }
        cell = table.Cell(row_index, col_index)
        vertical_value = cell.VerticalAlignment
        vertical_align = vertical_map.get(vertical_value, "unknown")
        return {
            "cell_vertical_align": vertical_align
        }

    def get_tables_format(self, doc):
        table_num = doc.Tables.Count
        formats = {}
        for index in range(table_num):
            table_index = index + 1
            table = self.get_table(doc, table_index)
            format = self.get_table_format(doc, table_index)
            if format.get("state") == "success":
                style_name = table.Cell(1, 1).Range.Text
                style_name = style_name.replace("\r\x07", "")
                if style_name == "":
                    continue
                format = format.get("properties")
                formats[str(table_index)] = format
        return formats

    def get_table_format(self, doc, table_index):
        attribution_dict = {
            # "column_width": self.__get_column_width,
            # "row_height": self.__get_row_height,
            "table_width": self.__get_table_width,
            # "table_height": self.__get_table_height,
            "text_wrapping": self.__get_text_wrapping,
            "pagination": self.__get_pagination,
            "alignment": self.__get_table_alignment,
            "left_indent": self.__get_left_indent,
        }
        try:
            # String conversion
            table_index = int(table_index)
            if table_index > 0:
                table = self.get_table(doc, table_index)
            else:
                print("table index must >= 0!")
                raise
            table_info = {
                # "column_width": None,
                # "row_height": None,
                "table_width": None,
                # "table_height": None,
                "text_wrapping": None,
                "pagination": None,
                "alignment": None,
                "left_indent": None,
            }
            # Fetch each property to read in order
            for attribution in attribution_dict:
                attribution_info_read_tool = attribution_dict.get(attribution)
                table_info[attribution] = attribution_info_read_tool(table)

            # Return success result
            return {"state": "success", "properties": table_info}
        except Exception as e:
            print(f"Get table format error! The detail is {e}")
            return {"state": "success", "properties": None, "exception": e}

    def get_table_infos(self, doc, before=0, after=0):
        try:
            table_info_result = []
            for table_index in range(1, doc.Tables.Count + 1):
                table_table_result = self.__get_table_info(doc, table_index=table_index, before=before, after=after)
                if table_table_result.get("state") == "success":
                    table_table_result.pop("state")
                    table_info_result.append(table_table_result)
            return table_info_result
        except Exception as e:
            print(f"Get table info error! The detail is: {e}")
            raise

    def __get_table_info(self, doc, table_index, before=0, after=0):
        """Get detailed information for the specified table.
        Return structure (example):
        {
            "table_index": 1,
            "cells": [
                {"row": 1, "col": 1, "paragraph": "Name", "paragraph_indices": [3]},
                {"row": 1, "col": 2, "paragraph": "Age", "paragraph_indices": [4]},
                ...
            ],
            "before_texts": ["table caption"],   # up to `before` non-empty paragraphs above the table
            "after_texts": ["para after table", ...]  # up to `after` non-empty paragraphs below the table
        }

        Principles:
        - before: do not use a simple Range-distance cutoff; take non-empty paragraphs immediately above the table, continuing upward if more are needed.
        - after: start from the first non-empty paragraph after the table and collect downward.
        - each cell's paragraph_indices: all paragraph indices that fall within the cell Range (1-based)."""
        try:
            # Base object
            tables = doc.Tables
            if table_index < 1 or table_index > tables.Count:
                raise IndexError(f"table_index {table_index} out of range (1..{tables.Count})")

            table = tables(table_index)
            all_paragraphs = list(doc.Paragraphs)  # 1..N, but this is Python list
            total_paras = len(all_paragraphs)

            table_start = table.Range.Start
            table_end = table.Range.End

            # --- 2. Page range ---
            start_page = table.Range.Information(constants.wdActiveEndPageNumber)
            end_page = table.Range.Duplicate
            end_page.Collapse(constants.wdCollapseEnd)
            end_page_number = end_page.Information(constants.wdActiveEndPageNumber)

            # === Find the last paragraph whose End <= table_start (candidate just above the table) ===
            last_para_before_table_idx = None
            for i, para in enumerate(all_paragraphs, start=1):
                # Stop at the first paragraph whose End exceeds table_start; the previous one is the above paragraph
                if para.Range.End <= table_start:
                    last_para_before_table_idx = i
                else:
                    break

            # === Build before_texts: collect up to `before` non-empty paragraphs upward from last_para_before_table_idx ===
            before_texts = []
            if before > 0 and last_para_before_table_idx:
                count = 0
                for idx in range(last_para_before_table_idx, 0, -1):  # Go upward from the paragraph immediately above
                    para_text = all_paragraphs[idx - 1].Range.Text.strip()
                    # Filter empty / carriage-return-only paragraphs
                    if para_text:
                        before_texts.append(para_text)
                        count += 1
                    # Stop once enough have been collected
                    if count >= before:
                        break
                before_texts.reverse()  # Restore document order

            # === Build after_texts: find first para whose Start >= table_end, then collect non-empty paragraphs ===
            first_para_after_table_idx = None
            for i, para in enumerate(all_paragraphs, start=1):
                if para.Range.Start >= table_end:
                    first_para_after_table_idx = i
                    break

            after_texts = []
            if after > 0 and first_para_after_table_idx:
                count = 0
                for idx in range(first_para_after_table_idx, total_paras + 1):
                    para_text = all_paragraphs[idx - 1].Range.Text.strip()
                    if para_text:
                        after_texts.append(para_text)
                        count += 1
                    if count >= after:
                        break

            # === Parse each cell: text + list of contained paragraph indices ===
            # For speed, pre-collect (start, end) tuples for each paragraph
            para_ranges = [(p.Range.Start, p.Range.End) for p in all_paragraphs]

            cells_info = []
            for r in range(1, table.Rows.Count + 1):
                for c in range(1, table.Columns.Count + 1):
                    cell = table.Cell(r, c)
                    raw_text = cell.Range.Text
                    # Strip trailing special characters from the cell (CR / cell marker)
                    text = raw_text.rstrip('\r\x07').strip()

                    cell_start = cell.Range.Start
                    cell_end = cell.Range.End

                    # Find all paragraph indices that fall within the cell range (1-based)
                    para_indices = []
                    # Iterate paragraph ranges; if Start >= cell_start and End <= cell_end, it belongs to the cell
                    for idx, (pstart, pend) in enumerate(para_ranges, start=1):
                        if pstart >= cell_start and pend <= cell_end:
                            para_indices.append(idx)

                    # Fallback: if none found (rare), try matching paragraph Start == cell_start
                    if not para_indices:
                        for idx, (pstart, pend) in enumerate(para_ranges, start=1):
                            if pstart == cell_start:
                                para_indices.append(idx)
                                break
                    cells_info.append({
                        "row": r,
                        "col": c,
                        "paragraph": text,
                        "paragraph_index": para_indices[0]  # May be [], [n], or [n, n+1, ...]
                    })

            return {
                "state": "success",
                "table_index": table_index,
                "page_range": {
                    "start_page": start_page,
                    "end_page": end_page_number
                },
                "cells": cells_info,
                "before_texts": before_texts,
                "after_texts": after_texts,
            }
        except Exception as e:
            print(f"Get Table info error! The detail is {e}")
            return {
                "state": "error",
                "table_index": table_index,
                "error": e
            }

    def __read_column_width(self, table, table_info, col_index: int = 1):
        width = self.__get_column_width(table, col_index)
        width_pt = width.get("width")
        table_info["column_width"]["width"]["value"]["pt"] = self.pt_to_convert(width_pt, "pt")
        table_info["column_width"]["width"]["value"]["mm"] = self.pt_to_convert(width_pt, "mm")
        table_info["column_width"]["width"]["value"]["cm"] = self.pt_to_convert(width_pt, "cm")
        table_info["column_width"]["width"]["value"]["inches"] = self.pt_to_convert(width_pt, "inches")
        table_info["column_width"]["rule"]["value"] = width.get("rule")
        return table_info

    def __read_row_height(self, table, table_info, row_index: int = 1):
        height = self.__get_row_height(table, row_index)
        height_pt = height.get("height")
        table_info["row_height"]["height"]["value"]["pt"] = self.pt_to_convert(height_pt, "pt")
        table_info["row_height"]["height"]["value"]["mm"] = self.pt_to_convert(height_pt, "mm")
        table_info["row_height"]["height"]["value"]["cm"] = self.pt_to_convert(height_pt, "cm")
        table_info["row_height"]["height"]["value"]["inches"] = self.pt_to_convert(height_pt, "inches")
        table_info["row_height"]["rule"]["value"] = height.get("rule")
        return table_info

    def __read_table_width(self, table, table_info):
        width = self.__get_table_width(table)
        table_info["table_width"]["value"] = width.get("width")
        return table_info

    def __read_text_wrapping(self, table, table_info):
        table_info["text_wrapping"]["value"] = table.Rows.WrapAroundText
        return table_info

    def __read_pagination(self, table, table_info):
        table_info["pagination"]["allow_break_across_pages"]["value"] = table.Rows.AllowBreakAcrossPages
        table_info["pagination"]["repeat_header"]["value"] = table.Rows(1).HeadingFormat if table.Rows.Count > 0 else 0
        table_info["pagination"]["keep_with_next"]["value"] = table.Range.ParagraphFormat.KeepWithNext
        table_info["pagination"]["page_break_before"]["value"] = table.Range.ParagraphFormat.PageBreakBefore
        return table_info

    def __read_table_alignment(self, table, table_info):
        alignment = self.__get_table_alignment(table)
        table_info["alignment"]["horizontal_align"]["value"] = alignment.get("horizontal_align")
        table_info["alignment"]["vertical_align"]["value"] = alignment.get("vertical_align")
        return table_info

    def __read_left_indent(self, table, table_info):
        left_indent = table.Rows.LeftIndent
        table_info["left_indent"]["value"]["pt"] = left_indent
        table_info["left_indent"]["value"]["mm"] = self.pt_to_convert(left_indent, "mm")
        table_info["left_indent"]["value"]["cm"] = self.pt_to_convert(left_indent, "cm")
        table_info["left_indent"]["value"]["inches"] = self.pt_to_convert(left_indent, "inches")
        return table_info

    def __read_cell_horizontal_align(self, table, table_info, row_index=1, col_index=1):
        table_info["cell_horizontal_align"]["value"] = self.__get_cell_vertical_alignment(table, row_index,
                                                                                          col_index).get(
            "cell_vertical_align")
        return table_info

    def read_table_properties(self, doc, table_index, params_list=[], language='zh', *args, **kwargs):
        # print("params_list: ",params_list)
        attribution_dict = {
            # "column_width": self.__read_column_width,
            # "row_height": self.__read_row_height,
            # "cell_horizontal_align": self.__read_cell_horizontal_align,
            "table_width": self.__read_table_width,
            # "table_height": self.__read_table_height,
            "text_wrapping": self.__read_text_wrapping,
            "pagination": self.__read_pagination,
            "alignment": self.__read_table_alignment,
            "left_indent": self.__read_left_indent,
        }
        # Load the read template
        template = self.config.get("properties_template")
        if language in ['zh', 'en']:
            table_info = copy.deepcopy(template.get(language))
        else:
            table_info = copy.deepcopy(template.get("zh"))
            print("Default Using Chinese")
        try:
            # String conversion
            table_index = int(table_index)
            if table_index > 0:
                table = self.get_table(doc, table_index)
            else:
                print("table index must >= 0!")
                raise
            if not params_list:
                # No property scope specified; read all by default
                params_list = list(attribution_dict.keys())
            else:
                # After a property scope is specified, remove unused keys from the template
                for attribution in attribution_dict.keys():
                    if attribution not in params_list:
                        table_info.pop(attribution)

            # Fetch each property to read in order
            for params in params_list:
                # Call parameters are supported
                if params in attribution_dict:
                    attribution_info_read_tool = attribution_dict.get(params)
                    table_info = attribution_info_read_tool(table, table_info)

            # Return success result
            return {"state": "success", "properties": table_info}

        except Exception as e:
            # Catch exceptions and return error information
            return {"state": "false", "table_index": table_index, "exception": str(e)}

    def get_table_properties(self, doc, table_index, params_list=[], language='zh', *args, **kwargs):

        # print("params_list: ",params_list)
        attribution_dict = {
            # "column_width": self.__get_column_width,
            # "row_height": self.__get_row_height,
            "table_width": self.__get_table_width,
            # "table_height": self.__get_table_height,
            "text_wrapping": self.__get_text_wrapping,
            "pagination": self.__get_pagination,
            "alignment": self.__get_table_alignment,
            "left_indent": self.__get_left_indent,
        }
        try:
            # String conversion
            table_index = int(table_index)
            if table_index > 0:
                table = self.get_table(doc, table_index)
            else:
                print("table index must >= 0!")
                raise
            table_info = {}
            # Fetch each property to read in order
            for params in params_list:
                # Call parameters are supported
                if params in attribution_dict:
                    attribution_info_read_tool = attribution_dict.get(params)
                    table_info[params] = attribution_info_read_tool(table)

            # Return success result
            return {"state": "success", "properties": table_info}

        except Exception as e:
            # Catch exceptions and return error information
            return {"state": "false", "table_index": table_index, "exception": str(e)}


if __name__ == '__main__':
    word = win32.DispatchEx("Word.Application")  # Or use Dispatch
    word.Visible = True  # Make visible (recommended when debugging)
    word_file_path = "./file/image_test.docx"
    word_file_path = os.path.join(ABS_DIR, word_file_path)
    # print(word_file_path)
    # Open an existing document
    try:
        # Open the document
        doc = word.Documents.Open(word_file_path)
        table_reader = TableReader()
        table = doc.Tables(1)

        # table.PreferredWidthType = constants.wdPreferredWidthPoints
        # table.PreferredWidth = 400
        # table.AllowAutoFit = False

        # doc.Save()
        # print(table.Cell(1, 1).Range.Font.Name)
        # print(table.Columns(1).Width)
        print(table_reader.read_table_properties(doc,1,[]))

        # print(doc.Tables(2).PreferredWidth)
        # print(table_reader.read_table_properties(doc, 1, []))
        # print(table_reader.read_table_properties(doc, 2, []))

        # print(table_reader.get_table_infos(doc,1,1,1))



    except Exception as e:
        print(f"操作失败：{str(e)}")
        raise

    finally:
        # Ensure resources are cleaned up
        if 'doc' in locals():
            doc.Close(SaveChanges=False)
        word.Quit()