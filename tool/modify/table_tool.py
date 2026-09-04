import win32com.client as win32
from win32com.client import constants
import os
from tool.basetool import BaseTool
from constant import ABS_DIR
from tool.basetool import ContextToolsConfig

class TableBaseTools():
    def __init__(self):
        # Create once during initialization and reuse
        self.excel_app = win32.Dispatch("Excel.Application")

    def convert_to_pt(self, value, unit):
        execl = self.excel_app
        """Convert spacing values from various units to points (pt)."""
        if value is None:
            return 0
        if unit in ["pt","point","磅"]:
            return float(value)
        elif unit in ["cm","centimeter","厘米"]:
            # print(execl.CentimetersToPoints(value))
            return execl.CentimetersToPoints(value)
        elif unit in ["mm","millimeter","毫米"]:
            return execl.CentimetersToPoints(value*0.1)
        elif unit in ["inches","英寸"]:
            return execl.InchesToPoints(value)
        else:
            raise ValueError(f"不支持的单位: {unit}")

    def set_row_height(self, table, row_index: int, height: float, unit: str = "cm", rule: str = "exactly"):
        """Set the height of a table row (affects all cells in that row).
        :param table: Word table object
        :param row_index: Row number; 0 means all rows, 1 means the first row
        :param height: Height value
        :param unit: Unit; one of:
                     "cm" - centimeters
                     "mm" - millimeters
                     "point" - points
                     "inches" - inches
        :param rule: Height rule; one of:
                     "auto" - automatic
                     "at_least" - minimum height
                     "exactly" - exact height"""
        try:
            # Unit conversion: normalize to points (pt)
            if rule == "auto":
                height_pt = 0
            else:
                height_pt = self.convert_to_pt(value=height,unit=unit)

            # Height rule mapping
            rule_map = {
                "auto": constants.wdRowHeightAuto,
                "at_least": constants.wdRowHeightAtLeast,
                "exactly": constants.wdRowHeightExactly
            }

            if rule not in rule_map:
                raise ValueError("高度规则必须是 auto/at_least/exactly")

            # Set row height
            if row_index == 0:  # All rows
                for row in table.Rows:
                    row.HeightRule = rule_map[rule]
                    if rule != "auto":
                        row.Height = height_pt
                target_desc = f"all rows -> {height} {unit} ({rule})"
            else:  # Specified row
                row = table.Rows(row_index)
                row.HeightRule = rule_map[rule]
                if rule != "auto":
                    row.Height = height_pt
                target_desc = f"row {row_index} -> {height} {unit} ({rule})"

            return {
                    "status": "success",
                    "message": f"Set {target_desc}"
            }

        except Exception as e:
            return {
                    "status": "error",
                    "message": str(e)
            }

    def set_column_width(self, table, col_index: int, width: float, unit: str = "cm", rule: str = "exactly"):
        """Set the width of a table column (affects all cells in that column).

        :param table: Word table object
        :param col_index: Column number; 0 means all columns, 1 means the first column
        :param width: Width value
        :param unit: Unit; one of:
                     "cm" - centimeters
                     "mm" - millimeters
                     "point" - points
                     "inches" - inches
        :param rule: Width rule; one of:
                     "auto" - automatic
                     "at_least" - minimum width (not strictly supported by Word)
                     "exactly" - exact width"""
        try:
            # Unit conversion: normalize to points (pt)
            if rule == "auto":
                width_pt = 0
            else:
                width_pt = self.convert_to_pt(value=width, unit=unit)

            # Width rule mapping
            # Word has no column-width rule equivalent to RowHeightRule; control via PreferredWidthType only
            width_type_map = {
                "auto": constants.wdPreferredWidthAuto,
                "exactly": constants.wdPreferredWidthPoints
            }

            if rule not in width_type_map and rule != "at_least":
                raise ValueError("width rule must be auto/at_least/exactly")

            # Set column width
            if col_index == 0:  # All columns
                for col in table.Columns:
                    if rule == "auto":
                        col.PreferredWidthType = width_type_map["auto"]
                        col.PreferredWidth = 0
                    else:
                        col.PreferredWidthType = width_type_map.get(rule, constants.wdPreferredWidthPoints)
                        col.PreferredWidth = width_pt
                target_desc = f"all columns -> {width} {unit} ({rule})"

            else:  # Specified column
                col = table.Columns(col_index)
                if rule == "auto":
                    col.PreferredWidthType = width_type_map["auto"]
                    col.PreferredWidth = 0
                else:
                    col.PreferredWidthType = width_type_map.get(rule, constants.wdPreferredWidthPoints)
                    col.PreferredWidth = width_pt
                target_desc = f"column {col_index} -> {width} {unit} ({rule})"

            return {
                    "status": "success",
                    "message": f"Set {target_desc}"
            }

        except Exception as e:
            return {
                    "status": "error",
                    "message": str(e)
            }

    def set_allow_break_across_pages(self, table, allow_break):
        """Set whether the table allows rows to break across pages."""
        try:
            table.Rows.AllowBreakAcrossPages = allow_break
            return {
                    "status": "success",
                    "message": f"Set allow_break_across_pages = {allow_break}"
            }
        except Exception as e:
            return {
                    "status": "error",
                    "message": str(e)
            }

    def set_repeat_header(self, table, repeat):
        """Set whether the table repeats its header row."""
        try:
            if table.Rows.Count > 0:
                table.Rows(1).HeadingFormat = repeat

            return {
                    "status": "success",
                    "message": f"Set repeat_header = {repeat}"
            }

        except Exception as e:
            return {
                    "status": "error",
                    "message": str(e)
            }

    def set_keep_with_next(self, table, keep):
        """Set the table paragraph "keep with next"."""
        try:
            table.Range.ParagraphFormat.KeepWithNext = keep
            return {
                    "status": "success",
                    "message": f"Set keep_with_next = {keep}"
            }

        except Exception as e:
            return {
                    "status": "error",
                    "message": str(e)
            }

    def set_page_break_before(self, table, enable):
        """Set whether to insert a page break before the table."""
        try:
            table.Range.ParagraphFormat.PageBreakBefore = enable
            return {
                    "status": "success",
                    "message": f"Set page_break_before = {enable}"
            }

        except Exception as e:
            return {
                    "status": "error",
                    "message": str(e)
            }

    def set_table_width(self, table, percent: float):
        """Set table width as a percentage of the page (the only stable, read/write-consistent approach).

        Design principles:
        - Force auto_content as the baseline
        - Only support percentage width
        - Do not freeze AllowAutoFit"""
        try:
            if not (0 < percent <= 100):
                raise ValueError("percent must be in (0, 100]")

            # Baseline calibration
            table.AllowAutoFit = True
            table.AutoFitBehavior(constants.wdAutoFitContent)

            # Percentage width (core)
            table.PreferredWidthType = constants.wdPreferredWidthPercent
            table.PreferredWidth = percent

            return {
                "status": "success",
                "message": f"Set table width to {percent}% of page"
            }

        except Exception as e:
            return {
                "status": "error",
                "message": str(e)
            }

    def set_table_height(self, table, height: float = None, unit: str = "cm"):
        """Set overall table height (use auto as baseline; implement via evenly distributed row heights).

        Design principles:
        1. First restore all rows to automatic height (content-driven)
        2. If height is provided, then set exact row heights on that baseline
        3. Do not expose height rule externally, to avoid semantic pollution

        :param table: Word table object
        :param height: Total table height; if None, only perform auto-height calibration
        :param unit: Unit; one of "cm", "mm", "inches", "pt"
        """
        try:
            # ---- Basic validation ----
            row_count = table.Rows.Count
            if row_count == 0:
                raise ValueError("Table has no rows, cannot set height")

            # ---- Step 1: Restore all rows to auto height (baseline) ----
            for row in table.Rows:
                row.HeightRule = constants.wdRowHeightAuto

            desc = "auto"

            # ---- Step 2: If total height is given, freeze to exact row height ----
            if height is not None and height > 0:
                height_pt = self.convert_to_pt(height, unit=unit)
                row_height = height_pt / row_count

                for row in table.Rows:
                    row.HeightRule = constants.wdRowHeightExactly
                    row.Height = row_height

                desc = f"{height}{unit} (each row {row_height:.2f} pt)"

            return {
                "status": "success",
                "message": f"Set table height with baseline auto, final={desc}"
            }

        except Exception as e:
            return {
                "status": "error",
                "message": str(e)
            }

    def set_cell_vertical_alignment(self, cell, alignment: str):
        """Set cell vertical alignment.
        Args:
            cell: Word cell object
            alignment (str): Vertical alignment; one of:
                             "top" / "center" / "bottom"
        Returns:
            {
                "vertical_alignment": {
                    "status": "success" / "error",
                    "message": "Operation result description"
                }
            }"""
        try:
            # Define alignment mapping (Word constants)
            alignment_map = {
                "TOP": 0,  # wdCellAlignVerticalTop
                "CENTER": 1,  # wdCellAlignVerticalCenter
                "BOTTOM": 3,  # wdCellAlignVerticalBottom
            }

            # Get the alignment value
            alignment_key = alignment_map.get(alignment.upper())
            if alignment_key is None:
                raise ValueError("Error Vertical Alignment, must be 'top', 'center' or 'bottom'")

            # Set alignment
            cell.VerticalAlignment = alignment_key

            return {
                    "status": "success",
                    "message": f"Cell vertical alignment set to {alignment.upper()}"
            }

        except Exception as e:
            return {
                    "status": "error",
                    "message": str(e)
            }

    def set_table_alignment(self, table, alignment: str):
        """Set overall horizontal alignment of the table (relative to the page).
        Args:
            table: Word table object
            alignment (str): Table alignment; one of:
                             "left" / "center" / "right"
        Returns:
            {
                "table_alignment": {
                    "status": "success" / "error",
                    "message": "Operation result description"
                }
            }"""
        try:
            # Alignment mapping (Word constants)
            alignment_map = {
                "LEFT": 0,  # wdAlignParagraphLeft
                "CENTER": 1,  # wdAlignParagraphCenter
                "RIGHT": 2,  # wdAlignParagraphRight
                "左对齐": 0,
                "居中": 1,
                "右对齐":2,
                "CENTERED":1,
            }

            align_key = alignment_map.get(alignment.upper())
            if align_key is None:
                raise ValueError("Error Alignment Type，must be 'left'、'center' or 'right'")

            # Set overall table alignment
            table.Rows.Alignment = align_key

            return {
                    "status": "success",
                    "message": f"Table alignment set to {alignment.upper()}"
            }

        except Exception as e:
            return {
                    "status": "error",
                    "message": str(e)
            }

    def set_table_text_wrapping(self, table, wrapping_style: int = 0):
        """Set table text wrapping style.

        :param table: Word table object
        :param wrapping_style: Wrapping style; one of:
                              "none" - no wrapping (default)
                              "around" - text wraps around"""
        try:
            table.Rows.WrapAroundText = wrapping_style
            style_desc = "no paragraph wrapping" if wrapping_style == 0 else "paragraph wrapping around"
            return {
                "status": "success",
                "message": f"Set table paragraph wrapping to '{wrapping_style}' ({style_desc})"
            }
        except Exception as e:
            return {
                "status": "error",
                "message": f"0 for no paragraph wrapping and -1 for paragraph wrapping around. The detail is {e}"
            }

    def set_table_left_indent(self, table, indent: float, unit: str = "cm"):
        """Set table left indent.

        :param table: Word table object
        :param indent: Indent distance
        :param unit: Unit; one of: "cm", "pt", "mm", "inches"
        """
        try:
            indent_pt = self.convert_to_pt(indent,unit)
            if indent_pt > 0:
                # First set table alignment to left so that left indent can be applied
                table.Rows.Alignment = 0  # 0 = wdAlignRowLeft
            # Set table left indent
            table.Rows.LeftIndent = indent_pt
            return {
                "status": "success",
                "message": f"Set table left indent to {indent} {unit} ({indent_pt:.2f} pt)"
            }
        except Exception as e:
            return {
                "status": "error",
                "message": str(e)
            }

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

    def get_cell(self, table, row_index, column_index):
        try:
            # Check if table object is valid
            if table is None:
                raise ValueError("Table object cannot be None")

            # Check if table object has Rows and Columns attributes
            if not hasattr(table, 'Rows') or not hasattr(table, 'Columns'):
                raise ValueError("The provided object is not a valid Word table object")

            # Validate row index
            if not isinstance(row_index, int) or row_index < 1:
                raise ValueError(f"Row index must be a positive integer, got: {row_index}")

            # Validate column index
            if not isinstance(column_index, int) or column_index < 1:
                raise ValueError(f"Column index must be a positive integer, got: {column_index}")

            # Check if row index is within range
            if row_index > table.Rows.Count:
                raise ValueError(f"Row index {row_index} is out of range. Table contains {table.Rows.Count} row(s)")

            # Check if column index is within range
            if column_index > table.Columns.Count:
                raise ValueError(
                    f"Column index {column_index} is out of range. Table contains {table.Columns.Count} column(s)")

            # Get specified cell (Word cell indexing starts from 1)
            cell = table.Cell(row_index, column_index)

            # Validate cell object
            if cell is None:
                raise Exception(f"Retrieved cell object is None for position: ({row_index}, {column_index})")

            return cell

        except ValueError as ve:
            # Re-raise ValueError with original type
            raise ve
        except Exception as e:
            # Wrap other exceptions with clearer error message
            error_msg = f"Failed to get cell (row:{row_index}, column:{column_index}): {str(e)}"
            raise Exception(error_msg) from e


class TableTools(BaseTool):
    def __init__(self,  pyconfig=ContextToolsConfig("/config/Tools/TableToolsConfig.yaml")):
        self.config = pyconfig.config
        self.name = self.config.get("name")
        self.tableTool = TableBaseTools()

    def __set_row_height(self, doc, table_index, row_index=0, height=1.0, unit="cm", rule="at_least"):
        table = self.tableTool.get_table(doc,table_index)
        status = self.tableTool.set_row_height(table,row_index=row_index,height=height,unit=unit,rule=rule)
        return {"row_height":status}

    def __set_column_width(self, doc, table_index, col_index=0, width=3.0, unit="cm",rule = "auto"):
        table = self.tableTool.get_table(doc, table_index)
        status = self.tableTool.set_column_width(table=table,col_index=col_index,width=width,unit=unit,rule=rule)
        return {"column_width":status}

    def __set_table_pagination(self, doc, table_index, allow_break_across_pages=None, repeat_header=None,
                             keep_with_next=None, page_break_before=None):
        table = self.tableTool.get_table(doc,table_index)
        result = {}
        if allow_break_across_pages is not None:
            status = self.tableTool.set_allow_break_across_pages(table,allow_break_across_pages)
            result["allow_break_across_page"] = status
        if repeat_header is not None:
            status = self.tableTool.set_repeat_header(table,repeat_header)
            result["repeat_header"] = status
        if keep_with_next is not None:
            status = self.tableTool.set_keep_with_next(table,keep_with_next)
            result["keep_with_next"] = status
        if page_break_before is not None:
            status = self.tableTool.set_page_break_before(table,page_break_before)
            result["page_break_before"] = status
        return result


    def __set_table_width(self,doc, table_index, width:float):
        table = self.tableTool.get_table(doc, table_index)
        status = self.tableTool.set_table_width(table, width)
        return {"table_width": status}

    def __set_cell_vertical_alignment(self, doc, table_index, row_index, col_index, alignment: str):
        table = self.tableTool.get_table(doc,table_index)
        cell = self.tableTool.get_cell(table,row_index,col_index)
        status = self.tableTool.set_cell_vertical_alignment(cell,alignment)
        return {"cell_vertical_alignment":status}

    def __set_table_vertical_alignment(self,doc, table_index,alignment: str):
        table = self.tableTool.get_table(doc, table_index)
        for i in range(1, table.Rows.Count + 1):
            for j in range(1, table.Columns.Count + 1):
                cell = table.Cell(i, j)
                status = self.tableTool.set_cell_vertical_alignment(cell,alignment)
        return {"table_alignment": status}

    def __set_table_horizontal_alignment(self,doc,table_index,alignment:str):
        table = self.tableTool.get_table(doc, table_index)
        status = self.tableTool.set_table_alignment(table, alignment)
        return {"table_alignment": status}


    def __set_table_alignment(self,doc, table_index, horizontal_align=None, vertical_align=None):
        results = {}
        if horizontal_align is not None:
            results["horizontal_alignment"] = self.__set_table_horizontal_alignment(doc,table_index,horizontal_align)
        if vertical_align is not None:
            results["vertical_alignment"] = self.__set_table_vertical_alignment(doc, table_index, vertical_align)
        return results

    def __set_table_left_indent(self, doc, table_index, indent:float, unit:str = "cm"):
        table = self.tableTool.get_table(doc, table_index)
        status = self.tableTool.set_table_left_indent(table,indent,unit)
        return {"table_left_indent": status}

    def __set_table_text_wrapping(self,doc,table_index, wrapping_style:int=0):
        table = self.tableTool.get_table(doc,table_index)
        status = self.tableTool.set_table_text_wrapping(table,wrapping_style)
        return {"text_wrapping":status}


    @BaseTool.register_tool({
        "en": {
            "function_description": "Set the width of tables",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "location_list": {"type": "list",
                               "description": "List of table indices. Indexing rules: 'all' = all tables; 1 = first table; 2 = second table."},
                "width": {
                    "type": "float",
                    "description": "The percentage value of the table width, e.g., 75 means the table occupies 75% of the page width."
                },
            }
        },
        "zh": {
            "function_description": "设置表格宽度",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "location_list": {"type": "list",
                               "description": "表格索引列表，索引规则：'all'=所有表格；1=第一个表格；2=第二个表格"},
                "width": {
                    "type": "float",
                    "description": "表格宽度的百分比数值，如75代表占页面75%"
                }
            }
        }
    })
    def set_table_width(self, doc, location_list, width):
        results = {}
        if 'all' in location_list:
            location_list = [i + 1 for i in range(doc.Tables.Count)]
        try:
            for table_index in location_list:
                status = self.__set_table_width(doc=doc, table_index=table_index, width=width)
                results = status
            doc.Save()
        except Exception as e:
            results["table_width"] = {
                "status": "error",
                "message": f"Failed to set row height, the detail is : {e}"}
        finally:
            return results

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set text wrapping for tables",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "location_list": {"type": "list",
                               "description": "List of table indices. Indexing rules: 'all' = all tables; 1 = first table; 2 = second table."},
                "wrapping_style": {
                    "type": "int",
                    "description": (
                            "Text wrapping style (integer type):\n"
                            " - 0: no text wrapping\n"
                            " - -1: enable text wrapping"
                    )
                }
            }
        },
        "zh": {
            "function_description": "设置表格文字环绕",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "location_list": {"type": "list",
                               "description": "表格索引列表，索引规则：'all'=所有表格；1=第一个表格；2=第二个表格"},
                "wrapping_style": {
                    "type": "int",
                    "description": (
                            "表格文字环绕（整数类型）：\n"
                            " - 0：无文字环绕\n"
                            " - -1：启用文字环绕"
                    )
                }
            }
        }
    })
    def set_table_text_wrapping(self,doc, location_list, wrapping_style=0):
        results = {}
        if 'all' in location_list:
            location_list = [i + 1 for i in range(doc.Tables.Count)]
        try:
            for table_index in location_list:
                status = self.__set_table_text_wrapping(doc=doc, table_index=table_index, wrapping_style=wrapping_style)
                results = status
            doc.Save()
        except Exception as e:
            results["text_wrapping"] = {
                "status": "error",
                "message": f"Failed to set text_wrapping, the detail is : {e}"}
        finally:
            return results

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set table pagination properties",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "location_list": {"type": "list",
                               "description": "List of table indices. Indexing rules: 'all' = all tables; 1 = first table; 2 = second table."},
                "allow_break_across_pages": {
                    "type": "int",
                    "description": "Allow breaking across pages (integer type): 0 = disabled, -1 = enabled"
                },
                "repeat_header": {
                    "type": "int",
                    "description": "Repeat header row on each page (integer type): 0 = disabled, -1 = enabled"
                },
                "keep_with_next": {
                    "type": "int",
                    "description": "Keep with next paragraph (integer type): 0 = disabled, -1 = enabled"
                },
                "page_break_before": {
                    "type": "int",
                    "description": "Page break before table (integer type): 0 = disabled, -1 = enabled"
                }
            }
        },
        "zh": {
            "function_description": "设置表格分页属性",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "location_list": {"type": "list",
                               "description": "表格索引列表，索引规则：'all'=所有表格；1=第一个表格；2=第二个表格"},
                "allow_break_across_pages": {
                    "type": "int",
                    "description": "允许跨行分页（整数类型）：0 = 关闭，-1 = 启用"
                },
                "repeat_header": {
                    "type": "int",
                    "description": "是否重复显示标题行（页眉行）：0 = 关闭，-1 = 启用"
                },
                "keep_with_next": {
                    "type": "int",
                    "description": "与下段同页（整数类型）：0 = 关闭，-1 = 启用"
                },
                "page_break_before": {
                    "type": "int",
                    "description": "段前分页（整数类型）：0 = 关闭，-1 = 启用"
                }
            }
        }
    })
    def set_table_pagination(self, doc, location_list, allow_break_across_pages=None, repeat_header=None,
                             keep_with_next=None, page_break_before=None):
        results = {}
        if 'all' in location_list:
            location_list = [i + 1 for i in range(doc.Tables.Count)]
        try:
            for table_index in location_list:
                status = self.__set_table_pagination(doc=doc, table_index=table_index, allow_break_across_pages=allow_break_across_pages, repeat_header=repeat_header,
                             keep_with_next=keep_with_next, page_break_before=page_break_before)
                results = status
            doc.Save()
        except Exception as e:
            results["text_wrapping"] = {
                "status": "error",
                "message": f"Failed to set text_wrapping, the detail is : {e}"}
        finally:
            return results

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set table alignment",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "location_list": {"type": "list",
                               "description": "List of table indices. Indexing rules: 'all' = all tables; 1 = first table; 2 = second table."},
                "horizontal_align": {
                    "type": "str",
                    "description": "Table horizontal alignment (string type): left = left alignment, center = center alignment, right = right alignment"
                },
                "vertical_align": {
                    "type": "str",
                    "description": "Table content vertical alignment (string type): top = top alignment, center = center alignment, bottom = bottom alignment"
                }
            }
        },
        "zh": {
            "function_description": "设置表格对齐方式",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "location_list": {"type": "list",
                               "description": "表格索引列表，索引规则：'all'=所有表格；1=第一个表格；2=第二个表格"},
                "horizontal_align": {
                    "type": "str",
                    "description": "表格水平对齐方式（字符串类型）：left = 左对齐，center = 居中，right = 右对齐"
                },
                "vertical_align": {
                    "type": "str",
                    "description": "表格内容垂直对齐方式（字符串类型）：top = 靠上，center = 居中，bottom = 底端对齐"
                }
            }
        }
    })
    def set_table_alignment(self, doc, location_list, horizontal_align=None,vertical_align=None):
        results = {}
        if 'all' in location_list:
            location_list = [i + 1 for i in range(doc.Tables.Count)]
        try:
            for table_index in location_list:
                status = self.__set_table_alignment(doc,table_index,horizontal_align=horizontal_align,vertical_align=vertical_align)
                results = status
            doc.Save()
        except Exception as e:
            results["table_alignment"] = {
                "status": "error",
                "message": f"Failed to set table alignment, the detail is : {e}"}
        finally:
            return results

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set table left indent",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "location_list": {"type": "list",
                               "description": "List of table indices. Indexing rules: 'all' = all tables; 1 = first table; 2 = second table."},
                "indent": {
                    "type": "float",
                    "description": "Left indent value"
                },
                "unit": {
                    "type": "str",
                    "description": (
                            "Indent unit, one of the following values:\n"
                            " - cm: centimeters\n"
                            " - mm: millimeters\n"
                            " - inches: inches\n"
                            " - pt: points"
                    )
                }
            }
        },
        "zh": {
            "function_description": "设置表格左缩进",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "location_list": {"type": "list",
                               "description": "表格索引列表，索引规则：'all'=所有表格；1=第一个表格；2=第二个表格"},
                "indent": {
                    "type": "float",
                    "description": "左缩进数值"
                },
                "unit": {
                    "type": "str",
                    "description": (
                            "缩进单位，可选值如下：\n"
                            " - cm：厘米\n"
                            " - mm：毫米\n"
                            " - inches：英寸\n"
                            " - pt：磅"
                    )
                }
            }
        }
    })
    def set_table_left_indent(self, doc, location_list, indent=0, unit = "cm"):
        results = {}
        if 'all' in location_list:
            location_list = [i + 1 for i in range(doc.Tables.Count)]
        try:
            for table_index in location_list:
                status = self.__set_table_left_indent(doc=doc,table_index=table_index,indent=indent,unit=unit)
                results = status
            doc.Save()
        except Exception as e:
            results["table_left_indent"] = {
                "status": "error",
                "message": f"Failed to set table alignment, the detail is : {e}"}
        finally:
            return results

    def set_cell_vertical_alignment(self, doc, location_list, row_index, col_index, alignment: str):
        results = {}
        if 'all' in location_list:
            location_list = [i + 1 for i in range(doc.Tables.Count)]
        try:
            for table_index in location_list:
                status = self.__set_cell_vertical_alignment(doc=doc,table_index=table_index,row_index=row_index, col_index=col_index, alignment=alignment)
                results = status
            doc.Save()
        except Exception as e:
            results["cell_vertical_alignment"] = {
                "status": "error",
                "message": f"Failed to set table alignment, the detail is : {e}"}
        finally:
            return results


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
        table_tool = TableTools()

        # Fully test all functions
        # test_table_format_functions(word_file_path)

        # print(doc.Tables(1).Rows(1).AllowBreakAcrossPages)
        # print(doc.Tables(1).Rows(1).VerticalAlignment)
        # table_tool.set_table_left_indent(doc,[1],1.0,"cm")

        # print(table_tool.set_table_alignment(doc,["all"],horizontal_align="center",vertical_align = "center" ))
        print(table_tool.set_table_pagination(doc,["all"],repeat_header=-1))
        doc.Save()
        # print("top:",doc.Tables(1).Cell(1,1).VerticalAlignment)
        # print("center:", doc.Tables(1).Cell(1, 2).VerticalAlignment)
        # print("bottom:", doc.Tables(1).Cell(1, 3).VerticalAlignment)
        # print("text wrapping:", doc.Tables(1).Rows.WrapAroundText)

    except Exception as e:
        print(f"操作失败：{str(e)}")
        raise

    finally:
        # Ensure resources are cleaned up
        if 'doc' in locals():
            doc.Close(SaveChanges=False)
        word.Quit()