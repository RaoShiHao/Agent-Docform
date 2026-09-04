from tool.basetool import BaseTool
from constant import ABS_DIR
import os
import yaml
from win32com.client import constants
import win32com.client as win32
import win32com
from tool.basetool import ContextToolsConfig

class FormatReaderTool(BaseTool):
    def __init__(self, pyconfig=ContextToolsConfig("/config/Tools/FormatReaderToolsConfig.yaml")):
        self.config = pyconfig.config
        self.name = self.config.get("name")

    def pt_to_convert(self, value, unit):
        value = float(value)
        cm_unit = 28.346456692913385
        inches_unit = 72.0
        """Convert spacing values from various units to points (pt)."""
        if value is None:
            return 0
        if unit == "pt" or unit == "point":
            return value
        elif unit == "cm":
            return round(value/cm_unit,2)
        elif unit == "mm":
            return round(10*value/cm_unit,2)
        elif unit == "inches":
            return round(value/inches_unit,2)
        else:
            raise ValueError(f"不支持的单位: {unit}")

    def color_int_to_hex(self,color_int: int) -> str:
        # Convert to unsigned 32-bit and keep the lowest 24 bits (RRGGBB)
        rgb = color_int & 0xFFFFFF
        return "#{:06X}".format(rgb)

    def get_headers_info(self, doc, section_index):
        """Safely get document header information (including checking whether the header exists)."""
        result = {
            'different_first_page': False,
            'different_odd_even': False,
            'header_distance': 0,
            'headers': {}
        }

        try:
            section = doc.Sections(section_index)
            page_setup = section.PageSetup

            # Basic header settings
            result['different_first_page'] = (page_setup.DifferentFirstPageHeaderFooter == -1)  # -1=True, 0=False
            result['different_odd_even'] = (page_setup.OddAndEvenPagesHeaderFooter == -1)
            result['header_distance'] = page_setup.HeaderDistance

            def safe_extract_header(header_type, name):
                """Safely extract header information (return empty data if the header does not exist)."""
                try:
                    header = section.Headers(header_type)
                    if not header.Exists:  # Critical check: whether the header actually exists
                        return None

                    range_ = header.Range
                    border = range_.Borders(win32.constants.wdBorderBottom)

                    return {
                        "text": range_.Text.strip(),
                        "name": range_.Font.Name,
                        "size": range_.Font.Size,
                        "alignment": range_.ParagraphFormat.Alignment,
                        "border_line": border.LineStyle if border.LineStyle != 0 else None  # 0 = no border
                    }
                except Exception as e:
                    print(f"读取页眉 {name} 失败: {str(e)}")
                    return None

            # Primary header (always attempt to read)
            if primary_info := safe_extract_header(win32.constants.wdHeaderFooterPrimary, "primary"):
                result['headers']['primary'] = primary_info

            # First-page header (read only when enabled)
            if result['different_first_page']:
                if first_info := safe_extract_header(win32.constants.wdHeaderFooterFirstPage, "first"):
                    result['headers']['first'] = first_info

            # Even-page header (read only when enabled)
            if result['different_odd_even']:
                if even_info := safe_extract_header(win32.constants.wdHeaderFooterEvenPages, "even"):
                    result['headers']['even'] = even_info

        except Exception as e:
            print(f"获取页眉信息时发生错误: {str(e)}")

        return result

    def get_footer_info(self, doc, section_index):
        """Safely get document footer information (including page-number settings)."""
        result = {
            'different_first_page': False,
            'different_odd_even': False,
            'footer_distance': 0,
            'page_number': {}
        }

        try:
            section = doc.Sections(section_index)
            page_setup = section.PageSetup

            # Basic footer settings
            result['different_first_page'] = (page_setup.DifferentFirstPageHeaderFooter == -1)  # -1=True, 0=False
            result['different_odd_even'] = (page_setup.OddAndEvenPagesHeaderFooter == -1)
            result['footer_distance'] = page_setup.FooterDistance

            def safe_extract_footer(footer_type, name):
                """Safely extract footer information (return None if the footer does not exist or has no page numbers)."""
                try:
                    footer = section.Footers(footer_type)
                    if not footer.Exists:  # Critical check: whether the footer exists
                        return None

                    footer_range = footer.Range
                    if footer_range.Fields.Count == 0:  # Check whether a page-number field exists
                        return None

                    page_numbers = footer.PageNumbers
                    return {
                        "format": page_numbers.NumberStyle,
                        "start": page_numbers.StartingNumber,
                        "continue": not page_numbers.RestartNumberingAtSection,
                        "alignment": footer_range.ParagraphFormat.Alignment,
                        "name": footer_range.Font.Name,
                        "size": footer_range.Font.Size
                    }
                except Exception as e:
                    print(f"读取页脚 {name} 失败: {str(e)}")
                    return None

            # Primary footer (always attempt to read)
            if primary_info := safe_extract_footer(win32.constants.wdHeaderFooterPrimary, "primary"):
                result["page_number"]['primary'] = primary_info

            # First-page footer (read only when enabled)
            if result['different_first_page']:
                if first_info := safe_extract_footer(win32.constants.wdHeaderFooterFirstPage, "first"):
                    result['page_number']['first'] = first_info

            # Even-page footer (read only when enabled)
            if result['different_odd_even']:
                if even_info := safe_extract_footer(win32.constants.wdHeaderFooterEvenPages, "even"):
                    result['page_number']['even'] = even_info

        except Exception as e:
            print(f"获取页脚信息时发生错误: {str(e)}")

        return result

    def get_page_properties(self, doc, section_index):
        """Read page properties of a Word document.
                :param doc: Word document object
                 section_index: Section index
                 content_x = 1 Number of pages of content to fetch
                :return: Operation result (status and page property info)"""
        try:
            # Get the page setup object
            page_setup = doc.PageSetup
            section = doc.Sections(section_index)
            # Get the TextColumns property for this section
            text_columns = section.PageSetup.TextColumns
            # Read page properties
            magin = {
                "TopMargin": page_setup.TopMargin,  # Top margin
                "BottomMargin": page_setup.BottomMargin,  # Bottom margin
                "LeftMargin": page_setup.LeftMargin,  # Left margin
                "RightMargin": page_setup.RightMargin,  # Right margin
            }
            gutter = {
                "Gutter": page_setup.Gutter,  # Gutter width
                "GutterPosition": page_setup.GutterPos,  # Gutter position (0: left, 1: top)
            }
            paper = {
                "PaperSize": page_setup.PaperSize,  # Paper size
                "PageWidth": page_setup.PageWidth,  # Page width
                "PageHeight": page_setup.PageHeight,  # Page height
                "Orientation": page_setup.Orientation,  # Page orientation (0: portrait, 1: landscape)
            }
            header = self.get_headers_info(doc, section_index)
            footer = self.get_footer_info(doc, section_index)
            grid = {
                "LayoutMode": page_setup.LayoutMode,  # Layout mode
                "LinesPage": page_setup.LinesPage,  # Lines per page
                "CharsLine": page_setup.CharsLine,  # Characters per page

            }
            columns = {
                "column_count": text_columns.Count,  # Number of columns
                "spacing": text_columns.Spacing,  # Column spacing
                "evenly_spaced": text_columns.EvenlySpaced,  # Whether columns are evenly distributed
                "line_between": text_columns.LineBetween,  # Whether to show column lines
                "column_width": text_columns(1).Width  # Column width
            }
            properties = {
                "magin": magin,
                "gutter": gutter,
                "paper": paper,
                "grid": grid,
                "columns": columns,
                "header": header,
                "footer": footer
            }

            rng = section.Range

            # Start page number
            start_rng = rng.Duplicate
            start_rng.Collapse(win32.constants.wdCollapseStart)
            start_page = start_rng.Information(win32.constants.wdActiveEndPageNumber)

            # End page number
            end_rng = rng.Duplicate
            end_rng.Collapse(win32.constants.wdCollapseEnd)
            end_page = end_rng.Information(win32.constants.wdActiveEndPageNumber)


            # Return success result
            return {"state": "success", "section_index": section_index,
                    "section_range": {"start": start_page, "end": end_page}, "page_format": properties}

        except Exception as e:
            # Catch exceptions and return error information
            return {"state": "false", "exception": str(e)}

    def get_section_properties(self,doc, section_index, content_x = 1):
        """Read page properties of a Word document.
        :param doc: Word document object
         section_index: Section index
         content_x = 1 Number of pages of content to fetch
        :return: Operation result (status and page property info)"""
        try:
            # Get the page setup object
            page_setup = doc.PageSetup
            section = doc.Sections(section_index)
            # Get the TextColumns property for this section
            text_columns = section.PageSetup.TextColumns
            # Read page properties
            magin = {
                "TopMargin": page_setup.TopMargin,  # Top margin
                "BottomMargin": page_setup.BottomMargin,  # Bottom margin
                "LeftMargin": page_setup.LeftMargin,  # Left margin
                "RightMargin": page_setup.RightMargin,  # Right margin
            }
            gutter = {
                "Gutter": page_setup.Gutter,  # Gutter width
                "GutterPosition": page_setup.GutterPos,  # Gutter position (0: left, 1: top)
            }
            paper = {
                "PaperSize": page_setup.PaperSize,  # Paper size
                "PageWidth": page_setup.PageWidth,  # Page width
                "PageHeight": page_setup.PageHeight,  # Page height
                "Orientation": page_setup.Orientation,  # Page orientation (0: portrait, 1: landscape)
            }
            header = self.get_headers_info(doc,section_index)
            footer = self.get_footer_info(doc,section_index)
            grid = {
                "LayoutMode": page_setup.LayoutMode, # Layout mode
                "LinesPage": page_setup.LinesPage,  # Lines per page
                "CharsLine": page_setup.CharsLine,  # Characters per page

            }
            columns = {
                "column_count":text_columns.Count, # Number of columns
                "spacing": text_columns.Spacing, # Column spacing
                "evenly_spaced": text_columns.EvenlySpaced, # Whether columns are evenly distributed
                "line_between": text_columns.LineBetween, # Whether to show column lines
                "column_width": text_columns(1).Width # Column width
            }
            properties = {
                "magin":magin,
                "gutter":gutter,
                "paper":paper,
                "grid":grid,
                "columns":columns,
                "header": header,
                "footer": footer
            }

            rng = section.Range

            # Start page number
            start_rng = rng.Duplicate
            start_rng.Collapse(win32.constants.wdCollapseStart)
            start_page = start_rng.Information(win32.constants.wdActiveEndPageNumber)

            # End page number
            end_rng = rng.Duplicate
            end_rng.Collapse(win32.constants.wdCollapseEnd)
            end_page = end_rng.Information(win32.constants.wdActiveEndPageNumber)

            content = self.get_section_first_x_pages(doc, section_index, start_page, end_page, x=content_x)
            # Return success result
            return {"state": "success", "section_index":section_index,"section_range":{"start":start_page,"end":end_page},"page_format": properties,"content":content}

        except Exception as e:
            # Catch exceptions and return error information
            return {"state": "false", "exception": str(e)}

    def get_font_properties(self, doc, paragraph_index):
        """Read all font properties of a paragraph in a Word document.
        :param doc: Word document object
        :param paragraph_index: Paragraph index (1-based)
        :return: Operation result (status and font property info)"""
        try:
            # Get the specified paragraph
            doc = doc
            paragraph = doc.Paragraphs(paragraph_index)
            # Get the paragraph font properties
            font = paragraph.Range.Font
            # print(dir(font))
            properties = {
                "Name": font.Name,  # Font name
                "NameAscii": font.NameAscii, # Western (ASCII) font
                "Size": font.Size,  # Font size
                "Bold": font.Bold,  # Whether bold
                "Italic": font.Italic,  # Whether italic
                "Underline": font.Underline,  # Underline style
                "Color": self.color_int_to_hex(font.Color),  # Font color
                "HighlightColor": font.Shading.BackgroundPatternColor,  # Background highlight color
                "StrikeThrough": font.StrikeThrough,  # Whether strikethrough
                "Subscript": font.Subscript,  # Whether subscript
                "Superscript": font.Superscript,  # Whether superscript
                "AllCaps": font.AllCaps,  # Whether all caps
                "SmallCaps": font.SmallCaps,  # Whether small caps
                "Spacing": font.Spacing,  # Character spacing
                "Scaling": font.Scaling,  # Character scaling
                "Emboss": font.Emboss,  # Whether emboss effect
                "Engrave": font.Engrave,  # Whether engrave effect
                "Shadow": font.Shadow,  # Whether shadow effect
            }
            # Return success result
            return {"state": "success", "properties": properties}

        except Exception as e:
            # Catch exceptions and return error information
            return {"state": "false", "exception": str(e)}

    def get_paragraph_properties(self, doc, index):
        """Get all formatting properties of the specified paragraph.
        :param index: Paragraph index (1-based)
        :return: Dictionary of paragraph properties (or error info on failure)"""
        try:
            # Validate the index
            if index < 1 or index > doc.Paragraphs.Count:
                raise ValueError(f"段落索引 {index} 超出有效范围（1-{doc.Paragraphs.Count}")

            # Get the paragraph object
            paragraph = doc.Paragraphs(index)
            fmt = paragraph.Range.ParagraphFormat

            align_key = {
                    0: "left", 1: "center", 2: "right",
                    3: "justify", 4: "distribute"
                }.get(fmt.Alignment, "unknown")

            # Get the paragraph start page number (start position)
            start_page = paragraph.Range.Information(3)

            # Get the paragraph end page number (end position)
            end_range = paragraph.Range.Duplicate  # Duplicate the Range object to avoid modifying the original paragraph
            end_range.Collapse(Direction=constants.wdCollapseEnd)  # Collapse to the end of the paragraph
            # end_range.Collapse(Direction=0)  # Collapse to the end of the paragraph
            end_page = end_range.Information(3)

            # Build the properties dictionary
            properties = {
                # Basic properties
                "text": paragraph.Range.Text.strip(),
                "style": paragraph.Style.NameLocal,
                "is_table": paragraph.Range.Tables.Count,
                # Alignment
                "alignment": align_key,
                # Outline level
                "outline_level": fmt.OutlineLevel,  # Outline level (1-10)
                # range page spans multiple pages
                "page_range": {"start_page": start_page, "end_page": end_page},
                # Spacing
                "spacing": {
                    # Line spacing
                    "line_spacing": {
                        "value": fmt.LineSpacing,
                        "rule": {
                            0: "single", 1: "1.5x", 2: "double",
                            4: "exact", 5: "multiple"
                        }.get(fmt.LineSpacingRule, "custom")
                    },
                    "before": fmt.SpaceBefore,  # Spacing before paragraph (pt)
                    "after": fmt.SpaceAfter,  # Spacing after paragraph (pt)
                },
                # Indentation
                "indent": {
                    "left": fmt.LeftIndent,  # Left indent (pt)
                    "right": fmt.RightIndent,  # Right indent (pt)
                    "first_line": fmt.FirstLineIndent  # First-line indent (positive) or hanging indent (negative)
                },
                # Pagination control
                "pagination": {
                    "widow_control": fmt.WidowControl,  # Widow/orphan control
                    "keep_with_next": fmt.KeepWithNext,  # Keep with next
                    "keep_together": fmt.KeepTogether,  # Keep lines together
                    "page_break_before": fmt.PageBreakBefore  # Page break before
                },
            }
            return {
                "state": "success",
                "properties": properties,
                "exception": None
            }
        except Exception as e:
            return {
                "state": "false",
                "properties": None,
                "exception": str(e)
            }

    def get_paragraph_format(self,doc,index):
        try:
            properies = {"paragraph":self.get_paragraph_properties(doc,index),"font":self.get_font_properties(doc,index)}
            return {"status": "success", "properties": properies, "exception": None}
        except Exception as e:
            print(f"Get Paragraph Format Error! The details is :{e}")
            return {"status":"false","properties":None,"exception":e}

    def get_image_properties(self,doc):
        """
        Read format information for all images in a Word document.

        Return format:
        {
            "state": "success/error",
            "images": [
                {
                    "index": int,             # Image index (1-based)
                    "width": float,           # Width (pt)
                    "height": float,          # Height (pt)
                    "scale_width": int,       # Width scale (%)
                    "scale_height": int,      # Height scale (%)
                    "lock_aspect_ratio": bool,# Whether aspect ratio is locked
                    "alignment": str,         # Alignment
                    "wrap_type": str,         # Text wrapping style
                    "horizontal_position": float, # Horizontal position (pt)
                    "vertical_position": float,   # Vertical position (pt)
                    "relative_horizontal_position": str, # Relative horizontal position
                    "relative_vertical_position": str    # Relative vertical position
                },
                ...
            ],
            "exception": None/str
        }
        """
        try:
            doc = doc
            images = []

            # Mapping table to convert constant values to readable strings
            alignment_map = {
                0: "left",
                1: "center",
                2: "right",
                3: "left",  # wdAlignParagraphJustify
                4: "distribute",
                7: "justify",
                -9999999: "undefined"
            }

            wrap_type_map = {
                0: "inline",  # wdWrapInline
                1: "square",  # wdWrapSquare
                2: "tight",  # wdWrapTight
                3: "through",  # wdWrapThrough
                4: "none",  # wdWrapNone
                5: "top-bottom",  # wdWrapTopBottom
                7: "inline"  # wdWrapInline
            }

            relative_position_map = {
                0: "margin",
                1: "page",
                2: "column",
                3: "character",
                4: "left_margin",
                5: "right_margin",
                6: "inside_margin",
                7: "outside_margin"
            }

            # Iterate over all shapes in the document (including images)
            for i, shape in enumerate(doc.InlineShapes, start=1):
                if shape.Type == 3:  # wdInlineShapePicture
                    # Get image format information
                    img_info = {
                        "index": i,
                        "page_number": shape.Range.Information(3),
                        "width": shape.Width,
                        "height": shape.Height,
                        "scale_width": shape.ScaleWidth,
                        "scale_height": shape.ScaleHeight,
                        "lock_aspect_ratio": shape.LockAspectRatio == -1,

                    }

                    # Get layout information (requires conversion to a Shape object)
                    try:
                        anchor = shape.ConvertToShape()
                        img_info.update({
                            "alignment": alignment_map.get(anchor.Anchor.ParagraphFormat.Alignment, "undefined"),
                            "wrap_type": wrap_type_map.get(anchor.WrapFormat.Type, "inline"),
                            "horizontal_position": anchor.Left,
                            "vertical_position": anchor.Top,
                            "relative_horizontal_position": relative_position_map.get(anchor.RelativeHorizontalPosition,
                                                                                      "margin"),
                            "relative_vertical_position": relative_position_map.get(anchor.RelativeVerticalPosition,
                                                                                    "margin")
                        })
                        anchor.Delete()  # Delete the temporarily converted shape
                    except Exception as e:
                        img_info.update({
                            "alignment": "inline",
                            "wrap_type": "inline",
                            "horizontal_position": 0,
                            "vertical_position": 0,
                            "relative_horizontal_position": "margin",
                            "relative_vertical_position": "margin"
                        })

                    images.append(img_info)

            # Return the result
            return {
                "state": "success",
                "images": images,
                "exception": None
            }

        except Exception as e:
            return {
                "state": "error",
                "images": [],
                "exception": str(e)
            }

    def get_all_tables_properties(self,doc):
        """Get modifiable properties of all tables in the document.

        Return format:
        {
            "success": True/False,
            "message": "Operation result description",
            "tables": [
                {
                    "table_index": int,
                    "rows": int,
                    "columns": int,
                    "cell_format": {
                        "font_size": float,
                        "bold": bool,
                        "font_color": str,
                        "alignment": str,
                        "vertical_alignment": str
                    },
                    "row_height": {
                        "height": float,
                        "unit": str,
                        "height_rule": str
                    },
                    "column_width": {
                        "width": float,
                        "unit": str
                    },
                    "table_size": {
                        "width": float,
                        "width_unit": str,
                        "height": float,
                        "height_unit": str
                    },
                    "pagination_settings": {
                        "allow_break_across_pages": bool,
                        "repeat_header": bool,
                        "keep_with_next": bool,
                        "page_break_before": bool
                    }
                },
                ...
            ]
        }"""
        try:
            tables = []

            # Reverse mapping for alignment
            alignment_reverse_map = {
                constants.wdAlignParagraphLeft: "left",
                constants.wdAlignParagraphCenter: "center",
                constants.wdAlignParagraphRight: "right",
                constants.wdAlignParagraphJustify: "justify",
                constants.wdAlignParagraphDistribute: "distribute"
            }

            # Reverse mapping for vertical alignment
            vertical_alignment_reverse_map = {
                constants.wdCellAlignVerticalTop: "top",
                constants.wdCellAlignVerticalCenter: "center",
                constants.wdCellAlignVerticalBottom: "bottom"
            }

            # Reverse mapping for height rule
            height_rule_reverse_map = {
                constants.wdRowHeightAuto: "auto",
                constants.wdRowHeightAtLeast: "at_least",
                constants.wdRowHeightExactly: "exact"
            }

            # Reverse mapping for width type
            width_type_reverse_map = {
                constants.wdPreferredWidthPoints: "pt",
                constants.wdPreferredWidthPercent: "percent",
                constants.wdPreferredWidthAuto: "auto"
            }

            # Iterate over all tables
            for i in range(1, doc.Tables.Count + 1):
                table = doc.Tables(i)

                # Get the table range object
                table_range = table.Range

                # Get cell format (use the first cell as a sample)
                sample_cell = table.Cell(1, 1).Range

                # Get font color (convert to hexadecimal)
                font_color = table_range.Font.Color
                if font_color != -1:  # -1 means automatic color
                    r = (font_color // 65536) % 256
                    g = (font_color // 256) % 256
                    b = font_color % 256
                    font_color_hex = f"#{r:02x}{g:02x}{b:02x}"
                else:
                    font_color_hex = "auto"

                # Get row height info (use the first row as a sample)
                sample_row = table.Rows(1)
                row_height_pt = sample_row.Height
                row_height_cm = row_height_pt / 28.35 if row_height_pt > 0 else 0

                # Get column width info (use the first column as a sample)
                sample_col = table.Columns(1)
                col_width_pt = sample_col.Width
                col_width_cm = col_width_pt / 28.35 if col_width_pt > 0 else 0

                # Build the table properties dictionary
                table_props = {
                    "index": i,
                    "rows": table.Rows.Count,
                    "columns": table.Columns.Count,
                    "cell_format": {
                        "font_size": table_range.Font.Size,
                        "bold": table_range.Font.Bold,
                        "font_color": font_color_hex,
                        "alignment": alignment_reverse_map.get(table_range.ParagraphFormat.Alignment, "unknown"),
                        "vertical_alignment": vertical_alignment_reverse_map.get(table.Cell(1, 1).VerticalAlignment,
                                                                                 "unknown")
                    },
                    "row_height": {
                        "height": row_height_cm,
                        "unit": "cm",
                        "height_rule": height_rule_reverse_map.get(sample_row.HeightRule, "unknown")
                    },
                    "column_width": {
                        "width": col_width_cm,
                        "unit": "cm"
                    },
                    "table_size": {
                        "width": table.PreferredWidth,
                        "width_unit": width_type_reverse_map.get(table.PreferredWidthType, "unknown"),
                        "height": sum(row.Height for row in table.Rows) / 28.35 if table.Rows.Count > 0 else 0,
                        "height_unit": "cm"
                    },
                    "pagination_settings": {
                        "allow_break_across_pages": table.Rows.AllowBreakAcrossPages,
                        "keep_with_next": table.Range.ParagraphFormat.KeepWithNext,
                        "page_break_before": table.Range.ParagraphFormat.PageBreakBefore
                    }
                }

                tables.append(table_props)

            return {
                "state": "success",
                "message": f"成功获取 {len(tables)} 个表格的属性",
                "tables": tables
            }

        except Exception as e:
            return {
                "state": "error",
                "message": f"获取表格属性时出错: {str(e)}",
                "tables": []
            }

    def get_section_first_x_pages(self, doc, section_index, start_page, end_page, x=1):
        """Get text content of the first x pages of a section (paragraph-granular; not a strict page cut).
        :param doc: Word document object
        :param section_index: Section index (1-based)
        :param start_page: Starting physical page number of the section
        :param end_page: Ending physical page number of the section
        :param x: First x pages
        :return: str"""
        section = doc.Sections(section_index)
        rng = section.Range
        texts = []
        # Compute the target page range
        target_page = min(start_page + x - 1, end_page)
        for para in rng.Paragraphs:
            prng = para.Range
            page = prng.Information(win32.constants.wdActiveEndPageNumber)
            if start_page <= page <= target_page:
                texts.append(prng.Text)
            elif page > target_page:
                break  # Stop once past the target page
        return "".join(texts)

    def read_paragraph_font_properties_base(self, style_format, font_fmt, para_fmt):
        # 7 basic font properties
        # style_format["font"]["basic"]["Name"]["value"] = font_fmt.Name
        # style_format["font"]["basic"]["NameFarEast"]["value"] = font_fmt.NameFarEast
        style_format["font"]["basic"]["NameAscii"]["value"] = font_fmt.NameAscii
        style_format["font"]["basic"]["Name"]["value"] = font_fmt.Name
        style_format["font"]["basic"]["Size"]["value"] = font_fmt.Size
        style_format["font"]["basic"]["Bold"]["value"] = font_fmt.Bold
        style_format["font"]["basic"]["Italic"]["value"] = font_fmt.Italic
        style_format["font"]["basic"]["Underline"]["value"] = font_fmt.Underline
        style_format["font"]["basic"]["Color"]["value"] = self.color_int_to_hex(font_fmt.Color)

        # 10 effects font properties
        style_format["font"]["effects"]["StrikeThrough"]["value"] = font_fmt.StrikeThrough
        style_format["font"]["effects"]["Subscript"]["value"] = font_fmt.Subscript
        style_format["font"]["effects"]["Superscript"]["value"] = font_fmt.Superscript
        style_format["font"]["effects"]["AllCaps"]["value"] = font_fmt.AllCaps
        style_format["font"]["effects"]["Spacing"]["value"] = font_fmt.Spacing
        style_format["font"]["effects"]["Scaling"]["value"] = font_fmt.Scaling
        style_format["font"]["effects"]["Emboss"]["value"] = font_fmt.Emboss
        style_format["font"]["effects"]["Engrave"]["value"] = font_fmt.Engrave
        style_format["font"]["effects"]["Shadow"]["value"] = font_fmt.Shadow
        style_format["font"]["effects"]["SmallCaps"]["value"] = font_fmt.SmallCaps

        # paragraph properties
        align_key = {
            0: "left", 1: "center", 2: "right",
            3: "justify", 4: "distribute"
        }.get(para_fmt.Alignment, "unknown")

        style_format["paragraph"]['alignment']["value"] = align_key
        style_format["paragraph"]['outlinelevel']["value"] = para_fmt.OutlineLevel

        # spacing
        style_format["paragraph"]["spacing"]['line_spacing']["spacing_rule"]['value'] = {
            0: "single", 1: "1.5x", 2: "double",
            4: "exact", 5: "multiple"
        }.get(para_fmt.LineSpacingRule, "custom")

        style_format["paragraph"]["spacing"]['line_spacing']["spacing_value"]['value'] = para_fmt.LineSpacing

        space_before = para_fmt.SpaceBefore
        style_format["paragraph"]["spacing"]['before_spacing']['value']["pt"] = self.pt_to_convert(space_before, "pt")
        style_format["paragraph"]["spacing"]['before_spacing']['value']["cm"] = self.pt_to_convert(space_before, "cm")
        style_format["paragraph"]["spacing"]['before_spacing']['value']["mm"] = self.pt_to_convert(space_before, "mm")
        style_format["paragraph"]["spacing"]['before_spacing']['value']["inches"] = self.pt_to_convert(space_before,
                                                                                                       "inches")

        space_after = para_fmt.SpaceAfter
        style_format["paragraph"]["spacing"]['after_spacing']['value']["pt"] = self.pt_to_convert(space_after, "pt")
        style_format["paragraph"]["spacing"]['after_spacing']['value']["cm"] = self.pt_to_convert(space_after, "cm")
        style_format["paragraph"]["spacing"]['after_spacing']['value']["mm"] = self.pt_to_convert(space_after, "mm")
        style_format["paragraph"]["spacing"]['after_spacing']['value']["inches"] = self.pt_to_convert(space_after,
                                                                                                      "inches")

        # indent
        character_unit = font_fmt.Size
        left_indent = abs(para_fmt.LeftIndent)
        style_format["paragraph"]["indent"]['left_indent']['value']["pt"] = self.pt_to_convert(left_indent, "pt")
        style_format["paragraph"]["indent"]['left_indent']['value']["cm"] = self.pt_to_convert(left_indent, "cm")
        style_format["paragraph"]["indent"]['left_indent']['value']["mm"] = self.pt_to_convert(left_indent, "mm")
        style_format["paragraph"]["indent"]['left_indent']['value']["inches"] = self.pt_to_convert(left_indent,
                                                                                                   "inches")
        style_format["paragraph"]["indent"]['left_indent']['value']["character"] = round(left_indent / character_unit)
        if para_fmt.LeftIndent >= 0:
            style_format["paragraph"]["indent"]['left_indent']['hanging'] = 0
        else:
            style_format["paragraph"]["indent"]['left_indent']['hanging'] = -1

        right_indent = abs(para_fmt.RightIndent)
        style_format["paragraph"]["indent"]['right_indent']['value']["pt"] = self.pt_to_convert(right_indent, "pt")
        style_format["paragraph"]["indent"]['right_indent']['value']["cm"] = self.pt_to_convert(right_indent, "cm")
        style_format["paragraph"]["indent"]['right_indent']['value']["mm"] = self.pt_to_convert(right_indent, "mm")
        style_format["paragraph"]["indent"]['right_indent']['value']["inches"] = self.pt_to_convert(right_indent,
                                                                                                    "inches")
        style_format["paragraph"]["indent"]['right_indent']['value']["character"] = round(right_indent / character_unit)
        if para_fmt.RightIndent >= 0:
            style_format["paragraph"]["indent"]['right_indent']['hanging'] = 0
        else:
            style_format["paragraph"]["indent"]['right_indent']['hanging'] = -1

        firstline_indent = abs(para_fmt.FirstLineIndent)
        style_format["paragraph"]["indent"]['firstline_indent']['value']["pt"] = self.pt_to_convert(firstline_indent,
                                                                                                    "pt")
        style_format["paragraph"]["indent"]['firstline_indent']['value']["cm"] = self.pt_to_convert(firstline_indent,
                                                                                                    "cm")
        style_format["paragraph"]["indent"]['firstline_indent']['value']["mm"] = self.pt_to_convert(firstline_indent,
                                                                                                    "mm")
        style_format["paragraph"]["indent"]['firstline_indent']['value']["inches"] = self.pt_to_convert(
            firstline_indent, "inches")
        style_format["paragraph"]["indent"]['firstline_indent']['value']["character"] = round(
            firstline_indent / character_unit)
        if para_fmt.FirstLineIndent >= 0:
            style_format["paragraph"]["indent"]['firstline_indent']['hanging'] = 0
        else:
            style_format["paragraph"]["indent"]['firstline_indent']['hanging'] = -1

        # pagination_control
        style_format["paragraph"]["pagination_control"]['widow_control']['value'] = para_fmt.WidowControl
        style_format["paragraph"]["pagination_control"]['keep_with_next']['value'] = para_fmt.KeepWithNext
        style_format["paragraph"]["pagination_control"]['keep_together']['value'] = para_fmt.KeepTogether
        style_format["paragraph"]["pagination_control"]['page_break_before']['value'] = para_fmt.PageBreakBefore

        return style_format

    def read_style_properties(self, doc, style_name,language='zh'):
        try:
            # Get or create the style
            style = doc.Styles(style_name)
            paragraphs = doc.Paragraphs
            # Check whether paragraphs exist
            if paragraphs.Count >= 1:
                para_fmt = paragraphs(1)
                para_fmt.Range.Style = style
                font_fmt = para_fmt.Range.Font
            else:
                # Insert a new paragraph and write default text
                new_paragraph = doc.Content.Paragraphs.Add()
                new_paragraph.Range.Text = "Style Reader Case"
                new_paragraph.Range.InsertParagraphAfter()  # Ensure the paragraph ends properly
                para_fmt = paragraphs(1)
                para_fmt.Range.Style = style
                font_fmt = para_fmt.Range.Font

            # font_fmt = style.Font
            # para_fmt = style.ParagraphFormat

            # Initialize styles to read
            style_format = self.config.get("style_properties_template")

            if language in['zh','en']:
                style_format = style_format.get(language)
            else:
                style_format = style_format.get("zh")
                print("Default Using Chinese")

            style_format = self.read_paragraph_font_properties_base(style_format, font_fmt,para_fmt)

            return style_format

        except Exception as e:
            print(e)
            return None

    def read_paragraph_properties(self, doc, para_index, language = 'zh'):
        try:
            # Validate the index
            if para_index < 1 or para_index > doc.Paragraphs.Count:
                raise ValueError(f"段落索引 {para_index} 超出有效范围（1-{doc.Paragraphs.Count}")

            # Get the paragraph object
            paragraph = doc.Paragraphs(para_index)
            para_fmt = paragraph.Range.ParagraphFormat
            font_fmt = paragraph.Range.Font
            style_format = self.config.get("paragraph_properties_template")

            if language in['zh','en']:
                style_format = style_format.get(language)
            else:
                style_format = style_format.get("zh")
                print("Default Using Chinese")

            # Build the properties dictionary
            properties = self.read_paragraph_font_properties_base(style_format=style_format,para_fmt=para_fmt,font_fmt=font_fmt)
            properties["font"]["basic"]["HighlightColorIndex"]["value"] = paragraph.Range.HighlightColorIndex
            return {
                "state": "success",
                "properties": properties,
                "exception": None
            }
        except Exception as e:
            return {
                "state": "false",
                "properties": None,
                "exception": str(e)
            }

    def read_page_properties(self, doc, section_index,language='zh',pop_None = True):
        try:
            # Get the page setup object
            if section_index == "all":
                page_setup = doc.PageSetup
                text_columns = doc.Sections(1).PageSetup.TextColumns
            else:
                section = doc.Sections(section_index)
                # Get the TextColumns property for this section
                page_setup = section.PageSetup
                text_columns = section.PageSetup.TextColumns

            page_info = self.config.get("page_properties_template")
            if language in['zh','en']:
                page_info = page_info.get(language)
            else:
                page_info = page_info.get("zh")
                print("Default Using Chinese")
            # Read page properties
            margin = {
                "top": page_setup.TopMargin,  # Top margin
                "bottom": page_setup.BottomMargin,  # Bottom margin
                "left": page_setup.LeftMargin,  # Left margin
                "right": page_setup.RightMargin,  # Right margin
            }
            # Fill in page margin properties
            for key,value in margin.items():
                page_info["margin"][key]["value"]["pt"] = value
                page_info["margin"][key]["value"]["cm"] = self.pt_to_convert(value,"cm")
                page_info["margin"][key]["value"]["mm"] =  self.pt_to_convert(value,"mm")
                page_info["margin"][key]["value"]["inches"] =  self.pt_to_convert(value,"inches")

            # Gutter / binding effect
            page_info["gutter"]["gutter"]["value"] = page_setup.Gutter  # Gutter width
            page_info["gutter"]["gutter_pos"]["value"] = page_setup.GutterPos  # Gutter width

            # Fill in paper values
            page_info["paper"]["size"]["value"] = page_setup.PaperSize # Paper size
            page_info["paper"]["orientation"]["value"] =  page_setup.Orientation  # Page orientation (0: portrait, 1: landscape)
            # Paper size
            paper = {
                "width": page_setup.PageWidth,  # Page width
                "height": page_setup.PageHeight,  # Page height
            }
            for key,value in paper.items():
                page_info["paper"][key]["value"]["pt"] = value
                page_info["paper"][key]["value"]["cm"] = self.pt_to_convert(value,"cm")
                page_info["paper"][key]["value"]["mm"] =  self.pt_to_convert(value,"mm")
                page_info["paper"][key]["value"]["inches"] =  self.pt_to_convert(value,"inches")

            # Fill in layout values
            page_info["grid"]["layout_mode"]["value"] = page_setup.LayoutMode  # Layout mode
            page_info["grid"]["lines_page"]["value"] = page_setup.LinesPage  # Lines per page
            page_info["grid"]["chars_line"]["value"] = page_setup.CharsLine  # Characters per page


            # columns: multi-column layout info
            page_info["columns"]["column_count"]["value"] = text_columns.Count  # Number of columns
            page_info["columns"]["evenly_spaced"]["value"] = text_columns.EvenlySpaced  # Whether columns are evenly distributed
            page_info["columns"]["line_between"]["value"] = text_columns.LineBetween  # Whether to show column lines
            columns = {
                "spacing": text_columns.Spacing,  # Column spacing
                "column_width": text_columns(1).Width  # Column width
            }
            for key,value in columns.items():
                page_info["columns"][key]["value"]["pt"] = value
                page_info["columns"][key]["value"]["cm"] = self.pt_to_convert(value,"cm")
                page_info["columns"][key]["value"]["mm"] = self.pt_to_convert(value,"mm")
                page_info["columns"][key]["value"]["inches"] = self.pt_to_convert(value,"inches")

            # Header
            header = self.get_headers_info(doc,section_index)
            header_dis = header.get("header_distance")

            # print(header)
            page_info["header"]["header_distance"]["value"]["pt"] = header_dis
            page_info["header"]["header_distance"]["value"]["cm"] = self.pt_to_convert(header_dis, "cm")
            page_info["header"]["header_distance"]["value"]["mm"] = self.pt_to_convert(header_dis, "mm")
            page_info["header"]["header_distance"]["value"]["inches"] = self.pt_to_convert(header_dis, "inches")
            page_info["header"]["different_first_page"]["value"] = header.get("different_first_page")
            page_info["header"]["different_odd_even"]["value"] = header.get("different_odd_even")


            for key in ["primary", "first","even"]:
                if key in header.get("headers"):
                    # print(key)
                    page_info["header"][key]['text']["value"] = header["headers"].get(key).get("text")
                    page_info["header"][key]['name']["value"] = header["headers"].get(key).get('name')
                    page_info["header"][key]['size']["value"] = header["headers"].get(key).get('size')
                    page_info["header"][key]['alignment']["value"] = header["headers"].get(key).get('alignment')
                    page_info["header"][key]['border_line']["value"] = header["headers"].get(key).get("border_line")
                else:
                    if pop_None:
                        page_info["header"].pop(key)


            # footer
            footer = self.get_footer_info(doc,section_index)
            footer_dis = footer.get("footer_distance")

            page_info["footer"]["footer_distance"]["value"]["pt"] = footer_dis
            page_info["footer"]["footer_distance"]["value"]["cm"] = self.pt_to_convert(footer_dis, "cm")
            page_info["footer"]["footer_distance"]["value"]["mm"] = self.pt_to_convert(footer_dis, "mm")
            page_info["footer"]["footer_distance"]["value"]["inches"] = self.pt_to_convert(footer_dis, "inches")
            page_info["footer"]["different_first_page"]["value"] = footer.get("different_first_page")
            page_info["footer"]["different_odd_even"]["value"] = footer.get("different_odd_even")

            for key in ["primary", "first", "even"]:
                if key in footer.get('page_number'):
                    page_info["footer"][key]['page_number']['format']["value"] = footer['page_number'].get(key).get('format')
                    page_info["footer"][key]['page_number']['start']["value"] = footer['page_number'].get(key).get('start')
                    page_info["footer"][key]['page_number']['continue']["value"] = footer['page_number'].get(key).get('continue')
                    page_info["footer"][key]['page_number']['alignment']["value"] = footer['page_number'].get(key).get('alignment')
                    page_info["footer"][key]['page_number']['name']["value"] = footer['page_number'].get(key).get('name')
                    page_info["footer"][key]['page_number']['size']["value"] = footer['page_number'].get(key).get('size')
                else:
                    if pop_None:
                        page_info["footer"].pop(key)
            # Return success result
            return {"state": "success", "section_index":section_index,"properties":  page_info}

        except Exception as e:
            # Catch exceptions and return error information
            return {"state": "false","section_index":section_index, "exception": str(e)}



if __name__ == '__main__':

    # modify = win32com.client.DispatchExEx("Word.Application")
    # print("DCOM & COM permissions configured correctly; Word application started")

    # modify = win32.DispatchEx("Word.Application")
    word = win32com.client.gencache.DispatchEx("Word.Application")
    word.Visible = True  # Make visible (recommended when debugging)
    # word_file_path = "experiment/theis_test_txt/theis_test_txt_init.docx"
    word_file_path = "./file/Word_test.docx"
    # word_file_path = "./file/Base.docx"
    # word_file_path = "./file/ch_gov.doc"
    # word_file_path = r"D:\pycharm_project\WPS_Agent\dataset\Template\Page\zh\normal\page3\base.docx"
    # word_file_path = "./file/acl2023.docx"

    # print(f"- {constants.wdUnderlineNone}：No underline;")
    # print(f"- {constants.wdUnderlineSingle}：Single underline;")
    # print(f"- {constants.wdUnderlineWords}：Words only underline;")
    # print(f"- {constants.wdUnderlineDouble}：Double underline;")
    # print(f"- {constants.wdUnderlineDotted}：Dotted underline;")
    # print(f"- {constants.wdUnderlineDash}：Dash underline;")
    # print(f"- {constants.wdUnderlineDashLong}：Dash long underline;")
    # print(f"- {constants.wdUnderlineWavy}：Wave underline;")
    # print(f"- {constants.wdUnderlineThick}：Thick underline;")

    word_file_path = os.path.join(ABS_DIR, word_file_path)
    # Open an existing document
    try:
        # Open the document
        doc = word.Documents.Open(word_file_path)
        reader_tool = FormatReaderTool()
        # print(doc.Paragraphs.Count)
        # for i in range(doc.Paragraphs.Count):
            # print(reader_tool.read_paragraph_properties(doc,i+1))

        # print(reader_tool.read_paragraph_properties(doc, 1))
        # print(reader_tool.read_font_properties(doc,9))
        # print(reader_tool.read_image_properties(doc))
        # print(reader_tool.get_all_tables_properties(doc))

        # print(reader_tool.get_page_properties(doc))
        # print(reader_tool.read_style_properties(doc,"MainBodyText"))
        # print(reader_tool.get_paragraph_properties(doc,1))
        # print(reader_tool.read_page_properties(doc,'all'))

        print(doc.PageSetup.Orientation)
        print(doc.Sections(1).PageSetup.Orientation)
        print(doc.Sections(2).PageSetup.Orientation)

        doc.PageSetup.Orientation = 0
        doc.Save()

        print(doc.PageSetup.Orientation)
        print(doc.Sections(1).PageSetup.Orientation)
        print(doc.Sections(2).PageSetup.Orientation)

        print(doc.PageSetup.PaperSize)
        print(doc.Sections(1).PageSetup.PaperSize)
        print(doc.Sections(2).PageSetup.PaperSize)

    except Exception as e:
        print(f"操作失败：{str(e)}")
        raise

    finally:
        # Ensure resources are cleaned up
        if 'doc' in locals():
            doc.Close(SaveChanges=False)
        word.Quit()