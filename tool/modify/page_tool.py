from tool.basetool import BaseTool
from constant import ABS_DIR
import os
from win32com.client import constants
import win32com.client as win32
from tool.basetool import ContextToolsConfig

class BasePageTools():
    def convert_to_pt(self, value, unit):
        execl = win32.Dispatch("Excel.Application")
        """Convert spacing values from various units to points (pt)."""
        if value is None:
            return 0
        if unit == "pt" or unit == "point" or unit == "磅":
            return float(value)
        elif unit == "cm" or unit == "centimeter" or unit == "厘米":
            return execl.CentimetersToPoints(value)
        elif unit == "mm" or unit == "millimeter" or unit == "毫米":
            return execl.CentimetersToPoints(value * 0.1)
        elif unit == "inches" or unit == "英寸":
            return execl.InchesToPoints(value)
        else:
            raise ValueError(f"不支持的单位: {unit}")

    def set_footer_header_layout(self, doc, section_index, header_distance=None, footer_distance=None, *args, **kwargs):
        """Set header/footer layout for a section (or globally), including different first page,
        odd/even pages switches, and header/footer distances from the edge.

        :param doc: Word document object
        :param section_index: Section index (int or 'all')
        :param header_distance: dict, header distance from top
        :param footer_distance: dict, footer distance from bottom
        :return: dict result"""
        results = {}
        try:
            if section_index == "all":
                page_setup = doc.PageSetup
            else:
                page_setup = doc.Sections(section_index).PageSetup

            # Header distance from edge
            if header_distance:
                try:
                    value = header_distance["value"]
                    unit = header_distance["unit"]
                    header_distance_pt = self.convert_to_pt(value,unit)
                    page_setup.HeaderDistance = header_distance_pt
                    results["header_distance"] = {
                        "status": "success",
                        "message": f"Header distance set to {header_distance_pt} pt ({value}{unit})."
                    }
                except Exception as e:
                    results["header_distance"] = {
                        "status": "error",
                        "message": f"Failed to set header distance, the detailed is: {str(e)}"
                    }

            # Footer distance from edge
            if footer_distance:
                try:
                    value = footer_distance["value"]
                    unit = footer_distance["unit"]
                    footer_distance_pt = self.convert_to_pt(value,unit)
                    page_setup.FooterDistance = footer_distance_pt
                    results["footer_distance"] = {
                        "status": "success",
                        "message": f"Footer distance set to {footer_distance_pt} pt ({value}{unit})."
                    }
                except Exception as e:
                    results["footer_distance"] = {
                        "status": "error",
                        "message": f"Failed to set footer distance, the detailed is: {str(e)}"
                    }

            doc.Save()
        except Exception as e:
            results["error"] = {
                "status": "error",
                "message": f"Failed to access section {section_index} setup, the detailed is: {str(e)}"
            }
        return results

    def __break_header_links(self,section, different_first_page, different_odd_even):
        # Primary Header: always disconnect
        header = section.Headers(win32.constants.wdHeaderFooterPrimary)
        if header.LinkToPrevious:
            header.LinkToPrevious = False

        # First Page Header
        if different_first_page == -1:
            header = section.Headers(win32.constants.wdHeaderFooterFirstPage)
            if header.LinkToPrevious:
                header.LinkToPrevious = False

        # Even Pages Header
        if different_odd_even == -1:
            header = section.Headers(win32.constants.wdHeaderFooterEvenPages)
            if header.LinkToPrevious:
                header.LinkToPrevious = False

    def __break_footer_links(self,section, different_first_page, different_odd_even):
        # Primary must be disconnected
        footer = section.Footers(win32.constants.wdHeaderFooterPrimary)
        if footer.LinkToPrevious:
            footer.LinkToPrevious = False
        # First page
        if different_first_page == -1:
            footer = section.Footers(win32.constants.wdHeaderFooterFirstPage)
            if footer.LinkToPrevious:
                footer.LinkToPrevious = False
        # Even pages
        if different_odd_even == -1:
            footer = section.Footers(win32.constants.wdHeaderFooterEvenPages)
            if footer.LinkToPrevious:
                footer.LinkToPrevious = False

    def set_header_content(self,doc,section_index,first=None, primary=None,even=None,
            different_first_page=0,different_odd_even=0, *args, **kwargs):
        results = {}
        try:
            # ===== 1. Determine the section =====
            if section_index == "all":
                sections = list(doc.Sections)
            else:
                sections = [doc.Sections(section_index)]
            for sec_idx, section in enumerate(sections, start=1):
                # ===== 2. Section-level switches (must be first) =====
                section.PageSetup.DifferentFirstPageHeaderFooter = (
                        different_first_page == -1
                )
                section.PageSetup.OddAndEvenPagesHeaderFooter = (
                        different_odd_even == -1
                )
                # Force-instantiate Header (very important)
                _ = section.Headers(win32.constants.wdHeaderFooterPrimary).Range
                if different_first_page == -1:
                    _ = section.Headers(win32.constants.wdHeaderFooterFirstPage).Range
                if different_odd_even == -1:
                    _ = section.Headers(win32.constants.wdHeaderFooterEvenPages).Range
                # ===== 3. Unconditionally disconnect inheritance (core fix) =====
                self.__break_header_links(
                    section,
                    different_first_page,
                    different_odd_even
                )
                header_map = {
                    "first": (
                        win32.constants.wdHeaderFooterFirstPage,
                        first,
                        different_first_page == -1
                    ),
                    "primary": (
                        win32.constants.wdHeaderFooterPrimary,
                        primary,
                        True
                    ),
                    "even": (
                        win32.constants.wdHeaderFooterEvenPages,
                        even,
                        different_odd_even == -1
                    ),
                }
                # ===== 4. Write Header content (or clear) =====
                for name, (header_type, config, enabled) in header_map.items():
                    if not enabled:
                        continue
                    try:
                        header = section.Headers(header_type)
                        rng = header.Range
                        # —— Always clear first, whether writing or not (avoid template pollution) ——
                        rng.Text = ""
                        if config:
                            rng.Text = config.get("paragraph", "")

                            fmt = config.get("format", {})
                            if "alignment" in fmt:
                                rng.ParagraphFormat.Alignment = fmt["alignment"]
                            if "name" in fmt:
                                rng.Font.Name = fmt["name"]
                            if "size" in fmt:
                                rng.Font.Size = fmt["size"]
                            if "border_line" in config:
                                rng.Borders(
                                    win32.constants.wdBorderBottom
                                ).LineStyle = config["border_line"]

                        results[f"section_{sec_idx}_{name}"] = {
                            "status": "success"
                        }
                    except Exception as e:
                        results[f"section_{sec_idx}_{name}"] = {
                            "status": "error",
                            "message": str(e)
                        }
            doc.Save()
        except Exception as e:
            results["error"] = {
                "status": "error",
                "message": str(e)
            }
        return results

    def set_footer_content(self,doc,section_index,first=None,primary=None,even=None,
                               different_first_page=0,different_odd_even=0,*args,**kwargs):
        results = {}
        try:
            # ===== 1. Determine the section =====
            if section_index == "all":
                sections = list(doc.Sections)
            else:
                sections = [doc.Sections(section_index)]
            for sec_idx, section in enumerate(sections, start=1):
                # ===== 2. Section-level switches (must be set first) =====
                section.PageSetup.DifferentFirstPageHeaderFooter = (
                        different_first_page == -1
                )
                section.PageSetup.OddAndEvenPagesHeaderFooter = (
                        different_odd_even == -1
                )
                # Force-instantiate Footer (prevent Word from skipping creation)
                _ = section.Footers(win32.constants.wdHeaderFooterPrimary).Range
                if different_first_page == -1:
                    _ = section.Footers(win32.constants.wdHeaderFooterFirstPage).Range
                if different_odd_even == -1:
                    _ = section.Footers(win32.constants.wdHeaderFooterEvenPages).Range
                # ===== 3. Unconditionally disconnect inheritance (critical fix) =====
                self.__break_footer_links(
                    section,
                    different_first_page,
                    different_odd_even
                )
                # ===== 4. Write footer content (if any) =====
                def apply_footer(footer_type, config):
                    footer = section.Footers(footer_type)
                    rng = footer.Range
                    # Completely clear the content
                    rng.Text = ""
                    try:
                        while rng.Fields.Count > 0:
                            rng.Fields(1).Delete()
                    except Exception:
                        pass
                    # PageNumbers only handles logic
                    page_numbers = footer.PageNumbers
                    page_numbers.RestartNumberingAtSection = not config.get(
                        "continue", True
                    )
                    page_numbers.StartingNumber = config.get("start", 1)
                    page_numbers.NumberStyle = config.get("format", 0)
                    # Insert the single PAGE Field
                    rng.Fields.Add(
                        Range=rng,
                        Type=win32.constants.wdFieldPage
                    )
                    # Paragraph
                    rng.ParagraphFormat.Alignment = config.get("alignment", 1)
                    # Font
                    rng.Font.Name = config.get("name", "Times New Roman")
                    rng.Font.Size = config.get("size", 12)
                footer_map = {
                    "first": (
                        win32.constants.wdHeaderFooterFirstPage,
                        first,
                        different_first_page == -1
                    ),
                    "primary": (
                        win32.constants.wdHeaderFooterPrimary,
                        primary,
                        True
                    ),
                    "even": (
                        win32.constants.wdHeaderFooterEvenPages,
                        even,
                        different_odd_even == -1
                    ),
                }
                for name, (footer_type, cfg, enabled) in footer_map.items():
                    if not enabled:
                        continue
                    try:
                        if cfg:
                            apply_footer(footer_type, cfg)
                        else:
                            # Explicitly clear (prevent leftover content)
                            rng = section.Footers(footer_type).Range
                            rng.Text = ""
                            try:
                                while rng.Fields.Count > 0:
                                    rng.Fields(1).Delete()
                            except Exception:
                                pass
                        results[f"section_{sec_idx}_{name}"] = {
                            "status": "success"
                        }
                    except Exception as e:
                        results[f"section_{sec_idx}_{name}"] = {
                            "status": "error",
                            "message": str(e)
                        }
            # ===== 5. Refresh the whole document and save =====
            doc.Fields.Update()
            doc.Save()

        except Exception as e:
            results["error"] = {
                "status": "error",
                "message": str(e)
            }
        return results

    def set_margin(self, doc, section_index, top, bottom, left, right, *args, **kwargs):
        """
        Set page margin properties for a Word document

        :param doc: Word document object
        :param top: Top margin (dict with value and unit)
        :param bottom: Bottom margin (dict with value and unit)
        :param left: Left margin (dict with value and unit)
        :param right: Right margin (dict with value and unit)
        :return: Dictionary containing results for each margin setting
        """
        results = {}

        try:
            if section_index == 'all':
                page_setup = doc.PageSetup
            else:
                page_setup = doc.Sections(section_index).PageSetup
            # Set top margin
            if top:
                try:
                    unit = top["unit"]
                    value = top["value"]
                    top_pt = self.convert_to_pt(value=value,unit=unit)
                    page_setup.TopMargin = top_pt
                    results["top_margin"] = {
                        "status": "success",
                        "message": f"Top margin set to {top} pt({value}{unit})."
                    }
                except Exception as e:
                    results["top_margin"] = {
                        "status": "error",
                        "message": f"Failed to set top margin, the detailed is: {str(e)}"
                    }

            # Set bottom margin
            if bottom:
                try:
                    unit = bottom["unit"]
                    value = bottom["value"]
                    bottom_pt = self.convert_to_pt(value=value,unit=unit)
                    page_setup.BottomMargin = bottom_pt
                    results["bottom_margin"] = {
                        "status": "success",
                        "message": f"Bottom margin set to {bottom_pt} pt ({value}{unit})."
                    }
                except Exception as e:
                    results["bottom_margin"] = {
                        "status": "error",
                        "message": f"Failed to access section {section_index} setup, the detailed is: {str(e)}"
                    }

            # Set left margin
            if left:
                try:
                    unit = left["unit"]
                    value = left["value"]
                    left_pt = self.convert_to_pt(value=value, unit=unit)
                    page_setup.LeftMargin = left_pt
                    results["left_margin"] = {
                        "status": "success",
                        "message": f"Left margin set to {left} pt ({value}{unit})."
                    }
                except Exception as e:
                    results["left_margin"] = {
                        "status": "error",
                        "message": f"Failed to set left margin, the detailed is: {str(e)}"
                    }

            # Set right margin
            if right:
                try:
                    unit = right["unit"]
                    value = right["value"]
                    right_pt = self.convert_to_pt(value=value, unit=unit)
                    page_setup.RightMargin= right_pt

                    results["right_margin"] = {
                        "status": "success",
                        "message": f"Right margin set to {right} pt ({value}{unit})."
                    }
                except Exception as e:
                    results["right_margin"] = {
                        "status": "error",
                        "message": f"Failed to set right margin, the detailed is: {str(e)}"
                    }

            doc.Save()
        except Exception as e:
            results["error"] = {
                "status": "error",
                "message": f"Failed to access section {section_index} setup, the detailed is: {str(e)}"
            }
        return results

    def set_gutter(self, doc, section_index, gutter=0, gutter_unit='pt', gutter_pos=0, *args, **kwargs):
        """
        Set gutter properties for a Word document

        :param doc: Word document object
        :param gutter: Gutter width in points
        :param gutter_pos: Gutter position (0: Left, 1: Top)
        :return: Dictionary containing results for each gutter setting
        """
        results = {}
        try:
            if section_index == 'all':
                page_setup = doc.PageSetup
            else:
                page_setup = doc.Sections(section_index).PageSetup
            # Set gutter width
            try:
                gutter_value = gutter
                gutter = self.convert_to_pt(gutter_value, gutter_unit)
                page_setup.Gutter = gutter
                # print("gutter_value:",gutter_value)
                # print("gutter:",gutter)
                results["gutter_width"] = {
                    "status": "success",
                    "message": f"Gutter width set to {gutter} points(from {gutter_value} {gutter_unit}) "
                }
            except Exception as e:
                results["gutter_width"] = {
                    "status": "error",
                    "message": f"Failed to set gutter width, the detailed is: {str(e)}"
                }

            # Set gutter position
            try:
                page_setup.GutterPos = gutter_pos
                position_text = "Left" if gutter_pos == 0 else "Top"
                results["gutter_position"] = {
                    "status": "success",
                    "message": f"Gutter position set to {position_text}"
                }
            except Exception as e:
                results["gutter_position"] = {
                    "status": "error",
                    "message": f"Failed to set gutter position, the detailed is: {str(e)}"
                }
            doc.Save()
        except Exception as e:
            results["error"] = {"status": "error",
                                "message": f"Failed to access section {section_index} setup, the detailed is: {str(e)}"}

        doc.Save()
        return results

    def set_paper(self, doc, section_index, size=None, width=None, height=None, orientation=None, *args, **kwargs):
        """
        Set paper properties for a Word document

        :param doc: Word document object
        :param size: Paper size constant
        :param width: page width (dict with value and unit)
        :param height: page height (dict with value and unit)
        :param orientation: page orientation (0: Portrait, 1: Landscape)
        :return: Dictionary containing results for each parameter setting
        """
        results = {}
        try:
            if section_index == 'all':
                page_setup = doc.PageSetup
            else:
                page_setup = doc.Sections(section_index).PageSetup

            # Set orientation
            if orientation is not None:
                try:
                    page_setup.Orientation = orientation
                    orientation_text = "Portrait" if orientation == 0 else "Landscape"
                    results["orientation"] = {
                        "status": "success",
                        "message": f"Orientation set to {orientation_text}"
                    }
                except Exception as e:
                    results["orientation"] = {
                        "status": "error",
                        "message": f"Failed to set orientation, the detailed is: {str(e)}"
                    }
            # Set paper size
            if size is not None:
                try:
                    page_setup.PaperSize = size
                    results["paper_size"] = {
                        "status": "success",
                        "message": f"Paper size set to {size}"
                    }
                except Exception as e:
                    results["paper_size"] = {
                        "status": "error",
                        "message": f"Failed to set paper size, the detailed is: {str(e)}"
                    }
            # Set page width
            if width:
                try:
                    value = width["value"]
                    unit = width["unit"]
                    width_pt = self.convert_to_pt(value,unit)
                    page_setup.PageWidth = width_pt
                    results["page_width"] = {
                        "status": "success",
                        "message": f"page width set to {width_pt} pt ({value}{unit})."
                    }
                except Exception as e:
                    results["page_width"] = {
                        "status": "error",
                        "message": f"Failed to set page width, the detailed is: {str(e)}"
                    }
            # Set page height
            if height:
                try:
                    value = height["value"]
                    unit = height["unit"]
                    height_pt = self.convert_to_pt(value,unit)
                    page_setup.PageHeight = height_pt
                    results["page_height"] = {
                        "status": "success",
                        "message": f"page height set to {height_pt} pt ({value}{unit})."
                    }
                except Exception as e:
                    results["page_height"] = {
                        "status": "error",
                        "message": f"Failed to set page height, the detailed is: {str(e)}"
                    }
            doc.Save()
        except Exception as e:
            results["error"] = {
                "status": "error",
                "message": f"Failed to access section {section_index} setup, the detailed is: {str(e)}"
            }
        return results

    def set_grid(self, doc, section_index, layout_mode=None, lines_page=None, chars_line=None, *args, **kwargs):
        """
        Set document grid properties for a Word document

        :param doc: Word document object
        :param layout_mode: Layout mode setting
        :param lines_page: Lines per page
        :param chars_line: Characters per line
        :return: Dictionary containing results for each parameter setting
        """
        results = {}
        if section_index == 'all':
            page_setup = doc.PageSetup
        else:
            page_setup = doc.Sections(section_index).PageSetup

        # Process layout_mode
        if layout_mode is not None:
            try:
                page_setup.LayoutMode = layout_mode
                results["layout_mode"] = {
                    "status": "success",
                    "message": f"1 Layout mode set to {page_setup.LayoutMode}"
                }
            except Exception as e:
                results["layout_mode"] = {
                    "status": "error",
                    "message": f"Failed to set layout mode, the detailed is : {str(e)}"
                }

        # Process chars_line
        if chars_line is not None and layout_mode > 0:
            try:
                page_setup.CharsLine = chars_line
                results["chars_line"] = {
                    "status": "success",
                    "message": f"Characters per line set to {chars_line}"
                }
            except Exception as e:
                results["chars_line"] = {
                    "status": "error",
                    "message": f"Failed to access section {section_index} setup, the detailed is: {str(e)}"
                }

        # Process lines_page
        if lines_page is not None and layout_mode > 0:
            try:
                page_setup.LinesPage = lines_page
                results["lines_page"] = {
                    "status": "success",
                    "message": f"Lines per page set to {lines_page}"
                }
            except Exception as e:
                results["lines_page"] = {
                    "status": "error",
                    "message": f"Failed to set lines per page, the detailed is : {str(e)}"
                }

        doc.Save()
        return results

    def set_columns(self, doc, section_index, column_count=1, evenly_spaced=0, column_width=None, spacing=None,
                    line_between=0, *args, **kwargs):
        """
        Set column layout for sections in a Word document

        :param doc: Word document object
        :param column_count: Number of columns
        :param evenly_spaced: Whether columns are evenly spaced (1=yes, 0=no)
        :param column_width: Column width settings (dict: value + unit)
        :param spacing: Column spacing settings (dict: value + unit)
        :param line_between: Whether to show line between columns (1=show, 0=hide)
        :return: Dictionary containing results for each parameter setting
        """
        results = {}
        try:
            if section_index == 'all':
                text_columns = doc.PageSetup.TextColumns
            else:
                section = doc.Sections(section_index)
                text_columns = section.PageSetup.TextColumns

            # Set column count
            try:
                text_columns.SetCount(column_count)
                results["column_count"] = {
                    "status": "success",
                    "message": f"Column count set to {column_count}"
                }
            except Exception as e:
                results["column_count"] = {
                    "status": "error",
                    "message": f"Failed to set column count, the detailed is: {str(e)}"
                }

            # Set evenly spaced
            try:
                if evenly_spaced:
                    text_columns.EvenlySpaced = evenly_spaced
                    results["evenly_spaced"] = {
                        "status": "success",
                        "message": f"Evenly spaced set to {evenly_spaced}"
                    }
            except Exception as e:
                results["evenly_spaced"] = {
                    "status": "error",
                    "message": f"Failed to set evenly spaced, the detailed is: {str(e)}"
                }

            # Set line between
            try:
                if line_between:
                    text_columns.LineBetween = line_between
                    results["line_between"] = {
                        "status": "success",
                        "message": f"Line between columns set to {line_between}"
                    }
            except Exception as e:
                results["line_between"] = {
                    "status": "error",
                    "message": f"Failed to set line between, the detailed is: {str(e)}"
                }

            # Set column width
            if column_width:
                try:
                    value = column_width["value"]
                    unit = column_width["unit"]
                    column_width_pt = self.convert_to_pt(value,unit=unit)
                    text_columns.Item(1).Width = column_width_pt
                    results["column_width"] = {
                        "status": "success",
                        "message": f"Column width set to {column_width_pt} pt ({value}{unit})."
                    }
                except Exception as e:
                    results["column_width"] = {
                        "status": "error",
                        "message": f"Failed to set column width, the detailed is: {str(e)}"
                    }

            # Set spacing
            if spacing:
                try:
                    value = spacing["value"]
                    unit = spacing["unit"]
                    spacing_pt = self.convert_to_pt(value, unit=unit)
                    text_columns.Spacing = spacing_pt
                    results["spacing"] = {
                        "status": "success",
                        "message": f"Column spacing set to {spacing_pt} pt ({value}{unit})."
                    }
                except Exception as e:
                    results["spacing"] = {
                        "status": "error",
                        "message": f"Failed to set column spacing, the detailed is: {str(e)}"
                    }
            doc.Save()
        except Exception as e:
            results["error"] = {
                "status": "error",
                "message": f"Failed to access section {section_index} setup, the detailed is: {str(e)}"
            }

        return results

    import win32com.client
    from win32com.client import constants

    def clean_header_content(self, doc, section_index):
        """Clear header content in a Word document.
        Args:
        doc: Word document object
        section_index: Section index to clear (1-based); 'all' clears all sections
        Returns:
        Dictionary containing results"""
        results = {}
        try:
            # Determine which sections to process
            if section_index == 'all':
                sections_to_process = range(1, doc.Sections.Count + 1)
            else:
                sections_to_process = [section_index]

            # Process each section
            for section_num in sections_to_process:
                try:
                    section = doc.Sections(section_num)

                    # Clear the primary header
                    if section.Headers(constants.wdHeaderFooterPrimary).Exists:
                        header_range = section.Headers(constants.wdHeaderFooterPrimary).Range
                        header_range.Delete()

                    results[f"section_{section_num}"] = {
                        "status": "success",
                        "message": f"第 {section_num} 节页眉已清除"
                    }

                except Exception as e:
                    results[f"section_{section_num}"] = {
                        "status": "error",
                        "message": f"清除第 {section_num} 节页眉时出错: {str(e)}"
                    }

            # Save the document
            doc.Save()
        except Exception as e:
            results["global_error"] = {
                "status": "error",
                "message": f"全局错误: {str(e)}"
            }

        return results


class PageTools(BaseTool):
    def __init__(self, pyconfig=ContextToolsConfig("/config/Tools/PageToolsConfig.yaml")):

        self.config = pyconfig.config
        self.name = self.config.get("name")
        self.excel = win32.Dispatch("Excel.Application")
        self.page_tool = BasePageTools()

    def convert_to_pt(self, value, unit):
        execl = win32.Dispatch("Excel.Application")
        """Convert spacing values from various units to points (pt)."""
        if value is None:
            return 0
        if unit == "pt" or unit == "point" or unit == "磅":
            return float(value)
        elif unit == "cm" or unit == "centimeter" or unit == "厘米":
            return execl.CentimetersToPoints(value)
        elif unit == "mm" or unit == "millimeter" or unit == "毫米":
            return execl.CentimetersToPoints(value * 0.1)
        elif unit == "inches" or unit == "英寸":
            return execl.InchesToPoints(value)
        else:
            raise ValueError(f"不支持的单位: {unit}")

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set the header and footer distances for Word document sections",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "section_list": {
                    "type": "list[int | str]",
                    "description": "Specify the target sections as a list. Each element represents a section identifier: an integer indicates the section number (starting from 1), and the string 'all' means apply to all sections. Examples: [1, 2] or ['all']."
                },
                "header_distance": {
                    "type": "dict",
                    "description": (
                            "Header distance from the top of the page(or the distance from the header to the page edge). Supported keys:\n"
                            "- value: numeric value\n"
                            "- unit: unit of measurement: 'pt', 'cm', 'mm', or 'inches'"
                    )
                },
                "footer_distance": {
                    "type": "dict",
                    "description": (
                            "Footer distance from the bottom of the page(or the distance from the header to the page edge).\n"
                            "Supported keys:\n"
                            "- value: numeric distance value\n"
                            "- unit: unit of measurement ('pt', 'cm', 'mm', or 'inches')"
                    )
                }
            }
        },
        "zh": {
            "function_description": "设置Word文档节页面中页眉距离和页脚距离的函数",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "section_list": {
                    "type": "list[int | str]",
                    "description": "指定要应用的分节列表。列表中的每个元素表示一个分节标识：整数表示第几节（从 1 开始计数），字符串 'all' 表示应用到所有节。例如：[1, 2] 或 ['all']。"
                },
                "header_distance": {
                    "type": "dict",
                    "description": (
                            "页眉距页面顶部的距离（或页眉距离页边缘距离），支持以下键：\n"
                            "- value: 数值\n"
                            "- unit: 单位：'pt'、'cm'、'mm'、'inches'"
                    )
                },
                "footer_distance": {
                    "type": "dict",
                    "description": (
                            "页脚距页面底部的距离（或页眉距离页边缘距离）。\n"
                            "支持的键：\n"
                            "- value：数值\n"
                            "- unit：计量单位，可选 'pt'、'cm'、'mm'、'inches'"
                    )
                },
            }
        }
    })
    def set_footer_header_layout(self, doc, section_list, header_distance=None, footer_distance=None, *args, **kwargs):
        result = None
        for section_index in section_list:
            result = self.page_tool.set_footer_header_layout(doc, section_index,header_distance,footer_distance)
        return result

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set header layout and content for the first section of a Word document",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "section_list": {
                    "type": "list[int | str]",
                    "description": "Specify the target sections as a list. Each element represents a section identifier: an integer indicates the section number (starting from 1), and the string 'all' means apply to all sections. Examples: [1, 2] or ['all']."
                },
                "first": {
                    "type": "dict",
                    "description": (
                            "First page header settings.\n"
                            "Only required when a different first page is enabled and the first page needs a header.\n"
                            "Supports the following keys:\n"
                            "- paragraph: Header text content\n"
                            "- format: Formatting dictionary containing:\n"
                            "  - alignment: Text alignment (0=left, 1=center, 2=right)\n"
                            "  - name: Font name (e.g., 'SimSun', 'Times New Roman')\n"
                            "  - size: Font size (integer)\n"
                            "- border_line: Header bottom border style (integer), supported values:\n"
                            "  0 = No border, 1 = Single solid, 2 = Dotted, 3 = Small dash,\n"
                            "  4 = Large dash, 5 = Dash dot, 6 = Dash dot dot, 7 = Double solid, 8 = Triple solid"
                    )
                },
                "primary": {
                    "type": "dict",
                    "description": (
                            "Primary header settings (for odd pages or unified header when odd/even differentiation is disabled).\n"
                            "Always required.\n"
                            "Supports the following keys:\n"
                            "- paragraph: Header text content\n"
                            "- format: Formatting dictionary containing:\n"
                            "  - alignment: Text alignment (0=left, 1=center, 2=right)\n"
                            "  - name: Font name (e.g., 'SimSun', 'Times New Roman')\n"
                            "  - size: Font size (integer)\n"
                            "- border_line: Header bottom border style (integer), supported values:\n"
                            "  0 = No border, 1 = Single solid, 2 = Dotted, 3 = Small dash,\n"
                            "  4 = Large dash, 5 = Dash dot, 6 = Dash dot dot, 7 = Double solid, 8 = Triple solid"
                    )
                },
                "even": {
                    "type": "dict",
                    "description": (
                            "Even page header settings.\n"
                            "Only required when different odd and even pages are enabled and a page number needs to be set for even-numbered pages.\n"
                            "Supports the following keys:\n"
                            "- paragraph: Header text content\n"
                            "- format: Formatting dictionary containing:\n"
                            "  - alignment: Text alignment (0=left, 1=center, 2=right)\n"
                            "  - name: Font name (e.g., 'SimSun', 'Times New Roman')\n"
                            "  - size: Font size (integer)\n"
                            "- border_line: Header bottom border style (integer), supported values:\n"
                            "  0 = No border, 1 = Single solid, 2 = Dotted, 3 = Small dash,\n"
                            "  4 = Large dash, 5 = Dash dot, 6 = Dash dot dot, 7 = Double solid, 8 = Triple solid"
                    )
                },
                "different_first_page": {
                    "type": "int",
                    "description": (
                            "Whether to use a different header for the first page. "
                            "Supported values:\n"
                            "- 0: do not enable different first page header\n"
                            "- -1: enable different first page header"
                    )
                },
                "different_odd_even": {
                    "type": "int",
                    "description": (
                            "Whether to use different headers for odd and even pages. "
                            "Supported values:\n"
                            "- 0: do not enable different odd/even headers\n"
                            "- -1: enable different odd/even headers"
                    )
                },
            }
        },
        "zh": {
            "function_description": "设置 Word 文档第一页节的页眉布局与内容",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "section_list": {
                    "type": "list[int | str]",
                    "description": "指定要应用的分节列表。列表中的每个元素表示一个分节标识：整数表示第几节（从 1 开始计数），字符串 'all' 表示应用到所有节。例如：[1, 2] 或 ['all']。"
                },
                "first": {
                    "type": "dict",
                    "description": (
                            "首页页眉设置。\n"
                            "仅当启用首页不同，且首页需要设置页眉时填写。\n"
                            "支持以下键：\n"
                            "- paragraph: 页眉文本内容\n"
                            "- format: 格式设置字典，包含以下键：\n"
                            "  - alignment: 对齐方式（0=左对齐，1=居中，2=右对齐）\n"
                            "  - name: 字体名称（如'宋体'、'Times New Roman'）\n"
                            "  - size: 字号大小（整数），单位为磅（pt）。中文常用字号对照情况：三号:16.0磅，小三:15.0磅，四号:14.0磅，小四:12.0磅，五号:10.5磅，小五:9.0磅\n"
                            "- border_line: 页眉底部边框样式（整数），支持的值：\n"
                            "  0 = 无边框，1 = 单实线，2 = 点线，3 = 小虚线，\n"
                            "  4 = 大虚线，5 = 点划线，6 = 点划点线，7 = 双实线，8 = 三实线"
                    )
                },
                "primary": {
                    "type": "dict",
                    "description": (
                            "常规页眉设置（用于奇数页或未启用奇偶页区分时的统一页眉）。\n"
                            "始终需要填写。\n"
                            "支持以下键：\n"
                            "- paragraph: 页眉文本内容\n"
                            "- format: 格式设置字典，包含以下键：\n"
                            "  - alignment: 对齐方式（0=左对齐，1=居中，2=右对齐）\n"
                            "  - name: 字体名称（如'宋体'、'Times New Roman'）\n"
                            "  - size: 字号大小（整数），单位为磅（pt）。中文常用字号对照情况：三号:16.0磅，小三:15.0磅，四号:14.0磅，小四:12.0磅，五号:10.5磅，小五:9.0磅\n"
                            "- border_line: 页眉底部边框样式（整数），支持的值：\n"
                            "  0 = 无边框，1 = 单实线，2 = 点线，3 = 小虚线，\n"
                            "  4 = 大虚线，5 = 点划线，6 = 点划点线，7 = 双实线，8 = 三实线"
                    )
                },
                "even": {
                    "type": "dict",
                    "description": (
                            "偶数页页眉设置。\n"
                            "仅当设置了启用奇偶页不同，且偶数页需要设置页码时填写。\n"
                            "支持以下键：\n"
                            "- paragraph: 页眉文本内容\n"
                            "- format: 格式设置字典，包含以下键：\n"
                            "  - alignment: 对齐方式（0=左对齐，1=居中，2=右对齐）\n"
                            "  - name: 字体名称（如'宋体'、'Times New Roman'）\n"
                            "  - size: 字号大小（整数），单位为磅（pt）。中文常用字号对照情况：三号:16.0磅，小三:15.0磅，四号:14.0磅，小四:12.0磅，五号:10.5磅，小五:9.0磅\n"
                            "- border_line: 页眉底部边框样式（整数），支持的值：\n"
                            "  0 = 无边框，1 = 单实线，2 = 点线，3 = 小虚线，\n"
                            "  4 = 大虚线，5 = 点划线，6 = 点划点线，7 = 双实线，8 = 三实线"
                    )
                },
                "different_first_page": {
                    "type": "int",
                    "description": (
                            "是否启用首页不同页眉。\n"
                            "支持的值：\n"
                            "- 0：不启用首页不同页眉\n"
                            "- -1：启用首页不同页眉"
                    )
                },
                "different_odd_even": {
                    "type": "int",
                    "description": (
                            "是否启用奇偶页不同页眉。\n"
                            "支持的值：\n"
                            "- 0：不启用奇偶页不同页眉\n"
                            "- -1：启用奇偶页不同页眉"
                    )
                }
            }
        }
    })
    def set_header_content(self,doc, section_list, first=None, primary=None, even=None,
                           different_first_page=0, different_odd_even=0, *args, **kwargs ):
        if different_first_page is True:
            different_first_page = -1
        elif different_first_page is False:
            different_first_page = 0
        if different_odd_even is True:
            different_odd_even = -1
        elif different_odd_even is False:
            different_odd_even = 0
        result = None
        for section_index in section_list:
            result = self.page_tool.set_header_content(doc, section_index, first, primary, even,
                           different_first_page, different_odd_even)
        return result

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set footer layout and page numbers for the first section of a Word document",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
               "section_list": {"type": "list[int | str]",
                                "description": "Specify the target sections as a list. Each element represents a section identifier: an integer indicates the section number (starting from 1), and the string 'all' means apply to all sections. Examples: [1, 2] or ['all']."
                                },
                "primary": {
                    "type": "dict",
                    "description": (
                            "Main footer settings (used for odd pages or unified footer when odd/even distinction is disabled). "
                            "This field is required.\n"
                            "Supported sub-key:\n"
                            "   - format: page number format code (int)\n"
                            "   - alignment: alignment (0 = left, 1 = center, 2 = right)\n"
                            "   - start: starting page number (int)\n"
                            "   - continue: whether to continue numbering (bool)\n"
                            "   - name: font name (e.g., 'Times New Roman')\n"
                            "   - size: font size in points (float or int)"
                    )
                },
                "first": {
                    "type": "dict",
                    "description": (
                            "Only required when a different first page is enabled and a page number needs to be set for the first page.\n"
                            "Supported sub-key:\n"
                            "   - format: page number format code (int)\n"
                            "   - alignment: alignment (0 = left, 1 = center, 2 = right)\n"
                            "   - start: starting page number (int)\n"
                            "   - continue: whether to continue numbering (bool)\n"
                            "   - name: font name (e.g., 'Times New Roman')\n"
                            "   - size: font size in points (float or int)"
                    )
                },
                "even": {
                    "type": "dict",
                    "description": (
                            "Only required when different odd and even pages are enabled and a page number needs to be set for even-numbered pages.\n"
                            "Supported sub-key:\n"
                            "  - format: Page number format code (integer, supported values are:)\n"
                            "      - 0 = Arabic numerals (\"1, 2, 3\")\n"
                            "      - 1 = Uppercase Roman numerals (\"I, II, III\")\n"
                            "      - 2 = Lowercase Roman numerals (\"i, ii, iii\")\n"
                            "      - 3 = Uppercase English letters (\"A, B, C\")\n"
                            "      - 4 = Lowercase English letters (\"a, b, c\")\n"
                            "      - 13 = Chinese counting numbers (\"一、二、三\")\n"
                            "      - 38 = Chinese capital numbers (\"壹、贰、叁\")\n"
                            "      - 30 = Chinese Heavenly Stems (\"甲、乙、丙\")\n"
                            "      - 31 = Chinese Earthly Branches (\"子、丑、寅\")\n"
                            "      - 57 = Arabic numerals with hyphens (e.g., \"- 1 -\")\n"
                            "  - alignment: alignment (0 = left, 1 = center, 2 = right)\n"
                            "  - start: starting page number (int)\n"
                            "  - continue: whether to continue numbering (bool)\n"
                            "  - name: font name (e.g., 'Times New Roman')\n"
                            "  - size: font size in points (float or int)"
                    )
                },
                "different_first_page": {
                    "type": "int",
                    "description": (
                            "Whether to use a different header for the first page. "
                            "Supported values:\n"
                            "- 0: do not enable different first page header\n"
                            "- -1: enable different first page header"
                    )
                },
                "different_odd_even": {
                    "type": "int",
                    "description": (
                            "Whether to use different headers for odd and even pages. "
                            "Supported values:\n"
                            "- 0: do not enable different odd/even headers\n"
                            "- -1: enable different odd/even headers"
                    )
                },
            }
        },
        "zh": {
            "function_description": "设置 Word 文档第一页节的页脚布局与页码信息",
            "params": {
                "doc": {"type": "object", "description": "Word 文档对象"},
                "section_list": {
                    "type": "list[int | str]",
                    "description": "指定要应用的分节列表。列表中的每个元素表示一个分节标识：整数表示第几节（从 1 开始计数），字符串 'all' 表示应用到所有节。例如：[1, 2] 或 ['all']。"
                },
                "primary": {
                    "type": "dict",
                    "description": (
                            "常规页脚设置（用于奇数页或未启用奇偶页区分时的统一页脚）。\n"
                            "始终需要填写。\n"
                            "支持以下子字段：\n"
                            "  - format：页码格式代码（整数，取值范围如下：）\n"
                            "      - 0 = 阿拉伯数字（1, 2, 3）\n"
                            "      - 1 = 大写罗马数字（I, II, III）\n"
                            "      - 2 = 小写罗马数字（i, ii, iii）\n"
                            "      - 3 = 大写英文字母（A, B, C）\n"
                            "      - 4 = 小写英文字母（a, b, c）\n"
                            "      - 13 = 中文数目编号（一、二、三）\n"
                            "      - 38 = 中文大写编号（壹、贰、叁）\n"
                            "      - 30 = 中文天干（甲、乙、丙）\n"
                            "      - 31 = 中文地支（子、丑、寅）\n"
                            "      - 57 = 阿拉伯数字带中横线（如 - 1 -）\n"
                            "  - alignment：页码对齐方式（0：左对齐，1：居中，2：右对齐）\n"
                            "  - start：起始页码（整数）\n"
                            "  - continue：是否延续页码（布尔值）\n"
                            "  - name：字体名称（如“宋体”、“Times New Roman”）\n"
                            "  - size：字体大小（单位：磅）"
                    )
                },
                "first": {
                    "type": "dict",
                    "description": (
                            "首页页脚设置。\n"
                            "仅当启用首页不同，且首页需要设置页码时需要填写。\n"
                            "支持以下子字段：\n"
                            "   - format：页码格式代码（整数）\n"
                            "   - alignment：页码对齐方式（0：左对齐，1：居中，2：右对齐）\n"
                            "   - start：起始页码（整数）\n"
                            "   - continue：是否延续页码（布尔值）\n"
                            "   - name：字体名称（如“宋体”、“Times New Roman”）\n"
                            "   - size：字体大小（单位：磅）"
                    )
                },
                "even": {
                    "type": "dict",
                    "description": (
                            "偶数页页脚设置。\n"
                            "仅当启用奇偶页不同，且偶数页需要设置页码时填写。\n"
                            "支持以下子字段：\n"
                            "   - format：页码格式代码（整数）\n"
                            "   - alignment：页码对齐方式（0：左对齐，1：居中，2：右对齐）\n"
                            "   - start：起始页码（整数）\n"
                            "   - continue：是否延续页码（布尔值）\n"
                            "   - name：字体名称（如“宋体”、“Times New Roman”）\n"
                            "   - size：字体大小（单位：磅）"
                    )
                },
                "different_first_page": {
                    "type": "int",
                    "description": "是否启用首页不同页脚：\n- 0 表示不启用\n- -1 表示启用"
                },
                "different_odd_even": {
                    "type": "int",
                    "description": "是否启用奇偶页不同页脚：\n- 0 表示不启用\n- -1 表示启用"
                },
            }
        }
    })
    def set_footer_content(self, doc, section_list, first=None, primary=None, even=None,
                           different_first_page=0, different_odd_even=0, *args, **kwargs):
        if different_first_page is True:
            different_first_page = -1
        elif different_first_page is False:
            different_first_page = 0
        if different_odd_even is True:
            different_odd_even = -1
        elif different_odd_even is False:
            different_odd_even = 0
        result = None
        for section_index in section_list:
            result = self.page_tool.set_footer_content(doc, section_index, first, primary, even,
                                                       different_first_page, different_odd_even)
        return result

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set the page margin properties of a Word document",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
               "section_list": {"type": "list[int | str]",
                                "description": "Specify the target sections as a list. Each element represents a section identifier: an integer indicates the section number (starting from 1), and the string 'all' means apply to all sections. Examples: [1, 2] or ['all']."
                                },
                "top": {
                    "type": "dict",
                    "description": (
                            "Top margin setting (optional). Supports the following keys:\n"
                            "- value: Numeric value of the top margin\n"
                            "- unit: Unit of measurement ('pt' for points, 'cm' for centimeters, "
                            "'mm' for millimeters, 'inches' for inches)"
                    )
                },
                "bottom": {
                    "type": "dict",
                    "description": (
                            "Bottom margin setting (optional). Supports the following keys:\n"
                            "- value: Numeric value of the bottom margin\n"
                            "- unit: Unit of measurement ('pt' for points, 'cm' for centimeters, "
                            "'mm' for millimeters, 'inches' for inches)"
                    )
                },
                "left": {
                    "type": "dict",
                    "description": (
                            "Left margin setting (optional). Supports the following keys:\n"
                            "- value: Numeric value of the left margin\n"
                            "- unit: Unit of measurement ('pt' for points, 'cm' for centimeters, "
                            "'mm' for millimeters, 'inches' for inches)"
                    )
                },
                "right": {
                    "type": "dict",
                    "description": (
                            "Right margin setting (optional). Supports the following keys:\n"
                            "- value: Numeric value of the right margin\n"
                            "- unit: Unit of measurement ('pt' for points, 'cm' for centimeters, "
                            "'mm' for millimeters, 'inches' for inches)"
                    )
                }
            }
        },
        "zh": {
            "function_description": "设置Word文档的页面边距属性",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "section_list": {
                    "type": "list[int | str]",
                    "description": "指定要应用的分节列表。列表中的每个元素表示一个分节标识：整数表示第几节（从 1 开始计数），字符串 'all' 表示应用到所有节。例如：[1, 2] 或 ['all']。"
                },
                "top": {
                    "type": "dict",
                    "description": (
                            "上边距配置（可选）。支持的键：\n"
                            "- value: 上边距数值\n"
                            "- unit: 计量单位（'pt' 表示磅，'cm' 表示厘米，'mm' 表示毫米，'inches' 表示英寸）"
                    )
                },
                "bottom": {
                    "type": "dict",
                    "description": (
                            "下边距配置（可选）。支持的键：\n"
                            "- value: 下边距数值\n"
                            "- unit: 计量单位（'pt' 表示磅，'cm' 表示厘米，'mm' 表示毫米，'inches' 表示英寸）"
                    )
                },
                "left": {
                    "type": "dict",
                    "description": (
                            "左边距配置（可选）。支持的键：\n"
                            "- value: 左边距数值\n"
                            "- unit: 计量单位（'pt' 表示磅，'cm' 表示厘米，'mm' 表示毫米，'inches' 表示英寸）"
                    )
                },
                "right": {
                    "type": "dict",
                    "description": (
                            "右边距配置（可选）。支持的键：\n"
                            "- value: 右边距数值\n"
                            "- unit: 计量单位（'pt' 表示磅，'cm' 表示厘米，'mm' 表示毫米，'inches' 表示英寸）"
                    )
                }
            }
        }
    })
    def set_margin(self, doc, section_list, top={}, bottom={}, left={}, right={}, *args, **kwargs):
        result = None
        for section_index in section_list:
            result = self.page_tool.set_margin(doc,section_index,top,bottom,left,right)
        return result

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set the page gutter properties of a Word document",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
               "section_list": {"type": "list[int | str]",
                                "description": "Specify the target sections as a list. Each element represents a section identifier: an integer indicates the section number (starting from 1), and the string 'all' means apply to all sections. Examples: [1, 2] or ['all']."
                                },
                "gutter": {"type": "float", "description": "Gutter width value"},
                "gutter_unit": {"type": "float", "description": "Gutter width unit, only support pt, cm, mm, inches"},
                "gutter_pos": {"type": "int", "description": "Gutter position (0: left, 1: top)"}
            }
        },
        "zh": {
            "function_description": "设置Word文档的装订线属性",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "section_list": {
                    "type": "list[int | str]",
                    "description": "指定要应用的分节列表。列表中的每个元素表示一个分节标识：整数表示第几节（从 1 开始计数），字符串 'all' 表示应用到所有节。例如：[1, 2] 或 ['all']。"
                },
                "gutter": {"type": "float", "description": "装订线宽度的值"},
                "gutter_unit": {"type": "str", "description": "装订线宽度的单位，仅支持pt，cm，mm，inches"},
                "gutter_pos": {"type": "int", "description": "装订线位置（0: 靠左, 1: 靠上）"}
            }
        }
    })
    def set_gutter(self,doc, section_list, gutter=0, gutter_unit='pt', gutter_pos=0, *args, **kwargs):
        result = None
        for section_index in section_list:
            result = self.page_tool.set_gutter(doc, section_index, gutter, gutter_unit, gutter_pos)
        return result

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set the paper size, height, and width of a Word document",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
               "section_list": {"type": "list[int | str]",
                                "description": "Specify the target sections as a list. Each element represents a section identifier: an integer indicates the section number (starting from 1), and the string 'all' means apply to all sections. Examples: [1, 2] or ['all']."
                                },
                "size": {
                    "type": "int",
                    "description": (
                            "Normal Paper size constants:\n"
                            "| 1: Tabloid, 11×17 inches | 2: Letter | 4: Legal | "
                            "5: Executive | 6: A3 | 7: A4 | 9: A5 | 11: B5 | 41: Custom size"
                    )
                },
                "width": {
                    "type": "dict",
                    "description": (
                            "Page width setting (optional). Supports the following keys:\n"
                            "- value: Numeric width value\n"
                            "- unit: Unit of measurement ('pt' for points, 'cm' for centimeters, "
                            "'mm' for millimeters, 'inches' for inches)"
                    )
                },
                "height": {
                    "type": "dict",
                    "description": (
                            "Page height setting (optional). Supports the following keys:\n"
                            "- value: Numeric height value\n"
                            "- unit: Unit of measurement ('pt' for points, 'cm' for centimeters, "
                            "'mm' for millimeters, 'inches' for inches)"
                    )
                },
                "orientation": {
                    "type": "int",
                    "description": "Page orientation (0: Portrait, 1: Landscape)"
                }
            }
        },
        "zh": {
            "function_description": "设置Word文档的纸张大小、高度和宽度",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "section_list": {
                    "type": "list[int | str]",
                    "description": "指定要应用的分节列表。列表中的每个元素表示一个分节标识：整数表示第几节（从 1 开始计数），字符串 'all' 表示应用到所有节。例如：[1, 2] 或 ['all']。"
                },
                "size": {
                    "type": "int",
                    "description": (
                            "常见纸张大小常量列表：\n"
                            "| 1: Tabloid（小报），11×17 英寸 | 2: 信纸（Letter）| 4: 法律纸 | "
                            "5: 行政纸 | 6: A3 | 7: A4 | 9: A5 | 11: B5 | 41: 自定义大小"
                    )
                },
                "width": {
                    "type": "dict",
                    "description": (
                            "页面宽度设置（可选）。支持的键：\n"
                            "- value: 宽度数值\n"
                            "- unit: 计量单位（'pt' 表示磅，'cm' 表示厘米，"
                            "'mm' 表示毫米，'inches' 表示英寸）"
                    )
                },
                "height": {
                    "type": "dict",
                    "description": (
                            "页面高度设置（可选）。支持的键：\n"
                            "- value: 高度数值\n"
                            "- unit: 计量单位（'pt' 表示磅，'cm' 表示厘米，"
                            "'mm' 表示毫米，'inches' 表示英寸）"
                    )
                },
                "orientation": {
                    "type": "int",
                    "description": "页面方向（0: 纵向, 1: 横向）"
                }
            }
        }
    })
    def set_paper(self,doc, section_list, size=None, width={}, height={}, orientation=None, *args, **kwargs):
        result = None
        for section_index in section_list:
            result = self.page_tool.set_paper(doc, section_index, size, width, height, orientation)
        return result

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set the document grid properties of a Word document",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
               "section_list": {"type": "list[int | str]",
                                "description": "Specify the target sections as a list. Each element represents a section identifier: an integer indicates the section number (starting from 1), and the string 'all' means apply to all sections. Examples: [1, 2] or ['all']."
                                },
                "layout_mode": {
                    "type": "int",
                    "description": "Layout mode (0: No grid, 1: Specify line and character grid, 2: Specify line grid only, 3: Align text to character grid)"
                },
                "lines_page": {
                    "type": "int",
                    "description": "Number of lines per page, range 1–43"
                },
                "chars_line": {
                    "type": "int",
                    "description": "Number of characters per line, range 1–48. When layout_mode is 3 (Align text to character grid), the range is 1–44"
                }
            }
        },
        "zh": {
            "function_description": "设置Word文档的文档网格属性",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "section_list": {
                    "type": "list[int | str]",
                    "description": "指定要应用的分节列表。列表中的每个元素表示一个分节标识：整数表示第几节（从 1 开始计数），字符串 'all' 表示应用到所有节。例如：[1, 2] 或 ['all']。"
                },
                "layout_mode": {"type": "int",
                                "description": "布局模式（0: 无网格, 1: 指定行和字符网格, 2: 只指定行网格, 3：文字对齐字符网络）"},
                "lines_page": {"type": "int", "description": "每页行数，范围1-43"},
                "chars_line": {"type": "int",
                               "description": "每行字符数，范围1-48。当layout_mode为3时（文字对齐字符网络），范围为1-44"}
            }
        }
    })
    def set_grid(self,doc, section_list, layout_mode=None, lines_page=None, chars_line=None, *args, **kwargs):
        result = None
        if not layout_mode:
            layout_mode = 0
        if layout_mode == 2:
            chars_line = None
        for section_index in section_list:
            result = self.page_tool.set_grid(doc, section_index, layout_mode, lines_page, chars_line)
        return result

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set the column layout of a Word document section",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
               "section_list": {
                   "type": "list[int | str]",
                   "description": "Specify the target sections as a list. Each element represents a section identifier: an integer indicates the section number (starting from 1), and the string 'all' means apply to all sections. Examples: [1, 2] or ['all']."
               },
                "column_count": {
                    "type": "int",
                    "description": "Number of columns to apply (e.g., 1 for single column, 2 for double)"
                },
                "evenly_spaced": {
                    "type": "int",
                    "description": (
                            "Whether the columns are evenly spaced:\n"
                            "-1: Enable evenly spaced columns (True)\n"
                            "0: Disable evenly spaced columns (False)"
                    )
                },
                "column_width": {
                    "type": "dict",
                    "description": (
                            "Width of the first column (optional). Supports the following keys:\n"
                            "- value: Width of the column (float)\n"
                            "- unit: Unit of measurement. Supported values:\n"
                            "  'pt' (points), 'cm' (centimeters), 'mm' (millimeters), 'inches' (inches)"
                    )
                },
                "spacing": {
                    "type": "dict",
                    "description": (
                            "Spacing between columns (optional). Supports the following keys:\n"
                            "- value: Space between columns (float)\n"
                            "- unit: Unit of measurement. Supported values:\n"
                            "  'pt' (points), 'cm' (centimeters), 'mm' (millimeters), 'inches' (inches)"
                    )
                },
                "line_between": {
                    "type": "int",
                    "description": (
                            "Whether to show a vertical line between columns:\n"
                            "-1: Show vertical separator line (True)\n"
                            "0: Do not show separator line (False)"
                    )
                }
            }
        },
        "zh": {
            "function_description": "设置Word文档某节的分栏布局",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "section_list": {
                    "type": "list[int | str]",
                    "description": "指定要应用的分节列表。列表中的每个元素表示一个分节标识：整数表示第几节（从 1 开始计数），字符串 'all' 表示应用到所有节。例如：[1, 2] 或 ['all']。"
                },
                "column_count": {
                    "type": "int",
                    "description": "分栏数（例如1表示单栏，2表示双栏）"
                },
                "evenly_spaced": {
                    "type": "int",
                    "description": (
                            "是否平均分配栏宽和间距：\n"
                            "-1：启用（平均分配）\n"
                            "0：关闭（可自定义栏宽与间距）"
                    )
                },
                "column_width": {
                    "type": "dict",
                    "description": (
                            "第一栏宽度设置（可选）。支持以下键：\n"
                            "- value：栏宽数值（float）\n"
                            "- unit：栏宽单位。支持的单位：\n"
                            "  'pt'（磅）、'cm'（厘米）、'mm'（毫米）、'inches'（英寸）"
                    )
                },
                "spacing": {
                    "type": "dict",
                    "description": (
                            "栏间距设置（可选）。支持以下键：\n"
                            "- value：间距数值（float）\n"
                            "- unit：栏间距单位。支持的单位：\n"
                            "  'pt'（磅）、'cm'（厘米）、'mm'（毫米）、'inches'（英寸）"
                    )
                },
                "line_between": {
                    "type": "int",
                    "description": (
                            "是否在栏之间显示分隔线：\n"
                            "-1：启用（显示分隔线）\n"
                            "0：关闭（不显示）"
                    )
                }
            }
        }
    })
    def set_columns(self, doc, section_list, column_count=1, evenly_spaced=0, column_width={}, spacing={},
                    line_between=0, *args, **kwargs ):
        result = None
        for section_index in section_list:
            result = self.page_tool.set_columns(doc,section_index,column_count,evenly_spaced,column_width,spacing,line_between)
        return result



if __name__ == '__main__':
    word = win32.DispatchEx("Word.Application")  # Or use Dispatch
    word.ActivePrinter = "Microsoft Print to PDF"  # Specify the printer
    word.Visible = True  # Make visible (recommended when debugging)
    word_file_path = "./file/Word_test.docx"
    word_file_path = os.path.join(ABS_DIR, word_file_path)

    # Open an existing document
    try:
        # Open the document
        doc = word.Documents.Open(word_file_path)
        pg_tool = PageTools()

        # print(pg_tool.set_margin(doc,[1],right={"value": 2.5,"unit": "cm" }))
        # print(pg_tool.set_footer_header_layout(doc, [1], header_distance={"value": 1.5,"unit": "cm"}))
        print(pg_tool.set_footer_content(doc,[1],different_first_page=0,different_odd_even=0,primary={
          "format": 0,
          "alignment": 1,
          "start": 1,
          "continue": False,
          "name": "仿宋",
          "size": 10
        }))

    except Exception as e:
        print(f"操作失败：{str(e)}")
        raise

    finally:
        if 'doc' in locals():
            doc.Close(SaveChanges=False)
        word.Quit()