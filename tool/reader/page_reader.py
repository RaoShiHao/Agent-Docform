from constant import ABS_DIR
import os,copy
from win32com.client import constants
import win32com.client as win32
from tool.file_trans import FileConverter
from tool.basetool import ContextToolsConfig,BaseTool

class PageReader(BaseTool):
    def __init__(self, pyconfig=ContextToolsConfig("/config/Tools/reader/page_reader_config.yaml")):
        self.config = pyconfig.config
        self.file_tool = FileConverter()

    def pt_to_convert(self, value, unit):
        value = float(value)
        # Speed up: read the converted value directly
        # execl = win32com.client.Dispatch("Excel.Application")
        # cm_unit = execl.CentimetersToPoints(1)
        # inches_unit = execl.InchesToPoints(1)
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

    def get_page_info(self, doc, save_dir, content_x):
        page_result_path = os.path.join(save_dir, "page_format.json")
        page_info_result = []
        for section_index in range(1, doc.Sections.Count + 1):
            section_page_result = self.get_section_properties(doc,section_index=section_index, content_x=content_x)
            if section_page_result.get("state") == "success":
                section_page_result.pop("state")
                page_info_result.append(section_page_result)
        self.file_tool.write_json_file(data=page_info_result,file_path=page_result_path)
        return page_info_result

    def __get_header_content_info(self, section):
        page_setup = section.PageSetup
        """Safely get document header information (including checking whether the header exists)."""
        result = {
            'different_first_page': page_setup.DifferentFirstPageHeaderFooter,  # -1=True, 0=False
            'different_odd_even': page_setup.OddAndEvenPagesHeaderFooter,
        }
        try:
            def safe_extract_header(header_type, name):
                """Safely extract header information (return empty data if the header does not exist)."""
                try:
                    header = section.Headers(header_type)
                    if not header.Exists:  # Critical check: whether the header actually exists
                        return None
                    range_ = header.Range
                    border = range_.Borders(win32.constants.wdBorderBottom)

                    return {
                        "paragraph": range_.Text.strip(),
                        "format":
                        {
                        "name": range_.Font.Name,
                        "size": range_.Font.Size,
                        "alignment": range_.ParagraphFormat.Alignment,
                         },
                        "border_line": border.LineStyle # 0 = no border
                    }
                except Exception as e:
                    print(f"读取页眉 {name} 失败: {str(e)}")
                    return None
            # Primary header (always attempt to read)
            if primary_info := safe_extract_header(win32.constants.wdHeaderFooterPrimary, "primary"):
                result['primary'] = primary_info

            # First-page header (read only when enabled)
            if result['different_first_page'] == -1:
                if first_info := safe_extract_header(win32.constants.wdHeaderFooterFirstPage, "first"):
                    result['first'] = first_info

            # Even-page header (read only when enabled)
            if result['different_odd_even'] == -1:
                if even_info := safe_extract_header(win32.constants.wdHeaderFooterEvenPages, "even"):
                    result['even'] = even_info

        except Exception as e:
            print(f"获取页眉信息时发生错误: {str(e)}")

        return result

    def __read_header_content_info(self, section, page_info, pop_None, *args, **kwargs):

        # Header
        header = self.__get_header_content_info(section)
        page_info["header_content"]["different_first_page"]["value"] = header.get("different_first_page")
        page_info["header_content"]["different_odd_even"]["value"] = header.get("different_odd_even")
        for key in ["primary", "first", "even"]:
            if key in header:
                # print(key)
                page_info["header_content"][key]['text']["value"] = header.get(key).get("paragraph")
                page_info["header_content"][key]['name']["value"] = header.get(key).get("format").get('name')
                page_info["header_content"][key]['size']["value"] = header.get(key).get("format").get('size')
                page_info["header_content"][key]['alignment']["value"] = header.get(key).get("format").get('alignment')
                page_info["header_content"][key]['border_line']["value"] = header.get(key).get("border_line")
            else:
                if pop_None:
                    page_info["header_content"].pop(key)
        return page_info

    def __get_footer_content_info(self, section):
        page_setup = section.PageSetup
        """Safely get document footer information (including page-number settings)."""
        result = {
            'different_first_page': page_setup.DifferentFirstPageHeaderFooter,  # -1=True, 0=False
            'different_odd_even': page_setup.OddAndEvenPagesHeaderFooter,
        }
        try:
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
                result['primary'] = primary_info

            # First-page footer (read only when enabled)
            if result['different_first_page'] == -1:
                if first_info := safe_extract_footer(win32.constants.wdHeaderFooterFirstPage, "first"):
                    result['first'] = first_info

            # Even-page footer (read only when enabled)
            if result['different_odd_even'] == -1:
                if even_info := safe_extract_footer(win32.constants.wdHeaderFooterEvenPages, "even"):
                    result['even'] = even_info

        except Exception as e:
            print(f"获取页脚信息时发生错误: {str(e)}")

        return result

    def __read_footer_content_info(self, section,page_info, pop_None, *args, **kwargs):
        # footer
        footer = self.__get_footer_content_info(section)
        page_info["footer_content"]["different_first_page"]["value"] = footer.get("different_first_page")
        page_info["footer_content"]["different_odd_even"]["value"] = footer.get("different_odd_even")
        # print(footer)
        for key in ["primary", "first", "even"]:
            if key in footer:
                page_info["footer_content"][key]['page_number']['format']["value"] = footer.get(key).get(
                    'format')
                page_info["footer_content"][key]['page_number']['start']["value"] = footer.get(key).get('start')
                page_info["footer_content"][key]['page_number']['continue']["value"] = footer.get(key).get(
                    'continue')
                page_info["footer_content"][key]['page_number']['alignment']["value"] = footer.get(key).get(
                    'alignment')
                page_info["footer_content"][key]['page_number']['name']["value"] = footer.get(key).get('name')
                page_info["footer_content"][key]['page_number']['size']["value"] = footer.get(key).get('size')
            else:
                if pop_None:
                    page_info["footer_content"].pop(key)

        return page_info

    def __get_footer_header_info(self, section):
        page_setup = section.PageSetup
        return {
            'footer_distance': page_setup.FooterDistance,
            'header_distance': page_setup.HeaderDistance
        }

    def __read_footer_header_info(self, section, page_info, *args, **kwargs):
        # Header
        footer_header_info = self.__get_footer_header_info(section)
        header_dis = footer_header_info.get("header_distance")
        # print(header)
        page_info["header_footer_layout"]["header_distance"]["value"]["pt"] = header_dis
        page_info["header_footer_layout"]["header_distance"]["value"]["cm"] = self.pt_to_convert(header_dis, "cm")
        page_info["header_footer_layout"]["header_distance"]["value"]["mm"] = self.pt_to_convert(header_dis, "mm")
        page_info["header_footer_layout"]["header_distance"]["value"]["inches"] = self.pt_to_convert(header_dis,
                                                                                                     "inches")
        footer_dis = footer_header_info.get("footer_distance")
        page_info["header_footer_layout"]["footer_distance"]["value"]["pt"] = footer_dis
        page_info["header_footer_layout"]["footer_distance"]["value"]["cm"] = self.pt_to_convert(footer_dis, "cm")
        page_info["header_footer_layout"]["footer_distance"]["value"]["mm"] = self.pt_to_convert(footer_dis, "mm")
        page_info["header_footer_layout"]["footer_distance"]["value"]["inches"] = self.pt_to_convert(footer_dis,
                                                                                                     "inches")
        return page_info

    def __get_margin_info(self,section):
        page_setup = section.PageSetup
        return {
            "TopMargin": page_setup.TopMargin,  # Top margin
            "BottomMargin": page_setup.BottomMargin,  # Bottom margin
            "LeftMargin": page_setup.LeftMargin,  # Left margin
            "RightMargin": page_setup.RightMargin,  # Right margin
        }

    def __read_margin_info(self, section, page_info, *args, **kwargs):
        # Read page properties
        page_setup = section.PageSetup
        margin = {
                "top": page_setup.TopMargin,  # Top margin
                "bottom": page_setup.BottomMargin,  # Bottom margin
                "left": page_setup.LeftMargin,  # Left margin
                "right": page_setup.RightMargin,  # Right margin
            }
        # Fill in page margin properties
        for key, value in margin.items():
            page_info["margin"][key]["value"]["pt"] = value
            page_info["margin"][key]["value"]["cm"] = self.pt_to_convert(value, "cm")
            page_info["margin"][key]["value"]["mm"] = self.pt_to_convert(value, "mm")
            page_info["margin"][key]["value"]["inches"] = self.pt_to_convert(value, "inches")
        return page_info

    def __get_paper_info(self,section):
        page_setup = section.PageSetup
        return {
            "PaperSize": page_setup.PaperSize,  # Paper size
            "PageWidth": page_setup.PageWidth,  # Page width
            "PageHeight": page_setup.PageHeight,  # Page height
            "Orientation": page_setup.Orientation,  # Page orientation (0: portrait, 1: landscape)
        }

    def __read_paper_info(self, section, page_info, *args, **kwargs):
        page_setup = section.PageSetup
        # Fill in paper values
        page_info["paper"]["size"]["value"] = page_setup.PaperSize  # Paper size
        page_info["paper"]["orientation"]["value"] = page_setup.Orientation  # Page orientation (0: portrait, 1: landscape)
        # Paper size
        paper = {
            "width": page_setup.PageWidth,  # Page width
            "height": page_setup.PageHeight,  # Page height
        }
        for key, value in paper.items():
            page_info["paper"][key]["value"]["pt"] = value
            page_info["paper"][key]["value"]["cm"] = self.pt_to_convert(value, "cm")
            page_info["paper"][key]["value"]["mm"] = self.pt_to_convert(value, "mm")
            page_info["paper"][key]["value"]["inches"] = self.pt_to_convert(value, "inches")
        return page_info

    def __get_grid_info(self,section):
        page_setup = section.PageSetup
        lay_mode = page_setup.LayoutMode  # Layout mode
        if lay_mode == 0:
            return {
                "LayoutMode": page_setup.LayoutMode,  # Layout mode
            }
        elif lay_mode == 2:
            return  {
            "LayoutMode": page_setup.LayoutMode,  # Layout mode
            "LinesPage": page_setup.LinesPage,  # Lines per page
        }
        else:
            return {
                "LayoutMode": page_setup.LayoutMode,  # Layout mode
                "LinesPage": page_setup.LinesPage,  # Lines per page
                "CharsLine": page_setup.CharsLine,  # Characters per page
            }

    def __read_grid_info(self, section, page_info, *args,**kwargs):
        page_setup = section.PageSetup
        lay_mode = page_setup.LayoutMode  # Layout mode
        # Fill in layout values
        page_info["grid"]["layout_mode"]["value"] = page_setup.LayoutMode  # Layout mode
        page_info["grid"]["lines_page"]["value"] = page_setup.LinesPage  # Lines per page
        page_info["grid"]["chars_line"]["value"] = page_setup.CharsLine  # Characters per page

        if lay_mode == 0:  # Layout mode
            page_info["grid"].pop("lines_page")
            page_info["grid"].pop("chars_line")
        if lay_mode == 2:
            page_info["grid"].pop("chars_line")
        return page_info

    def __get_column_info(self,section):
        # Get the TextColumns property for this section
        page_setup = section.PageSetup
        text_columns = page_setup.TextColumns
        return {
                "column_count": text_columns.Count,  # Number of columns
                "spacing": text_columns.Spacing,  # Column spacing
                "evenly_spaced": text_columns.EvenlySpaced,  # Whether columns are evenly distributed
                "line_between": text_columns.LineBetween,  # Whether to show column lines
                # "first_column_width": text_columns(1).Width  # width of the first column
            }

    def __read_column_info(self,section,page_info,*args,**kwargs):
        text_columns = section.PageSetup.TextColumns
        # columns: multi-column layout info
        page_info["columns"]["column_count"]["value"] = text_columns.Count  # Number of columns
        page_info["columns"]["evenly_spaced"]["value"] = text_columns.EvenlySpaced  # Whether columns are evenly distributed
        page_info["columns"]["line_between"]["value"] = text_columns.LineBetween  # Whether to show column lines
        columns = {
            "spacing": text_columns.Spacing,  # Column spacing
            "first_column_width": text_columns(1).Width  # Column width
        }
        for key, value in columns.items():
            page_info["columns"][key]["value"]["pt"] = value
            page_info["columns"][key]["value"]["cm"] = self.pt_to_convert(value, "cm")
            page_info["columns"][key]["value"]["mm"] = self.pt_to_convert(value, "mm")
            page_info["columns"][key]["value"]["inches"] = self.pt_to_convert(value, "inches")

        return page_info

    def __get_gutter_info(self,section):
        page_setup = section.PageSetup
        return {
            "Gutter": page_setup.Gutter,  # Gutter width
            "GutterPosition": page_setup.GutterPos,  # Gutter position (0: left, 1: top)
        }

    def __read_gutter_info(self, section, page_info,*args,**kwargs):
        page_setup = section.PageSetup
        page_info["gutter"]["gutter"]["value"]['pt'] = page_setup.Gutter  # Gutter width
        page_info["gutter"]["gutter"]["value"]['cm'] = self.pt_to_convert(page_setup.Gutter ,"cm")# Gutter width
        page_info["gutter"]["gutter"]["value"]['mm'] = self.pt_to_convert(page_setup.Gutter,"mm")# Gutter width
        page_info["gutter"]["gutter"]["value"]['inches'] =self.pt_to_convert(page_setup.Gutter ,"inches")# Gutter width
        page_info["gutter"]["gutter_pos"]["value"] = page_setup.GutterPos  # Gutter width
        return page_info

    def __get_section_x_pages(self, section, start_page, end_page, x=1):
        """Get text content of the first x pages of a section (paragraph-granular; not a strict page cut).
        :param doc: Word document object
        :param section_index: Section index (1-based)
        :param start_page: Starting physical page number of the section
        :param end_page: Ending physical page number of the section
        :param x: First x pages
        :return: str"""
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

    def get_page_properties(self, doc, section_index, params_list=[]):
        """Read page properties of a Word document.
                :param doc: Word document object
                 section_index: Section index
                 params_list Property list for the section; if omitted, returns all info rather than empty
                :return: Operation result (status and page property info)"""
        properties = {}
        attribution_dict = {
            "margin": self.__get_margin_info,
            "gutter": self.__get_gutter_info,
            "paper": self.__get_paper_info,
            "grid": self.__get_grid_info,
            "columns": self.__get_column_info,
            "header_footer_layout": self.__get_footer_header_info,
            "footer_content": self.__get_footer_content_info,
            "header_content": self.__get_header_content_info
        }
        if not params_list:
            params_list = attribution_dict.keys()
        try:
            # Get the page section object
            section = doc.Sections(section_index)
            for params in params_list:
                # Call parameters are supported
                if params in attribution_dict:
                    attribution_info_get_tool = attribution_dict.get(params)
                    attribution_info = attribution_info_get_tool(section)
                    properties[params] = attribution_info
            # Return success result
            return {"state": "success", "properties": properties}

        except Exception as e:
            # Catch exceptions and return error information
            return {"state": "false", "exception": str(e)}

    def get_section_properties(self, doc, section_index, params_list=[], content_x=1):
        """Get formatting properties of one document section plus text from the first x pages.
        :param doc: Word document object
         section_index: Section index
         content_x = 1 Number of pages of content to fetch
        :return: Operation result (status and page property info)"""
        try:
            # Get the page setup object
            section = doc.Sections(section_index)
            properties = self.get_page_properties(doc,section_index,params_list)

            rng = section.Range
            # Start page number
            start_rng = rng.Duplicate
            start_rng.Collapse(win32.constants.wdCollapseStart)
            start_page = start_rng.Information(win32.constants.wdActiveEndPageNumber)

            # End page number
            end_rng = rng.Duplicate
            end_rng.Collapse(win32.constants.wdCollapseEnd)
            end_page = end_rng.Information(win32.constants.wdActiveEndPageNumber)

            content = self.__get_section_x_pages(section, start_page, end_page, x=content_x)
            # Return success result
            return {"state": "success", "section_index": section_index,
                    "section_range": {"start": start_page, "end": end_page-1}, "page_format": properties,
                    "content": content}

        except Exception as e:
            # Catch exceptions and return error information
            return {"state": "false", "exception": str(e)}

    def read_page_properties(self,doc, section_index, params_list=[], language = 'zh', pop_None = True, *args,**kwargs):
        try:
            section_index = int(section_index)
            # Get the page setup object
            section = doc.Sections(section_index)
            attribution_dict = {
                "margin": self.__read_margin_info,
                "gutter": self.__read_gutter_info,
                "paper": self.__read_paper_info,
                "grid": self.__read_grid_info,
                "columns": self.__read_column_info,
                "header_footer_layout": self.__read_footer_header_info,
                "footer_content": self.__read_footer_content_info,
                "header_content": self.__read_header_content_info
            }
            # Load the read template
            template = self.config.get("properties_template")
            if language in['zh','en']:
                page_info = copy.deepcopy(template.get(language))
            else:
                page_info = copy.deepcopy(template.get("zh"))
                print("Default Using Chinese")

            if not params_list:
                # No property scope specified; read all by default
                params_list = list(attribution_dict.keys())
            else:
                # After a property scope is specified, remove unused keys from the template
                for attribution in attribution_dict.keys():
                    if attribution not in params_list:
                        page_info.pop(attribution)

            # print(page_info)
            # Fetch each property to read in order
            for params in params_list:
                # Call parameters are supported
                if params in attribution_dict:
                    # print("=" * 50)
                    # print(params)
                    attribution_info_read_tool = attribution_dict.get(params)
                    page_info = attribution_info_read_tool(section,page_info,pop_None)
                    # print(page_info)

            # Return success result
            return {"state": "success", "properties":  page_info}

        except Exception as e:
            # Catch exceptions and return error information
            return {"state": "false","section_index":section_index, "exception": str(e)}


if __name__ == '__main__':

    word = win32.DispatchEx("Word.Application")
    word.Visible = True  # Make visible (recommended when debugging)
    word_file_path = "./file/Word_test.docx"
    # word_file_path = "./file/Base.docx"
    # word_file_path = "./file/ch_gov.doc"


    word_file_path = os.path.join(ABS_DIR, word_file_path)
    # Open an existing document
    try:
        # Open the document
        doc = word.Documents.Open(word_file_path)
        reader_tool = PageReader()

        print(reader_tool.read_page_properties(doc,1,['grid']))

    except Exception as e:
        print(f"操作失败：{str(e)}")
        raise

    finally:
        # Ensure resources are cleaned up
        if 'doc' in locals():
            doc.Close(SaveChanges=False)
        word.Quit()