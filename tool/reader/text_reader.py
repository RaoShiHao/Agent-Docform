from constant import ABS_DIR
import os,copy,re
from win32com.client import constants
import win32com
from tool.file_trans import FileConverter
from tool.basetool import ContextToolsConfig,BaseTool

class TextReader(BaseTool):
    def __init__(self, pyconfig=ContextToolsConfig("/config/Tools/reader/text_reader_config.yaml")):
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

    def color_int_to_hex(self, color_int: int) -> str:
        # Extract BGR components
        b = (color_int >> 16) & 0xFF  # Blue
        g = (color_int >> 8) & 0xFF  # Green
        r = color_int & 0xFF  # Red
        # Combine into RGB
        return "#{:02X}{:02X}{:02X}".format(r, g, b)

    def __get_base_font_info(self,range_obj,*args,**kwargs):
        font = range_obj.Font
        return {
            "Name": font.Name,  # Font name
            "NameAscii": font.NameAscii,  # Western (ASCII) font
            "NameFarEast": font.NameFarEast,  # Chinese font
            "Size": font.Size,  # Font size
            "Bold": font.Bold,  # Whether bold
            "Italic": font.Italic,  # Whether italic
            "Underline": font.Underline,  # Underline style
            "Color": self.color_int_to_hex(font.Color),  # Font color
            "HighlightColor": font.Shading.BackgroundPatternColor # Background highlight color
        }

    def __read_base_font_info(self, range_obj, text_info, *args,**kwargs):
        font_fmt = range_obj.Font
        text_info["base_font"]["NameAscii"]["value"] = font_fmt.NameAscii
        text_info["base_font"]["NameFarEast"]["value"] = font_fmt.NameFarEast
        text_info["base_font"]["Name"]["value"] = font_fmt.Name
        text_info["base_font"]["Size"]["value"] = font_fmt.Size
        text_info["base_font"]["Bold"]["value"] = font_fmt.Bold
        text_info["base_font"]["Italic"]["value"] = font_fmt.Italic
        text_info["base_font"]["Underline"]["value"] = font_fmt.Underline
        text_info["base_font"]["Color"]["value"] = self.color_int_to_hex(font_fmt.Color)
        text_info["base_font"]["HighlightColorIndex"]["value"] = range_obj.HighlightColorIndex
        return text_info

    def __get_advanced_font_info(self,range_obj,*args,**kwargs):
        font = range_obj.Font
        return {
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

    def __read_advanced_font_info(self, range_obj, text_info, *args,**kwargs):
        font_fmt = range_obj.Font
        # 10 advanced_font font properties
        text_info["advanced_font"]["StrikeThrough"]["value"] = font_fmt.StrikeThrough
        text_info["advanced_font"]["Subscript"]["value"] = font_fmt.Subscript
        text_info["advanced_font"]["Superscript"]["value"] = font_fmt.Superscript
        text_info["advanced_font"]["AllCaps"]["value"] = font_fmt.AllCaps
        text_info["advanced_font"]["Spacing"]["value"] = font_fmt.Spacing
        text_info["advanced_font"]["Scaling"]["value"] = font_fmt.Scaling
        text_info["advanced_font"]["Emboss"]["value"] = font_fmt.Emboss
        text_info["advanced_font"]["Engrave"]["value"] = font_fmt.Engrave
        text_info["advanced_font"]["Shadow"]["value"] = font_fmt.Shadow
        text_info["advanced_font"]["SmallCaps"]["value"] = font_fmt.SmallCaps
        return text_info
    def __get_outlinelevel_info(self, range, *args,**kwargs):
        fmt = range.ParagraphFormat
        return {
            # Outline level
            "outlinelevel": fmt.OutlineLevel,  # Outline level (1-10)
        }

    def __read_outlinelevel_info(self, range_obj, text_info, *args,**kwargs):
        
        para_fmt = range_obj.ParagraphFormat
        text_info['outlinelevel']["value"] = para_fmt.OutlineLevel
        return text_info

    def __get_alignment_info(self, range_obj, *args, **kwargs):
        fmt = range_obj.ParagraphFormat
        align_key = {
            0: "left", 1: "center", 2: "right",
            3: "justify", 4: "distribute"
        }.get(fmt.Alignment, "unknown")

        return {
            # Alignment
            "alignment": align_key,
        }

    def __read_alignment_info(self, range_obj, text_info, *args,**kwargs):
        para_fmt = range_obj.ParagraphFormat
        # paragraph properties
        align_key = {
            0: "left", 1: "center", 2: "right",
            3: "justify", 4: "distribute"
        }.get(para_fmt.Alignment, "unknown")
        text_info['alignment']["value"] = align_key
        return text_info

    def __get_pagination_control_info(self,  range_obj, *args, **kwargs):
        fmt = range_obj.ParagraphFormat
        return {
                "widow_control": fmt.WidowControl,  # Widow/orphan control
                "keep_with_next": fmt.KeepWithNext,  # Keep with next
                "keep_together": fmt.KeepTogether,  # Keep lines together
                "page_break_before": fmt.PageBreakBefore  # Page break before
        }

    def __read_pagination_control_info(self,range_obj, text_info, *args,**kwargs):
        
        para_fmt = range_obj.ParagraphFormat
        # pagination_control
        text_info["pagination_control"]['widow_control']['value'] = para_fmt.WidowControl
        text_info["pagination_control"]['keep_with_next']['value'] = para_fmt.KeepWithNext
        text_info["pagination_control"]['keep_together']['value'] = para_fmt.KeepTogether
        text_info["pagination_control"]['page_break_before']['value'] = para_fmt.PageBreakBefore
        return text_info
    
    def __get_spacing_info(self,  range_obj, *args, **kwargs):
        fmt = range_obj.ParagraphFormat
        return {
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
        }
    
    def __read_spacing_info(self,range_obj, text_info, *args,**kwargs):
        para_fmt = range_obj.ParagraphFormat
        
        # spacing
        text_info["spacing"]['line_spacing']["spacing_rule"]['value'] = {
            0: "single", 1: "1.5x", 2: "double",
            4: "exact", 5: "multiple"
        }.get(para_fmt.LineSpacingRule, "custom")
        text_info["spacing"]['line_spacing']["spacing_value"]['value'] = para_fmt.LineSpacing

        space_before = para_fmt.SpaceBefore
        text_info["spacing"]['before_spacing']['value']["pt"] = self.pt_to_convert(space_before, "pt")
        text_info["spacing"]['before_spacing']['value']["cm"] = self.pt_to_convert(space_before, "cm")
        text_info["spacing"]['before_spacing']['value']["mm"] = self.pt_to_convert(space_before, "mm")
        text_info["spacing"]['before_spacing']['value']["inches"] = self.pt_to_convert(space_before,"inches")
        space_after = para_fmt.SpaceAfter
        text_info["spacing"]['after_spacing']['value']["pt"] = self.pt_to_convert(space_after, "pt")
        text_info["spacing"]['after_spacing']['value']["cm"] = self.pt_to_convert(space_after, "cm")
        text_info["spacing"]['after_spacing']['value']["mm"] = self.pt_to_convert(space_after, "mm")
        text_info["spacing"]['after_spacing']['value']["inches"] = self.pt_to_convert(space_after,"inches")
        
        return text_info
    def __get_indent_info(self, range_obj, *args, **kwargs):

        fmt = range_obj.ParagraphFormat
        return {
                "left": fmt.LeftIndent,  # Left indent (pt)
                "right": fmt.RightIndent,  # Right indent (pt)
                "first_line": fmt.FirstLineIndent  # First-line indent (positive) or hanging indent (negative)
        }
    
    def __read_indent_info(self,range_obj,text_info, *args,**kwargs):
        
        # indent
        character_unit = range_obj.Font.Size
        indent = self.__get_indent_info(range_obj)
        left_indent = indent.get("left")
        right_indent = indent.get("right")
        firstline_indent = indent.get("first_line")

        if left_indent >= 0:
            text_info["indent"]['left_indent']['hanging'] = 0
        else:
            text_info["indent"]['left_indent']['hanging'] = -1
        left_indent = abs(left_indent)
        text_info["indent"]['left_indent']['value']["pt"] = self.pt_to_convert(left_indent, "pt")
        text_info["indent"]['left_indent']['value']["cm"] = self.pt_to_convert(left_indent, "cm")
        text_info["indent"]['left_indent']['value']["mm"] = self.pt_to_convert(left_indent, "mm")
        text_info["indent"]['left_indent']['value']["inches"] = self.pt_to_convert(left_indent,"inches")
        text_info["indent"]['left_indent']['value']["character"] = round(left_indent / character_unit)

        if right_indent >= 0:
            text_info["indent"]['right_indent']['hanging'] = 0
        else:
            text_info["indent"]['right_indent']['hanging'] = -1

        right_indent = abs(right_indent)
        text_info["indent"]['right_indent']['value']["pt"] = self.pt_to_convert(right_indent, "pt")
        text_info["indent"]['right_indent']['value']["cm"] = self.pt_to_convert(right_indent, "cm")
        text_info["indent"]['right_indent']['value']["mm"] = self.pt_to_convert(right_indent, "mm")
        text_info["indent"]['right_indent']['value']["inches"] = self.pt_to_convert(right_indent,"inches")
        text_info["indent"]['right_indent']['value']["character"] = round(right_indent / character_unit)

        if firstline_indent >= 0:
            text_info["indent"]['firstline_indent']['hanging'] = 0
        else:
            text_info["indent"]['firstline_indent']['hanging'] = -1
        firstline_indent = abs(firstline_indent)
        text_info["indent"]['firstline_indent']['value']["pt"] = self.pt_to_convert(firstline_indent,
                                                                                                    "pt")
        text_info["indent"]['firstline_indent']['value']["cm"] = self.pt_to_convert(firstline_indent,
                                                                                                    "cm")
        text_info["indent"]['firstline_indent']['value']["mm"] = self.pt_to_convert(firstline_indent,
                                                                                                    "mm")
        text_info["indent"]['firstline_indent']['value']["inches"] = self.pt_to_convert(
            firstline_indent, "inches")
        text_info["indent"]['firstline_indent']['value']["character"] = round(
            firstline_indent / character_unit)

        return text_info

    def get_text_properties(self, doc, paragraph_index, params_list=[], *args,**kwargs):
        properties = {}
        attribution_dict = {
            "base_font": self.__get_base_font_info,
            "advanced_font": self.__get_advanced_font_info,
            "outlinelevel": self.__get_outlinelevel_info,
            "alignment": self.__get_alignment_info,
            "pagination_control": self.__get_pagination_control_info,
            "spacing": self.__get_spacing_info,
            "indent": self.__get_indent_info
        }
        if not params_list:
            params_list = attribution_dict.keys()
        try:
            if paragraph_index == 0:
                range_obj = doc.Application.Selection.Range
            elif paragraph_index >0:
                range_obj = doc.Paragraphs(paragraph_index).Range
            else:
                print("paragraph index must >= 0!")
                raise
            for params in params_list:
                # Call parameters are supported
                if params in attribution_dict:
                    attribution_info_get_tool = attribution_dict.get(params)
                    attribution_info = attribution_info_get_tool(range_obj)
                    properties[params] = attribution_info
            # Return success result
            return {"state": "success", "properties": properties}
        except Exception as e:
            # Catch exceptions and return error information
            return {"state": "false", "exception": str(e)}

    def __get_target_ranges_by_index(self, doc, paragraph_index, start, length):
        """Locate a Range object by character index (compatible with tables and body text)."""
        try:
            if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                print(f"[Error] Paragraph index {paragraph_index} out of range.")
                return None

            para = doc.Paragraphs(paragraph_index)
            rng = para.Range

            # Check whether the paragraph is inside a table
            wdWithInTable = 12
            if rng.Information(wdWithInTable):
                # Re-fetch paragraphs inside the cell to avoid empty Characters
                cell = rng.Cells(1)
                inner_para = cell.Range.Paragraphs(1)
                rng = inner_para.Range

            para_len = rng.Characters.Count
            if start < 1 or start > para_len or start + length - 1 > para_len:
                print(
                    f"[Error] Invalid char range: paragraph has {para_len} chars, requested {start}-{start + length - 1}.")
                return None

            target_range = rng.Characters(start).Duplicate
            target_range.End = rng.Characters(start + length - 1).End
            return target_range

        except Exception as e:
            print(f"[Error] get_target_range_by_index failed: {e}")
            return None

    def __get_target_ranges_by_text(self, doc, paragraph_index, target_text, match_index=1):
        """Locate a Range object by text match (compatible with tables and body text).
        match_index may be int or "all"."""
        try:
            if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                print(f"[Error] Paragraph index {paragraph_index} out of range.")
                return None

            para = doc.Paragraphs(paragraph_index)
            rng = para.Range

            # Check whether inside a table
            wdWithInTable = 12
            if rng.Information(wdWithInTable):
                cell = rng.Cells(1)
                inner_para = cell.Range.Paragraphs(1)
                rng = inner_para.Range

            matches = []
            current_start = rng.Start

            while current_start < rng.End:
                search_rng = rng.Duplicate
                search_rng.Start = current_start
                find = search_rng.Find
                find.Text = target_text
                find.Forward = True
                find.MatchCase = False

                if not find.Execute():
                    break

                matches.append(search_rng.Duplicate)
                current_start = search_rng.End

                if isinstance(match_index, int) and len(matches) >= match_index:
                    break

            if not matches:
                print(f"[Info] No match for text '{target_text}' in paragraph {paragraph_index}")
                return None

            if match_index == "all":
                return matches
            elif isinstance(match_index, int) and match_index <= len(matches):
                return matches[match_index - 1]
            else:
                return None

        except Exception as e:
            print(f"[Error] get_target_range_by_text failed: {e}")
            return None

    def __get_target_ranges_by_regex(self, doc, paragraph_index, regex_pattern, match_index=1):
        """Locate a Range object by regex match (compatible with tables and body text).
        match_index may be int or "all"."""
        try:
            if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                print(f"[Error] Paragraph index {paragraph_index} out of range.")
                return None

            para = doc.Paragraphs(paragraph_index)
            rng = para.Range

            # Check whether inside a table
            wdWithInTable = 12
            if rng.Information(wdWithInTable):
                cell = rng.Cells(1)
                inner_para = cell.Range.Paragraphs(1)
                rng = inner_para.Range

            text = rng.Text
            matches = list(re.finditer(regex_pattern, text))
            if not matches:
                print(f"[Info] No regex match for '{regex_pattern}' in paragraph {paragraph_index}")
                return None

            # Support match_index = int or "all"
            result_ranges = []
            for i, m in enumerate(matches):
                start = rng.Start + m.start()
                end = rng.Start + m.end()
                result_ranges.append(doc.Range(Start=start, End=end))

            if match_index == "all":
                return result_ranges
            elif isinstance(match_index, int) and 1 <= match_index <= len(result_ranges):
                return result_ranges[match_index - 1]
            else:
                return None

        except Exception as e:
            print(f"[Error] get_target_range_by_regex failed: {e}")
            return None

    def get_target_ranges(self, doc, paragraph_list, mode, text_params=None):
        """Unified interface for obtaining Range objects.

        Args:
            doc: Word document object
            paragraph_list: List of paragraph indices or a list containing 'all'
            mode: Location mode; one of 'index', 'text', 'regex'
            params: Parameter dict; fields depend on the mode

        Returns:
            Dictionary containing the target Range object(s) and info"""
        if text_params is None:
            text_params = {}
        try:
            if mode == 'index':
                return self.__get_target_ranges_by_index(
                    doc, paragraph_list,
                    text_params.get('start', 1),
                    text_params.get('length', 1)
                )
            elif mode == 'text':
                return self.__get_target_ranges_by_text(
                    doc, paragraph_list,
                    text_params.get('text', ''),
                )
            elif mode == 'regex':
                return self.__get_target_ranges_by_regex(
                    doc, paragraph_list,
                    text_params.get('pattern', ''),
                )
            else:
                return {
                    "status": "error",
                    "message": f"Unsupported mode: {mode}",
                    "ranges": [],
                    "target_ranges": []
                }

        except Exception as e:
            return {
                "status": "error",
                "message": f"Error in range selection: {str(e)}",
                "ranges": [],
                "target_ranges": []
            }

    def read_text_properties(self, doc, paragraph_index, params_list=[], language='zh',*args,**kwargs):
        # print(function_list)
        map_dict = {
            "partial_advanced_font": "advanced_font",
            "partial_base_font": "base_font"
        }
        params_list = [map_dict.get(item) for item in params_list if item in map_dict]
        # print(params_list)
        attribution_dict = {
            "base_font": self.__read_base_font_info,
            "advanced_font": self.__read_advanced_font_info,
            "outlinelevel": self.__read_outlinelevel_info,
            "alignment": self.__read_alignment_info,
            "pagination_control": self.__read_pagination_control_info,
            "spacing": self.__read_spacing_info,
            "indent": self.__read_indent_info,
        }
        # Load the read template
        template = self.config.get("properties_template")
        if language in ['zh', 'en']:
            text_info = copy.deepcopy(template.get(language))
        else:
            text_info = copy.deepcopy(template.get("zh"))
            print("Default Using Chinese")
        try:
            # String conversion
            paragraph_index = int(paragraph_index)
            if paragraph_index == 0:
                range_obj = doc.Application.Selection.Range
            elif paragraph_index > 0:
                range_obj = doc.Paragraphs(paragraph_index).Range
            else:
                print("paragraph index must >= 0!")
                raise
            if not params_list:
                # No property scope specified; read all by default
                params_list = list(attribution_dict.keys())
            else:
                # After a property scope is specified, remove unused keys from the template
                for attribution in attribution_dict.keys():
                    if attribution not in params_list:
                        text_info.pop(attribution)

            # Fetch each property to read in order
            for params in params_list:
                # Call parameters are supported
                if params in attribution_dict:
                    attribution_info_read_tool = attribution_dict.get(params)
                    text_info = attribution_info_read_tool(range_obj, text_info)

            # Return success result
            return {"state": "success", "properties": text_info}

        except Exception as e:
            # Catch exceptions and return error information
            return {"state": "false", "paragraph_index": paragraph_index, "exception": str(e)}

    def __read_partial_advanced_font_info(self,doc,paragraph_index, mode, params,text_info,*args,**kwargs):
        range_obj = self.get_target_ranges(doc,paragraph_index,mode,params)
        return self.__read_advanced_font_info(range_obj,text_info)

    def __read_partial_base_font_info(self,doc,paragraph_index, mode, text_params,text_info,*args,**kwargs):
        range_obj = self.get_target_ranges(doc,paragraph_index,mode,text_params)
        return self.__read_base_font_info(range_obj,text_info)

    def read_partial_text_properties(self, doc, paragraph_index, mode_list, function_list=[], language = 'zh'):
        results = []
        for mode in mode_list:
            results.append({"mode":mode.get("mode"),"params":mode.get("params"),
                            "properties": self.__single_read_partial_text_properties(doc,paragraph_index,mode.get("mode"),mode.get("params"),
                                                                                     function_list,language)})
        return results

    def __single_read_partial_text_properties(self, doc, paragraph_index, mode, text_params, function_list=[], language = 'zh'):
        # print(function_list)
        map_dict = {
            "partial_advanced_font":"advanced_font",
            "partial_base_font":"base_font"
        }
        function_list = [map_dict.get(item) for item in function_list]
        attribution_dict = {
            "advanced_font": self.__read_partial_advanced_font_info,
            "base_font": self.__read_partial_base_font_info,
        }
        # Load the read template
        template = self.config.get("partial_text_properties_template")
        if language in ['zh', 'en']:
            text_info = copy.deepcopy(template.get(language))
        else:
            text_info = copy.deepcopy(template.get("zh"))
            print("Default Using Chinese")
        try:
            # String conversion
            paragraph_index = int(paragraph_index)
            if not function_list:
                # No property scope specified; read all by default
                function_list = list(attribution_dict.keys())
            else:
                # After a property scope is specified, remove unused keys from the template
                for attribution in attribution_dict.keys():
                    if attribution not in function_list:
                        # print(attribution)
                        text_info.pop(attribution)

            # Fetch each property to read in order
            for function_params in function_list:
                # Call parameters are supported
                if function_params in attribution_dict:
                    attribution_info_read_tool = attribution_dict.get(function_params)
                    text_info = attribution_info_read_tool(doc, paragraph_index, mode, text_params, text_info)
            # Return success result
            return {"state": "success", "properties": text_info}

        except Exception as e:
            # Catch exceptions and return error information
            return {"state": "false", "paragraph_index": paragraph_index, "exception": str(e)}

    def get_paragraph_info(self, doc, index,only_key = 'all'):
        """Get all formatting properties of the specified paragraph.
        :param index: Paragraph index (1-based)
        :return: Dictionary of paragraph properties (or error info on failure)"""
        try:
            # Validate the index
            if index < 1 or index > doc.Paragraphs.Count:
                raise ValueError(f"段落索引 {index} 超出有效范围（1-{doc.Paragraphs.Count}")
            # Get the paragraph object
            paragraph = doc.Paragraphs(index)
            font = paragraph.Range.Font
            fmt = paragraph.Range.ParagraphFormat
            # Get the paragraph start page number (start position)
            start_page = paragraph.Range.Information(3)
            # Get the paragraph end page number (end position)
            end_range = paragraph.Range.Duplicate  # Duplicate the Range object to avoid modifying the original paragraph
            end_range.Collapse(Direction=constants.wdCollapseEnd)  # Collapse to the end of the paragraph
            # end_range.Collapse(Direction=0)  # Collapse to the end of the paragraph
            end_page = end_range.Information(3)

            is_table = paragraph.Range.Tables.Count
            if is_table:
                return {"state": "success","properties": None,}

            if only_key == 'all':
                # Build the properties dictionary
                properties = {
                    "index": index,
                    # Basic properties
                    "text": paragraph.Range.Text.strip(),
                    "style": paragraph.Style.NameLocal,
                    # Outline level
                    "outlinelevel": fmt.OutlineLevel,  # Outline level (1-10)
                    # range page spans multiple pages
                    "page_range": {"start_page": start_page, "end_page": end_page},
                    "font": {
                        "bold": font.Bold == -1,
                        "name": font.Name,
                        "size": font.Size,
                        "italic": font.Italic == -1,
                    },
                    "spacing":
                    {
                        "spacing_rule":fmt.LineSpacingRule,
                        "spacing_value":fmt.LineSpacing,
                        "spacing_before": fmt.SpaceBefore,
                        "spacing_after":fmt.SpaceAfter
                    },
                    "indent":
                    {
                    "firstline_indent": fmt.FirstLineIndent,
                    "right_indent": fmt.RightIndent,
                    "left_indent": fmt.LeftIndent
                    }
                }
            else:
                properties = {
                "index": index,
                # Basic properties
                "text": paragraph.Range.Text.strip(),
                "page_range": {"start_page": start_page, "end_page": end_page},
            }
            return {
                "state": "success",
                "properties": properties,
            }
        except Exception as e:
            return {
                "state": "false",
                "properties": None,
                "exception": str(e)
            }

    def get_paragraphs_format(self,doc):
        formats = {}
        paragraph_num = doc.Paragraphs.Count
        for index in range(paragraph_num):
            paragraph_index = index+1
            format = self.get_format(doc,paragraph_index)
            if format.get("state") == "success":
                style_name = doc.Paragraphs(paragraph_index).Range.Text
                style_name = style_name.replace("\r","")
                style_name = style_name.replace("\x07", "")
                # Do not add info for empty paragraphs
                if style_name == "":
                    continue
                settings = format.get("properties")
                formats[str(paragraph_index)] = settings
        return formats

    def get_format(self, doc, paragraph_index, *args,**kwargs):
        properties = {}
        attribution_dict = {
            "base_font": self.__get_base_font_info,
            "advanced_font": self.__get_advanced_font_info,
            "outlinelevel": self.__get_outlinelevel_info,
            "alignment": self.__get_alignment_info,
            "pagination_control": self.__get_pagination_control_info,
            "spacing": self.__get_spacing_info,
            "indent": self.__get_indent_info
        }
        try:
            if paragraph_index == 0:
                range_obj = doc.Application.Selection.Range
            elif paragraph_index >0:
                range_obj = doc.Paragraphs(paragraph_index).Range
            else:
                print("paragraph index must >= 0!")
                raise

            for attribution,get_property_function in attribution_dict.items():
                properties[attribution] = get_property_function(range_obj)
            # Return success result
            return {"state": "success", "properties": properties}
        except Exception as e:
            # Catch exceptions and return error information
            return {"state": "false", "exception": str(e)}
if __name__ == '__main__':

    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = True  # Make visible (recommended when debugging)
    # word_file_path = "./file/Word_test.docx"
    word_file_path = "./file/Base.docx"
    # word_file_path = "./file/ch_gov.doc"
    # word_file_path = "experiment/Base/Base.docx"
    word_file_path = os.path.join(ABS_DIR, word_file_path)
    # Open an existing document
    try:
        # Open the document
        doc = word.Documents.Open(word_file_path)
        reader_tool = TextReader()
        # print(reader_tool.get_text_properties(doc,3,["base_font"]))
        # print(reader_tool.read_text_properties(doc, 1, ["base_font", "alignment"]))
        # location_list = [i + 1 for i in range(doc.Paragraphs.Count)]
        # print(location_list)
        # print(reader_tool.get_paragraph_info(doc, 1,))
        # print(reader_tool.get_paragraph_info(doc, 10, ))

        # print(reader_tool.read_text_properties(doc,1,[]))

        # print(reader_tool.read_partial_text_properties(doc,1,"text",{"text":"气"},[]))
        print(reader_tool.read_partial_text_properties(doc, 62, [{'mode': 'text', 'params': {'text': '血', 'match_index': 1}}],['partial_base_font']))


    except Exception as e:
        print(f"操作失败：{str(e)}")
        raise

    finally:
        # Ensure resources are cleaned up
        if 'doc' in locals():
            doc.Close(SaveChanges=False)
        word.Quit()