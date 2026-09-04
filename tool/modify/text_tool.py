from tool.basetool import BaseTool
from constant import ABS_DIR
import os, re
import yaml
from win32com.client import constants
import win32com.client as win32
import win32com
from tool.basetool import ContextToolsConfig

class TextTools(BaseTool):
    def __init__(self, pyconfig=ContextToolsConfig("/config/Tools/TextToolsConfig.yaml")):
        self.config = pyconfig.config
        self.name = self.config.get("name")
        self.excel = win32com.client.Dispatch("Excel.Application")

    def convert_to_pt(self, value, unit):
        execl = self.excel
        """Convert spacing values from various units to points (pt)."""
        if value is None:
            return 0
        if unit == "pt" or unit == "point":
            return float(value)
        elif unit == "cm":
            return execl.CentimetersToPoints(value)
        elif unit == "mm":
            return execl.CentimetersToPoints(value*0.1)
        elif unit == "inches":
            return execl.InchesToPoints(value)
        else:
            raise ValueError(f"不支持的单位: {unit}")

    def is_caption_paragraph(self,paragraph):
        """Determine whether a paragraph is a table/figure caption (regex-based)."""
        text = paragraph.Range.Text.strip()
        # Match common Chinese/English caption formats (e.g. Table 1: / Figure 1:)
        pattern = r'^((表|图|表格|Table|Figure)[\s]*\d+[\s]*[:：])'
        return bool(re.match(pattern, text))

    def format_natural_paragraph(self, doc, paragraph_index):
        """Apply formatting to the N-th natural-language paragraph
        (automatically skips tables, images, blank paragraphs, and captions).

        Args:
            doc: Word document object
            paragraph_index: Natural paragraph index (1-based)

        Returns:
            Word Range object, or None if the paragraph is not found"""
        try:
            # Filter real text paragraphs (four-way filter)
            text_paragraphs = [
                para for para in doc.Paragraphs
                if (para.Range.Tables.Count == 0 and  # Not a table
                    para.Range.InlineShapes.Count == 0 and  # Not an image
                    para.Range.Text.strip() and  # Not a blank paragraph
                    not self.is_caption_paragraph(para))  # Not a table/figure caption
            ]
            # Safety check
            if not text_paragraphs:
                return None
            if paragraph_index < 1 or paragraph_index > len(text_paragraphs):
                return None
            # Return the Range object of the target paragraph
            return text_paragraphs[paragraph_index - 1].Range
        except Exception:
            return None


    @BaseTool.register_tool({
        "en": {
            "function_description": "Set basic font properties for multiple paragraphs in a Word document",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "location_list": {"type": "list","description": "List of document location indices. Indexing rules: all = all paragraphs; 0 = currently selected region/cursor position; 1 = first body paragraph; 2 = second body paragraph (indices exclude empty paragraphs, tables, images, figure/table captions, and other non-body elements)."},
                "setting": {"type": "dict", "description": """Dictionary of font properties to set. Supported fields:
                      - "Name": "Font name (str). The default font setting field used to control both Chinese and Western characters simultaneously. When the user simply specifies “set the font to XXX”, this field should be used with priority, in line with common Word UI behavior."
                      - "NameFarEast": "Chinese font name (str). Used to explicitly specify the font for Chinese characters only; it affects Chinese text only. This field is used when Chinese and Western fonts need to be differentiated, or when the Chinese font needs to be overridden independently."
                      - "NameAscii": "Western font name (str). Used to explicitly specify the font for Western characters (letters, digits, and other ASCII characters). This field is used when an explicit distinction between Chinese and Western characters is required."
                      - "Size": Font size (int).
                      - "Bold": Whether to apply bold (int, 0 for no effect, -1 for effect).
                      - "Italic": Whether to apply italic (int, 0 for no effect, -1 for effect).
                      - "Underline": Underline style (int, see below for supported values):
                      0：No underline; 1：Single underline; 2：Words only underline; 3：Double underline; 4：Dotted underline; 7：Dash underline; 39：Dash long underline; 11：Wave underline; 6：Thick underline.
                      - "Color": Font color (str, hexadecimal color, e.g., "#FF0000").
                      - "HighlightColor": int, Font HighlightColor,  must be one of the following values: 0 (wdNoHighlight), 1 (wdBlack), 3 (wdTurquoise), 4 (wdBrightGreen), 5 (wdPink), 6 (wdRed), 7 (wdYellow), 9 (wdDarkBlue), 10 (wdTeal), 11 (wdGreen), 12 (wdViolet), 13 (wdDarkRed), 14 (wdDarkYellow), 15 (wdGray50), 16 (wdGray25)"""}
            }
        },
        "zh": {
            "function_description": "设置Word文档中多个段落的基本字体属性",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "location_list": {"type": "list","description": "文档位置索引列表，索引规则： 'all'=所有段落； 0=当前选中区域/光标位置；1=第一个正文段落；2=第二个正文段落（索引已过滤空段落、表格、图片、图表标题等非正文元素）"},
                "setting": {"type": "dict", "description": """包含字体属性的字典，支持的字段如下：
                      - "Name": 字体名称（str），默认字体设置字段，用于同时控制中英文字符的字体显示， 当用户仅说明“字体为 XXX”时，应优先使用该字段，以符合 Word 的常规使用习惯。。
                      - "NameFarEast": 中文字体名称（类型：字符串），用于**显式指定中文字体**，仅影响中文字符；  当需要区分中英文字体，或对中文字体进行单独覆盖时使用。 
                      - "NameAscii": 西文字体名称（英文、数字等）(str)，用于**显式指定西文字体**（英文、数字等 ASCII 字符）； 当需要强调中英文字符区分时使用。  
                      - "Size": 字体大小（int），单位为磅（pt）。中文常用字号对照情况：初号:42.0磅，小初:36.0磅，一号:26.0磅，小一:24.0磅，二号:22.0磅，小二:18.0磅，三号:16.0磅，小三:15.0磅，四号:14.0磅，小四:12.0磅，五号:10.5磅，小五:9.0磅，六号:7.5磅，小六:6.5磅，七号:5.5磅，八号:5.0磅。
                      - "Bold": 是否加粗（int, 0为无效果，-1为有效果）。
                      - "Italic": 是否倾斜（int, 0为无效果，-1为有效果）。
                      - "Underline": 下划线样式（int, 支持的取值如下）：
                          0：无下划线; 1：单线下划线; 2：仅文字下划线; 3：双线下划线; 4：点线下划线; 7：短线下划线（短横线样式的下划线）; 39：长线下划线（长横线样式的下划线）; 11：波浪线下划线; 6：粗线下划线。
                      - "Color": 字体颜色（str, 十六进制颜色，如 "#FF0000"）。
                     - "HighlightColor": 字体高亮背景（int）， 必须为下列枚举值之一：0（无高亮），1（黑色），3（青绿色），4（亮绿色），5（粉色），6（红色），7（黄色），9（深蓝色），10（墨绿色），11（绿色），12（紫色），13（深红色），"14（深黄色），15（50%灰），16（25%灰）"""}
            }
        }
    })
    def set_base_font(self, doc, location_list, setting={}):
        """Set font properties for specified paragraphs or the selection in a Word document.

        Args:
            doc: Word document object
            location_list: Location list; 0 means current selection, positive ints are paragraph indices
            setting: Dictionary of font properties

        Returns:
            Dictionary containing the operation result"""
        results = {}
        if 'all' in location_list:
            location_list = [i+1 for i in range(doc.Paragraphs.Count)]
        try:
            for location in location_list:
                # Handle the current selection (location = 0)
                if location == 0:
                    range_obj = doc.Application.Selection.Range
                elif location > 0:
                    if location > doc.Paragraphs.Count:
                        raise IndexError(f"Paragraph index {location} exceeds document length ({doc.Paragraphs.Count})")
                    range_obj = doc.Paragraphs(location).Range
                else:
                    raise ValueError("Location must be non-negative integer")

                font = range_obj.Font
                font_name = setting.get("Name",None)
                font_en_name = setting.get("NameAscii", None)
                font_zh_name = setting.get("NameFarEast", None)

                # Default handling logic
                if font_zh_name is None and font_name is not None:
                    setting["NameFarEast"] = font_name
                if font_en_name is None and font_name is not None:
                    # UI behavior: if no Western font is specified, English follows the Chinese font
                    setting["NameAscii"] = font_name

                try:
                    # Work around issues where some Chinese fonts are unsupported
                    font.NameFarEast = setting.get("NameFarEast")
                except Exception as e:
                    pass


                # Set font properties
                for attr in ["Size","Bold", "Italic", "Underline","Name","NameAscii"]:
                    if attr in setting:
                        if setting[attr] is not None:
                            setattr(font, attr, setting[attr])

                if "Color" in setting and setting["Color"]:
                    hex_color = setting["Color"].lstrip("#")
                    r = int(hex_color[0:2], 16)
                    g = int(hex_color[2:4], 16)
                    b = int(hex_color[4:6], 16)
                    # Swap R and B components
                    font.Color = b * 65536 + g * 256 + r

                # Set highlight
                if "HighlightColor" in setting:
                    highlight = setting["HighlightColor"]
                    valid_colors = {0, 1, 3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16}
                    if highlight in valid_colors:
                        range_obj.HighlightColorIndex = highlight
                    else:
                        raise ValueError(f"HighlightColor not support: {setting['HighlightColor']}")

            doc.Save()
            results["base_font"] = {
                "status": "success",
                "message": "Font properties set successfully"
            }
        except Exception as e:
            results["base_font"] = {
                "status": "error",
                "message": f"Failed to set font properties: {str(e)}"
            }
        return results

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set advanced font effects for multiple paragraphs in a Word document",
            "params": {
                "doc": {
                    "type": "object",
                    "description": "Word document object"
                },
                "location_list": {"type": "list",
                                  "description": "List of document location indices. Indexing rules: all = all paragraphs; 0 = currently selected region/cursor position; 1 = first body paragraph; 2 = second body paragraph (indices exclude empty paragraphs, tables, images, figure/table captions, and other non-body elements)."},
                "setting": {
                    "type": "dict",
                    "description": """A dictionary containing font effects. Supported fields include:
    - "StrikeThrough": Strikethrough effect (int, 0 = disabled, -1 = enabled).
    - "Subscript": Subscript (int, 0 = disabled, -1 = enabled).
    - "Superscript": Superscript (int, 0 = disabled, -1 = enabled).
    - "AllCaps": All capital letters (int, 0 = disabled, -1 = enabled).
    - "SmallCaps": Small caps (int, 0 = disabled, -1 = enabled).
    - "Spacing": Character spacing (float, in points. 0 = normal, positive = expanded spacing, negative = condensed spacing).
    - "Scaling": Character width scaling (int, range: 1–600, representing percentage of normal width).
    - "Emboss": Emboss effect (int, 0 = disabled, -1 = enabled).
    - "Engrave": Engrave effect (int, 0 = disabled, -1 = enabled).
    - "Shadow": Shadow effect (int, 0 = disabled, -1 = enabled)."""
                }
            }
        },
        "zh": {
            "function_description": "设置Word文档中多个段落的高级字体效果",
            "params": {
                "doc": {
                    "type": "object",
                    "description": "Word文档对象"
                },
                "location_list": {"type": "list",
                                  "description": "文档位置索引列表，索引规则：'all'=所有段落； 0=当前选中区域/光标位置；1=第一个正文段落；2=第二个正文段落（索引已过滤空段落、表格、图片、图表标题等非正文元素）"},
                "setting": {
                    "type": "dict",
                    "description": """包含字体效果的字典，支持以下字段：
    - "StrikeThrough": 删除线效果（int，0 表示无，-1 表示启用）。
    - "Subscript": 下标效果（int，0 表示无，-1 表示启用）。
    - "Superscript": 上标效果（int，0 表示无，-1 表示启用）。
    - "AllCaps": 全部大写（int，0 表示无，-1 表示启用）。
    - "SmallCaps": 小型大写（int，0 表示无，-1 表示启用）。
    - "Spacing": 字符间距（float，单位为 point，0 表示正常间距，正值表示加宽，负值表示紧缩）。
    - "Scaling": 字符缩放比例（int，取值范围为 1–600，表示字符宽度的百分比）。
    - "Emboss": 浮雕效果（int，0 表示无，-1 表示启用）。
    - "Engrave": 雕刻效果（int，0 表示无，-1 表示启用）。
    - "Shadow": 阴影效果（int，0 表示无，-1 表示启用）。"""
                }
            }
        }
    })
    def set_advanced_font(self, doc, location_list, setting={}):
        """
        Set advanced font effects for multiple paragraphs or selection in Word document

        Args:
            doc: Word document object
            location_list: List of positions (0 for current selection, positive integers for natural paragraph indices)
            setting: Dictionary containing font effects
        Returns:
            Dictionary with operation results
        """
        results = {}
        if 'all' in location_list:
            location_list = [i+1 for i in range(doc.Paragraphs.Count)]
        # Supported fields and their descriptions
        supported_fields = {
            "StrikeThrough": "Strike-through",
            "Subscript": "Subscript",
            "Superscript": "Superscript",
            "AllCaps": "All caps",
            "SmallCaps": "Small caps",
            "Spacing": "Character spacing",
            "Scaling": "Character scaling",
            "Emboss": "Emboss",
            "Engrave": "Engrave",
            "Shadow": "Shadow"
        }

        try:
            for paragraph_index in location_list:
                # Get the target range
                if paragraph_index == 0:
                    range_obj = doc.Application.Selection.Range
                elif paragraph_index > 0:
                    range_obj = doc.Paragraphs(paragraph_index).Range
                    if range_obj is None:
                        raise ValueError(f"Position {paragraph_index} not found or invalid natural paragraph")
                else:
                    raise ValueError("Paragraph index must be non-negative integer")

                font = range_obj.Font
                for attr, desc in supported_fields.items():
                    if attr not in setting:
                        continue
                    value = setting[attr]
                    if attr == "Scaling":
                        if not (1 <= value <= 600):
                            raise ValueError("Scaling must be between 1 and 600")
                    setattr(font, attr, value)
            results["advanced_font"] = {
                "status": "success",
                "message": "Advanced font effects set successfully"
            }
            doc.Save()
        except Exception as e:
            results["advanced_font"] = {
                "status": "error",
                "message": f"Failed to set advanced font effects: {str(e)}"
            }
        return results

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set the outline level for a specific paragraph style in a Word document.",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "location_list": {"type": "list",
                                  "description": "List of document location indices. Indexing rules: all = all paragraphs; 0 = currently selected region/cursor position; 1 = first body paragraph; 2 = second body paragraph (indices exclude empty paragraphs, tables, images, figure/table captions, and other non-body elements)."},
                "outlinelevel": {
                    "type": "int",
                    "description": """Outline level to assign to the style:
        - 1 to 9: Correspond to heading levels 1 to 9.
        - 10: Body text (no outline level)."""
                }
            }
        },
        "zh": {
            "function_description": "为指定段落样式设置大纲级别（可新建或修改样式）",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "location_list": {"type": "list",
                                  "description": "文档位置索引列表，索引规则：'all'=所有段落； 0=当前选中区域/光标位置；1=第一个正文段落；2=第二个正文段落（索引已过滤空段落、表格、图片、图表标题等非正文元素）"},
                "outlinelevel": {
                    "type": "int",
                    "description": """要分配给样式的大纲级别：
        - 1 到 9：对应 Word 的标题级别（1为最高级，9为最低级）。
        - 10：设置为正文文本（无大纲级别）。"""
                }
            }
        }
    })
    def set_outlinelevel(self, doc, location_list, outlinelevel):
        results = {}
        if 'all' in location_list:
            location_list = [i+1 for i in range(doc.Paragraphs.Count)]
        # Validate outline level first (before processing any paragraphs)
        try:
            outlinelevel = int(outlinelevel)
            if outlinelevel < 1 or outlinelevel > 10:
                raise ValueError("Outline level must be between 1-9 (heading levels) or 10 (body text)")
            level_description = "Body text" if outlinelevel == 10 else f"Heading level {outlinelevel}"
        except Exception as e:
            # If outline level is invalid, return error
            results["outlinelevel"] = {
                'status': 'error',
                'message': f"Invalid outline level: {str(e)}"
            }
            doc.Save()
            return results

        try:
            for location in location_list:
                if location == 0:
                    range_obj = doc.Application.Selection.Range
                elif location > 0:
                    if location > doc.Paragraphs.Count:
                        raise IndexError(f"Paragraph index {location} exceeds document length ({doc.Paragraphs.Count})")
                    range_obj = doc.Paragraphs(location).Range
                else:
                    raise ValueError("Location must be non-negative integer")

                # Set outline level for the range
                para_fmt = range_obj.ParagraphFormat
                para_fmt.OutlineLevel = outlinelevel

            results["outlinelevel"] = {
                "status": "success",
                "message": f"Set to {level_description}"
            }
            doc.Save()
        except Exception as e:
            results["outlinelevel"] = {
                "status": "error",
                "message": f"Failed to set outline level: {str(e)}"
            }

        return results

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set the alignment type of a paragraph style",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "location_list": {"type": "list",
                                  "description": "List of document location indices. Indexing rules: all = all paragraphs; 0 = currently selected region/cursor position; 1 = first body paragraph; 2 = second body paragraph (indices exclude empty paragraphs, tables, images, figure/table captions, and other non-body elements)."},
                "alignment": {
                    "type": "str",
                    "description": (
                            "Alignment type, one of the following values:\n"
                            " - left: align text to the left\n"
                            " - center: center text\n"
                            " - right: align text to the right\n"
                            " - justify: justify text to both left and right edges\n"
                            " - distribute: evenly distribute text by adjusting character spacing"
                    )
                }
            }
        },
        "zh": {
            "function_description": "设置某段落样式的对齐方式",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "location_list": {"type": "list",
                                  "description": "文档位置索引列表，索引规则：'all'=所有段落； 0=当前选中区域/光标位置；1=第一个正文段落；2=第二个正文段落（索引已过滤空段落、表格、图片、图表标题等非正文元素）"},
                "alignment": {
                    "type": "str",
                    "description": (
                            "对齐方式，可选值如下：\n"
                            " - left：左对齐，段落靠左边排列\n"
                            " - center：居中对齐，段落在页面中居中\n"
                            " - right：右对齐，段落靠右边排列\n"
                            " - justify：两端对齐，段落左右两端同时对齐\n"
                            " - distribute：分散对齐，自动调整字符间距使各行长度相等"
                    )
                }
            }
        }
    })
    def set_alignment(self, doc, location_list, alignment):
        results = {}
        if 'all' in location_list:
            location_list = [i+1 for i in range(doc.Paragraphs.Count)]
        # Define alignment mapping
        alignment_map = {
            "left": constants.wdAlignParagraphLeft,
            "center": constants.wdAlignParagraphCenter,
            "right": constants.wdAlignParagraphRight,
            "justify": constants.wdAlignParagraphJustify,
            "distribute": constants.wdAlignParagraphDistribute,
            "左对齐": constants.wdAlignParagraphLeft,
            "居中": constants.wdAlignParagraphCenter,
            "右对齐": constants.wdAlignParagraphRight,
            "两端对齐": constants.wdAlignParagraphJustify,
            "分散对齐": constants.wdAlignParagraphDistribute,
        }

        # Validate alignment first
        if alignment not in alignment_map:
            results["alignment"] = {
                "status": "error",
                "message": f"Invalid alignment type: {alignment}"
            }
            return results

        alignment_desc = {
            constants.wdAlignParagraphLeft: "Left aligned",
            constants.wdAlignParagraphCenter: "Center aligned",
            constants.wdAlignParagraphRight: "Right aligned",
            constants.wdAlignParagraphJustify: "Justified",
            constants.wdAlignParagraphDistribute: "Distributed"
        }[alignment_map[alignment]]

        try:
            for location in location_list:
                if location == 0:
                    range_obj = doc.Application.Selection.Range
                elif location > 0:
                    if location > doc.Paragraphs.Count:
                        raise IndexError(
                            f"Paragraph index {location} exceeds document length ({doc.Paragraphs.Count})"
                        )
                    range_obj = doc.Paragraphs(location).Range
                else:
                    raise ValueError("Location must be non-negative integer")

                # Set alignment for the range
                para_fmt = range_obj.ParagraphFormat
                para_fmt.Alignment = alignment_map[alignment]

            results["alignment"] = {
                "status": "success",
                "message": alignment_desc
            }

        except Exception as e:
            results["alignment"] = {
                "status": "error",
                "message": f"Failed to set alignment: {str(e)}"
            }

        # Save document (non-critical operation)
        try:
            doc.Save()
        except Exception as e:
            results["error"] = {
                "status": "error",
                "message": f"Failed to set alignment: {str(e)}"
            }

        return results

    @BaseTool.register_tool({
        "en": {
            "function_description": "Control pagination settings for a paragraph style",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "location_list": {"type": "list",
                                  "description": "List of document location indices. Indexing rules: all = all paragraphs; 0 = currently selected region/cursor position; 1 = first body paragraph; 2 = second body paragraph (indices exclude empty paragraphs, tables, images, figure/table captions, and other non-body elements)."},
                "widow_control": {
                    "type": "int",
                    "description": "Widow control. Prevents the first or last line of a paragraph from appearing alone at the top or bottom of a page1.\nValue guide: 0 disables this option, -1 enables it."
                },
                "keep_with_next": {
                    "type": "int",
                    "description": "Keep with next. Ensures the current paragraph stays on the same page1 as the following paragraph.\nValue guide: 0 disables this option, -1 enables it."
                },
                "keep_together": {
                    "type": "int",
                    "description": "Keep lines together. Prevents page1 breaks within the paragraph, keeping all lines on the same page1.\nValue guide: 0 disables this option, -1 enables it."
                },
                "page_break_before": {
                    "type": "int",
                    "description": "Page break before. Inserts a page1 break before the current paragraph.\nValue guide: 0 disables this option, -1 enables it."
                }
            }
        },
        "zh": {
            "function_description": "设置某段落样式的分页控制属性",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "location_list": {"type": "list",
                                  "description": "文档位置索引列表，索引规则：'all'=所有段落； 0=当前选中区域/光标位置；1=第一个正文段落；2=第二个正文段落（索引已过滤空段落、表格、图片、图表标题等非正文元素）"},
                "widow_control": {"type": "int",
                                  "description": "孤行控制，防止段落的第一行或最后一行单独出现在页面顶部或底部。\n取值说明：0 表示关闭该功能，-1 表示启用该功能。"},
                "keep_with_next": {"type": "int",
                                   "description": "与下段同页，确保当前段落与下一段落在同一页。\n取值说明：0 表示关闭该功能，-1 表示启用该功能。"},
                "keep_together": {"type": "int",
                                  "description": "段中不分页，防止段落中间分页。\n取值说明：0 表示关闭该功能，-1 表示启用该功能。"},
                "page_break_before": {"type": "int",
                                      "description": "段前分页，在当前段落前插入分页符。\n取值说明：0 表示关闭该功能，-1 表示启用该功能。"}
            }
        }
    })
    def set_pagination_control(self, doc, location_list, widow_control=None, keep_with_next=None,
                               keep_together=None, page_break_before=None):
        """
        Set pagination control properties for multiple locations in a Word document

        Args:
            doc: Word document object
            location_list: List of positions (0 for current selection, positive integers for paragraph indices)
            widow_control: Control widow/orphan lines (prevents first/last line of paragraph appearing alone)
            keep_with_next: Keep paragraph with next paragraph
            keep_together: Keep lines in paragraph together (no page break within)
            page_break_before: Page break before paragraph

        Returns:
            dict: results in the format:
            {
                "widow_control": {"status": str, "message": str},
                "keep_with_next": {"status": str, "message": str},
                "keep_together": {"status": str, "message": str},
                "page_break_before": {"status": str, "message": str}
            }
        """
        results = {}
        if 'all' in location_list:
            location_list = [i+1 for i in range(doc.Paragraphs.Count)]
        try:
            for location in location_list:
                # Handle current selection (location = 0)
                if location == 0:
                    range_obj = doc.Application.Selection.Range
                # Handle specified paragraph (location > 0)
                elif location > 0:
                    if location > doc.Paragraphs.Count:
                        raise IndexError(f"Paragraph index {location} exceeds document length ({doc.Paragraphs.Count})")
                    range_obj = doc.Paragraphs(location).Range
                else:
                    raise ValueError("Location must be non-negative integer")

                # Set pagination properties
                para_fmt = range_obj.ParagraphFormat

                if widow_control is not None:
                    try:
                        para_fmt.WidowControl = widow_control
                        results["widow_control"] = {
                            "status": "success",
                            "message": f"Set to {'on' if widow_control else 'off'}"
                        }
                    except Exception as e:
                        results["widow_control"] = {
                            "status": "error",
                            "message": f"Failed to set WidowControl: {str(e)}"
                        }

                if keep_with_next is not None:
                    try:
                        para_fmt.KeepWithNext = keep_with_next
                        results["keep_with_next"] = {
                            "status": "success",
                            "message": f"Set to {'on' if keep_with_next else 'off'}"
                        }
                    except Exception as e:
                        results["keep_with_next"] = {
                            "status": "error",
                            "message": f"Failed to set KeepWithNext: {str(e)}"
                        }

                if keep_together is not None:
                    try:
                        para_fmt.KeepTogether = keep_together
                        results["keep_together"] = {
                            "status": "success",
                            "message": f"Set to {'on' if keep_together else 'off'}"
                        }
                    except Exception as e:
                        results["keep_together"] = {
                            "status": "error",
                            "message": f"Failed to set KeepTogether: {str(e)}"
                        }

                if page_break_before is not None:
                    try:
                        para_fmt.PageBreakBefore = page_break_before
                        results["page_break_before"] = {
                            "status": "success",
                            "message": f"Set to {'on' if page_break_before else 'off'}"
                        }
                    except Exception as e:
                        results["page_break_before"] = {
                            "status": "error",
                            "message": f"Failed to set PageBreakBefore: {str(e)}"
                        }

        except Exception as e:
            results["error"] = {
                "status": "error",
                "message": f"Failed to apply pagination control: {str(e)}"
            }

        # Save document (non-critical operation)
        try:
            doc.Save()
        except Exception as e:
            results["error"] = {
                "status": "error",
                "message": f"Failed to save document: {str(e)}"
            }

        return results

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set line, before, and after spacing for a paragraph style",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "location_list": {"type": "list",
                                  "description": "List of document location indices. Indexing rules:all = all paragraphs; 0 = currently selected region/cursor position; 1 = first body paragraph; 2 = second body paragraph (indices exclude empty paragraphs, tables, images, figure/table captions, and other non-body elements)."},
               "line_spacing": {
                    "type": "dict",
                    "description": (
                            "Line spacing config (optional). Supported keys:\n"
                            "- spacing_rule: 'single' | '1.5' | 'double' | 'exact' | 'multiple'\n"
                            "- value: spacing value in pt (only needed for 'exact' or 'multiple')"
                    )
                },
                "before_spacing": {
                    "type": "dict",
                    "description": (
                            "Spacing before paragraph (optional). Supported keys:\n"
                            "- value: numeric value\n"
                            "- unit: unit of value: 'pt', 'cm', 'mm','inches'"
                    )
                },
                "after_spacing": {
                    "type": "dict",
                    "description": (
                            "Spacing after paragraph (optional). Supported keys:\n"
                            "- value: numeric value\n"
                            "- unit: unit of value: 'pt', 'cm', 'mm','inches'"
                    )
                }
            }
        },
        "zh": {
            "function_description": "设置段落样式的行间距、段前间距和段后间距",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "location_list": {"type": "list",
                                  "description": "文档位置索引列表，索引规则：'all'=所有段落； 0=当前选中区域/光标位置；1=第一个正文段落；2=第二个正文段落（索引已过滤空段落、表格、图片、图表标题等非正文元素）"},
                "line_spacing": {
                    "type": "dict",
                    "description": (
                            "行间距配置（可选）。支持的键：\n"
                            "- spacing_rule: 'single'（单倍）| '1.5'（1.5倍）| 'double'（双倍）| 'exact'（固定值）| 'multiple'（多倍）\n"
                            "- value: 行间距数值（仅当使用 exact 或 multiple 时需要）"
                    )
                },
                "before_spacing": {
                    "type": "dict",
                    "description": (
                            "段前间距配置（可选）。支持的键：\n"
                            "- value: 数值\n"
                            "- unit: 单位：'pt'、'cm'、'mm'、'inches'"
                    )
                },
                "after_spacing": {
                    "type": "dict",
                    "description": (
                            "段后间距配置（可选）。支持的键：\n"
                            "- value: 数值\n"
                            "- unit: 单位：'pt'、'cm'、'mm'、'inches'"
                    )
                }
            }
        }
    })
    def set_spacing(self, doc, location_list, line_spacing=None, before_spacing=None, after_spacing=None):
        """
        Set spacing properties for multiple locations in a Word document

        Args:
            doc: Word document object
            location_list: List of positions (0 for current selection, positive integers for paragraph indices)
            line_spacing: Dictionary containing line spacing settings:
                - "spacing_rule": "single"/"1.5"/"double"/"exact"/"multiple"
                - "value": Value for exact/multiple spacing
            before_spacing: {"value": number, "unit": "pt"/"cm"/"in"/...}
            after_spacing: {"value": number, "unit": "pt"/"cm"/"in"/...}

        Returns:
            dict:
            {
                "line_spacing": {"status": str, "message": str},
                "before_spacing": {"status": str, "message": str},
                "after_spacing": {"status": str, "message": str},
                "error": {"status": "error", "message": "..."}  # optional
            }
        """
        results = {}
        if 'all' in location_list:
            location_list = [i+1 for i in range(doc.Paragraphs.Count)]
        # Define spacing rule mapping
        rule_map = {
            "single": constants.wdLineSpaceSingle,
            "1.5": constants.wdLineSpace1pt5,
            "double": constants.wdLineSpaceDouble,
            "exact": constants.wdLineSpaceExactly,
            "multiple": constants.wdLineSpaceMultiple
        }

        # Validate line spacing rule
        if line_spacing and line_spacing.get("spacing_rule") not in rule_map:
            results["line_spacing"] = {
                "status": "error",
                "message": f"Invalid line spacing rule: {line_spacing.get('spacing_rule')}"
            }
            return results

        try:
            for location in location_list:
                # Handle current selection (location = 0)
                if location == 0:
                    range_obj = doc.Application.Selection.Range
                elif location > 0:
                    if location > doc.Paragraphs.Count:
                        raise IndexError(f"Paragraph index {location} exceeds document length ({doc.Paragraphs.Count})")
                    range_obj = doc.Paragraphs(location).Range
                else:
                    raise ValueError("Location must be non-negative integer")

                # Set spacing properties
                para_fmt = range_obj.ParagraphFormat

                # Line spacing
                if line_spacing:
                    try:
                        rule = line_spacing["spacing_rule"]
                        value = line_spacing.get("value", 1)

                        para_fmt.LineSpacingRule = rule_map[rule]

                        if rule == "exact":
                            spacing_value = float(value)
                            para_fmt.LineSpacing = spacing_value
                            results["line_spacing"] = {
                                "status": "success",
                                "message": f"Set to exact {spacing_value} pt"
                            }
                        elif rule == "multiple":
                            spacing_value = float(value * 12)
                            para_fmt.LineSpacing = spacing_value
                            results["line_spacing"] = {
                                "status": "success",
                                "message": f"Set to multiple ({value}x) {spacing_value} pt"
                            }
                        else:
                            results["line_spacing"] = {
                                "status": "success",
                                "message": f"Set to {rule} spacing"
                            }
                    except Exception as e:
                        results["line_spacing"] = {
                            "status": "error",
                            "message": f"Failed to set line spacing: {str(e)}"
                        }

                # Before spacing
                if before_spacing:
                    try:
                        val = before_spacing["value"]
                        unit = before_spacing.get("unit", "pt")
                        spacing_pt = self.convert_to_pt(val, unit)
                        para_fmt.SpaceBefore = spacing_pt

                        results["before_spacing"] = {
                            "status": "success",
                            "message": f"Set to {val} {unit} ({spacing_pt} pt)"
                        }
                    except Exception as e:
                        results["before_spacing"] = {
                            "status": "error",
                            "message": f"Failed to set before spacing: {str(e)}"
                        }

                # After spacing
                if after_spacing:
                    try:
                        val = after_spacing["value"]
                        unit = after_spacing.get("unit", "pt")
                        spacing_pt = self.convert_to_pt(val, unit)
                        para_fmt.SpaceAfter = spacing_pt

                        results["after_spacing"] = {
                            "status": "success",
                            "message": f"Set to {val} {unit} ({spacing_pt} pt)"
                        }
                    except Exception as e:
                        results["after_spacing"] = {
                            "status": "error",
                            "message": f"Failed to set after spacing: {str(e)}"
                        }

        except Exception as e:
            results["error"] = {
                "status": "error",
                "message": f"Failed to apply spacing settings: {str(e)}"
            }

        # Save document (non-critical operation)
        try:
            doc.Save()
        except Exception as e:
            results["error"] = {
                "status": "error",
                "message": f"Failed to save document: {str(e)}"
            }

        return results

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set paragraph indentation (left, right, first line/hanging) for a specified paragraph style",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "location_list": {"type": "list",
                                  "description": "List of document location indices. Indexing rules: all = all paragraphs; 0 = currently selected region/cursor position; 1 = first body paragraph; 2 = second body paragraph (indices exclude empty paragraphs, tables, images, figure/table captions, and other non-body elements)."},
                "left_indent": {
                    "type": "dict",
                    "description": (
                            "Left indent settings. Keys:\n"
                            "- 'value' (float): the indentation value\n"
                            "- 'unit' (string): the unit of the value, one of ['point', 'cm', 'mm','inches']\n"
                            "- 'hanging' (int): set to -1 for hanging indent (value will be negative), 0 for normal"
                    )
                },
                "right_indent": {
                    "type": "dict",
                    "description": (
                            "Right indent settings. Keys:\n"
                            "- 'value' (float): the indentation value\n"
                            "- 'unit' (string): the unit of the value, one of ['point', 'cm', 'mm','inches']\n"
                            "- 'hanging' (int): set to -1 for hanging indent (value will be negative), 0 for normal"
                    )
                },
                "firstline_indent": {
                    "type": "dict",
                    "description": (
                            "First line indent settings. Keys:\n"
                            "- 'value' (float): the indentation value\n"
                            "- 'unit' (string): the unit of the value, one of ['point', 'cm', 'mm', 'character','inches']\n"
                            "- 'hanging' (int): 0 means indent first line (positive value), -1 means hanging indent (value will be negative)"
                    )
                }
            }
        },
        "zh": {
            "function_description": "设置指定段落样式的缩进（左缩进、右缩进、首行或悬挂缩进）",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "location_list": {"type": "list",
                                  "description": "文档位置索引列表，索引规则：'all'=所有段落； 0=当前选中区域/光标位置；1=第一个正文段落；2=第二个正文段落（索引已过滤空段落、表格、图片、图表标题等非正文元素）"},
                "left_indent": {
                    "type": "dict",
                    "description": (
                            "左缩进设置，包含以下键：\n"
                            "- 'value' (float): 缩进的数值\n"
                            "- 'unit' (string): 单位，可选值为 ['point', 'cm', 'mm','inches']\n"
                            "- 'hanging' (int): 设置为-1表示悬挂缩进（将使用负值），0表示普通左缩进"
                    )
                },
                "right_indent": {
                    "type": "dict",
                    "description": (
                            "右缩进设置，包含以下键：\n"
                            "- 'value' (float): 缩进的数值\n"
                            "- 'unit' (string): 单位，可选值为 ['point', 'cm', 'mm', 'inches']\n"
                            "- 'hanging' (int): 设置为-1表示悬挂缩进（将使用负值），0表示普通右缩进"
                    )
                },
                "firstline_indent": {
                    "type": "dict",
                    "description": (
                            "首行或悬挂缩进设置，包含以下键：\n"
                            "- 'value' (float): 缩进的数值\n"
                            "- 'unit' (string): 单位，可选值为 ['point', 'cm', 'mm', 'character','inches']\n"
                            "- 'hanging' (int): 0 表示首行缩进（正值），-1 表示悬挂缩进（负值）"
                    )
                }
            }
        }
    })
    def set_indent(self, doc, location_list, left_indent=None, right_indent=None, firstline_indent=None):
        """
        Set indentation properties for multiple locations in a Word document

        Args:
            doc: Word document object
            location_list: List of positions (0 for current selection, positive integers for paragraph indices)
            left_indent: Dictionary containing left indent settings:
                - "value": Numeric value
                - "unit": Unit of measurement ("pt", "cm", "in", "character", etc.)
                - "hanging": Optional flag for hanging indent (-1 for hanging)
            right_indent: Dictionary containing right indent settings:
                - "value": Numeric value
                - "unit": Unit of measurement ("pt", "cm", "in", "character", etc.)
                - "hanging": Optional flag for hanging indent (-1 for hanging)
            firstline_indent: Dictionary containing first line indent settings:
                - "value": Numeric value
                - "unit": Unit of measurement ("pt", "cm", "in", "character", etc.)
                - "hanging": Optional flag for hanging indent (-1 for hanging)

        Returns:
            dict: {
                "left_indent": {"status": "...", "message": "..."} (if requested),
                "right_indent": {"status": "...", "message": "..."} (if requested),
                "firstline_indent": {"status": "...", "message": "..."} (if requested),
                "error": {"status": "error", "message": "..."} (optional)
            }
        """
        results = {}
        if 'all' in location_list:
            location_list = [i+1 for i in range(doc.Paragraphs.Count)]
        # Unit conversion helper (kept as original)
        def to_point(value, unit, font_size=12):
            if unit in ["point", "pt", "cm", "mm", "inches"]:
                return self.convert_to_pt(value, unit)
            elif unit == "character":
                return font_size * value
            else:
                raise ValueError(f"Unsupported unit: {unit}")

        try:
            for location in location_list:
                try:
                    # Handle current selection (location = 0)
                    if location == 0:
                        range_obj = doc.Application.Selection.Range
                        font_size = range_obj.Font.Size
                    # Handle specified paragraph (location > 0)
                    elif location > 0:
                        if location > doc.Paragraphs.Count:
                            raise IndexError(
                                f"Paragraph index {location} exceeds document length ({doc.Paragraphs.Count})")
                        range_obj = doc.Paragraphs(location).Range
                        font_size = range_obj.Font.Size
                    else:
                        # Preserve original error semantics: location must be a non-negative integer
                        raise ValueError("Location must be non-negative integer")

                    # Set indentation properties (kept original logic)
                    para_fmt = range_obj.ParagraphFormat

                    # Left indent
                    if left_indent:
                        try:
                            val = left_indent['value']
                            unit = left_indent.get('unit', 'pt')
                            pt_val = to_point(val, unit, font_size)
                            if left_indent.get('hanging', 0) == -1:
                                pt_val = -abs(pt_val)
                            para_fmt.LeftIndent = abs(pt_val)

                            indent_type = "Hanging" if pt_val < 0 else "Left"
                            results['left_indent'] = {
                                'status': 'success',
                                'message': f"{indent_type} indent set to {abs(pt_val)} pt (from {val} {unit})"
                            }
                        except Exception as e:
                            results['left_indent'] = {
                                'status': 'error',
                                'message': f"Failed to set left indent: {str(e)}"
                            }

                    # Right indent
                    if right_indent:
                        try:
                            val = right_indent['value']
                            unit = right_indent.get('unit', 'pt')
                            pt_val = to_point(val, unit, font_size)
                            if right_indent.get('hanging', 0) == -1:
                                pt_val = -abs(pt_val)
                            para_fmt.RightIndent = abs(pt_val)

                            indent_type = "Hanging" if pt_val < 0 else "Right"
                            results['right_indent'] = {
                                'status': 'success',
                                'message': f"{indent_type} indent set to {abs(pt_val)} pt (from {val} {unit})"
                            }
                        except Exception as e:
                            results['right_indent'] = {
                                'status': 'error',
                                'message': f"Failed to set right indent: {str(e)}"
                            }

                    # First line indent
                    if firstline_indent:
                        try:
                            val = firstline_indent['value']
                            unit = firstline_indent.get('unit', 'pt')
                            pt_val = to_point(val, unit, font_size)
                            if firstline_indent.get('hanging', 0) == -1:
                                pt_val = -abs(pt_val)
                            para_fmt.FirstLineIndent = abs(pt_val)

                            indent_type = "Hanging" if pt_val < 0 else "First line"
                            results['firstline_indent'] = {
                                'status': 'success',
                                'message': f"{indent_type} indent set to {abs(pt_val)} pt (from {val} {unit})"
                            }
                        except Exception as e:
                            results['firstline_indent'] = {
                                'status': 'error',
                                'message': f"Failed to set first line indent: {str(e)}"
                            }

                    # Continue with the next location (keep per-location processing behavior)

                except Exception as e:
                    # If a location raises (e.g. out of range or negative), write to results["error"] and break
                    results["error"] = {
                        'status': 'error',
                        'message': f"Error processing location: {str(e)}"
                    }
                    break

        except Exception as e:
            # Catch unexpected outer errors
            results["error"] = {
                'status': 'error',
                'message': f"Failed to apply indent settings: {str(e)}"
            }

        # Save document (non-critical operation)
        try:
            doc.Save()
        except Exception as e:
            results["error"] = {
                'status': 'error',
                'message': f"Failed to save document: {str(e)}"
            }

        return results



def get_fontsize_table():
    file_path = os.path.join(ABS_DIR, "file/base_font.docx")
    context_tool = TextTools()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True  # Set to True to observe the operation process
    doc = word.Documents.Open(file_path)
    location_list = [i + 1 for i in range(doc.Paragraphs.Count)]
    for location in location_list:
        paragraph = doc.Paragraphs(location)
        text = paragraph.Range.Text[:2]
        font_size = paragraph.Range.Font.Size
        print(f"{text}:{font_size}磅，",end="")

    word.Quit()

if __name__ == '__main__':

    file_path = os.path.join(ABS_DIR,"file/Word_test.docx")
    context_tool = TextTools()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True  # Set to True to observe the operation process
    doc = word.Documents.Open(file_path)
    print(doc.Paragraphs(1).Range.Text)
    # doc.Paragraphs(1).Range.Font.Name = "Times New Roman"
    # status = context_tool.set_base_font(doc,[1,2],setting = {
    #     "Name": "楷体",
    #     "NameFarEast": "宋体",
    #     "NameAscii": "Times New Roman",
    #     "Size": 10.5
    #   })
    # print(status)
    # status = context_tool.set_spacing(doc,[1,2], line_spacing= {"spacing_rule": "exact",
    #     "value": 20
    #   })
    # print(status)
    word.Quit()
    # status = context_tool.set_outlinelevel(doc,[1,2],1)
    # Set alignment for current selection and paragraphs 1 and 3
    # status = context_tool.set_indent(doc, [1], firstline_indent = {"value":2,"unit":"character"})
    # print(status)

    # context_agent = AppAgent(os.path.join(ABS_DIR,"config/Agents/text_agent.yaml"))
    # prompt = context_agent.function_call_prompt("请将第二段字体加粗，文本居中，字体设置为24磅")

    # print(prompt)




