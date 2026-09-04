from constant import ABS_DIR
import os,copy,re
from win32com.client import constants
import win32com
from tool.file_trans import FileConverter
from tool.basetool import ContextToolsConfig,BaseTool

class ImageReader(BaseTool):
    def __init__(self, pyconfig=ContextToolsConfig("/config/Tools/reader/image_reader_config.yaml")):
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
    def get_image(self, doc, image_index):
        """Get an embedded image (InlineShape) from a Word document.
        :param doc: Word document object
        :param image_index: Image index (1-based)
        :return: InlineShape object or error information"""
        try:
            total = doc.InlineShapes.Count
            if total == 0:
                raise ValueError("文档中没有嵌入式图片。")
            if not (1 <= image_index <= total):
                raise IndexError(f"图片索引 {image_index} 超出范围（1 - {total}）。")
            image = doc.InlineShapes(image_index)
            return image
        except Exception as e:
            return None
    def pt_to_percent(self,value, PageSetup):
        page_width = PageSetup.PageWidth - PageSetup.LeftMargin - PageSetup.RightMargin
        percent = round(value / page_width * 100,2)
        return percent

    def __get_size(self, image,*args,**kwargs):
        width_pt = image.Width
        height_pt = image.Height
        lock_aspect_ratio = image.LockAspectRatio
        return {"width": width_pt, "height": height_pt, "lock_aspect_ratio": lock_aspect_ratio, }

    def __read_size(self, image, doc, image_info,*args,**kwargs):
        width_pt = image.Width
        height_pt = image.Height
        lock_aspect_ratio = image.LockAspectRatio

        image_info["size"]['width']['value']["pt"] = self.pt_to_convert(width_pt, "pt")
        image_info["size"]['width']['value']["cm"] = self.pt_to_convert(width_pt, "cm")
        image_info["size"]['width']['value']["mm"] = self.pt_to_convert(width_pt, "mm")
        image_info["size"]['width']['value']["inches"] = self.pt_to_convert(width_pt, "inches")
        image_info["size"]['width']['value']["percent"] = self.pt_to_percent(width_pt, doc.PageSetup)
        
        image_info["size"]['height']['value']["pt"] = self.pt_to_convert(height_pt, "pt")
        image_info["size"]['height']['value']["cm"] = self.pt_to_convert(height_pt, "cm")
        image_info["size"]['height']['value']["mm"] = self.pt_to_convert(height_pt, "mm")
        image_info["size"]['height']['value']["inches"] = self.pt_to_convert(height_pt, "inches")

        image_info["size"]["lock_aspect_ratio"]['value']= lock_aspect_ratio

        return image_info
    def __get_alignment(self, image,*args,**kwargs):
        # Get the paragraph containing the image
        paragraph = image.Range.Paragraphs(1)
        alignment_value = paragraph.Alignment
        # Reverse-map to a string
        mapping = {
            0: "left",
            1: "center",
            2: "right",
            3: "justify"
        }
        alignment = mapping.get(alignment_value, "unknown")
        return {
            "alignment": alignment
        }

    def __read_alignment(self, image, image_info,*args,**kwargs):
        # Get the paragraph containing the image
        paragraph = image.Range.Paragraphs(1)
        alignment_value = paragraph.Alignment
        # Reverse-map to a string
        mapping = {
            0: "left",
            1: "center",
            2: "right",
            3: "justify"
        }
        alignment = mapping.get(alignment_value, "unknown")

        image_info["alignment"]["value"] = alignment
        return image_info

    def __get_pagination(self, image,*args,**kwargs):
        paragraph = image.Range.ParagraphFormat
        return {
            "keep_with_next": paragraph.KeepWithNext,
            "keep_together": paragraph.KeepTogether,
            "page_break_before": paragraph.PageBreakBefore
        }

    def __read_pagination(self, image, image_info,*args,**kwargs):
        paragraph = image.Range.ParagraphFormat
        image_info['pagination']["keep_with_next"]["value"] = paragraph.KeepWithNext
        image_info['pagination']["keep_together"]["value"] = paragraph.KeepTogether
        image_info['pagination']["page_break_before"]["value"] = paragraph.PageBreakBefore
        return image_info

    def read_image_properties(self, doc, image_index, params_list=[], language='zh',*args,**kwargs):
        # print(function_list)
        # print(params_list)
        attribution_dict = {
            "size": self.__read_size,
            "pagination": self.__read_pagination,
            "alignment": self.__read_alignment,
        }
        # Load the read template
        template = self.config.get("properties_template")
        if language in ['zh', 'en']:
            image_info = copy.deepcopy(template.get(language))
        else:
            image_info = copy.deepcopy(template.get("zh"))
            print("Default Using Chinese")
        try:
            # String conversion
            image_index = int(image_index)
            if image_index > 0:
                image = self.get_image(doc,image_index)
            else:
                print("image index must >= 0!")
                raise
            if not params_list:
                # No property scope specified; read all by default
                params_list = list(attribution_dict.keys())
            else:
                # After a property scope is specified, remove unused keys from the template
                for attribution in attribution_dict.keys():
                    if attribution not in params_list:
                        image_info.pop(attribution)

            # Fetch each property to read in order
            for params in params_list:
                # Call parameters are supported
                if params in attribution_dict:
                    attribution_info_read_tool = attribution_dict.get(params)
                    image_info = attribution_info_read_tool(image=image, image_info=image_info,doc=doc)
                    # print(image_info)
                    # print(params)

            # Return success result
            return {"state": "success", "properties": image_info}

        except Exception as e:
            # Catch exceptions and return error information

            return {"state": "false", "image_index": image_index, "exception": str(e)}

    def get_images_info(self, doc, before=0, after=0, *args,**kwargs):
        try:
            image_info_result = []
            for image_index in range(1, doc.InlineShapes.Count + 1):
                image_image_result = self.__get_image_info(doc, image_index=image_index, before=before, after=after)
                if image_image_result.get("state") == "success":
                    image_image_result.pop("state")
                    image_info_result.append(image_image_result)
            return image_info_result
        except Exception as e:
            print(f"Get image info error! The detail is: {e}")
            raise

    def __get_image_info(self, doc, image_index, before=0, after=0, *args, **kwargs):
        """Get detailed information for the specified image (page number, surrounding text, etc.)."""
        try:
            all_paragraphs = list(doc.Paragraphs)

            # --- 1. Get the image object ---
            image = self.get_image(doc, image_index)
            img_range = image.Range

            # --- 2. Page number information ---
            start_page = img_range.Information(constants.wdActiveEndPageNumber)

            # --- 3. Determine paragraph index ---
            img_start = img_range.Start
            img_para_index = None

            for i, para in enumerate(all_paragraphs, start=1):
                if para.Range.Start <= img_start <= para.Range.End:
                    img_para_index = i
                    break

            # --- 4. Get surrounding paragraphs (improved) ---
            before_paras = []
            after_paras = []
            current_para_text = ""

            if img_para_index:
                # Get current paragraph text (excluding the image itself)
                current_para = all_paragraphs[img_para_index - 1]
                current_para_text = self.__get_paragraph_text_without_image(current_para, img_range)

                # Collect paragraphs before the image
                start_before_idx = max(0, img_para_index - 1 - before)
                before_count = 0
                i = img_para_index - 2  # Start from the paragraph before the current one

                while i >= start_before_idx and before_count < before:
                    if i >= 0:  # Ensure the index is valid
                        para_text = all_paragraphs[i].Range.Text
                        cleaned_text = self.__clean_paragraph_text(para_text)
                        if cleaned_text:
                            before_paras.insert(0, cleaned_text)  # Insert in order
                            before_count += 1
                    i -= 1

                # Collect paragraphs after the image — keep searching if filtered paragraphs are encountered
                after_count = 0
                i = img_para_index  # Start from the paragraph after the current one

                while i < len(all_paragraphs) and after_count < after:
                    para_text = all_paragraphs[i].Range.Text
                    cleaned_text = self.__clean_paragraph_text(para_text)

                    if cleaned_text:
                        after_paras.append(cleaned_text)
                        after_count += 1
                    # If this paragraph is filtered out, check the next one without incrementing the count
                    # This ensures we collect the requested number of valid paragraphs

                    i += 1

                # Debug information
                # print(f"Debug - paragraph index of image: {img_para_index}")
                # print(f"Debug - collected {len(before_paras)} valid paragraphs before")
                # print(f"Debug - collected {len(after_paras)} valid paragraphs after")
                # print(f"Debug - after-paragraph content: {after_paras}")

            # --- 5. Basic image properties ---
            image_size = self.__get_size(image)
            alignment = self.__get_alignment(image)

            # --- 6. Aggregate results ---
            return {
                "state": "success",
                "image_index": image_index,
                "page_number": start_page,
                "current_paragraph": current_para_text,
                "before_paragraphs": before_paras,
                "after_paragraphs": after_paras,
                "size": image_size,
                "alignment": alignment
            }
        except Exception as e:
            print(f"Get image information error! The detail is {e}")
            return {
                "state": "error",
                "message": str(e)
            }

    def __clean_paragraph_text(self, text):
        """Clean paragraph text by removing special characters and blank paragraphs."""
        if not text:
            return ""

        # Remove special characters such as newlines and tabs
        cleaned = text.replace('\r', '').replace('\x07', '').replace('\x0b', '').strip()

        # Filter out paragraphs that contain only special characters
        special_chars = {'/', '|', '-', '*', ' ', '\t', '\n', '\r', '\x0c', '\x00'}
        if all(c in special_chars or c.isspace() for c in cleaned):
            return ""

        # Filter out overly short paragraphs (possibly separators)
        if len(cleaned) <= 1 and cleaned in special_chars:
            return ""

        return cleaned

    def __get_paragraph_text_without_image(self, paragraph, image_range):
        """Get paragraph text while excluding the text range of the specified image."""
        try:
            para_range = paragraph.Range
            para_text = para_range.Text

            # If the image range is within the paragraph, remove the image-corresponding text
            if (image_range.Start >= para_range.Start and
                    image_range.End <= para_range.End):
                # Compute the image position within the paragraph text
                start_pos = image_range.Start - para_range.Start
                end_pos = image_range.End - para_range.Start

                # Remove the text portion corresponding to the image
                para_text = para_text[:start_pos] + para_text[end_pos:]

            return self.__clean_paragraph_text(para_text)
        except Exception as e:
            print(f"Error extracting paragraph text without image: {e}")
            return self.__clean_paragraph_text(paragraph.Range.Text)

    def __get_image_format(self, image, key_list, *args,**kwargs):
        try:
            property_get_dict = {
                "size": self.__get_size,
                "pagination": self.__get_pagination,
                "alignment": self.__get_alignment
             }
            format_dict = {}
            for key in key_list:
                if key in property_get_dict:
                    format_dict[key] = property_get_dict.get(key)(image)
            result = {
                "status":"success",
                "properties":format_dict
            }
        except Exception as e:
            print(f"Image format get Error! The detail is {e}")
            result = {
                "status": "error",
                "exception": e
            }
        finally:
            return result

    def get_image_format(self, doc, image_index, key_list = ["size", "pagination", "alignment"], *args,**kwargs):
        image = self.get_image(doc,image_index)
        return self.__get_image_format(image,key_list)

    def get_images_format(self,doc):
        image_num = doc.InlineShapes.Count
        formats = {}
        for index in range(image_num):
            image_index = index+1
            format = self.get_image_format(doc,image_index)
            if format.get("status") == "success":
                formats[str(image_index)] = format.get("properties")
        return formats

if __name__ == '__main__':
    import win32com.client as win32
    word = win32.DispatchEx("Word.Application")  # Or use Dispatch
    word.Visible = True  # Make visible (recommended when debugging)
    # word_file_path = "./file/Word_test.docx"
    word_file_path = "./file/Base.docx"
    word_file_path = os.path.join(ABS_DIR, word_file_path)
    # print(word_file_path)
    # Open an existing document
    try:
        # Open the document
        doc = word.Documents.Open(word_file_path)
        image_reader = ImageReader()
        # print(image_reader.read_image_properties(doc, 1, []))

        print(image_reader.get_images_info(doc,"",1,1))

    except Exception as e:
        print(f"操作失败：{str(e)}")
        raise

    finally:
        # Ensure resources are cleaned up
        if 'doc' in locals():
            doc.Close(SaveChanges=False)
        word.Quit()