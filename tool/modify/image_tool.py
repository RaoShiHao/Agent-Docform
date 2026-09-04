import win32com.client as win32
import os
from tool.basetool import BaseTool
from constant import ABS_DIR
from tool.basetool import ContextToolsConfig

class ImageBaseTools():
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
            return execl.CentimetersToPoints(value)
        elif unit in ["mm","millimeter","毫米"]:
            return execl.CentimetersToPoints(value*0.1)
        elif unit in ["inches","英寸"]:
            return execl.InchesToPoints(value)
        else:
            raise ValueError(f"不支持的单位: {unit}")

    def set_size(
            self,
            image,
            width: float = None,
            height: float = None,
            unit: str = "pt",
            lock_aspect_ratio: int = -1,
            doc=None
    ):
        try:
            # ===== Set lock aspect ratio =====
            image.LockAspectRatio = lock_aspect_ratio
            msgs = []
            # ===== Lock aspect ratio =====
            if lock_aspect_ratio == -1:
                if width is not None:
                    width_pt = self.convert_to_pt(width, unit)
                    image.Width = width_pt
                    msgs.append(f"width={width}{unit} ({round(width_pt, 2)}pt)")
                elif height is not None:
                    height_pt = self.convert_to_pt(height, unit)
                    image.Height = height_pt
                    msgs.append(f"height={height}{unit} ({round(height_pt, 2)}pt)")
                else:
                    return {
                        "status": "success",
                        "message": "No size parameter provided, nothing changed (aspect ratio locked)"
                    }
            # ===== Do not lock aspect ratio =====
            else:
                if width is not None:
                    width_pt = self.convert_to_pt(width, unit)
                    image.Width = width_pt
                    msgs.append(f"width={width}{unit} ({round(width_pt, 2)}pt)")
                if height is not None:
                    height_pt = self.convert_to_pt(height, unit)
                    image.Height = height_pt
                    msgs.append(f"height={height}{unit} ({round(height_pt, 2)}pt)")
                if not msgs:
                    return {
                        "status": "success",
                        "message": "No size parameter provided, nothing changed (aspect ratio unlocked)"
                    }
            return {
                "status": "success",
                "message": f"Set picture size: {', '.join(msgs)}, lock_aspect_ratio={lock_aspect_ratio}"
            }
        except Exception as e:
            return {
                "status": "error",
                "message": str(e)
            }

    def set_size_by_percent(self,image,percent: float, doc=None):
        """Scale by percentage based on the current size.
        :param image: Image object
        :param percent: Scale factor (e.g. 1.1 means enlarge by 10% from current size)"""
        try:
            # 1. Force-lock aspect ratio so the other dimension follows when one is changed
            image.LockAspectRatio = -1

            # 2. Get the image's current absolute width (usually in pt)
            current_width = image.Width

            # 3. Compute the target width
            # If current is 100pt and percent is 1.5, target is 150pt
            target_width = current_width * percent

            # 4. Apply the new size
            image.Width = target_width

            return {
                "status": "success",
                "message": f"Resized by {percent * 100}%: {round(current_width, 2)}pt -> {round(target_width, 2)}pt (Locked Aspect Ratio)"
            }
        except Exception as e:
            return {
                "status": "error",
                "message": f"Failed to scale by percent: {str(e)}"
            }

    def set_alignment(self,image,alignment:str):
        try:
            # Get the paragraph containing the image
            paragraph = image.Range.Paragraphs(1)
            # Set alignment
            alignment = alignment.lower()
            if alignment in ["left","左对齐"]:
                paragraph.Alignment = 0  # wdAlignParagraphLeft
            elif alignment in ["center","居中对齐"]:
                paragraph.Alignment = 1  # wdAlignParagraphCenter
            elif alignment in ["right","右对齐"]:
                paragraph.Alignment = 2  # wdAlignParagraphRight
            else:
                raise ValueError(f"No supported alignment type: {alignment}, only: left, center, right")

            return {"state": "success", "message": f"Set the image alignment to be {alignment}"}
        except Exception as e:
            return {"state": "error", "message": str(e)}

    def set_keep_with_next(self,image, keep_with_next=None):
        try:
            paragraph = image.Range.Paragraphs(1)
            if keep_with_next is not None:
                paragraph.KeepWithNext = -1 if keep_with_next else 0
            return {"status":"success","message":f"Set image keep_with_next = {keep_with_next}"}
        except Exception as e:
            return {
                    "status": "error",
                    "message": str(e)
            }
    def set_keep_together(self, image, keep_together=None):
        try:
            paragraph = image.Range.Paragraphs(1)
            if keep_together is not None:
                paragraph.KeepTogether = -1 if keep_together else 0
            return {"status": "success", "message": f"Set image keep_together = {keep_together}"}
        except Exception as e:
            return {
                "status": "error",
                "message": str(e)
            }
    def set_page_break_before(self, image, page_break_before=None):
        try:
            paragraph = image.Range.Paragraphs(1)
            if page_break_before is not None:
                paragraph.PageBreakBefore = -1 if page_break_before else 0
            return {"status": "success", "message": f"Set image page_break_before = {page_break_before}"}
        except Exception as e:
            return {
                "status": "error",
                "message": str(e)
            }
    def set_pagination(self, image, keep_with_next=None, keep_together=None,
                             page_break_before=None):
        try:
            result = {}
            if keep_with_next is not None:
                result["keep_with_next"] = self.set_keep_with_next(image,keep_with_next)
            if keep_together is not None:
                result["keep_together"] = self.set_keep_together(image,keep_together)
            if page_break_before is not None:
                result["page_break_before"] = self.set_page_break_before(image,page_break_before)
            return result
        except Exception as e:
            return {"state": "false", "message": str(e)}
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


class ImageTools(BaseTool):
    def __init__(self, pyconfig=ContextToolsConfig("/config/Tools/ImageToolsConfig.yaml")):
        self.config = pyconfig.config
        self.name = self.config.get("name")
        self.image_tool = ImageBaseTools()

    def __set_size(self, doc, image_index, width: float = None, height:float=None,unit: str = "pt", lock_aspect_ratio: int = -1):
        image = self.image_tool.get_image(doc,image_index)
        status = self.image_tool.set_size(image=image,width=width,height=height,unit=unit,lock_aspect_ratio=lock_aspect_ratio,doc=doc)
        return {"image_size":status}

    def __set_alignment(self,doc, image_index, alignment):
        image = self.image_tool.get_image(doc, image_index)
        status = self.image_tool.set_alignment(image, alignment)
        return {"image_alignment": status}

    def __set_pagination(self,doc, image_index, keep_with_next=True, keep_together=True,
                             page_break_before=False):
        image = self.image_tool.get_image(doc,image_index)
        status = self.image_tool.set_pagination(image,keep_with_next=keep_with_next,keep_together=keep_together,page_break_before=page_break_before)
        return status

    def __set_size_by_percent(self,doc, image_index, percent):
        image = self.image_tool.get_image(doc,image_index)
        status = self.image_tool.set_size_by_percent(image,percent)
        return status

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set the size of a specific figure by index",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "image_list": {"type": "list",
                               "description": "Document image position index list. Indexing rules: 'all' = all images; 1 = first image; 2 = second image"},
                "width": { "type": "float", "description": "Width value; set to None if no width setting is specified." },
                "height": {"type": "float",
                           "description": "Height value; set to None if no height setting is specified."},
                "unit": {"type": "str", "description": "Unit of the size value (pt/cm/mm/inches/percent)"},
                "lock_aspect_ratio": {"type": "int",
                                      "description": "Whether to lock the aspect ratio (0 = disable, -1 = enable)"}
            }
        },
        "zh": {
            "function_description": "通过索引设置指定图片的大小",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "image_list": {"type": "list",
                               "description": "文档图片位置索引列表。索引规则：'all' = 所有图片；1 = 第一张图片；2 = 第二张图片"},
                "width": {"type": "float", "description": "宽度值，未说明需要设置时填None"},
                "height": {"type": "float", "description": "高度值，未说明需要设置时填None"},
                "unit": {"type": "str", "description": "大小值的单位（pt/cm/mm/inches）"},
                "lock_aspect_ratio": {"type": "int", "description": "是否锁定长宽比（0 = 关闭，-1 = 启用）"}
            }
        }
    })
    def set_size(self, doc, image_list, width=None, height=None, unit="pt", lock_aspect_ratio=-1,*args, **kwargs):
        results = {}
        if 'all' in image_list:
            image_list = [i + 1 for i in range(doc.InlineShapes.Count)]
        try:
            for image_index in image_list:
                status = self.__set_size(doc=doc, image_index=image_index, width=width, height=height,
                                                 unit=unit, lock_aspect_ratio=lock_aspect_ratio)
                results = status
            doc.Save()
        except Exception as e:
            results["image_size"] = {
                "status": "error",
                "message": f"Failed to set image size, the detail is : {e}"}
        finally:
            return results

    @BaseTool.register_tool({
        "en": {
            "function_description": "Set the alignment of a specific figure by index",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "image_list": {"type": "list","description": "Document image position index list. Indexing rules: 'all' = all images; 1 = first image; 2 = second image"},
                "alignment": {"type": "str", "description": "Alignment type (left/center/right)"}
            }
        },
        "zh": {
            "function_description": "通过索引设置指定图片的对齐方式",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "image_list": {"type": "list", "description": "文档图片位置索引列表。索引规则：'all' = 所有图片；1 = 第一张图片；2 = 第二张图片"},
                "alignment": {"type": "str", "description": "对齐方式（left/center/right）"}
            }
        }
    })
    def set_alignment(self, doc, image_list, alignment="center",*args, **kwargs):
        results = {}
        if 'all' in image_list:
            image_list = [i + 1 for i in range(doc.InlineShapes.Count)]
        try:
            for image_index in image_list:
                status = self.__set_alignment(doc=doc, image_index=image_index, alignment=alignment)
                results = status
            doc.Save()
        except Exception as e:
            results["image_alignment"] = {
                "status": "error",
                "message": f"Failed to set image alignment, the detail is : {e}"}
        finally:
            return results

    @BaseTool.register_tool({
    "en": {
        "function_description": "Set pagination for images in Word document",
        "params": {
            "doc": {"type": "object", "description": "Word document object"},
            "image_list": {"type": "list","description": "Document image position index list. Indexing rules: 'all' = all images; 1 = first image; 2 = second image"},
            "keep_with_next": {"type": "int", "description": "Keep with next paragraph (0 = disable, -1 = enable)", "default": 0},
            "keep_together": {"type": "int", "description": "Keep image content together (0 = disable, -1 = enable)", "default": 0},
            "page_break_before": {"type": "int", "description": "Force page break before (0 = disable, -1 = enable)", "default": 0}
        }
    },
    "zh": {
        "function_description": "设置Word文档中图片的分页控制",
        "params": {
            "doc": {"type": "object", "description": "Word文档对象"},
            "image_list": {"type": "list","description": "文档图片位置索引列表，索引规则：'all'=所有图片；1=第一个图片；2=第二个图片"},
            "keep_with_next": {"type": "int", "description": "与下段同页 (0 = 关闭, -1 = 启用)", "default": 0},
            "keep_together": {"type": "int", "description": "禁止跨页断行 (0 = 关闭, -1 = 启用)", "default": 0},
            "page_break_before": {"type": "int", "description": "段前分页 (0 = 关闭, -1 = 启用)", "default": 0}
        }
    }
})
    def set_pagination(self, doc, image_list, keep_with_next=0, keep_together=0,
                             page_break_before=0,*args, **kwargs):
        results = {}
        if 'all' in image_list:
            image_list = [i + 1 for i in range(doc.InlineShapes.Count)]
        try:
            for image_index in image_list:
                status = self.__set_pagination(doc=doc, image_index=image_index, keep_together=keep_together,keep_with_next=keep_with_next,page_break_before=page_break_before)
                results = status
            doc.Save()
        except Exception as e:
            results["image_pagination"] = {
                "status": "error",
                "message": f"Failed to set image pagination, the detail is : {e}"}
        finally:
            return results

    @BaseTool.register_tool({
        "en": {
            "function_description": "Scale specific figures by a percentage relative to their current size",
            "params": {
                "doc": {"type": "object", "description": "Word document object"},
                "image_list": {"type": "list",
                               "description": "List of document image indices. Rules: 'all' = all images; 1 = first image; 2 = second image"},
                "percent": {"type": "float",
                            "description": "Scaling factor relative to the CURRENT size (e.g., 1.5 = 150% of current size, 0.8 = 80% of current size). Default is 1."}
            }
        },
        "zh": {
            "function_description": "按百分比缩放指定图片（基于当前实际大小）",
            "params": {
                "doc": {"type": "object", "description": "Word文档对象"},
                "image_list": {"type": "list",
                               "description": "文档图片位置索引列表。索引规则：'all' = 所有图片；1 = 第一张图片；2 = 第二张图片"},
                "percent": {"type": "float",
                            "description": "相对于图片【当前实际大小】的缩放比例（例如：1.5表示放大至现在的150%，0.8表示缩小至现在的80%）。默认为1。"}
            }
        }
    })
    def set_size_by_percent(self, doc, image_list, percent=1):
        results = {}
        if 'all' in image_list:
            image_list = [i + 1 for i in range(doc.InlineShapes.Count)]
        try:
            for image_index in image_list:
                status = self.__set_size_by_percent(doc=doc, image_index=image_index, percent=percent)
                results = status
            doc.Save()
        except Exception as e:
            results["image_size_by_percent"] = {
                "status": "error",
                "message": f"Failed to set image size by percent, the detail is : {e}"}
        finally:
            return results


if __name__ == '__main__':
    word = win32.DispatchEx("Word.Application")  # Or use Dispatch
    word.Visible = True  # Make visible (recommended when debugging)
    word_file_path = "./file/image_test.docx"
    word_file_path = os.path.join(ABS_DIR, word_file_path)
    print(word_file_path)
    # Open an existing document
    try:
        # Open the document
        doc = word.Documents.Open(word_file_path)

        image = doc.InlineShapes(1)

        fg_tool = ImageTools()

        # Test centering
        # fg_tool.set_alignment(doc,[1],"right")

        print(fg_tool.set_size(doc,[1],height = 65, width = None, unit="mm", lock_aspect_ratio=-1))
        print(fg_tool.set_alignment(doc, [1], alignment='left'))


    except Exception as e:
        print(f"操作失败：{str(e)}")
        raise

    finally:
        # Ensure resources are cleaned up
        if 'doc' in locals():
            doc.Save()
            doc.Close(SaveChanges=False)
        word.Quit()