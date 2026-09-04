import yaml
from llm.openai_client import OpenaiLLMImageClient
from constant import ABS_DIR
from tool.evaluation.format_reader import FormatReaderTool
import win32com.client as win32
import os,copy
from pathlib import Path
from tool.file_trans import FileConverter
from tool.word_com import create_word_app, open_document, release_word
from tool.reader.page_reader import PageReader
from tool.reader.text_reader import TextReader
from tool.reader.table_reader import TableReader
from tool.reader.image_reader import ImageReader

class DocInform():
    def __init__(self):
        self.file_tool = FileConverter()
        self.reader_tool = FormatReaderTool()
        self.page_reader = PageReader()
        self.text_reader = TextReader()
        self.table_reader = TableReader()
        self.image_reader = ImageReader()
        self.use_cache = True

    def docx_to_images(self, doc_path, save_dir):
        """Convert a docx document page-by-page into images and save them."""
        save_dir = os.path.join(ABS_DIR, save_dir)
        # Doc to PDF
        self.file_tool.create_folder(save_dir)
        pdf_path = self.file_tool.docx_to_pdf(doc_path, save_dir)
        # PDF to images
        image_dir = self.file_tool.pdf_to_images(pdf_path, save_dir, dpi=300)
        return image_dir

    def info_get(self, doc_path, save_dir, scope='page',only_key = 'all',is_vision = True,**kwargs):
        if self.use_cache:
            os.makedirs(save_dir, exist_ok=True)
        else:
            self.file_tool.prepare_directory(save_dir)
        if is_vision:
            self.docx_to_images(doc_path=doc_path,save_dir=os.path.join(ABS_DIR,save_dir))
        valid_scopes = ["page", "text", "table", "image"]
        if scope not in valid_scopes:
            raise ValueError(f"Invalid scope: {scope}. Must be one of {valid_scopes}")
        word = create_word_app(visible=False)
        doc = None
        try:
            doc = open_document(word, doc_path)
            get_info_dict = {
            "page":self.get_page_info,
            "text":self.get_text_info,
            "table":self.get_tables_info,
            "image":self.get_image_info
            }
            return get_info_dict[scope](doc,save_dir=os.path.join(ABS_DIR,save_dir,"meta"),only_key=only_key,**kwargs)
        except Exception as e:
            print(f"Meta data extraction failed: {str(e)}")
            raise
        finally:
            release_word(word, doc, save_changes=False)

    def get_page_info(self, doc, save_dir, content_x,*args,**kwargs):
        page_result_path = os.path.join(save_dir, "page_format.json")
        page_info_result = []
        for section_index in range(1, doc.Sections.Count + 1):
            section_page_result = self.page_reader.get_section_properties(doc,section_index=section_index,content_x=content_x)
            if section_page_result.get("state") == "success":
                section_page_result.pop("state")
                page_info_result.append(section_page_result)
        self.file_tool.write_json_file(data=page_info_result,file_path=page_result_path)
        return page_info_result

    def get_image_info(self, doc, save_dir,*args,**kwargs):
        image_result_path = os.path.join(save_dir, "image_format.json")
        images_info_result = self.image_reader.get_images_info(doc=doc,save_dir=save_dir,**kwargs)
        self.file_tool.write_json_file(data=images_info_result, file_path=image_result_path)
        return images_info_result

    def get_tables_info(self, doc, save_dir,*args,**kwargs):
        table_result_path = os.path.join(save_dir,"table_format.json")
        table_infos_result = self.table_reader.get_table_infos(doc, before = kwargs.get('before',0),after=kwargs.get('after',0))
        self.file_tool.write_json_file(data=table_infos_result, file_path=table_result_path)
        return table_infos_result

    def get_text_info(self, doc, save_dir, only_key = 'all',*args,**kwargs):
        text_result_path = os.path.join(save_dir, "text_format.json")
        text_info_result = []
        for paragraph_index in range(1, doc.Paragraphs.Count + 1):
            text_result = self.text_reader.get_paragraph_info(doc, index=paragraph_index,only_key=only_key)
            if text_result.get("state") == "success" and text_result.get("properties"):
                text_info_result.append(text_result)
        text_info_by_page = self.get_text_info_result_by_page(text_info_result)
        self.file_tool.write_json_file(data=text_info_by_page, file_path=text_result_path)
        return text_info_by_page


    def get_text_info_result_by_page(self, text_info_result):
        text_info_by_page = {}
        for text_info in text_info_result:
            # Use a deep copy so each entry is a fully independent data copy
            text_info_copy = copy.deepcopy(text_info).get("properties")
            # print(text_info_copy)
            # Get the page range and remove the page_range field
            page_range = text_info_copy.pop("page_range")
            # print(page_range)
            page_start = page_range.get("start_page")
            page_end = page_range.get("end_page")

            # Handle the start page
            if page_start not in text_info_by_page:
                text_info_by_page[page_start] = []
            text_info_by_page[page_start].append(text_info_copy)

            # If start and end pages differ, also add to the end page
            if page_start != page_end:
                if page_end not in text_info_by_page:
                    text_info_by_page[page_end] = []
                text_info_by_page[page_end].append(text_info_copy)

        return text_info_by_page



