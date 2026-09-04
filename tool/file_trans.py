import win32com.client
from pdf2image import convert_from_path
import pypandoc
from constant import ABS_DIR
from pathlib import Path
import os, json
import re,shutil
from datetime import datetime

class FileConverter:
    def __init__(self, base_dir=None):
        """
        Initialize the converter; optionally set a base directory for relative path resolution.
        :param base_dir: Project root directory; defaults to None (use current working directory).
        """
        self.ABS_DIR = os.path.abspath(base_dir) if base_dir else os.getcwd()


    def resolve_mixed_path(self, input_path, base_dir=None):
        """
        Resolve paths that mix absolute paths with relative path symbols.
        :param input_path: Path that may contain `./` or `../` (e.g. `D:\project\./images`).
        :param base_dir: Base directory (defaults to the current working directory).
        :return: Normalized absolute path.
        """
        # Convert to a Path object and expand relative path symbols
        path = Path(input_path.replace('/', os.sep))  # Normalize to the system path separator
        if not path.is_absolute() and base_dir:
            base = Path(base_dir)
            return str((base / path).resolve())
        # Resolve relative symbols inside an absolute path
        return str(path.resolve())

    def _resolve_path(self, *path_parts):
        combined = os.path.join(*path_parts)
        # Handle mixed absolute/relative path cases
        if any(sym in combined for sym in ('./', '../')):
            return self.resolve_mixed_path(combined, self.ABS_DIR)
        # Standard path resolution
        abs_path = os.path.join(self.ABS_DIR, combined)
        return os.path.abspath(abs_path)


    def md_to_docx(self, input_md, output_docx=None, image_dir=None):
        """
        Convert a Markdown file to a Word document.
        :param input_md: Input .md file path (relative or absolute).
        :param output_docx: Output .docx path (defaults to the input file name).
        :param image_dir: Image directory (relative or absolute).
        """
        try:
            # Resolve paths
            abs_input = self._resolve_path(input_md)
            abs_output = os.path.abspath(output_docx) if output_docx else \
                os.path.join(self.ABS_DIR, os.path.splitext(os.path.basename(input_md))[0] + ".docx")
            abs_image_dir = self._resolve_path(image_dir) if image_dir else os.path.dirname(abs_input)
            # print(abs_input)
            # print(abs_image_dir)
            # Convert the document
            pypandoc.convert_file(
                abs_input,
                "docx",
                outputfile=abs_output,
                format="markdown",
                extra_args=[
                    f"--resource-path={abs_image_dir}",
                    "--extract-media=images"
                ]
            )
            print(f"MD to DOCX succeeded: {abs_output}")
            return abs_output
        except Exception as e:
            print(f"MD to DOCX failed: {e}")
            return None

    def docx_to_pdf(self, input_docx, output_pdf=None):
        """
        Convert a Word document to PDF.
        :param input_docx: Input .docx file path (relative or absolute).
        :param output_pdf: Output .pdf path (defaults to the input file name).
        """
        import tempfile
        from tool.word_com import create_word_app, open_document, release_word

        word = None
        doc = None
        tmp_docx = None
        try:
            abs_input = self._resolve_path(input_docx)
            abs_output = os.path.join(output_pdf, os.path.splitext(os.path.basename(input_docx))[0] + ".pdf") if output_pdf else \
                os.path.join(self.ABS_DIR, os.path.splitext(os.path.basename(input_docx))[0] + ".pdf")
            # Convert a temp copy so the live working docx is never locked by this Word.
            fd, tmp_docx = tempfile.mkstemp(suffix=".docx")
            os.close(fd)
            shutil.copy2(abs_input, tmp_docx)

            word = create_word_app(visible=False)
            doc = open_document(word, tmp_docx, read_only=True)
            # Prefer ExportAsFixedFormat: fewer interactive prompts than SaveAs PDF.
            doc.ExportAsFixedFormat(OutputFileName=abs_output, ExportFormat=17, OpenAfterExport=False)
            print(f"DOCX to PDF succeeded: {abs_output}")
            return abs_output
        except Exception as e:
            print(f"DOCX to PDF failed: {e}")
            raise
        finally:
            release_word(word, doc, save_changes=False)
            if tmp_docx and os.path.isfile(tmp_docx):
                try:
                    os.remove(tmp_docx)
                except Exception:
                    pass

    def pdf_to_images(self, input_pdf, output_folder=None, dpi=300):
        """
        Convert a PDF to a set of images.
        :param input_pdf: Input .pdf file path (relative or absolute).
        :param output_folder: Output directory (defaults to a folder named after the PDF).
        :param dpi: Image resolution (default 300).
        """
        try:
            abs_input = self._resolve_path(input_pdf)
            abs_output = os.path.join(output_folder, "doc_images") if output_folder else \
                os.path.join(self.ABS_DIR, "doc_images")
            os.makedirs(abs_output, exist_ok=True)
            # Convert PDF pages to images
            images = convert_from_path(abs_input, dpi=dpi)
            for i, image in enumerate(images):
                image.save(os.path.join(abs_output, f"page_{i + 1}.png"), "PNG")
            print(f"PDF to images succeeded, saved to: {abs_output}")
            return abs_output
        except Exception as e:
            print(f"PDF to images failed: {e}")
            raise

    def read_json_file(self, file_path, encoding='utf-8'):
        """
        Safely read a JSON file (handles encoding and format errors).

        Args:
            file_path (str): Path to the JSON file.
            encoding (str): File encoding (default utf-8).

        Returns:
            dict/list/None: Parsed data on success, None on failure.
        """
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"[Error] File not found: {file_path}")
        except json.JSONDecodeError as e:
            print(f"[Error] Invalid JSON format: {file_path} reason: {e.msg} (at: line {e.lineno}, column {e.colno})")
        except UnicodeDecodeError:
            print(f"[Error] Encoding issue; try another encoding (e.g. gbk)")
        except Exception as e:
            print(f"[Error] Read failed: {str(e)}")
        return None


    def write_json_file(self, data, file_path, encoding='utf-8', indent=4):
        """
        Safely write a JSON file (creates directories and formats output).

        Args:
            data (dict/list): Python data structure to write.
            file_path (str): Output file path.
            encoding (str): File encoding (default utf-8).
            indent (int): Indent spaces (None for compact format).

        Returns:
            bool: True on success, False on failure.
        """
        try:
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            with open(file_path, 'w', encoding=encoding) as f:
                json.dump(data, f, ensure_ascii=False, indent=indent)
            print(f"{file_path} save successful")
            return True
        except TypeError:
            print("[Error] Data contains non-serializable types")
        except PermissionError:
            print(f"[Error] No write permission: {file_path}")
        except Exception as e:
            print(f"[Error] Write failed: {str(e)}")
        return False

    def create_folder(self, path):
        """
        Create a folder (supports multi-level directory creation).

        Args:
            path (str): Directory path to create.

        Returns:
            bool: True on success, False on failure.
        """
        try:
            # Normalize the path (handle mixed forward/backslash separators)
            normalized_path = os.path.normpath(path)

            # Create the directory with pathlib (handles nested directories)
            folder = Path(normalized_path)
            folder.mkdir(parents=True, exist_ok=True)

            print(f"Folder created successfully: {normalized_path}")
            return True
        except Exception as e:
            print(f"Folder creation failed: {str(e)}")
            return False

    def list_exctract(self, list_result):
        try:
            """Extract all JSON fragments wrapped between ```json and ``` from a string."""
            pattern = r'```json(.*?)```'
            matches = re.findall(pattern, list_result, re.DOTALL)
            # Strip leading/trailing whitespace (e.g. newlines)
            match_result = [match.strip() for match in matches]
            data = json.loads(match_result[0])
            # print(data)
        except Exception as e:
            print(f"Funcall Extract Error! The detail is {e}")
            data = []
        return data

    def prepare_directory(self,dir_path):
        """
        Prepare a directory: clear its contents if it exists, otherwise create it.

        Args:
            dir_path: Target directory path.
        """
        # Check whether the directory exists
        if os.path.exists(dir_path):
            # Check that it is a directory
            if os.path.isdir(dir_path):
                # Iterate over and delete all contents under the directory
                for item in os.listdir(dir_path):
                    item_path = os.path.join(dir_path, item)
                    try:
                        # If it is a file or symlink, delete it directly
                        if os.path.isfile(item_path) or os.path.islink(item_path):
                            os.unlink(item_path)
                        # If it is a directory, delete recursively
                        elif os.path.isdir(item_path):
                            shutil.rmtree(item_path)
                    except Exception as e:
                        print(f"Failed to delete {item_path}: {e}")
        else:
            # Directory does not exist; create it (including any parent directories)
            try:
                os.makedirs(dir_path, exist_ok=True)
                print(f"Directory {dir_path} created successfully")
            except Exception as e:
                print(f"Failed to create directory {dir_path}: {e}")

    def save_error(self, save_dir, content="Json Parser Error!"):
        now = datetime.now()
        file_name = f"JsonParserError_{now.year}-{now.month}-{now.day}-{now.hour}-{now.minute}-{now.second}.json"
        self.write_json_file(data={"Error":content},file_path=os.path.join(save_dir,file_name))


    def save_fail(self, save_dir):
        file_name = f"TaskFail.json"
        self.write_json_file(data={"Error":"Task Failed!"},file_path=os.path.join(save_dir,file_name))

# Usage example
if __name__ == "__main__":
    # Initialize (ABS_DIR is computed from the current script directory)
    converter = FileConverter(base_dir=ABS_DIR)
    # Example conversion pipeline
    try:
        # 1. MD to DOCX
        docx_path = converter.md_to_docx(
            "./file/Word_case.md",
            image_dir="./file"  # Relative image path: point to the parent of the image folder, not the image folder itself
        )
        # 2. DOCX to PDF
        pdf_path = converter.docx_to_pdf(docx_path)
        # 3. PDF to images
        converter.pdf_to_images(pdf_path, dpi=300)

    except Exception:
        print("Conversion pipeline interrupted")
