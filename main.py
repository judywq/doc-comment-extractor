from dataclasses import dataclass
import os
import json
from datetime import datetime
import argparse
from typing import Dict, List, Optional
from xml.etree import ElementTree
import zipfile
import logging
from exception import StartTokenNotFound

logger = logging.getLogger(__name__)

DEBUG = True

@dataclass
class Comment:
    id: str
    para_id: str
    para_id_parent: Optional[str]
    author: str
    date: str
    comment_text: str
    highlighted_text: str
    start: int
    end: int
    
    def get_dict(self) -> Dict:
        return {
            "start": self.start,
            "end": self.end,
            "highlighted_text": self.highlighted_text,
            "comment_text": self.comment_text,
            "author": self.author,
            "date": self.date
        }

@dataclass
class HighlightRange:
    """Highlight range with relative start position."""
    comment_id: str
    absolute_start: int
    section_start: int
    texts: List[str]
    
    def __init__(self, comment_id: str, absolute_start: int):
        self.comment_id = comment_id
        self.absolute_start = absolute_start
        self.section_start = -1
        self.texts = []
    
    def append(self, text: str):
        self.texts.append(text)
    
    def get_text(self) -> str:
        return ''.join(self.texts)
    
    def get_relative_start(self) -> int:
        return self.absolute_start - self.section_start

@dataclass
class Section:
    start: int
    end: int
    raw_text: str
    stripped_text: str

class CommentExtractor:
    def __init__(self, start_token: str, end_token: str):
        self.start_token = start_token
        self.end_token = end_token
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'
        }

    def extract_text_between_tokens(self, text: str) -> Section:
        """Extract text between start and end tokens."""
        try:
            start_idx = text.index(self.start_token) + len(self.start_token)
            end_idx = text.index(self.end_token, start_idx)
            raw_text = text[start_idx:end_idx]
            lstripped_text = raw_text.lstrip()
            stripped_text = lstripped_text.rstrip()
            blank_chars_before_start_token = len(raw_text) - len(lstripped_text)
            blank_chars_before_end_token = len(lstripped_text) - len(stripped_text)
            
            return Section(start=start_idx + blank_chars_before_start_token, 
                           end=end_idx - blank_chars_before_end_token, 
                           raw_text=raw_text, 
                           stripped_text=stripped_text)
        except ValueError:
            return None

   
    
    def _read_docx_file(self, file_path: str) -> tuple[ElementTree.Element, ElementTree.Element, ElementTree.Element]:
        try:
            with zipfile.ZipFile(file_path) as zip_ref:
                # Read XML files
                comments_root = self._read_xml_files(zip_ref, 'word/comments.xml')
                comments_extend_root = self._read_xml_files(zip_ref, 'word/commentsExtended.xml')
                doc_root = self._read_xml_files(zip_ref, 'word/document.xml')
                return comments_root, comments_extend_root, doc_root
        except Exception as e:
            logger.error("Error reading %s: %s", file_path, str(e))
            return None, None, None

    def _read_xml_files(self, zip_ref, inner_file_name: str) -> ElementTree.Element:
        """Read and parse XML files from the Word document."""
        try:
            xml = zip_ref.read(inner_file_name)
            return ElementTree.fromstring(xml)
        except KeyError as e:
            raise ValueError(f"Required XML file not found in document: {e}")

    def _extract_highlight_ranges(self, doc_root) -> Dict[str, HighlightRange]:
        """Extract comment ranges and their corresponding text from document with position info."""
        current_range_ids = []
        comment_id_to_range = {}
        full_text = ""
        is_first_paragraph = True
                
        for elem in doc_root.iter():                
            if elem.tag == f'{{{self.namespaces["w"]}}}commentRangeStart':
                comment_id = elem.get(f'{{{self.namespaces["w"]}}}id')
                current_range_ids.append(comment_id)
                hr = HighlightRange(comment_id=comment_id, 
                                    absolute_start=len(full_text))
                comment_id_to_range[comment_id] = hr
            elif elem.tag == f'{{{self.namespaces["w"]}}}commentRangeEnd':
                comment_id = elem.get(f'{{{self.namespaces["w"]}}}id')
                if comment_id in comment_id_to_range:
                    current_range_ids.remove(comment_id)
            elif elem.tag == f'{{{self.namespaces["w"]}}}t':
                if elem.text:
                    for comment_id in current_range_ids:
                        hr = comment_id_to_range[comment_id]
                        hr.append(elem.text)
                    full_text += elem.text
            elif elem.tag == f'{{{self.namespaces["w"]}}}p':
                if is_first_paragraph:
                    is_first_paragraph = False
                else:
                    # Add new line to all highlight ranges, if not the first paragraph
                    new_line = "\n"
                    full_text += new_line
                    for comment_id in current_range_ids:
                        hr = comment_id_to_range[comment_id]
                        hr.append(new_line)

        return comment_id_to_range, full_text

    def _get_sub_comments(self, comments_extend_root) -> set[str]:
        """Get set of paragraph IDs that are replies to other comments."""
        sub_comments = set()
        for comment_ex in comments_extend_root.findall('.//w15:commentEx', self.namespaces):
            para_id = comment_ex.get(f'{{{self.namespaces["w15"]}}}paraId')
            parent_para_id = comment_ex.get(f'{{{self.namespaces["w15"]}}}paraIdParent', None)                    
            if parent_para_id is not None:
                sub_comments.add(para_id)
        return sub_comments

    def _extract_comment_text(self, comment_node) -> str:
        """Extract text from a comment node."""
        comment_text = []
        for p in comment_node.findall('.//w:t', self.namespaces):
            if p.text:
                comment_text.append(p.text)
        return ' '.join(comment_text)

    def extract_comments_from_docx(self, docx_path: str) -> List[Comment]:
        """Extract comments directly from the Word document's XML structure."""
        try:
            # Read XML files
            comments_root, comments_extend_root, doc_root = self._read_docx_file(docx_path)
            
            # Extract comment ranges and their text
            comment_id_to_range, full_text = self._extract_highlight_ranges(doc_root)
            
            section = self.extract_text_between_tokens(full_text)
            
            # Get sub-comments (replies)
            sub_comments = self._get_sub_comments(comments_extend_root)
            
            # Process main comments
            comments = []
            for comment_node in comments_root.findall('.//w:comment', self.namespaces):
                comment_id = comment_node.get(f'{{{self.namespaces["w"]}}}id')
                para = comment_node.find('.//w:p', self.namespaces)
                
                if para is None:
                    continue
                    
                para_id = para.get(f'{{{self.namespaces["w14"]}}}paraId')
                if para_id in sub_comments:
                    continue  # Skip reply comments
                
                highlighted_range = comment_id_to_range.get(comment_id, None)
                if highlighted_range is None:
                    logger.warning("Skip comment. No highlighted text found for comment %s", comment_id)
                    continue
                
                highlighted_range.section_start = section.start
                
                pos = highlighted_range.get_relative_start()
                if pos < 0:
                    logger.warning("Skip comment. Comment %s appears before start token", comment_id)
                    continue
                
                comment = Comment(
                    id=comment_id,
                    para_id=para_id,
                    para_id_parent=None,
                    author=comment_node.get(f'{{{self.namespaces["w"]}}}author', ''),
                    date=comment_node.get(f'{{{self.namespaces["w"]}}}date', datetime.now().isoformat()),
                    comment_text=self._extract_comment_text(comment_node),
                    highlighted_text=highlighted_range.get_text(),
                    start=pos,
                    end=pos + len(highlighted_range.get_text())
                )
                comments.append(comment.get_dict())
            
            return comments, section
            
        except Exception as e:
            logger.error("Error extracting comments: %s", str(e))
            return [], None

    def process_document(self, file_path: str) -> Optional[Dict]:
        """Process a single document and return the JSON structure."""
        result = {
            "revised_essay": None,
            "comments": []
        }
        try:            
            # Process comments
            comments, section = self.extract_comments_from_docx(file_path)
            revised_essay = section.stripped_text
            if not revised_essay:
                logger.warning("Could not find tokens in %s", file_path)
                return result
            
            result["revised_essay"] = revised_essay
            result["comments"] = comments
            return result
            
        except Exception as e:
            logger.error("Error processing %s: %s", file_path, str(e))
            return result

def process_folder(input_folder: str, output_folder: str, start_token: str, end_token: str, 
                  include_author: bool = False, include_date: bool = False):
    """Process all Word documents in the input folder and save results to output folder."""
    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    
    extractor = CommentExtractor(start_token, end_token)
    
    # Process each .docx file in the input folder
    for idx, filename in enumerate(os.listdir(input_folder)):
        if filename.endswith('.docx'):
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}.json")
            logger.info("Processing %s", filename)
            
            result = extractor.process_document(input_path)
            
            # Filter out author and date if not requested
            if result and not include_author:
                for comment in result['comments']:
                    comment.pop('author', None)
            if result and not include_date:
                for comment in result['comments']:
                    comment.pop('date', None)
                    
            comments = result['comments']
            logger.info("There are %s comments in %s", len(comments), filename)
            if result:
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(result, f, ensure_ascii=False, indent=2)
                logger.info("Processed %s -> %s", filename, output_path)
        if DEBUG:
            break

def main():
    parser = argparse.ArgumentParser(description='Extract comments from Word documents')
    parser.add_argument('--input_folder', help='Folder containing Word documents')
    parser.add_argument('--output_folder', help='Folder to save JSON files')
    parser.add_argument('--start_token', help='Start token for text extraction')
    parser.add_argument('--end_token', help='End token for text extraction')
    parser.add_argument('--author', action='store_true', default=False,
                      help='Include author field in comments (default: False)')
    parser.add_argument('--date', action='store_true', default=False,
                      help='Include date field in comments (default: False)')
    
    args = parser.parse_args()
    
    process_folder(args.input_folder, args.output_folder, args.start_token, args.end_token,
                  include_author=args.author, include_date=args.date)

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    main()
