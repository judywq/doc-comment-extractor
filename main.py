from dataclasses import dataclass
import os
import json
from datetime import datetime
from docx import Document
import argparse
from typing import Dict, List, Optional
from xml.etree import ElementTree
import zipfile
import logging

DEBUG = True
logger = logging.getLogger(__name__)

@dataclass
class Comment:
    id: str
    para_id: str
    para_id_parent: Optional[str]
    author: str
    date: str
    comment_text: str
    highlighted_text: str


class CommentExtractor:
    def __init__(self, start_token: str, end_token: str):
        self.start_token = start_token
        self.end_token = end_token
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'
        }

    def extract_text_between_tokens(self, text: str) -> Optional[str]:
        """Extract text between start and end tokens."""
        try:
            start_idx = text.index(self.start_token) + len(self.start_token)
            end_idx = text.index(self.end_token, start_idx)
            return text[start_idx:end_idx].strip()
        except ValueError:
            return None

    def get_document_text(self, doc) -> str:
        """Extract full text from the document."""
        return " ".join(paragraph.text for paragraph in doc.paragraphs)

    def _read_xml_files(self, zip_ref) -> tuple[ElementTree.Element, ElementTree.Element, ElementTree.Element]:
        """Read and parse XML files from the Word document."""
        try:
            comments_xml = zip_ref.read('word/comments.xml')
            comments_extended_xml = zip_ref.read('word/commentsExtended.xml')
            document_xml = zip_ref.read('word/document.xml')
            
            return (
                ElementTree.fromstring(comments_xml),
                ElementTree.fromstring(comments_extended_xml),
                ElementTree.fromstring(document_xml)
            )
        except KeyError as e:
            raise ValueError(f"Required XML file not found in document: {e}")

    def _extract_comment_ranges(self, doc_root) -> Dict[str, str]:
        """Extract comment ranges and their corresponding text from document."""
        current_range_ids = []
        comment_id_to_text = {}
        
        for elem in doc_root.iter():
            if elem.tag == f'{{{self.namespaces["w"]}}}commentRangeStart':
                range_id = elem.get(f'{{{self.namespaces["w"]}}}id')
                current_range_ids.append(range_id)
                comment_id_to_text[range_id] = []
            elif elem.tag == f'{{{self.namespaces["w"]}}}commentRangeEnd':
                range_id = elem.get(f'{{{self.namespaces["w"]}}}id')
                if range_id in current_range_ids:
                    comment_id_to_text[range_id] = ''.join(comment_id_to_text[range_id])
                    current_range_ids.remove(range_id)
            elif elem.tag == f'{{{self.namespaces["w"]}}}t':
                if elem.text and len(current_range_ids) > 0:
                    for range_id in current_range_ids:
                        comment_id_to_text[range_id].append(elem.text)
        
        return comment_id_to_text

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
            with zipfile.ZipFile(docx_path) as zip_ref:
                # Read XML files
                comments_root, comments_extend_root, doc_root = self._read_xml_files(zip_ref)
                
                # Extract comment ranges and their text
                comment_id_to_text = self._extract_comment_ranges(doc_root)
                
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
                    
                    highlighted_text = comment_id_to_text.get(comment_id, None)
                    if highlighted_text is None:
                        logger.warning("Skip comment. No highlighted text found for comment %s", comment_id)
                        continue
                    comment = Comment(
                        id=comment_id,
                        para_id=para_id,
                        para_id_parent=None,
                        author=comment_node.get(f'{{{self.namespaces["w"]}}}author', ''),
                        date=comment_node.get(f'{{{self.namespaces["w"]}}}date', datetime.now().isoformat()),
                        comment_text=self._extract_comment_text(comment_node),
                        highlighted_text=highlighted_text
                    )
                    comments.append(comment)
                
                return comments
                
        except zipfile.BadZipFile:
            logger.error("Error: %s is not a valid Word document", docx_path)
            return []
        except Exception as e:
            logger.error("Error extracting comments: %s", str(e))
            return []

    def process_comments(self, doc_path: str, revised_essay: str) -> List[Dict]:
        """Process comments and calculate their positions relative to revised essay."""
        processed_comments = []
        raw_comments = self.extract_comments_from_docx(doc_path)
        
        for comment in raw_comments:
            highlighted_text = comment.highlighted_text
            if highlighted_text:
                # Find the position of the highlighted text in the revised essay
                start = revised_essay.find(highlighted_text)
                if start != -1:
                    end = start + len(highlighted_text)
                    
                    processed_comments.append({
                        "start": start,
                        "end": end,
                        "highlighted_text": highlighted_text,
                        "comment_text": comment.comment_text,
                        "author": comment.author,
                        "date": comment.date
                    })
        
        return processed_comments

    def process_document(self, file_path: str) -> Optional[Dict]:
        """Process a single document and return the JSON structure."""
        result = {
            "revised_essay": None,
            "comments": []
        }
        try:
            doc = Document(file_path)
            
            # Get full document text
            full_text = self.get_document_text(doc)
            
            # Extract text between tokens
            revised_essay = self.extract_text_between_tokens(full_text)
            if not revised_essay:
                logger.warning("Could not find tokens in %s", file_path)
                return result
            
            # Process comments
            comments = self.process_comments(file_path, revised_essay)
            
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
    for filename in os.listdir(input_folder):
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
