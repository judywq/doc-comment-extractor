from dataclasses import dataclass
import os
import json
from datetime import datetime
from docx import Document
import argparse
from typing import Dict, List, Optional
from xml.etree import ElementTree
import zipfile

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

    def extract_comments_from_docx(self, docx_path: str) -> List[Dict]:
        """Extract comments directly from the Word document's XML structure."""
        comments = []
        
        try:
            with zipfile.ZipFile(docx_path) as zip_ref:
                try:
                    # Read both comments.xml and commentsExtended.xml
                    comments_xml = zip_ref.read('word/comments.xml')
                    comments_extended_xml = zip_ref.read('word/commentsExtended.xml')
                    document_xml = zip_ref.read('word/document.xml')
                    
                    comments_root = ElementTree.fromstring(comments_xml)
                    comments_extend_root = ElementTree.fromstring(comments_extended_xml)
                    doc_root = ElementTree.fromstring(document_xml)
                    
                    
                    # First pass: find all comment ranges and their text
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

                    # Create a list of sub-comments
                    sub_comments = []
                    for comment_ex in comments_extend_root.findall('.//w15:commentEx', self.namespaces):
                        para_id = comment_ex.get(f'{{{self.namespaces["w15"]}}}paraId')
                        parent_para_id = comment_ex.get(f'{{{self.namespaces["w15"]}}}paraIdParent', None)                    
                        if parent_para_id is not None:
                            sub_comments.append(para_id)
                    
                    for comment_node in comments_root.findall('.//w:comment', self.namespaces):
                        comment_id = comment_node.get(f'{{{self.namespaces["w"]}}}id')
                        para = comment_node.find('.//w:p', self.namespaces)
                        if para is not None:
                            para_id = para.get(f'{{{self.namespaces["w14"]}}}paraId')
                            if para_id in sub_comments:
                                # Skip sub-comments (replies)
                                continue
                            
                            comment_text = []
                            for p in comment_node.findall('.//w:t', self.namespaces):
                                if p.text:
                                    comment_text.append(p.text)                            
                            comment = Comment(
                                id=comment_id,
                                para_id=para_id,
                                para_id_parent=None,
                                author=comment_node.get(f'{{{self.namespaces["w"]}}}author', ''),
                                date=comment_node.get(f'{{{self.namespaces["w"]}}}date', datetime.now().isoformat()),
                                comment_text=' '.join(comment_text),
                                highlighted_text=comment_id_to_text.get(comment_id, 'ERROR: No highlighted text found')
                            )
                            comments.append(comment)

                except KeyError as e:
                    print(f"No comments found in document: {e}")
                    return []
                
        except Exception as e:
            print(f"Error extracting comments: {str(e)}")
            return []
            
        return comments

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
        try:
            doc = Document(file_path)
            
            # Get full document text
            full_text = self.get_document_text(doc)
            
            # Extract text between tokens
            revised_essay = self.extract_text_between_tokens(full_text)
            if not revised_essay:
                print(f"Warning: Could not find tokens in {file_path}")
                return None
            
            # Process comments
            comments = self.process_comments(file_path, revised_essay)
            
            return {
                "revised_essay": revised_essay,
                "comments": comments
            }
            
        except Exception as e:
            print(f"Error processing {file_path}: {str(e)}")
            return None

def process_folder(input_folder: str, output_folder: str, start_token: str, end_token: str):
    """Process all Word documents in the input folder and save results to output folder."""
    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    
    extractor = CommentExtractor(start_token, end_token)
    
    # Process each .docx file in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith('.docx'):
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}.json")
            
            result = extractor.process_document(input_path)
            comments = result['comments']
            print(f"Number of comments: {len(comments)}")
            if result:
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(result, f, ensure_ascii=False, indent=2)
                print(f"Processed {filename} -> {output_path}")
        if DEBUG:
            break

def main():
    parser = argparse.ArgumentParser(description='Extract comments from Word documents')
    parser.add_argument('--input_folder', help='Folder containing Word documents')
    parser.add_argument('--output_folder', help='Folder to save JSON files')
    parser.add_argument('--start_token', help='Start token for text extraction')
    parser.add_argument('--end_token', help='End token for text extraction')
    
    args = parser.parse_args()
    
    process_folder(args.input_folder, args.output_folder, args.start_token, args.end_token)

if __name__ == "__main__":
    main()
