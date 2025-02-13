import os
import json
from datetime import datetime
from docx import Document
import argparse
from typing import Dict, List, Optional
from xml.etree import ElementTree
import zipfile

DEBUG = True

class CommentExtractor:
    def __init__(self, start_token: str, end_token: str):
        self.start_token = start_token
        self.end_token = end_token
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
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
                # Read comments from word/comments.xml if it exists
                try:
                    comments_xml = zip_ref.read('word/comments.xml')
                    root = ElementTree.fromstring(comments_xml)
                    
                    # Read document.xml to get the comment anchors
                    document_xml = zip_ref.read('word/document.xml')
                    doc_root = ElementTree.fromstring(document_xml)
                    
                    # Create a mapping of comment IDs to their ranges
                    comment_ranges = {}
                    for elem in doc_root.findall('.//w:commentRangeStart', self.namespaces):
                        comment_id = elem.get(f'{{{self.namespaces["w"]}}}id')
                        comment_ranges[comment_id] = {
                            'start': elem,
                            'text': []
                        }
                    
                    # Find comment ends and collect text between start and end
                    for elem in doc_root.iter():
                        if elem.tag == f'{{{self.namespaces["w"]}}}commentRangeEnd':
                            comment_id = elem.get(f'{{{self.namespaces["w"]}}}id')
                            if comment_id in comment_ranges:
                                comment_ranges[comment_id]['end'] = elem
                    
                    # Process each comment
                    for comment in root.findall('.//w:comment', self.namespaces):
                        comment_id = comment.get(f'{{{self.namespaces["w"]}}}id')
                        author = comment.get(f'{{{self.namespaces["w"]}}}author', '')
                        date = comment.get(f'{{{self.namespaces["w"]}}}date', datetime.now().isoformat())
                        
                        # Get comment text
                        comment_text = []
                        for p in comment.findall('.//w:t', self.namespaces):
                            if p.text:
                                comment_text.append(p.text)
                        
                        comments.append({
                            'id': comment_id,
                            'author': author,
                            'date': date,
                            'comment_text': ' '.join(comment_text),
                            'highlighted_text': ''  # Will be filled later
                        })
                
                except KeyError:
                    # No comments in document
                    return []
                
        except Exception as e:
            print(f"Error extracting comments: {str(e)}")
            return []
            
        return comments

    def process_comments(self, doc_path: str, revised_essay: str) -> List[Dict]:
        """Process comments and calculate their positions relative to revised essay."""
        processed_comments = []
        raw_comments = self.extract_comments_from_docx(doc_path)
        
        # Get the full document text to help with positioning
        doc = Document(doc_path)
        full_text = self.get_document_text(doc)
        
        for comment in raw_comments:
            # Try to find the commented text in the revised essay
            # This is a simplified approach - in a real implementation,
            # you might need more sophisticated text matching
            paragraphs = [p.text for p in doc.paragraphs]
            
            for para in paragraphs:
                if para.strip():
                    # Look for this paragraph in the revised essay
                    para_pos = revised_essay.find(para)
                    if para_pos != -1:
                        # Found the paragraph, now we can calculate relative positions
                        start = para_pos
                        end = start + len(para)
                        
                        processed_comments.append({
                            "start": start,
                            "end": end,
                            "highlighted_text": para,
                            "comment_text": comment['comment_text'],
                            "author": comment['author'],
                            "date": comment['date']
                        })
                        break
        
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
