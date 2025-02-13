import os
import json
from datetime import datetime
from docx import Document
import argparse
from typing import Dict, List, Optional

class CommentExtractor:
    def __init__(self, start_token: str, end_token: str):
        self.start_token = start_token
        self.end_token = end_token

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

    def process_comments(self, doc, revised_essay: str) -> List[Dict]:
        """Process comments and calculate their positions relative to revised essay."""
        comments = []
        
        # Get all comments from the document
        for comment in doc.comments:
            # Get the parent paragraph of the comment
            parent = comment._element.parent
            
            # Get the commented text
            highlighted_text = comment.parent.text if hasattr(comment, 'parent') else ""
            
            # Find the position in the revised essay
            start = revised_essay.find(highlighted_text)
            if start != -1:
                end = start + len(highlighted_text)
                
                comments.append({
                    "start": start,
                    "end": end,
                    "highlighted_text": highlighted_text,
                    "comment_text": comment.text,
                    "author": comment.author,
                    "date": comment._element.date.isoformat() if hasattr(comment._element, 'date') else datetime.now().isoformat()
                })
        
        return comments

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
            comments = self.process_comments(doc, revised_essay)
            
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
