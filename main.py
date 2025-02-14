import os
import json
import argparse
import logging
from comment_extractor import CommentExtractor, ExtractConfig
from formatters.html_formatter import HtmlFormatter

logger = logging.getLogger(__name__)
DEBUG = False


def process_folder(input_folder: str, output_folder: str, config: ExtractConfig):
    """Process all Word documents in the input folder and save results to output folder."""
    # Create separate folders for json and html outputs
    json_output_folder = os.path.join(output_folder, 'json')
    html_output_folder = os.path.join(output_folder, 'html')
    os.makedirs(json_output_folder, exist_ok=True)
    os.makedirs(html_output_folder, exist_ok=True)
    
    extractor = CommentExtractor(config)
    html_formatter = HtmlFormatter()
    
    for idx, filename in enumerate(os.listdir(input_folder)):
        if filename.endswith('.docx') and not filename.startswith('~$'):
            input_path = os.path.join(input_folder, filename)
            base_name = os.path.splitext(filename)[0]
            json_output_path = os.path.join(json_output_folder, f"{base_name}.json")
            html_output_path = os.path.join(html_output_folder, f"{base_name}.html")
            logger.info("Processing %s", filename)
            
            result = extractor.process_document(input_path)
            comments = result['comments']
            logger.info("There are %s comments in %s", len(comments), filename)
            
            if result:
                # Save JSON output
                with open(json_output_path, 'w', encoding='utf-8') as f:
                    json.dump(result, f, ensure_ascii=False, indent=2)
                logger.info("Processed %s -> %s", filename, json_output_path)
                
                # Generate and save HTML output
                html_content = html_formatter.format(result)
                with open(html_output_path, 'w', encoding='utf-8') as f:
                    f.write(html_content)
                logger.info("Generated HTML %s", html_output_path)
                
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
    
    config = ExtractConfig(
        start_token=args.start_token,
        end_token=args.end_token,
        include_author=args.author,
        include_date=args.date
    )
    process_folder(args.input_folder, args.output_folder, config)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    main()
