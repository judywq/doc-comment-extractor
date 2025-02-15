import os
import json
import argparse
import logging
from comment_extractor import CommentExtractor, ExtractConfig
from formatters.html_formatter import HtmlFormatter

logger = logging.getLogger(__name__)
DEBUG = False


def process_folder(input_folder: str, output_folder: str, config: ExtractConfig, out_formats: set[str]):
    """Process all Word documents in the input folder and save results to output folder."""
    # Create output folders only for requested formats
    output_folders = {}
    for fmt in out_formats:
        output_folders[fmt] = os.path.join(output_folder, fmt)
        os.makedirs(output_folders[fmt], exist_ok=True)
    
    extractor = CommentExtractor(config)
    html_formatter = HtmlFormatter() if 'html' in out_formats else None
    
    for idx, filename in enumerate(os.listdir(input_folder)):
        if filename.endswith('.docx') and not filename.startswith('~$'):
            input_path = os.path.join(input_folder, filename)
            base_name = os.path.splitext(filename)[0]
            logger.info("Processing %s", filename)
            
            result = extractor.process_document(input_path)
            comments = result['comments']
            logger.info("There are %s comments in %s", len(comments), filename)
            
            if result:
                # Save JSON output if requested
                if 'json' in out_formats:
                    json_output_path = os.path.join(output_folders['json'], f"{base_name}.json")
                    with open(json_output_path, 'w', encoding='utf-8') as f:
                        json.dump(result, f, ensure_ascii=False, indent=2)
                    logger.info("Generated JSON %s", json_output_path)
                
                # Generate and save HTML output if requested
                if 'html' in out_formats:
                    html_output_path = os.path.join(output_folders['html'], f"{base_name}.html")
                    html_content = html_formatter.format(result)
                    with open(html_output_path, 'w', encoding='utf-8') as f:
                        f.write(html_content)
                    logger.info("Generated HTML %s", html_output_path)
                
        if DEBUG:
            break


def main():
    parser = argparse.ArgumentParser(description='Extract comments from Word documents')
    parser.add_argument('--input_folder', help='Folder containing Word documents')
    parser.add_argument('--output_folder', help='Folder to save output files')
    parser.add_argument('--start_token', help='Start token for text extraction')
    parser.add_argument('--end_token', help='End token for text extraction')
    parser.add_argument('--author', action='store_true', default=False,
                      help='Include author field in comments (default: False)')
    parser.add_argument('--date', action='store_true', default=False,
                      help='Include date field in comments (default: False)')
    parser.add_argument('--out_formats', default='json',
                      help='Comma-separated list of output formats (json,html) (default: json)')
    
    args = parser.parse_args()
    
    # Parse output formats
    out_formats = {fmt.strip().lower() for fmt in args.out_formats.split(',')}
    valid_formats = {'json', 'html'}
    invalid_formats = out_formats - valid_formats
    if invalid_formats:
        parser.error(f"Invalid output format(s): {', '.join(invalid_formats)}. Valid formats are: {', '.join(valid_formats)}")
    
    config = ExtractConfig(
        start_token=args.start_token,
        end_token=args.end_token,
        include_author=args.author,
        include_date=args.date
    )
    process_folder(args.input_folder, args.output_folder, config, out_formats)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    main()
