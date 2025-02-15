import os
import argparse
import logging
from comment_extractor import CommentExtractor, ExtractConfig
from formatters.html_formatter import HtmlFormatter
from formatters.json_formatter import JsonFormatter
from formatters.formatter_factory import FormatterFactory

logger = logging.getLogger(__name__)
DEBUG = False


def process_folder(input_folder: str, output_folder: str, config: ExtractConfig, out_formats: set[str]):
    """Process all Word documents in the input folder and save results to output folder."""
    # Initialize formatters based on requested formats
    formatters = {
        fmt: FormatterFactory.get_formatter(fmt)
        for fmt in out_formats
    }
    
    extractor = CommentExtractor(config)
    
    for idx, filename in enumerate(os.listdir(input_folder)):
        if filename.endswith('.docx') and not filename.startswith('~$'):
            input_path = os.path.join(input_folder, filename)
            logger.info("Processing %s", filename)
            
            result = extractor.process_document(input_path)
            comments = result['comments']
            logger.info("There are %s comments in %s", len(comments), filename)
            
            if result:
                # Save output in each requested format
                for fmt, formatter in formatters.items():
                    output_path = formatter.save(result, output_folder, filename)
                    logger.info("Generated %s: %s", fmt.upper(), output_path)
                
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
    valid_formats = FormatterFactory.get_valid_formats()
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
