from typing import Dict
from .base import BaseFormatter
import os

class HtmlFormatter(BaseFormatter):
    """Formatter that converts comment data to HTML with styled tooltips."""
    
    def format(self, json_data: Dict) -> str:
        essay_text = json_data["revised_essay"]
        comments = sorted(json_data["comments"], key=lambda x: x["start"])
        
        if not essay_text or not comments:
            return ""
        
        return self._generate_html(essay_text, comments)
    
    def _generate_html(self, essay_text: str, comments: list) -> str:
        html = self._get_html_template()
        result_text = []
        current_pos = 0
        
        for comment in comments:
            # Add text before the comment
            result_text.append(self._process_text(essay_text[current_pos:comment["start"]]))
            
            # Add highlighted text with tooltip
            highlighted = self._process_text(essay_text[comment["start"]:comment["end"]])
            comment_text = self._process_text(comment["comment_text"])
            result_text.append(
                f'<span class="highlighted">{highlighted}'
                f'<span class="tooltip">{comment_text}</span></span>'
            )
            
            current_pos = comment["end"]
        
        # Add remaining text
        result_text.append(self._process_text(essay_text[current_pos:]))
        
        # Join all text and close HTML tags
        html += "".join(result_text)
        html += "\n</body>\n</html>"
        
        return html
    
    def _process_text(self, text: str) -> str:
        """Convert newlines to <br> tags and escape HTML special characters."""
        return text.replace('\n', '<br><br>')
    
    def _get_html_template(self) -> str:
        return """
        <html>
        <head>
        <style>
            body { 
                font-family: Arial, sans-serif;
                line-height: 1.6;
                max-width: 800px;
                margin: 200px auto;
                padding: 0 20px;
            }
            .highlighted {
                background-color: #fff3cd;
                position: relative;
                cursor: pointer;
                display: inline-block;
                white-space: pre-wrap;
            }
            .tooltip {
                visibility: hidden;
                background-color: #333;
                color: white;
                text-align: left;
                padding: 8px;
                border-radius: 4px;
                position: absolute;
                z-index: 1;
                width: 200px;
                font-size: 14px;
                bottom: 100%;
                left: 50%;
                transform: translateX(-50%);
                margin-bottom: 5px;
            }
            .highlighted:hover .tooltip {
                visibility: visible;
            }
            .tooltip::after {
                content: "";
                position: absolute;
                top: 100%;
                left: 50%;
                margin-left: -5px;
                border-width: 5px;
                border-style: solid;
                border-color: #333 transparent transparent transparent;
            }
        </style>
        </head>
        <body>
        """ 
    
    def save(self, data: dict, output_path: str) -> None:
        """Format and save the data to an HTML file."""
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(self.format(data)) 