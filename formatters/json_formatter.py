import json
import os
from typing import Dict, Any

class JsonFormatter:
    def format(self, data: Dict[str, Any]) -> str:
        """Convert the data to a JSON string."""
        return json.dumps(data, ensure_ascii=False, indent=2)
    
    def save(self, data: Dict[str, Any], output_path: str) -> None:
        """Format and save the data to a JSON file."""
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(self.format(data)) 