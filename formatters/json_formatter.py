import json
from typing import Dict, Any
from .base import BaseFormatter

class JsonFormatter(BaseFormatter):
    def format(self, data: Dict[str, Any]) -> str:
        """Convert the data to a JSON string."""
        return json.dumps(data, ensure_ascii=False, indent=2)
