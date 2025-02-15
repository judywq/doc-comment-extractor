from typing import Dict, Type
from .base import BaseFormatter, FormatterConfig
from .html_formatter import HtmlFormatter
from .json_formatter import JsonFormatter

class FormatterFactory:
    _formatters: Dict[str, tuple[Type[BaseFormatter], FormatterConfig]] = {
        'json': (JsonFormatter, FormatterConfig(extension='.json', output_subfolder='json')),
        'html': (HtmlFormatter, FormatterConfig(extension='.html', output_subfolder='html'))
    }

    @classmethod
    def get_formatter(cls, name: str) -> BaseFormatter:
        """Get a formatter instance by name."""
        if name not in cls._formatters:
            raise ValueError(f"Unknown formatter: {name}. Valid formatters are: {', '.join(cls._formatters.keys())}")
        
        formatter_class, config = cls._formatters[name]
        return formatter_class(config)

    @classmethod
    def get_valid_formats(cls) -> set[str]:
        """Get set of valid formatter names."""
        return set(cls._formatters.keys()) 
