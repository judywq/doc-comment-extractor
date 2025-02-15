from abc import ABC, abstractmethod
from typing import Dict

class BaseFormatter(ABC):
    """Base class for formatters that convert extracted comments to different formats."""
    
    @abstractmethod
    def format(self, data: Dict) -> str:
        """Convert the input data to the target format."""
        pass 
    
    @abstractmethod
    def save(self, data: Dict, output_path: str) -> None:
        """Save the formatted data to a file."""
        pass
