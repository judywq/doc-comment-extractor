import os
from abc import ABC, abstractmethod
from dataclasses import dataclass


@dataclass
class FormatterConfig:
    extension: str
    output_subfolder: str


class BaseFormatter(ABC):
    """Base class for formatters that convert extracted comments to different formats."""

    def __init__(self, config: FormatterConfig):
        self.config = config

    @abstractmethod
    def format(self, data: dict) -> str:
        """Format the data into a string."""
        pass

    def save(self, data: dict, base_output_dir: str, filename: str) -> str:
        """
        Save the formatted data to a file.
        Returns the full output path.
        """
        output_path = self.get_output_path(base_output_dir, filename)

        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        # Format and save the data
        formatted_data = self.format(data)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(formatted_data)

        return output_path

    def get_output_path(self, base_output_dir: str, filename: str) -> str:
        """Construct the full output path for the file."""
        base_name = os.path.splitext(filename)[0]
        return os.path.join(
            base_output_dir,
            self.config.output_subfolder,
            f"{base_name}{self.config.extension}",
        )
