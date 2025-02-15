from typing import Dict
from .base import BaseFormatter
from setting import (
    ESSAY_TEXT_KEY,
    COMMENTS_KEY,
    COMMENT_START_KEY,
    COMMENT_END_KEY,
    COMMENT_TEXT_KEY,
)


class XmlFormatter(BaseFormatter):
    """Formatter that converts comment data to XML with inline comment tags."""

    def format(self, json_data: Dict) -> str:
        essay_text = json_data[ESSAY_TEXT_KEY]
        comments = json_data[COMMENTS_KEY]

        if not essay_text or not comments:
            return ""

        # Create a list of all positions where we need to insert tags
        positions = []
        for i, comment in enumerate(comments):
            # Add start tag position
            positions.append(
                (comment[COMMENT_START_KEY], "start", i, comment[COMMENT_TEXT_KEY])
            )
            # Add end tag position
            positions.append((comment[COMMENT_END_KEY], "end", i, None))

        # Sort positions by index, with end tags coming before start tags at same position
        positions.sort(key=lambda x: (x[0], 0 if x[1] == "end" else 1))

        # Build the XML string
        result = []
        current_pos = 0

        for pos, tag_type, comment_id, comment_text in positions:
            # Add text before the tag
            result.append(essay_text[current_pos:pos])

            # Add the tag
            if tag_type == "start":
                result.append(
                    f'<comment-start id="{comment_id}" data="{self._escape_xml(comment_text)}"/>'
                )
            else:
                result.append(f'<comment-end id="{comment_id}"/>')

            current_pos = pos

        # Add remaining text
        result.append(essay_text[current_pos:])

        return "".join(result)

    def _escape_xml(self, text: str) -> str:
        """Escape special characters for XML."""
        if text is None:
            return ""
        return (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&apos;")
        )
