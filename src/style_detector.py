from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os
import re

class MarkdownStyleDetector:
    def __init__(self):
        self.required_styles = set()  # Stores required styles
        self.current_level = 1  # Default paragraph level

    def scan_markdown_styles(self, markdown_path):
        """Scans a Markdown file and prints the required Word styles."""
        self.required_styles.clear()  # Reset required styles

        with open(markdown_path, 'r', encoding='utf-8') as md_file:
            lines = md_file.readlines()

        for line in lines:
            line = line.strip()

            if not line:
                continue  # Skip empty lines

            # Detect headers (H1 to H6)
            match = re.match(r'^(#{1,6})\s*(.*)', line)
            if match:
                level = len(match.group(1))
                self.required_styles.add(f"Heading {level}")
                self.current_level = level  # Set paragraph level
                continue

            # Detect lists
            if re.match(r'^[-*+]\s+', line):
                self.required_styles.add("List Bullet")
                continue
            elif re.match(r'^\d+\.\s+', line):
                self.required_styles.add("List Number")
                continue

            # Detect blockquotes
            if line.startswith(">"):
                self.required_styles.add("Quote")
                continue

            # Detect inline code
            if re.search(r'`.+?`', line):
                self.required_styles.add("Code")
                continue

            # Normal paragraph text, assign to last known level
            self.required_styles.add(f"Body Text {self.current_level}")

        # Debug print required styles
        print("\nRequired Styles for Markdown Conversion:")
        for style in sorted(self.required_styles):
            print(f"- {style}")

    def get_all_styles(self, template_path):
        """Lists all styles, including hidden ones, in the Word template."""
        doc = Document(template_path)
        styles = {s.name for s in doc.styles}

        # Try accessing latent styles (hidden styles)
        if hasattr(doc.styles, 'latent_styles'):
            latent_styles = doc.styles.latent_styles
            for style_id in latent_styles.element.iterchildren():
                style_name = style_id.get("w:styleName")
                if style_name:
                    styles.add(style_name)  # Add hidden styles to detected list


        print("\n✅ All Available Styles in Template:")
        for style in sorted(styles):
            print(f"- {style}")

        return styles

    def ensure_styles_exist(self, template_path):
       
        """ 1st Checks if the Word template exists and is not locked before processing."""
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template not found: {template_path}")

        try:
            doc = Document(template_path)  # Load Word template
        except PermissionError:
            raise PermissionError(f"❌ The template file is currently open. Please close '{template_path}' and try again.")
        
        """ 2nd Checks and creates missing styles in the Word template."""
        existing_styles = self.get_all_styles(template_path)
        missing_styles = [s for s in self.required_styles if s not in existing_styles]
        
        """ 3rd Creates missing styles in the Word template."""

        if missing_styles:
            print("\n⚠️ Missing Styles Detected. Creating them in the Word Template...")
            for style in missing_styles:
                print(f"- Creating: {style}")
                self._create_style(doc, style)
                
        """ 4th Saves the updated template with new styles."""
        # Save the updated template with new styles
        updated_template_path = template_path.replace(".dotx", "_updated.dotx").replace(".dotm", "_updated.dotm")
        doc.save(updated_template_path)
        print(f"\n✅ Updated Word Template saved as: {updated_template_path}")

        return updated_template_path  # Return new template path

        print("\n✅ All required styles are available in the Word template.")
        return template_path  # Return original template if nothing was changed

    def _create_style(self, doc, style_name):
        """Creates a missing style with default formatting."""
        styles = doc.styles

        if style_name.startswith("Heading"):
            level = int(style_name.split(" ")[1])
            new_style = styles.add_style(style_name, 1)  # 1 = Paragraph Style
            new_style.font.bold = True
            new_style.font.size = Pt(16 - (level * 2))  # Reduce size for deeper levels
            new_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        elif style_name.startswith("Body Text"):
            new_style = styles.add_style(style_name, 1)
            new_style.font.size = Pt(11)
            new_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        elif style_name == "Quote":
            new_style = styles.add_style(style_name, 1)
            new_style.font.italic = True
            new_style.font.size = Pt(11)
            new_style.paragraph_format.left_indent = Pt(15)
            new_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        elif style_name == "List Bullet":
            new_style = styles.add_style(style_name, 1)
            new_style.paragraph_format.left_indent = Pt(10)

        elif style_name == "List Number":
            new_style = styles.add_style(style_name, 1)
            new_style.paragraph_format.left_indent = Pt(10)

        elif style_name == "Code":
            new_style = styles.add_style(style_name, 1)
            new_style.font.name = "Courier New"
            new_style.font.size = Pt(10)
            new_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

# Now, missing styles will be created automatically before the conversion begins.
