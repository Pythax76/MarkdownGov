import os

# Define the project structure
project_name = "MarkdownGov"
folders = [
    f"{project_name}/src",
    f"{project_name}/templates",
    f"{project_name}/docs"
]
files = {
    f"{project_name}/main.py": "# Main script for MarkdownGov",
    f"{project_name}/src/__init__.py": "# Init file for src package",
    f"{project_name}/src/converter.py": "# Markdown to Word conversion logic",
    f"{project_name}/src/metadata.py": "# Document metadata handling",
    f"{project_name}/templates/default.dotx": "",  # Placeholder for Word template
    f"{project_name}/docs/README.md": "# MarkdownGov Documentation",
    f"{project_name}/requirements.txt": "python-docx\nmarkdown",
    f"{project_name}/setup.py": "# Setup script for packaging (optional)",
    f"{project_name}/.gitignore": "*.pyc\n__pycache__/\n.env\n"
}

# Create directories
for folder in folders:
    os.makedirs(folder, exist_ok=True)

# Create files
for filepath, content in files.items():
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(content)

# Workspace is now created.
project_name
