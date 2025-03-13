import os
import aspose.words as aw

def get_docx_metadata(file_path):
    doc = aw.Document(file_path)
    properties = doc.built_in_document_properties
    word_count = properties.words
    page_count = properties.pages
    
    # Extract titles and headings
    titles_and_headings = []
    for node in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
        paragraph = node.as_paragraph()
        if paragraph is not None and paragraph.paragraph_format.style_identifier in [aw.StyleIdentifier.HEADING1, aw.StyleIdentifier.HEADING2, aw.StyleIdentifier.HEADING3]:
            titles_and_headings.append(paragraph.get_text().strip())
    
    metadata = {
        "Title": properties.title,
        "Author": properties.author,
        "Word Count": word_count,
        "Page Count": page_count,
        "Titles and Headings": titles_and_headings
    }
    
    return metadata

def main():
    docx_folder = 'docx'
    with open('aspose.md', 'w') as md_file:
        for filename in os.listdir(docx_folder):
            if filename.endswith('.docx'):
                file_path = os.path.join(docx_folder, filename)
                metadata = get_docx_metadata(file_path)
                md_file.write(f"## Metadata for {filename}\n")
                for key, value in metadata.items():
                    md_file.write(f"**{key}**: {value}\n")
                md_file.write("\n")

if __name__ == "__main__":
    main()