from docx import Document

def update_hyperlinks(doc_path, link_mapping):
    # Load the document
    doc = Document(doc_path)

    # Access the document's relationship part for hyperlinks
    rels = doc.part.rels

    # Loop through each relationship
    for rel in rels.values():
        if "hyperlink" in rel.target_ref:  # Check if it is a hyperlink
            old_url = rel.target_ref
            if old_url in link_mapping:
                # Replace with the new URL
                rel._target = link_mapping[old_url]

    # Save the updated document
    doc.save('updated_document.docx')

# Example usage
link_mapping = {
    'http://oldlink1.com': 'http://newlink1.com',
    'http://oldlink2.com': 'http://newlink2.com'
}
update_hyperlinks('example.docx', link_mapping)
