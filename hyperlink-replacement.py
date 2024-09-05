from docx import Document

def update_hyperlinks(doc_path, old_url, new_url):
    # Load the document
    doc = Document(doc_path)

    # Loop through paragraphs to find and replace hyperlinks
    for paragraph in doc.paragraphs:
        if paragraph.hyperlinks:  # Check if the paragraph contains a hyperlink
            for run in paragraph.runs:
                if old_url in run.text:  # Check if the run contains the old URL
                    run.text = run.text.replace(old_url, new_url)

    # Save the updated document
    doc.save('updated_document.docx')

# Usage
update_hyperlinks('test.docx', 'https://docs.anl.gov/main/groups/intranet/@shared/@lms/documents/procedure/lms-proc-52.pdf', 'https://my.anl.gov/esb/view/LMS-PROC-52')
