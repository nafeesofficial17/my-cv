
# Install GTK+ for Windows (Recommended for WeasyPrint)

Download and install GTK3 runtime from: 
    https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases

Run the installer and make sure to check the option to add GTK to your PATH



# follow these step:

Create venv file:
    python -m venv venv
Activate venv file:
    . .\venv\Scripts\activate  

install requirements:
    pip install -r requirements.txt

run script:
    python .\html_to_pdf_script.py

run DOCX script:
    python .\html_to_docx_script.py

note:
    PDF is built from weasyprint.html (dark sidebar). DOCX is built from docx.html (light sidebar), a separate copy
    with the same content, so make content changes in BOTH files.
    DOCX export rebuilds the PDF layout with Word-native elements: the sidebar on every page, skill tags, the stats
    strip, project cards, right-aligned dates and clickable links.
    Both scripts write m-nafees-YYYYMMDD.pdf / .docx. Colours and sizes in html_to_docx_script.py mirror the CSS in
    docx.html, so if you restyle it, update the constants at the top of the script too.
    Close the .docx in Word before regenerating; otherwise the script saves a -1, -2 ... copy next to it.