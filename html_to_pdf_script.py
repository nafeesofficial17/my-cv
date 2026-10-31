# import pdfkit
from weasyprint import HTML, CSS
from datetime import datetime

file_name = "m-nafees-" + str(datetime.now().date()).replace("-", "") + ".pdf"
# file_name = "mohd-nafees-" + str(datetime.now().date()).replace("-", "") + ".pdf"


# def html_to_pdf():

#     # Path to HTML file
#     html_file = "pdfkit.html"
    
#     # Output PDF file
#     output_pdf = file_name

#     # PDF options (footer, page size, etc.)
#     options = {
#         "page-size": "A4",
#         "margin-top": "10mm",
#         "margin-right": "10mm",
#         "margin-bottom": "10mm",
#         "margin-left": "10mm",
#         "encoding": "UTF-8",
#         "no-outline": None,
#         "enable-local-file-access": None,
#         "print-media-type": None,
#         "dpi": 300,
#         "image-dpi": 300,
#         "image-quality": 100,
#         "zoom": 1.0,
#         "disable-smart-shrinking": None,
#     }



#     # Configure pdfkit (update wkhtmltopdf path if needed)
#     # Example path on Linux: /usr/bin/wkhtmltopdf
#     # Example path on Windows: r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
#     path_wkhtmltopdf = '/usr/bin/wkhtmltopdf'  
#     config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

#     # Convert HTML file to PDF
#     pdfkit.from_file(html_file, output_pdf, configuration=config, options=options)

#     print(f"✅ PDF generated successfully: {output_pdf}")


if __name__ == "__main__":

    # this code pdfkit
    # to generate pdf using pdfkit you need to install wkhtmltopdf package in your system
    # html_to_pdf()

    # this code for weasyprint
    HTML("weasyprint.html").write_pdf(
    # HTML("modify_weasyprint.html").write_pdf(
        file_name,
        presentational_hints=True
        # ,
        # stylesheets=[CSS(string="""
        #     @page { size: A4; margin: 20mm 15mm; }
        #     @media print {
        #         html, body, .container, .main, .sidebar, .block, p, span, li {
        #         hyphens: none !important;
        #         -webkit-hyphens: none !important;
        #         -moz-hyphens: none !important;
        #         }
        #         .block p {
        #         text-align: left !important;
        #         word-break: normal !important;
        #         overflow-wrap: break-word !important;
        #         white-space: normal !important;
        #         }
        #     }
        #     """)]
        )
