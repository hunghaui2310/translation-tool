# from PyPDF2 import PdfFileWriter, PdfReader
import PyPDF2
import textract

if __name__ == '__main__':

    reader = PyPDF2.PdfReader('input.pdf')
    # pdf_writer = PdfFileWriter()

    # Iterate over each page in the PDF
    for page in reader.pages:
        # Replace text here
        # text = page.extract_text().replace('old text', 'new text')
        # page.set_text(text)
        # pdf_writer.add_page(page)

        # Replace text here
        # new_text = text.replace('old text', 'new text')
        text = textract.process('path/to/pdf/file', method='pdfminer')
        print(text)
        # pdf_page.clearContent()
        # pdf_page.appendToPage(PyPDF2.pdf.ContentStream(new_text.encode('utf-8'), pdf_page.pdf))
        # pdf_writer.addPage(pdf_page)

    # Write the output PDF
    # with open('output.pdf', 'wb') as f:
    #     pdf_writer.write(f)