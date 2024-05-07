import logging
import pdfplumber
from docx import Document


logging.basicConfig(filename='pdf_to_docx.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def convert_pdf_to_docx(pdf_file, docx_file):
    try:
        pdf = pdfplumber.open(pdf_file)
        doc = Document()

        for page in pdf.pages:
            text = page.extract_text()
            doc.add_paragraph(text)

        doc.save(docx_file)
        logging.info(f'PDF-файл "{pdf_file}" успешно преобразован в документ Word "{docx_file}"')
    except Exception as e:
        logging.error(f'Ошибка при преобразовании PDF в DOCX: {e}')


if __name__ == "__main__":
    pdf_file = input("Введите имя PDF-файла для преобразования: ")
    docx_file = input("Введите имя DOCX-файла для сохранения результата: ")
    convert_pdf_to_docx(pdf_file, docx_file)
