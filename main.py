import os
import openpyxl
import platform
from docx import Document
from dotenv import load_dotenv


load_dotenv()
workbook = openpyxl.load_workbook(filename=os.getenv('VARIABLES_FILE'))
sheet = workbook.worksheets[0]


def get_variables():
    variables_dictionary = {}

    # get variables list
    header_iterator = 0
    for column in sheet.iter_cols(1, sheet.max_column):
        variables_dictionary[column[0].value] = header_iterator
        header_iterator += 1

    return variables_dictionary


if __name__ == '__main__':
    os.makedirs(os.getenv('DOCX_OUTPUT_DIRECTORY'), exist_ok=True)

    variables = get_variables()

    for row in sheet.iter_rows(min_row=2):
        # variables should always include File column
        output_docx = f'{os.getenv('DOCX_OUTPUT_DIRECTORY')}{row[variables['File']].value}.docx'
        output_pdf = f'{os.getenv('PDF_OUTPUT_DIRECTORY')}{row[variables['File']].value}.pdf'
        doc = Document(os.getenv('TEMPLATE_FILE'))
        for paragraph in doc.paragraphs:
            for variable in [*variables.keys()]:
                placeholder = f'[[{variable}]]'
                if placeholder in paragraph.text:
                    inline = paragraph.runs
                    for i in range(len(inline)):
                        if placeholder in inline[i].text:
                            text = inline[i].text.replace(placeholder, row[variables[variable]].value)
                            inline[i].text = text
        doc.save(output_docx)
        if platform.system() == 'Linux':
            # TODO
            print('Linux support to be done')
        else:
            os.system(f'rocketpdf parsedocxs {os.getenv('DOCX_OUTPUT_DIRECTORY')}')

