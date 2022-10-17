import glob
import os
from email.mime import base
from pathlib import Path

import pandas as pd
from docxtpl import DocxTemplate


def main ():

    base_dir = Path(__file__).parent
    word_template_path =  base_dir / glob.glob('./*docx')
    excel_path = base_dir  / glob.glob('./*xlsx')
    output_dir = base_dir / 'OUTPUTS'

    output_dir.mkdir(exist_ok=True)

    df = pd.read_excel(excel_path)

    print(df.to_dict(orient='records'))

    filter = input('Do you want to filter by what coolumn?')

    value_filter = input('Choose the value of the corresponding filter to generate the report:')

    # file_name = input('Choose the name for the report file:')

    for record in df.to_dict(orient='records'):
        if record[filter] == value_filter:
            doc = DocxTemplate(word_template_path)
            doc.render(record)
            output_path = output_dir / f"{value_filter}.docx"
            doc.save(output_path)
        else:
            ValueError()

if __name__ == '__main__':
    main()

# Marcos Aurélio Chaves