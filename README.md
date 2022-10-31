# Generate Word Template From Excel
(en)This apps reads the informations of the first table in a excel file and complete the template word document
(pt-BR)Esse app lê informações dentro do arquivo excel na primeira tabela e transfere para o template word.

## Requirements

Firts of all you need the one of the latest python version installed

Access the directory that you save the "requirements.txt" and open the CMD or Terminal. \
\
And then, run the following command:

```
pip install -r requirements.txt
```
After this, it will install the dependencies required to run the python file. \
\
Make sure that they are installed by running `pip freeze`.

## How to make your word template

To make your word template, you simply have to add the {{column name}} in this format in the word document, the column name needs to be the same that exists in the excel file, but the column name can not have number in the beginning and can not have special characters in any position or spaces, for this objectivo use "_", see the example templates. \
\
Note: There are no script, so you can delete the example template files, the only rule is that you can't add more than one excel(.xlsx) file or more than one word (.docx) file in the same directory as the word_automation.py file.

## How to run the app

Firts, make sure you have all requirements installed, then run in word_automation.py in your terminal.\
\
Then, choose the column do you want to filter, and choose the value of some cell that you want, so you will get that line as result, and it will return the template document filled with the line's information from excel, and that it.\
To generate other files, simply repeat and choose other cells.