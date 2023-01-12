from zipfile import ZipFile

with ZipFile('Температура.docx', 'r') as myzip:
    myzip.extractall(path='Архив')