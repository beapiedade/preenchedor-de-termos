from extractor import Extractor

def extractFiles(excel_bytes):
    excel_data = Extractor.from_path(excel_bytes)
    
    if not excel_data:
        return []
    
    headers = ['VAR_CURSO', 'VAR_PROCESSO', 'VAR_NOME', 'VAR_REGIS']
    dictionaries = [dict(zip(headers, row)) for row in excel_data]

    return dictionaries

__main__.extractFiles = extractFiles