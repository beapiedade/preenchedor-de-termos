import io 
from extractor import Extractor

def extractFiles(excel_bytes):
    excel_bytes = excel_bytes.to_py()
    excel_buffer = io.BytesIO(excel_bytes)
    excel_data = Extractor.from_buffer(excel_buffer)
    final_data, errors = Extractor.check_data(excel_data)

    if not final_data:
        return []
    
    headers = ['VAR_CURSO', 'VAR_PROCESSO', 'VAR_NOME', 'VAR_REGIS']
    data_dict = [dict(zip(headers, row)) for row in final_data]
    errors_dict = {i: row for i, row in enumerate(errors) if row}

    return data_dict, errors_dict