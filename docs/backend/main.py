from extractor import Extractor
from factory import Factory

def generate(excel_bytes, template_bytes):
    data, error = Extractor.from_path(excel_bytes)
    
    usable_files = {}

    if data:
        for item in data:
            try:
                content = Factory.create_document(template_bytes, item)

                file_name = str(item[2]) if item[2] else 'documento_sem_nome'
                file_name = f"{file_name.replace(' ', '_').lower()}.docx"
                usable_files[file_name] = content

            except Exception:
                error.append(item)

    unusable_files = [str(error[2]) if error and len(error) > 2 else "Registro inválido" for error in error]

    return {
        "arquivos": usable_files,
        "erros": unusable_files
    }