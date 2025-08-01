from docxtpl import DocxTemplate

class Factory:
    @staticmethod
    def create_document(destiny_path, model_path, data):
        file_name = f"{destiny_path}/{data[2].replace(' ', '_').lower()}.docx"
        model = DocxTemplate(model_path)

        context = {
            'VAR_CURSO': data[0],
            'VAR_PROCESSO': data[1],
            'VAR_NOME': data[2],
            'VAR_REGIS': data[3]
        }

        model.render(context)
        model.save(file_name)