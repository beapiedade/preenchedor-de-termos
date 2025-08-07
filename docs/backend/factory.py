from docxtpl import DocxTemplate
import io 

class Factory:
    @staticmethod
    def create_document(model_path, data):
        model = DocxTemplate(model_path)

        context = {
            'VAR_CURSO': data[0],
            'VAR_PROCESSO': data[1],
            'VAR_NOME': data[2],
            'VAR_REGIS': data[3]
        }

        model.render(context)

        doc_buffer = io.BytesIO()
        model.save(doc_buffer)
        doc_buffer.seek(0)
        
        return doc_buffer.getvalue()
        