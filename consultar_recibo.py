from app import app
from models import db, ReciboGerado
from docx import Document
import io

def consultar_recibo(recibo_id):
    # Inicializa o contexto da aplicação
    with app.app_context():
        # Consulta o recibo pelo ID
        recibo = ReciboGerado.query.get(recibo_id)
        if recibo:
            print(f"Recibo ID: {recibo.id}")
            print(f"Número do Recibo: {recibo.numero_recibo}")
            print(f"Nome do Cliente: {recibo.cliente_nome}")
            print(f"Valor: {recibo.valor}")
            print(f"Data de Geração: {recibo.data_geracao}")
            # Se precisar ver o conteúdo do documento
            doc = Document(io.BytesIO(recibo.documento_blob))
            for paragrafo in doc.paragraphs:
                print(paragrafo.text)
        else:
            print("Recibo não encontrado.")

# Chame a função com o ID do recibo que deseja consultar
consultar_recibo(46)
