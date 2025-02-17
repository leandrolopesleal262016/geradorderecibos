from flask import Flask, render_template, request, jsonify, send_file
from docx import Document
import io
import pandas as pd
from datetime import datetime
import os
import zipfile

app = Flask(__name__)
fornecedores_df = None
documentos_gerados = []  # Declaração global

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_clientes', methods=['GET'])
def get_clientes():
    global fornecedores_df
    try:
        if fornecedores_df is None:
            print("Carregando CSV...")
            fornecedores_df = pd.read_csv('Fornecedores.csv', sep=';', encoding='windows-1252', skiprows=4)
            
        print("Dados do DataFrame:")
        print(f"Total de registros: {len(fornecedores_df)}")
        print(f"Colunas: {fornecedores_df.columns.tolist()}")
        
        # Filtragem e processamento
        fornecedores_df['CPF/CNPJ'] = fornecedores_df['CPF/CNPJ'].astype(str).str.replace(r'\D', '', regex=True)
        
        # Log dos documentos
        print("\nAmostra de CPF/CNPJ:")
        print(fornecedores_df['CPF/CNPJ'].head())
        
        empresas_df = fornecedores_df[fornecedores_df['CPF/CNPJ'].str.len() > 11]
        pessoas_df = fornecedores_df[fornecedores_df['CPF/CNPJ'].str.len() == 11]
        
        empresas = empresas_df['Razão social'].dropna().tolist()
        pessoas = pessoas_df['Razão social'].dropna().tolist()
        
        print("\nDetalhes da separação:")
        print(f"Empresas encontradas: {len(empresas)}")
        print(f"Pessoas encontradas: {len(pessoas)}")
        print("\nPessoas físicas identificadas:")
        for pessoa in pessoas:
            print(f"- {pessoa}")
            
        response_data = {
            'empresas': empresas,
            'pessoas': pessoas
        }
        # print("\nDados enviados para frontend:", response_data)
        
        return jsonify(response_data)
        
    except Exception as e:
        print(f"Erro detalhado: {str(e)}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500
@app.route('/generate_receipts_bulk', methods=['POST'])
def generate_receipts_bulk():
    global documentos_gerados
    try:
        clientes_selecionados = request.json.get('clientes', [])
        documentos_gerados = []
        preview_content = []

        for cliente_nome in clientes_selecionados:
            cliente_data = fornecedores_df[fornecedores_df['Razão social'] == cliente_nome].iloc[0].to_dict()
            
            # Converte valores numéricos para string
            valor = str(float(cliente_data.get('amount', '0.00')))
            
            content = [
                "guilhermebeijo@bencato.com.br - CNPJ: 26.149.105/0001-09 - www.bencato.com.br",
                "Relatório de recibos",
                f"RECIBO Nº {str(cliente_data.get('document_number', '0001'))} - parcela única",
                f"VALOR: R$ {valor}",
                f"ADMINISTRATIVO - {cliente_nome}",
                f"Recebi(emos) a quantia de R$ {valor}",
                f"na forma de pagamento {str(cliente_data.get('payment_method', ''))}",
                f"correspondente a {str(cliente_data.get('description', ''))}",
                "e para maior clareza firmo(amos) o presente.",
                f"Bauru, {datetime.now().strftime('%d de %B de %Y')}",
                cliente_nome.upper(),
                str(cliente_data.get('CPF/CNPJ', ''))
            ]
            
            # Resto do código permanece igual
            doc = Document()
            
            # Adiciona o conteúdo ao documento
            for line in content:
                paragraph = doc.add_paragraph()
                run = paragraph.add_run(str(line))
                if "RECIBO" in line or "VALOR" in line:
                    run.bold = True

            # Salva o documento
            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            documentos_gerados.append((cliente_nome, doc_buffer.getvalue()))

            # Adiciona à preview
            preview_content.append({
                'nome': cliente_nome,
                'conteudo': content
            })

        return jsonify({
            'preview': preview_content,
            'status': 'success'
        })

    except Exception as e:
        print(f"Erro: {str(e)}")
        return jsonify({'error': str(e)}), 500@app.route('/download_recibos', methods=['GET'])
def download_recibos():
    global documentos_gerados
    try:
        if not documentos_gerados:
            return jsonify({'error': 'Nenhum documento disponível para download'}), 404

        # Cria um buffer em memória para o ZIP
        zip_buffer = io.BytesIO()
        
        # Cria o arquivo ZIP com modo de escrita binária
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for nome, doc_content in documentos_gerados:
                # Cria um buffer temporário para cada documento
                doc_buffer = io.BytesIO(doc_content)
                
                # Nome do arquivo limpo (remove caracteres especiais)
                nome_arquivo = "".join(c for c in nome if c.isalnum() or c in (' ', '-', '_'))
                
                # Adiciona o documento ao ZIP
                zip_file.writestr(f"recibo_{nome_arquivo}.docx", doc_buffer.getvalue())
                
                # Fecha o buffer temporário
                doc_buffer.close()

        # Prepara o buffer para leitura
        zip_buffer.seek(0)
        
        # Envia o arquivo
        response = send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='recibos.zip'
        )
        
        # Adiciona headers para evitar cache
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        
        return response

    except Exception as e:
        print(f"Erro no download: {str(e)}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)