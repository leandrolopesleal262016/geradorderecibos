from flask import Flask, render_template, request, jsonify, send_file
from docx import Document
from docx.shared import Cm, Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import pandas as pd
from datetime import datetime
import os
import zipfile
import locale
from datetime import datetime
data_atual = datetime.now().strftime('%d/%m/%Y')

app = Flask(__name__)
fornecedores_df = None
documentos_gerados = []  # Declaração global

def traduzir_mes(mes_en):
    meses = {
        'January': 'janeiro',
        'February': 'fevereiro',
        'March': 'março',
        'April': 'abril',
        'May': 'maio',
        'June': 'junho',
        'July': 'julho',
        'August': 'agosto',
        'September': 'setembro',
        'October': 'outubro',
        'November': 'novembro',
        'December': 'dezembro'
    }
    return meses.get(mes_en, mes_en)

def formatar_data_atual():
    data_atual = datetime.now()
    mes_en = data_atual.strftime('%B')
    mes_pt = traduzir_mes(mes_en)
    return data_atual.strftime(f'%d de {mes_pt} de %Y')

def processar_modelo(modelo_conteudo, dados_cliente):
    texto = modelo_conteudo.replace('{cliente_nome}', dados_cliente['nome'])
    texto = texto.replace('{valor}', dados_cliente['valor'])
    texto = texto.replace('{valor_extenso}', dados_cliente['valor_extenso'])
    texto = texto.replace('{numero_documento}', dados_cliente['numero_documento'])
    texto = texto.replace('{data}', dados_cliente['data'])
    return texto.split('\n')

def valor_por_extenso(valor):
    valor = float(valor)
    if valor == 0:
        return 'zero reais'
        
    int_value = int(valor)
    decimal_value = int((valor - int_value) * 100)

    unidades = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove']
    dezenas = ['', 'dez', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa']
    centenas = ['', 'cento', 'duzentos', 'trezentos', 'quatrocentos', 'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos']
    
    extenso = []
    
    # Parte inteira
    if int_value > 0:
        if int_value == 1:
            extenso.append('um real')
        else:
            # Lógica para converter o número
            if int_value >= 100:
                centena = int(int_value / 100)
                extenso.append(centenas[centena])
                int_value = int_value % 100
            
            if int_value >= 10:
                dezena = int(int_value / 10)
                extenso.append(dezenas[dezena])
                int_value = int_value % 10
            
            if int_value > 0:
                extenso.append(unidades[int_value])
            
            extenso.append('reais')
    
    # Parte decimal
    if decimal_value > 0:
        extenso.append('e')
        if decimal_value >= 10:
            dezena = int(decimal_value / 10)
            extenso.append(dezenas[dezena])
            decimal_value = decimal_value % 10
        
        if decimal_value > 0:
            extenso.append(unidades[decimal_value])
        
        extenso.append('centavos')
    
    return ' '.join(extenso)

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
        dados = request.json
        clientes_selecionados = dados.get('clientes', [])
        valor = dados.get('valor', '0.00')
        documentos_gerados = []
        preview_content = []
        
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        data_atual = datetime.now().strftime('%d de %B de %Y')

        for cliente_nome in clientes_selecionados:
            cliente_data = fornecedores_df[fornecedores_df['Razão social'] == cliente_nome].iloc[0].to_dict()
            
            doc = Document()
            
            # Configuração das margens
            sections = doc.sections
            for section in sections:
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)
            
            # Adiciona logo e informações da empresa em uma tabela
            header_table = doc.add_table(rows=1, cols=2)
            header_table.autofit = False
            header_table.columns[0].width = Inches(2.0)
            
            # Célula da logo
            logo_cell = header_table.cell(0, 0)
            logo_paragraph = logo_cell.paragraphs[0]
            logo_run = logo_paragraph.add_run()
            logo_run.add_picture('static/images/logo.png', width=Inches(2.0))
            
            # Célula das informações da empresa
            info_cell = header_table.cell(0, 1)
            info_cell.text = "BEIJO E MATOS CONSTRUÇÕES E ENGENHARIA LTDA\nJoaquim da Silva Martha, 12-53 - Sala 3 - Altos da Cidade - Bauru/SP\nguilhermebeijo@bencato.com.br - CNPJ: 26.149.105/0001-09 - www.bencato.com.br"
            
            doc.add_paragraph()  # Espaço
            
            # Linha do recibo e valor
            recibo_table = doc.add_table(rows=1, cols=2)
            recibo_cell = recibo_table.cell(0, 0)
            recibo_cell.text = f"RECIBO Nº {str(cliente_data.get('document_number', '0001'))} - parcela única"
            valor_cell = recibo_table.cell(0, 1)
            valor_cell.text = f"VALOR: R$ {valor}"
            
            # Informações do cliente e descrição
            p = doc.add_paragraph()
            p.add_run(f"ADMINISTRATIVO - {cliente_nome}\n").bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            valor_extenso = valor_por_extenso(valor)

            p.add_run(f"Recebi(emos) a quantia de R$ {valor} ({valor_extenso}) na forma de pagamento em dinheiro, correspondente a serviços prestados e para maior clareza firmo(amos) o presente.")
            
            # Data alinhada à direita
            data_paragraph = doc.add_paragraph()
            data_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            data_paragraph.add_run(f"Bauru, {data_atual}")
            
            doc.add_paragraph()  # Espaço
            
            # Assinatura centralizada
            assinatura_paragraph = doc.add_paragraph()
            assinatura_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            assinatura_paragraph.add_run("_" * 50 + "\n")
            assinatura_paragraph.add_run(cliente_nome.upper() + "\n")
            assinatura_paragraph.add_run(str(cliente_data.get('CPF/CNPJ', '')))

            # Salva o documento
            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            documentos_gerados.append((cliente_nome, doc_buffer.getvalue()))

            # Preview content mantém o formato anterior para compatibilidade
            preview_content.append({
                'nome': cliente_nome,
                'conteudo': [str(p.text) for p in doc.paragraphs]
            })

        return jsonify({
            'preview': preview_content,
            'status': 'success'
        })

    except Exception as e:
        print(f"Erro: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/download_recibos', methods=['GET'])
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
        zip_buffer.seek(0)# Prepara o buffer para leitura
                
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


