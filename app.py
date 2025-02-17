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
    # Limpa o valor removendo pontos e trocando vírgula por ponto
    valor_limpo = str(valor).replace('.', '').replace(',', '.')
    valor = float(valor_limpo)
    
    int_value = int(valor)
    decimal_value = int(round((valor - int_value) * 100))  # Arredonda para evitar erros de precisão

    unidades = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove']
    dezenas = ['', 'dez', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa']
    dezenas_especiais = ['dez', 'onze', 'doze', 'treze', 'quatorze', 'quinze', 'dezesseis', 'dezessete', 'dezoito', 'dezenove']
    centenas = ['', 'cento', 'duzentos', 'trezentos', 'quatrocentos', 'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos']
    
    extenso = []
    extenso_centavos = ''
    
    # Parte inteira
    if int_value >= 1000:
        milhares = int(int_value / 1000)
        if milhares == 1:
            extenso.append('um mil')
        else:
            extenso.extend(numero_para_extenso(milhares))
            extenso.append('mil')
        int_value = int_value % 1000
        if int_value > 0 and int_value < 100:
            extenso.append('e')
    
    if int_value >= 100:
        centena = int(int_value / 100)
        if centena == 1 and int_value % 100 == 0:
            extenso.append('cem')
        else:
            extenso.append(centenas[centena])
        int_value = int_value % 100
        if int_value > 0:
            extenso.append('e')
    
    if int_value >= 10 and int_value <= 19:
        extenso.append(dezenas_especiais[int_value - 10])
    else:
        if int_value >= 10:
            dezena = int(int_value / 10)
            extenso.append(dezenas[dezena])
            int_value = int_value % 10
            if int_value > 0:
                extenso.append('e')
        
        if int_value > 0:
            extenso.append(unidades[int_value])
    
    extenso.append('reais')
    
    # Parte dos centavos
    if decimal_value > 0:
        extenso.append('e')
        if decimal_value >= 10 and decimal_value <= 19:
            extenso_centavos = dezenas_especiais[decimal_value - 10]
        else:
            dezena_centavos = int(decimal_value / 10)
            unidade_centavos = decimal_value % 10
            
            if dezena_centavos > 0:
                extenso_centavos = dezenas[dezena_centavos]
                if unidade_centavos > 0:
                    extenso_centavos += ' e ' + unidades[unidade_centavos]
            elif unidade_centavos > 0:
                extenso_centavos = unidades[unidade_centavos]
        
        extenso.append(extenso_centavos)
        extenso.append('centavos')
    
    return ' '.join(extenso)

def numero_para_extenso(numero):
    unidades = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove']
    dezenas = ['', 'dez', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa']
    
    extenso = []
    if numero >= 10:
        dezena = int(numero / 10)
        extenso.append(dezenas[dezena])
        numero = numero % 10
        if numero > 0:
            extenso.append('e')
    
    if numero > 0:
        extenso.append(unidades[numero])
    
    return extenso
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
                modelo_selecionado = dados.get('modelo')
                clientes_selecionados = dados.get('clientes', [])
                valor = dados.get('valor', '0.00')
                numero_documento = dados.get('numero_documento', '')
                data = dados.get('data', datetime.now().strftime('%d/%m/%Y'))

                documentos_gerados = []
                preview_content = []

                # Recupera o modelo do localStorage
                modelo_texto = request.json.get('modeloConteudo')

                for cliente_nome in clientes_selecionados:
                    # Verifica se o cliente existe no DataFrame
                    cliente_filtrado = fornecedores_df[fornecedores_df['Razão social'] == cliente_nome]

                    if cliente_filtrado.empty:
                        print(f"Cliente não encontrado: {cliente_nome}")
                        continue

                    cliente_data = cliente_filtrado.iloc[0].to_dict()

                    doc = Document()
                    # Configuração das margens
                    sections = doc.sections
                    for section in sections:
                        section.left_margin = Inches(1)
                        section.right_margin = Inches(1)
                    # Ajuste das colunas e formatação do texto
                    header_table = doc.add_table(rows=1, cols=2)
                    header_table.autofit = False
                    header_table.columns[0].width = Inches(1.2)  # Coluna do logo 40% menor
                    header_table.columns[1].width = Inches(5.8)  # Coluna do texto maior

                    # Logo
                    logo_cell = header_table.cell(0, 0)
                    logo_paragraph = logo_cell.paragraphs[0]
                    logo_run = logo_paragraph.add_run()
                    logo_run.add_picture('static/images/logo.png', width=Inches(1.2))

                    # Texto da empresa em cinza
                    info_cell = header_table.cell(0, 1)
                    info_paragraph = info_cell.paragraphs[0]
                    info_run = info_paragraph.add_run("BEIJO E MATOS CONSTRUÇÕES E ENGENHARIA LTDA\nJoaquim da Silva Martha, 12-53 - Sala 3 - Altos da Cidade - Bauru/SP\nguilhermebeijo@bencato.com.br - CNPJ: 26.149.105/0001-09 - www.bencato.com.br")
                    info_run.font.color.rgb = RGBColor(128, 128, 128)  # Cor cinza
                    info_run.font.size = Pt(11)

                    doc.add_paragraph()  # Espaço após o cabeçalho

                    # Aplica o modelo editado
                    texto_formatado = modelo_texto.format(
                        cliente_nome=cliente_nome,
                        valor=valor,
                        valor_extenso=valor_por_extenso(valor),
                        numero_documento=numero_documento,
                        data=data
                    )

                    # Divide o texto em linhas para o documento
                    linhas_recibo = texto_formatado.split('\n')

                    # Adiciona as linhas do recibo ao documento
                    for linha in linhas_recibo:
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        p.add_run(linha)

                    doc.add_paragraph()  # Espaço antes da data

                    # Data em português
                    data_atual = datetime.now()
                    mes_pt = traduzir_mes(data_atual.strftime('%B'))
                    data_formatada = f"Bauru, {data_atual.day} de {mes_pt} de {data_atual.year}"

                    # Data com espaço adicional
                    data_paragraph = doc.add_paragraph()
                    data_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    data_paragraph.add_run(data_formatada)

                    # Adiciona um enter após a data
                    doc.add_paragraph()
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

                    preview_content.append({
                        'nome': cliente_nome,
                        'conteudo': linhas_recibo
                    })

                if not preview_content:
                    return jsonify({'error': 'Nenhum cliente válido encontrado'}), 400

                return jsonify({
                    'preview': preview_content,
                    'status': 'success'
                })

            except Exception as e:
                print(f"Erro detalhado: {str(e)}")
                import traceback
                print(f"Traceback: {traceback.format_exc()}")
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


