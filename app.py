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
from models import db, ReceiptSequence, ReciboGerado, ModeloRecibo
app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///recibos.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db.init_app(app)

data_atual = datetime.now().strftime('%d/%m/%Y')
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
        'October': 'outubro',        'November': 'novembro',
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
    # Converte para float caso seja string
    if isinstance(valor, str):
        valor = float(valor.replace('.', '').replace(',', '.'))
    
    # Separa parte inteira e decimal
    parte_inteira = int(valor)
    parte_decimal = int(round((valor - parte_inteira) * 100))

    # Arrays com palavras
    unidades = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove']
    dezenas = ['', 'dez', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa']
    dezenas_especiais = ['dez', 'onze', 'doze', 'treze', 'quatorze', 'quinze', 'dezesseis', 'dezessete', 'dezoito', 'dezenove']
    centenas = ['', 'cento', 'duzentos', 'trezentos', 'quatrocentos', 'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos']

    extenso = []

    # Trata valor zero
    if parte_inteira == 0 and parte_decimal == 0:
        return 'zero reais'

    # Processa milhares
    if parte_inteira >= 1000:
        milhares = parte_inteira // 1000
        if milhares == 1:
            extenso.append('um mil')
        else:
            if milhares >= 100:
                centena_milhar = milhares // 100
                extenso.append(centenas[centena_milhar])
                milhares = milhares % 100
                if milhares > 0:
                    extenso.append('e')
            
            if milhares >= 10:
                if milhares >= 10 and milhares <= 19:
                    extenso.append(dezenas_especiais[milhares - 10])
                else:
                    dezena_milhar = milhares // 10
                    extenso.append(dezenas[dezena_milhar])
                    milhares = milhares % 10
                    if milhares > 0:
                        extenso.append('e')
                        extenso.append(unidades[milhares])
            elif milhares > 0:
                extenso.append(unidades[milhares])
            
            extenso.append('mil')

        parte_inteira = parte_inteira % 1000
        if parte_inteira > 0:
            if parte_inteira < 100:
                extenso.append('e')
            else:
                extenso.append('')

    # Processa centenas
    if parte_inteira >= 100:
        centena = parte_inteira // 100
        if centena == 1 and parte_inteira % 100 == 0:
            extenso.append('cem')
        else:
            extenso.append(centenas[centena])
        parte_inteira = parte_inteira % 100
        if parte_inteira > 0:
            extenso.append('e')

    # Processa dezenas e unidades
    if parte_inteira >= 10 and parte_inteira <= 19:
        extenso.append(dezenas_especiais[parte_inteira - 10])
    else:
        if parte_inteira >= 10:
            extenso.append(dezenas[parte_inteira // 10])
            parte_inteira = parte_inteira % 10
            if parte_inteira > 0:
                extenso.append('e')
        if parte_inteira > 0:
            extenso.append(unidades[parte_inteira])

    extenso.append('reais')

    # Processa centavos
    if parte_decimal > 0:
        extenso.append('e')
        if parte_decimal >= 10 and parte_decimal <= 19:
            extenso.append(dezenas_especiais[parte_decimal - 10])
        else:
            dezena = parte_decimal // 10
            unidade = parte_decimal % 10
            if dezena > 0:
                extenso.append(dezenas[dezena])
                if unidade > 0:
                    extenso.append('e')
                    extenso.append(unidades[unidade])
            elif unidade > 0:
                extenso.append(unidades[unidade])
        extenso.append('centavos')

    return ' '.join(extenso)
def numero_para_extenso(numero):
    unidades = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove']
    dezenas = ['', 'dez', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa']
    
    extenso = []
    
    # Verifica se o número está dentro do intervalo válido
    if numero < 0 or numero > 99:
        return ['número fora do intervalo válido']
    
    if numero >= 10:
        dezena = int(numero / 10)
        if dezena < len(dezenas):  # Verifica se o índice é válido
            extenso.append(dezenas[dezena])
            numero = numero % 10
            if numero > 0:
                extenso.append('e')
    
    if numero > 0 and numero < len(unidades):
        extenso.append(unidades[numero])
    
    return extenso

def get_next_receipt_number():
    with app.app_context():
        seq = ReceiptSequence.query.first()
        if not seq:
            seq = ReceiptSequence(last_number=0)
            db.session.add(seq)
        
        seq.last_number += 1
        db.session.commit()
        return f"{seq.last_number:05d}"
    
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

        # Obter e formatar a data atual
        data_atual = datetime.now()
        data_formatada = data_atual.strftime('%d/%m/%Y')

        # Converter valor para formato correto
        valor_str = dados.get('valor', '0,00')
        valor_limpo = valor_str.replace('.', '').replace(',', '.')
        valor_float = float(valor_limpo)
    
        # Formata o valor para exibição
        valor_formatado = f"{valor_float:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

        documentos_gerados = []
        preview_content = []

        # Recupera e atualiza o modelo do texto
        modelo_texto = request.json.get('modeloConteudo')
        modelo_texto = modelo_texto.replace('{numero_documento}', '{numero_recibo}')

        for cliente_nome in clientes_selecionados:
            numero_recibo = get_next_receipt_number()

            cliente_filtrado = fornecedores_df[fornecedores_df['Razão social'] == cliente_nome]
            if cliente_filtrado.empty:
                continue

            cliente_data = cliente_filtrado.iloc[0].to_dict()
            documento_cliente = cliente_data.get('CPF/CNPJ', '')

            texto_formatado = modelo_texto.format(
                cliente_nome=cliente_nome,
                valor=valor_formatado,
                valor_extenso=valor_por_extenso(valor_float),
                numero_recibo=numero_recibo,
                data=data_formatada,
                documento_cliente=documento_cliente
            )
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

        # Salvar no banco
        doc_buffer = io.BytesIO()
        doc.save(doc_buffer)

        recibo = ReciboGerado(
            numero_recibo=numero_recibo,
            modelo_id=int(modelo_selecionado),  # Convertendo para inteiro
            cliente_nome=cliente_nome,
            valor=valor_float,  # Usando o valor convertido
            documento_blob=doc_buffer.getvalue()
        )
        db.session.add(recibo)
        db.session.commit()

        documentos_gerados.append((cliente_nome, doc_buffer.getvalue()))
        preview_content.append({
            'nome': cliente_nome,
            'conteudo': linhas_recibo
        })

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
        
        # Cria o arquivo ZIP com modo de escrita binária e compressão
        with zipfile.ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED) as zip_file:
            for nome, doc_content in documentos_gerados:
                # Cria um buffer temporário para cada documento
                doc_buffer = io.BytesIO(doc_content)
                doc_buffer.seek(0)
                
                # Nome do arquivo limpo (remove caracteres especiais)
                nome_arquivo = "".join(c for c in nome if c.isalnum() or c in (' ', '-', '_'))
                
                # Adiciona o documento ao ZIP usando o buffer do documento
                zip_file.writestr(f"recibo_{nome_arquivo}.docx", doc_buffer.getvalue())
                
                # Fecha o buffer temporário
                doc_buffer.close()

        # Prepara o buffer para leitura
        zip_buffer.seek(0)
        
        # Obtém o tamanho do buffer
        zip_size = zip_buffer.getbuffer().nbytes
                
        # Envia o arquivo com tamanho específico
        response = send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='recibos.zip'
        )
        
        # Adiciona headers específicos para download
        response.headers["Content-Length"] = zip_size
        response.headers["Content-Type"] = "application/zip"
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        
        return response

    except Exception as e:
        print(f"Erro no download: {str(e)}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500
    
@app.route('/historico_recibos')
def historico_recibos():
    recibos = ReciboGerado.query.order_by(ReciboGerado.data_geracao.desc()).all()
    return render_template('historico.html', recibos=recibos)

@app.route('/consulta_recibos')
def consulta_recibos():
    recibos = ReciboGerado.query.order_by(ReciboGerado.data_geracao.desc()).all()
    return render_template('consulta_recibos.html', recibos=recibos)

@app.route('/download_recibo/<int:recibo_id>')
def download_recibo(recibo_id):
    recibo = ReciboGerado.query.get_or_404(recibo_id)
    
    return send_file(
        io.BytesIO(recibo.documento_blob),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name=f'recibo_{recibo.numero_recibo}.docx'
    )

@app.route('/visualizar_recibo/<int:recibo_id>')
def visualizar_recibo(recibo_id):
    recibo = ReciboGerado.query.get_or_404(recibo_id)
    return render_template('visualizar_recibo.html', recibo=recibo)


@app.route('/modelos', methods=['GET'])
def listar_modelos():
    modelos = ModeloRecibo.query.all()
    return jsonify([{
        'id': m.id,
        'nome': m.nome,
        'conteudo': m.conteudo
    } for m in modelos])

@app.route('/modelos', methods=['POST'])
def salvar_modelo():
    dados = request.json
    modelo = ModeloRecibo(
        nome=dados['nome'],
        conteudo=dados['conteudo']
    )
    db.session.add(modelo)
    db.session.commit()
    return jsonify({'id': modelo.id, 'message': 'Modelo salvo com sucesso'})

@app.route('/modelos/<int:modelo_id>', methods=['PUT'])
def atualizar_modelo(modelo_id):
    dados = request.json
    modelo = ModeloRecibo.query.get_or_404(modelo_id)
    modelo.nome = dados['nome']
    modelo.conteudo = dados['conteudo']
    db.session.commit()
    return jsonify({'message': 'Modelo atualizado com sucesso'})

def init_db():    
    with app.app_context():
        # Criar tabelas se não existirem
        db.create_all()        
        # Criar modelos padrão se não existirem
        if not ModeloRecibo.query.first():
            modelos_padrao = [
                {
                    'nome': 'Modelo Emitente',
                    'conteudo': 'RECIBO Nº {numero_recibo}...'
                },
                {
                    'nome': 'Modelo Destinatário',
                    'conteudo': 'RECIBO Nº {numero_recibo}...'
                },
                {
                    'nome': 'Modelo Personalizado',
                    'conteudo': 'RECIBO Nº {numero_recibo}...'
                }
            ]
            for modelo in modelos_padrao:
                db.session.add(ModeloRecibo(**modelo))
            db.session.commit()

if __name__ == '__main__':
    with app.app_context():
        init_db()
    app.run(debug=True)

