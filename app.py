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
from models import db, ReceiptSequence, ReciboGerado, ModeloRecibo, Cliente

# Create Flask app first
app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///recibos.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db.init_app(app)

# Define helper function
def get_document_content(blob):
    doc = Document(io.BytesIO(blob))
    conteudo = []
    for paragrafo in doc.paragraphs:
        if paragrafo.text.strip():
            conteudo.append(paragrafo.text.strip())
    return conteudo

# Register function in Jinja environment
app.jinja_env.globals.update(get_document_content=get_document_content)

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
    try:
        # Busca empresas e pessoas do banco de dados
        empresas = Cliente.query.filter_by(tipo='empresa').order_by(Cliente.razao_social).all()
        pessoas = Cliente.query.filter_by(tipo='pessoa').order_by(Cliente.razao_social).all()
        
        response_data = {
            'empresas': [empresa.razao_social for empresa in empresas],
            'pessoas': [pessoa.razao_social for pessoa in pessoas]
        }
        
        print(f"Total de empresas encontradas: {len(response_data['empresas'])}")
        print(f"Total de pessoas encontradas: {len(response_data['pessoas'])}")
        
        return jsonify(response_data)
        
    except Exception as e:
        print(f"Erro ao buscar clientes: {str(e)}")
        return jsonify({'error': str(e)}), 500
    
@app.route('/generate_receipts_bulk', methods=['POST'])
def generate_receipts_bulk():
    global documentos_gerados
    try:
        dados = request.json
        modelo_selecionado = dados.get('modelo')
        clientes_selecionados = dados.get('clientes', [])

        data_atual = datetime.now()
        data_formatada = data_atual.strftime('%d/%m/%Y')

        valor_str = dados.get('valor', '0,00')
        valor_limpo = valor_str.replace('.', '').replace(',', '.')
        valor_float = float(valor_limpo)
        valor_formatado = f"{valor_float:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

        documentos_gerados = []
        preview_content = []

        modelo_texto = request.json.get('modeloConteudo')
        modelo_texto = modelo_texto.replace('{numero_documento}', '{numero_recibo}')

        for cliente_nome in clientes_selecionados:
            numero_recibo = get_next_receipt_number()

            # Busca cliente no banco de dados
            cliente = Cliente.query.filter_by(razao_social=cliente_nome).first()
            if not cliente:
                continue

            texto_formatado = modelo_texto.format(
                cliente_nome=cliente.razao_social,
                valor=valor_formatado,
                valor_extenso=valor_por_extenso(valor_float),
                numero_recibo=numero_recibo,
                data=data_formatada,
                documento_cliente=cliente.cpf_cnpj
            )

            # Criação do documento Word
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
            # Assinatura atualizada usando dados do modelo Cliente
            assinatura_paragraph = doc.add_paragraph()
            assinatura_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            assinatura_paragraph.add_run("_" * 50 + "\n")
            assinatura_paragraph.add_run(cliente.razao_social.upper() + "\n")
            assinatura_paragraph.add_run(cliente.cpf_cnpj)
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

@app.route('/add_cliente', methods=['POST'])
def add_cliente():
    try:
        data = request.json
        
        # Verifica se já existe cliente com este documento
        cliente_existente = Cliente.query.filter_by(cpf_cnpj=data['cpf_cnpj']).first()
        if cliente_existente:
            return jsonify({'error': 'CPF/CNPJ já cadastrado'}), 400
            
        novo_cliente = Cliente(
            razao_social=data['razao_social'],
            cpf_cnpj=data['cpf_cnpj'],
            tipo=data['tipo']
        )
        
        db.session.add(novo_cliente)
        db.session.commit()
        
        return jsonify({'message': 'Cliente cadastrado com sucesso'}), 201
        
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

@app.route('/consulta_clientes')
def consulta_clientes():
    clientes = Cliente.query.order_by(Cliente.razao_social).all()
    return render_template('consulta_clientes.html', clientes=clientes)

@app.route('/delete_cliente/<int:cliente_id>', methods=['DELETE'])
def delete_cliente(cliente_id):
    try:
        cliente = Cliente.query.get_or_404(cliente_id)
        db.session.delete(cliente)
        db.session.commit()
        return jsonify({'success': True})
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'error': str(e)})
    
@app.route('/atualizar_recibo', methods=['POST'])
def atualizar_recibo():
    print("Iniciando atualização do recibo")
    try:
        dados = request.json
        recibo_id = dados.get('recibo_id')
        conteudo_novo = dados.get('conteudo')
    
        print(f"Recibo ID: {recibo_id}")
        print(f"Novo conteúdo: {conteudo_novo}")
    
        recibo = ReciboGerado.query.get_or_404(recibo_id)
    
        # Cria novo documento
        doc = Document()
    
        # Configurações do documento
        sections = doc.sections
        for section in sections:
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)
    
        # Adiciona cabeçalho
        header_table = doc.add_table(rows=1, cols=2)
        header_table.autofit = False
        header_table.columns[0].width = Inches(1.2)
        header_table.columns[1].width = Inches(5.8)
    
        # Logo
        logo_cell = header_table.cell(0, 0)
        logo_run = logo_cell.paragraphs[0].add_run()
        logo_run.add_picture('static/images/logo.png', width=Inches(1.2))
    
        # Informações da empresa
        info_cell = header_table.cell(0, 1)
        info_run = info_cell.paragraphs[0].add_run()
        info_run.text = "BEIJO E MATOS CONSTRUÇÕES E ENGENHARIA LTDA\n"
        info_run.text += "Joaquim da Silva Martha, 12-53 - Sala 3 - Altos da Cidade - Bauru/SP\n"
        info_run.text += "guilhermebeijo@bencato.com.br - CNPJ: 26.149.105/0001-09 - www.bencato.com.br"
        info_run.font.color.rgb = RGBColor(128, 128, 128)
        info_run.font.size = Pt(11)
    
        # Adiciona conteúdo atualizado
        doc.add_paragraph()  # Espaço após cabeçalho
    
        for linha in conteudo_novo:
            if linha.strip():
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.add_run(linha.strip())
    
        # Salva o documento
        doc_buffer = io.BytesIO()
        doc.save(doc_buffer)
        doc_buffer.seek(0)
    
        # Atualiza no banco
        recibo.documento_blob = doc_buffer.getvalue()
        db.session.commit()
    
        print("Recibo atualizado com sucesso")
    
        return jsonify({
            'status': 'sucesso',
            'mensagem': 'Recibo atualizado com sucesso'
        })
    
    except Exception as e:
        print(f"Erro na atualização: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return jsonify({
            'status': 'erro',
            'mensagem': str(e)
        }), 500
        
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

from flask import abort
from docx import Document
import io
import json

@app.route('/debug_recibo/<int:recibo_id>')
def debug_recibo(recibo_id):
    try:
        print(f"Buscando recibo ID: {recibo_id}")
        recibo = ReciboGerado.query.get(recibo_id)
        
        if not recibo:
            print(f"Recibo {recibo_id} não encontrado")
            return jsonify({'erro': f'Recibo {recibo_id} não encontrado'}), 404
            
        print(f"Recibo encontrado: {recibo.numero_recibo}")
        
        # Lê o documento Word
        doc = Document(io.BytesIO(recibo.documento_blob))
        
        # Extrai conteúdo
        conteudo = []
        for paragrafo in doc.paragraphs:
            texto = paragrafo.text.strip()
            if texto:
                conteudo.append(texto)
                print(f"Parágrafo encontrado: {texto}")
        
        dados = {
            'id': recibo.id,
            'numero_recibo': recibo.numero_recibo,
            'cliente_nome': recibo.cliente_nome,
            'valor': str(recibo.valor),
            'data_geracao': recibo.data_geracao.strftime('%d/%m/%Y %H:%M'),
            'conteudo': conteudo
        }
        
        print("Dados do recibo:", json.dumps(dados, indent=2))
        return jsonify(dados)
        
    except Exception as e:
        print(f"Erro ao debugar recibo: {str(e)}")
        return jsonify({'erro': str(e)}), 500
    
if __name__ == '__main__':
    with app.app_context():
        init_db()
    app.run(debug=True)

