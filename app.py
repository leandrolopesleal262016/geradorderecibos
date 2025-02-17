from flask import Flask, render_template, request, jsonify, send_file
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import pandas as pd
from datetime import datetime
import os
import zipfile
import locale

app = Flask(__name__)
fornecedores_df = None
documentos_gerados = []  # Declaração global

def valor_por_extenso(valor):
    if valor == 0:
        return "zero reais"
        
    unidades = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"]
    dezenas = ["", "dez", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"]
    dez_a_dezenove = ["dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"]
    centenas = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"]
    
    reais = int(valor)
    centavos = int((valor * 100) % 100)
    
    if reais == 0 and centavos == 0:
        return "zero reais"
        
    texto = []
    
    # Processa reais
    if reais > 0:
        if reais == 1:
            texto.append("um real")
        else:
            # Processa milhares
            milhares = reais // 1000
            if milhares > 0:
                if milhares == 1:
                    texto.append("mil")
                else:
                    texto.append(valor_por_extenso(milhares) + " mil")
            
            # Processa centenas, dezenas e unidades
            resto = reais % 1000
            if resto > 0:
                if resto < 100:
                    if milhares > 0:
                        texto.append("e")
                if resto == 100:
                    texto.append("cem")
                else:
                    c = resto // 100
                    d = (resto % 100) // 10
                    u = resto % 10
                    
                    if c > 0:
                        texto.append(centenas[c])
                    if d > 0:
                        if c > 0:
                            texto.append("e")
                        if d == 1 and u > 0:
                            texto.append(dez_a_dezenove[u])
                        else:
                            texto.append(dezenas[d])
                            if u > 0:
                                texto.append("e")
                                texto.append(unidades[u])
                    elif u > 0:
                        if c > 0:
                            texto.append("e")
                        texto.append(unidades[u])
            
            texto.append("reais")
    
    # Processa centavos
    if centavos > 0:
        if reais > 0:
            texto.append("e")
        if centavos == 1:
            texto.append("um centavo")
        else:
            d = centavos // 10
            u = centavos % 10
            
            if d == 1:
                if u > 0:
                    texto.append(dez_a_dezenove[u])
                else:
                    texto.append(dezenas[d])
            else:
                if d > 0:
                    texto.append(dezenas[d])
                    if u > 0:
                        texto.append("e")
                        texto.append(unidades[u])
                else:
                    texto.append(unidades[u])
            texto.append("centavos")
    
    return " ".join(texto)

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
        clientes_selecionados = request.json.get('clientes', [])
        modelo_selecionado = request.json.get('modelo', '1')
        documentos_gerados = []
        preview_content = []

        for cliente_nome in clientes_selecionados:
            cliente_data = fornecedores_df[fornecedores_df['Razão social'] == cliente_nome].iloc[0].to_dict()
            
            # Converte valores numéricos para string
            valor = str(float(cliente_data.get('amount', '0.00')))
            
            # Cabeçalho comum para ambos os modelos
            header = [
                "Beijo e Matos Construções e Engenharia LTDA",
                "Joaquim da Silva Martha, 12-53 - Sala 3 - Altos da Cidade - Bauru/SP",
                "guilhermebeijo@bencato.com.br - CNPJ: 26.149.105/0001-09 - www.bencato.com.br",
                "Relatório de recibos"
            ]

            if modelo_selecionado == '1':
                # Modelo 1 - Emitente (como na primeira imagem)
                content = [
                    *header,
                    f"RECIBO Nº {str(cliente_data.get('document_number', '0001'))} - parcela única",
                    f"VALOR: R$ {valor}",
                    f"ADMINISTRATIVO - {cliente_nome}",
                    f"Recebi(emos) a quantia de R$ {valor} ({valor_por_extenso(float(valor))})",
                    f"na forma de pagamento {str(cliente_data.get('payment_method', 'Conciliação'))}",
                    f"correspondente a {str(cliente_data.get('description', 'VALE ALIMENTAÇÃO'))} (documento número {str(cliente_data.get('document_number', '0001'))} parcela única)",
                    "e para maior clareza firmo(amos) o presente.",
                    f"Bauru, {datetime.now().strftime('%d de %B de %Y')}",
                    cliente_nome.upper(),
                    str(cliente_data.get('CPF/CNPJ', ''))
                ]
            else:
                # Modelo 2 - Destinatário (como na segunda imagem)
                content = [
                    *header,
                    f"RECIBO Nº {str(cliente_data.get('document_number', '0001'))} - parcela única",
                    f"VALOR: R$ {valor}",
                    f"{cliente_nome}",
                    f"Recebemos de BENCATO CONSTRUCOES LTDA a quantia de R$ {valor} ({valor_por_extenso(float(valor))})",
                    f"na forma de pagamento {str(cliente_data.get('payment_method', 'Conciliação'))}, correspondente a",
                    f"PARCELA ADM OBRA - CONFORME CONTRATO (documento número {str(cliente_data.get('document_number', '0001'))} parcela única)",
                    "e para maior clareza firmo(amos) o presente.",
                    f"Bauru, {datetime.now().strftime('%d de %B de %Y')}",
                    "BEIJO E MATOS CONSTRUCOES E ENGENHARIA LTDA",
                    "26.149.105/0001-09"
                ]
            
            # Criar documento Word
            doc = Document()
            
            # Configurar margens do documento
            section = doc.sections[0]
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)
            section.top_margin = Cm(2.5)
            section.bottom_margin = Cm(2.5)

            # Criar tabela para cabeçalho
            table = doc.add_table(rows=1, cols=2)
            table.autofit = False
            table.allow_autofit = False
            
            # Configurar larguras das colunas
            table.columns[0].width = Cm(6)  # Largura maior para o logo
            table.columns[1].width = Cm(12)  # Ajuste da largura do texto
            
            # Células da tabela
            logo_cell = table.cell(0, 0)
            text_cell = table.cell(0, 1)
            
            # Adicionar logo
            logo_path = os.path.join('static', 'images', 'logo.png')
            if os.path.exists(logo_path):
                logo_paragraph = logo_cell.paragraphs[0]
                logo_run = logo_paragraph.add_run()
                logo_run.add_picture(logo_path, width=Cm(6))  # Dobro do tamanho
            
            # Adicionar texto do cabeçalho
            text_paragraph = text_cell.paragraphs[0]
            text_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            # Nome da empresa
            empresa_run = text_paragraph.add_run("Beijo e Matos Construções e Engenharia LTDA\n")
            empresa_run.font.size = Pt(10)
            empresa_run.font.bold = True
            
            # Endereço
            endereco_run = text_paragraph.add_run("Joaquim da Silva Martha, 12-53 - Sala 3 - Altos da Cidade - Bauru/SP\n")
            endereco_run.font.size = Pt(10)
            
            # Email e site
            contato_run = text_paragraph.add_run("guilhermebeijo@bencato.com.br - CNPJ: 26.149.105/0001-09 - www.bencato.com.br")
            contato_run.font.size = Pt(10)
            
            # Adicionar "Relatório de recibos" em azul
            doc.add_paragraph()  # Espaço após o cabeçalho
            relatorio_paragraph = doc.add_paragraph()
            relatorio_run = relatorio_paragraph.add_run("Relatório de recibos")
            relatorio_run.font.color.rgb = RGBColor(0, 70, 127)  # Azul corporativo
            relatorio_run.font.size = Pt(11)
            
            doc.add_paragraph()  # Espaço após o título
            
            # Criar tabela para número do recibo e valor
            header_table = doc.add_table(rows=1, cols=2)
            header_table.autofit = False
            header_table.allow_autofit = False
            
            # Configurar larguras das colunas do cabeçalho
            header_table.columns[0].width = Cm(10)
            header_table.columns[1].width = Cm(8)
            
            # Adicionar número do recibo e valor
            recibo_cell = header_table.cell(0, 0)
            valor_cell = header_table.cell(0, 1)
            
            recibo_paragraph = recibo_cell.paragraphs[0]
            recibo_run = recibo_paragraph.add_run(f"RECIBO Nº {str(cliente_data.get('document_number', '0001'))} - parcela única")
            recibo_run.bold = True
            recibo_run.font.size = Pt(12)
            
            valor_paragraph = valor_cell.paragraphs[0]
            valor_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            valor_run = valor_paragraph.add_run(f"VALOR: R$ {valor}")
            valor_run.bold = True
            valor_run.font.size = Pt(12)
            
            # Adicionar nome do cliente
            if modelo_selecionado == '1':
                cliente_paragraph = doc.add_paragraph()
                cliente_run = cliente_paragraph.add_run(f"ADMINISTRATIVO - {cliente_nome}")
                cliente_run.font.size = Pt(11)
            
            # Adicionar texto justificado em uma linha
            doc.add_paragraph()  # Espaço antes do texto
            texto_principal = doc.add_paragraph()
            texto_principal.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
            if modelo_selecionado == '1':
                texto = f"Recebi(emos) a quantia de R$ {valor} ({valor_por_extenso(float(valor))}) na forma de pagamento {str(cliente_data.get('payment_method', 'Conciliação'))}, correspondente a {str(cliente_data.get('description', 'VALE ALIMENTAÇÃO'))} (documento número {str(cliente_data.get('document_number', '0001'))} parcela única) e para maior clareza firmo(amos) o presente."
            else:
                texto = f"Recebemos de BENCATO CONSTRUCOES LTDA a quantia de R$ {valor} ({valor_por_extenso(float(valor))}) na forma de pagamento {str(cliente_data.get('payment_method', 'Conciliação'))}, correspondente a PARCELA ADM OBRA - CONFORME CONTRATO (documento número {str(cliente_data.get('document_number', '0001'))} parcela única) e para maior clareza firmo(amos) o presente."
            
            texto_run = texto_principal.add_run(texto)
            texto_run.font.size = Pt(11)
            
            # Adicionar data
            data_atual = datetime.now().strftime('%d de %B de %Y')
            data_paragraph = doc.add_paragraph()
            data_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            data_run = data_paragraph.add_run(f"Bauru, {data_atual}")
            data_run.font.size = Pt(11)
            
            # Adicionar espaço antes da linha de assinatura
            doc.add_paragraph()
            doc.add_paragraph()
            
            # Adicionar linha para assinatura
            linha_assinatura = doc.add_paragraph()
            linha_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
            linha_run = linha_assinatura.add_run("_" * 50)
            
            # Adicionar nome abaixo da linha
            nome_assinatura = doc.add_paragraph()
            nome_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if modelo_selecionado == '1':
                nome_run = nome_assinatura.add_run(cliente_nome.upper())
                cpf_cnpj = doc.add_paragraph()
                cpf_cnpj.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cpf_cnpj.add_run(str(cliente_data.get('CPF/CNPJ', '')))
            else:
                nome_run = nome_assinatura.add_run("BEIJO E MATOS CONSTRUCOES E ENGENHARIA LTDA")
                cpf_cnpj = doc.add_paragraph()
                cpf_cnpj.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cpf_cnpj.add_run("26.149.105/0001-09")
                
                # Adicionar espaçamento após o cabeçalho
                if i == 3:
                    paragraph.space_after = Pt(20)

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