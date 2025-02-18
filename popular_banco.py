import pandas as pd
from models import db, Cliente
from app import app
import re

def populate_clients_from_csv():
    with app.app_context():
        # Primeiro, vamos limpar a tabela de clientes
        Cliente.query.delete()
        
        # Lê o arquivo CSV
        fornecedores_df = pd.read_csv('Fornecedores.csv', sep=';', encoding='windows-1252', skiprows=4)
        
        # Processa e insere no banco de dados
        for _, row in fornecedores_df.iterrows():
            # Pega o CPF/CNPJ e remove caracteres não numéricos
            cpf_cnpj = re.sub(r'\D', '', str(row['CPF/CNPJ']))
            
            # Verifica se tem razão social e CPF/CNPJ válidos
            if pd.isna(row['Razão social']) or not cpf_cnpj:
                continue
                
            # Determina o tipo baseado no tamanho do documento
            tipo = 'empresa' if len(cpf_cnpj) > 11 else 'pessoa'
            
            try:
                # Verifica se já existe um cliente com este CPF/CNPJ
                cliente_existente = Cliente.query.filter_by(cpf_cnpj=cpf_cnpj).first()
                
                if not cliente_existente:
                    cliente = Cliente(
                        razao_social=row['Razão social'].strip(),
                        cpf_cnpj=cpf_cnpj,
                        tipo=tipo
                    )
                    db.session.add(cliente)
                    print(f"Cliente adicionado: {cliente.razao_social}")
                    
            except Exception as e:
                print(f"Erro ao processar cliente {row['Razão social']}: {str(e)}")
                continue
        
        try:
            db.session.commit()
            print("Dados importados com sucesso!")
        except Exception as e:
            db.session.rollback()
            print(f"Erro ao salvar no banco de dados: {str(e)}")

if __name__ == '__main__':
    with app.app_context():
        db.create_all()  # Garante que as tabelas existam
        populate_clients_from_csv()
