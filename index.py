import os
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
import fdb
import pandas as pd
from openpyxl import load_workbook

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui' 

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Obter o caminho do banco de dados do formulário
        db_path = request.form['db_path']
        
        if not os.path.exists(db_path):
            flash('Caminho inválido. Por favor, insira um caminho válido para o arquivo do banco de dados.')
            return redirect(url_for('index'))
        
        try:
            # Conectar ao banco de dados
            con = fdb.connect(dsn=f'localhost:{db_path}', user='SYSDBA', password='masterkey')
            
            # Executar a consulta
            query = """
            SELECT 
                '' AS "Código",  -- Deixando em branco
                p.DESCRICAO AS "Descrição",  -- Descrição do produto
                p.REFFABRICANTE AS "Referência",  -- Referência do fabricante
                '' AS "Cód. Auxiliar",  -- Deixando em branco
                f.NOME AS "Fornecedor",  -- Nome do fornecedor
                '' AS "Fornecedor exclusivo",  -- Deixando em branco
                '' AS "Comprador",  -- Deixando em branco
                e.EMPRESA AS "Empresa",  -- Nome da empresa da tabela EMPRESA
                'SIM' AS "Contabiliza saldo em estoque",  -- Contabiliza saldo em estoque
                '' AS "Indisponível para venda",  -- Deixando em branco
                g.DESCRICAO AS "Setor",  -- Descrição do grupo como Setor
                sg.DESCRICAO AS "Linha",  -- Descrição do subgrupo como Linha
                '' AS "Marca",  -- Deixando em branco
                '' AS "Coleção",  -- Deixando em branco
                '' AS "Espessura",  -- Deixando em branco
                fam.DESCRICAO AS "Classificação",  -- Descrição da família como Classificação
                gr.DESCRICAO AS "Tamanho",  -- Descrição da grade como Tamanho
                '' AS "Cores",  -- Deixando em branco
                p.UNIDADESAIDA AS "Unidade de venda",  -- Unidade de venda
                p.FATORCONVERSAO AS "Múltiplo de venda",  -- Múltiplo de venda
                'R$' AS "Moeda",  -- Moeda
                cp.PRECOCUSTO AS "Custo com ICMS (R$)",  -- Custo com ICMS
                cp.DESCONTOPERC AS "Desconto (%)",  -- Desconto percentual
                '' AS "Acréscimo (%)",  -- Deixando em branco
                cp.ALIQIPI AS "IPI (%)",  -- Alíquota de IPI
                cp.FRETEVLR AS "Frete (R$)",  -- Valor do frete
                '' AS "Despesas acessórias (R$)",  -- Deixando em branco
                '' AS "Substituição tributária (R$)",  -- Deixando em branco
                '' AS "Diferencial ICMS (R$)",  -- Deixando em branco
                '' AS "Mark-up (%)",  -- Deixando em branco
                p.PRECO AS "Preço de venda R$",  -- Preço de venda
                'S' AS "Permite desconto",  -- Permite desconto
                '' AS "Comissão %",  -- Deixando em branco
                '' AS "Configuração tributária",  -- Deixando em branco
                c.CODIGONCM AS "NCM",  -- Código NCM
                p.CODCEST AS "CEST",  -- Código CEST
                '' AS "Produto supérfluo",  -- Deixando em branco
                '' AS "Tipo de item",  -- Deixando em branco
                '' AS "Origem da mercadoria",  -- Deixando em branco
                '' AS "Regime de Incidência PIS e COFINS",  -- Deixando em branco
                '' AS "Produto é brinde",  -- Deixando em branco
                '' AS "Produto de catálogo",  -- Deixando em branco
                '' AS "Descrição de catálogo",  -- Deixando em branco
                '' AS "Disponível na loja virtual",  -- Deixando em branco
                '' AS "Exige controle",  -- Deixando em branco
                '' AS "Tipo de controle",  -- Deixando em branco
                '' AS "Tamanho controle",  -- Deixando em branco
                p.PESO AS "Peso bruto (kg)",  -- Peso bruto do produto
                '' AS "Peso líquido (kg)",  -- Deixando em branco
                '' AS "Descrição complementar?",  -- Deixando em branco
                '' AS "Altura (frete)",  -- Deixando em branco
                '' AS "Largura (frete)",  -- Deixando em branco
                '' AS "Comprimento (frete)",  -- Deixando em branco
                '' AS "Altura",  -- Deixando em branco
                '' AS "Largura",  -- Deixando em branco
                '' AS "Comprimento",  -- Deixando em branco
                '' AS "Importado por balança",  -- Deixando em branco
                p.ESTMINIMO AS "Quantidade mínima",  -- Quantidade mínima
                p.ESTMAXIMO AS "Quantidade máxima",  -- Quantidade máxima
                '' AS "Quantidade compra",  -- Deixando em branco
                p.OBSERVACAO AS "Observação",  -- Observações
                p.REFFABRICANTE AS "Código de barras",  -- Código de barras
                '' AS "Características"  -- Deixando em branco
            FROM 
                PRODUTO p
            JOIN 
                FORNECE f ON p.PRINCIPALFORNEC = f.CODFORNEC
            JOIN 
                CLASFISC c ON p.CODCLASFIS = c.CODCLASFIS
            JOIN 
                GRUPROD g ON p.CODGRUPO = g.CODGRUPO
            JOIN 
                SUBGRUP sg ON p.CODSUBGRUPO = sg.CODSUBGRUPO
            JOIN 
                COMPPROD cp ON p.CODPROD = cp.CODPROD
            JOIN 
                GRADE gr ON cp.CODGRADE = gr.CODGRADE
            JOIN 
                FAMILIA fam ON p.CODFAMILIA = fam.CODIGO
            JOIN 
                EMPRESA e ON 1=1  -- Join com a tabela EMPRESA, assumindo que só há uma linha
            WHERE 
                p.ATIVO = 'S';
            """
            df = pd.read_sql(query, con)
            
            # Exportar para Excel
            file_path = 'Produtos.xlsx'
            df.to_excel(file_path, index=False, engine='openpyxl')

            # Ajustar a largura das colunas
            wb = load_workbook(file_path)
            ws = wb.active

            for column_cells in ws.columns:
                max_length = 0
                column = column_cells[0].column_letter 
                for cell in column_cells:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column].width = adjusted_width

            wb.save(file_path)
            
            # Fechar a conexão
            con.close()
            
            # Retornar o arquivo para download
            return send_file(file_path, as_attachment=True)
        except Exception as e:
            flash(f'Erro ao processar: {e}')
            return redirect(url_for('index'))
        
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
