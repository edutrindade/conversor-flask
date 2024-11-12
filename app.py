import os
from flask import Flask, request, render_template, send_file
import fdb
import pandas as pd
from openpyxl import load_workbook

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Obter o caminho do banco de dados do formulário
        db_path = request.form['db_path']
        
        print(f'Caminho do banco de dados recebido: {db_path}')
        # Conectar ao banco de dados
        con = fdb.connect(dsn=f'localhost:{db_path}', user='SYSDBA', password='masterkey')
        
        # Executar a consulta
        query = """
        SELECT 
            '' AS "Código",
            p.DESCRICAO AS "Descrição",
            p.REFFABRICANTE AS "Código de Barras",
            c.CODIGONCM AS "NCM",
            f.NOME AS "Fornecedor",
            'SIM' AS "Contabiliza saldo em estoque",
            g.DESCRICAO AS "Setor",
            sg.DESCRICAO AS "Linha",
            gr.DESCRICAO AS "Tamanho",
            fam.DESCRICAO AS "Classificação",
            'LIDER MODAS' AS "Empresa",
            'R$' AS "Moeda",
            cp.PRECOCUSTO AS "Custo com ICMS (R$)",
            cp.DESCONTOPERC AS "Desconto (%)",
            cp.ALIQIPI AS "IPI (%)",
            cp.FRETEVLR AS "Frete (R$)",
            p.ESTMINIMO AS "Quantidade mínima",
            p.ESTMAXIMO AS "Quantidade máxima",
            p.UNIDADESAIDA AS "Unidade de Venda",
            p.UNIDADEENT AS "Unidade de Compra",
            p.FATORCONVERSAO AS "Múltiplo de venda",
            p.PESO AS "Peso bruto",
            p.PRECO AS "Preço de venda R$",
            'S' AS "Permite desconto",
            p.OBSERVACAO AS "Observação"
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
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
