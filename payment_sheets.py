import pandas as pd

file_path = '.../Vendas.xlsx' #Add your path to the excel file
payment_sheet_name = 'Pagamentos'
file_path_resultado = '.../resultado.xlsx' #Add the path where you created the file using the sales_sheets.py script

pagamento_df = pd.read_excel(file_path, sheet_name=payment_sheet_name)
resultado_df = pd.read_excel(file_path_resultado)

result_dict = resultado_df.set_index('Nome do Vendedor')['Comissao_Paga'].to_dict()

differences_data = []

for index, row in pagamento_df.iterrows():
    vendedor = row['Nome do Vendedor']
    comissao_recebida = row['Comissão']
    comissao_esperada = result_dict.get(vendedor, 0)  # If salesman not foud, assumes 0
    difference = comissao_recebida - comissao_esperada
    if difference != 0:
        differences_data.append({
            'Nome do Vendedor': vendedor,
            'Comissao_Recebida': comissao_recebida,
            'Comissao_Esperada': comissao_esperada,
            'Diferença': difference
        })

differences_df = pd.DataFrame(differences_data)

differences_file_path = '.../diferencas_comissao.xlsx' #Add the path where you want to save the excel file
differences_df.to_excel(differences_file_path, index=False)

print("Tabela de diferenças de comissões salva em", differences_file_path)
