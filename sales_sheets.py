import pandas as pd

file_path = '.../Vendas.xlsx' #Add your path to the excel file

sales_sheet = pd.read_excel(file_path, sheet_name='Vendas')
payment_sheet = pd.read_excel(file_path, sheet_name='Pagamentos')

#Correcting incorrect data in the “Valor da Venda” column from the 17th line onwards
def clean_data(value):
    cleaned_value = ''.join(filter(str.isdigit, str(value)))
    return int(cleaned_value) if clean_data else None

sales_sheet['Valor da Venda'] = sales_sheet['Valor da Venda'].apply(clean_data)
sales_sheet.loc[17:, 'Valor da Venda'] = sales_sheet.loc[17:, 'Valor da Venda'] / 100

sales_sheet['Comissao_Venda'] = sales_sheet['Valor da Venda'] * 0.10
sales_sheet['Comissao_Marketing'] = 0
sales_sheet['Comissao_Gerente'] = 0


# Apply the rules
for index, row in sales_sheet.iterrows():
    if row['Canal de Venda'] == 'Online':
        marketing_comission = row['Comissao_Venda'] * 0.20
        sales_sheet.at[index, 'Comissao_Marketing'] = marketing_comission

total_sales_comission = sales_sheet.groupby('Nome do Vendedor')['Comissao_Venda'].sum()

resultado = pd.DataFrame(columns=['Nome do Vendedor', 'Comissao_Total', 'Comissao_Paga'])

for salesman, commission in total_sales_comission.items():
    manager_commission = 0
    if commission >= 1500:
        manager_commission = commission * 0.10
    
    paid_comission = commission - manager_commission - sales_sheet[sales_sheet['Nome do Vendedor'] == salesman]['Comissao_Marketing'].sum()
    
    novo_registro = pd.DataFrame({
        'Nome do Vendedor': [salesman],
        'Comissao_Total': [commission],
        'Comissao_Paga': [paid_comission]
    })
    
    if not novo_registro.empty:
        resultado = pd.concat([resultado, novo_registro], ignore_index=True)

resultado_file_path = '.../resultado.xlsx' #Add the path where you want to save the excel file
resultado.to_excel(resultado_file_path, index=False)

print("Cálculo de comissões concluído e salvo em", resultado_file_path)
