import pandas as pd

class Getnet:
    def iniciar(planilha1, planilha2, operadora, unidade, data):
        df_notas = pd.read_excel(planilha1)
        df_vendas = pd.read_excel(planilha2, skiprows=7)

        df_notas1 = df_notas[['Autorização','No. Titulo', 'NF Eletr','Nome Cliente', 'Parcela', 'Vlr.Titulo']]

        df_notas1.rename(columns={
            'Autorização': 'AUT',
            'No. Titulo': 'RPS',
            'NF Eletr': 'NF',
            'Nome Cliente': 'BANDEIRA1',
            'Parcela':'PARCELA',
            'Vlr.Titulo':'VALOR' 
        },inplace=True)

        df_notas1['BANDEIRA'] = df_notas1['BANDEIRA1'].str.extract(r' - ([^ (]+)')

        df_notas1.drop(columns='BANDEIRA1', inplace=True)

        exel_notas = df_notas1.groupby('AUT', as_index=False).agg({
            'VALOR': 'sum',
            'PARCELA': 'count',
            'BANDEIRA': 'first',
            'RPS': 'first',
            'NF': 'first'
        }, inplace=True)

        df_vendas = df_vendas[['Número da Autorização','Valor Bruto','Total de Parcelas','Cartões', 'Valor Líquido']]

        df_vendas.rename(columns={
            'Número da Autorização':'AUT',
            'Valor Bruto':'VALOR BRUTO',
            'Total de Parcelas':'PARCELAS',
            'Valor Líquido':'VALOR LIQUIDO',
            'Cartões': 'BANDEIRA1',
        },inplace=True)

        df_vendas['BANDEIRA'] = df_vendas['BANDEIRA1'].replace({
            'VISA CRÉDITO': 'VISA',
            'VISA DÉBITO': 'VISA',
            'MASTERCARD CRÉDITO': 'MASTER',
            'MASTERCARD DÉBITO': 'MASTER',
            'AMEX CRÉDITO': 'AMEX',
            'AMEX DÉBITO': 'AMEX',
            'ELO CRÉDITO': 'ELO',
            'ELO DÉBITO': 'ELO',
            'HIPERCARD CRÉDITO': 'HIPER',
            'HIPERCARD DÉBITO': 'HIPER'
        })


        xlsx_vendas = df_vendas.drop(columns='BANDEIRA1')

        mesclagem = pd.merge(exel_notas, xlsx_vendas, on = 'AUT')

        mesclagem['VALOR'] = mesclagem['VALOR'].astype(float)

        mesclagem['VALOR'] = mesclagem['VALOR'].round(2)


        def verificar_nota(linha):
            if linha['BANDEIRA_x'] != linha['BANDEIRA_y']:
                return 'Bandeira Divergente'
            elif linha['PARCELA'] != linha['PARCELAS']:
                return 'Parcela Divergente'
            elif abs(linha['VALOR'] - linha['VALOR LIQUIDO']) > 0.10:
                return 'Valor Divergente'
            elif pd.isna(linha['NF']):
                return 'Sem NF'
            else:
                return 'OK'
            
        mesclagem['STATUS'] = mesclagem.apply(verificar_nota, axis=1)

        resultado = pd.merge(exel_notas, mesclagem)

        nome = f'{operadora}_{unidade}_{data}.xlsx'

        resultado.to_excel(nome, index=False)

        return resultado