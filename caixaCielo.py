import pandas as pd

class Cielo:
    def iniciar(planilha1, planilha2, operadora, unidade, data):

        df_notas = pd.read_excel(planilha1)
        df_vendas = pd.read_excel(planilha2, skiprows=9)


        df_notas1 = df_notas.copy()


        df_notas1[['Num. NSU','No. Titulo', 'NF Eletr','Nome Cliente', 'Parcela', 'Vlr.Titulo']]

        df_notas1.rename(columns={
            'Num. NSU':'NSU',
            'No. Titulo': 'RPS',
            'NF Eletr': 'NF',
            'Nome Cliente': 'BANDEIRA1',
            'Parcela':'PARCELA',
            'Vlr.Titulo':'VALOR' 
        },inplace=True)

        df_notas1['BANDEIRA'] = df_notas1['BANDEIRA1'].str.extract(r' - ([^ (]+)')

        df_notas1.drop(columns='BANDEIRA1', inplace=True)

        excel_notas = df_notas1.groupby('NSU', as_index=False).agg({
            'VALOR': 'sum',
            'PARCELA': 'count',
            'BANDEIRA': 'first',
            'RPS': 'first',
            'NF': 'first'
        }, inplace=True)

        df_vendas1 = df_vendas.copy()


        df_vendas1 = df_vendas1[['NSU/DOC', 'Valor bruto','Quantidade total de parcelas','Bandeira', 'Valor líquido']]

        df_vendas1.rename(columns={
            'NSU/DOC':'NSU', 
            'Valor bruto':'VALOR BRUTO',
            'Quantidade total de parcelas': 'PARCELAS',
            'Bandeira': 'BANDEIRA1',
            'Valor líquido':'VALOR LIQUIDO',
        },inplace=True)

        df_vendas1['BANDEIRA'] = df_vendas1['BANDEIRA1'].replace({
            'Visa': 'VISA',
            'Mastercard': 'MASTER',
            'American Express': 'AMEX',
            'Elo': 'ELO',
            'Hipercard': 'HIPER'
        })

        df_vendas1['PARCELAS'] = df_vendas1['PARCELAS'].fillna(1)


        mesclagem = pd.merge(excel_notas, df_vendas1, on = 'NSU')


        mesclagem['VALOR'] = mesclagem['VALOR'].astype(float)

        mesclagem['VALOR'] = mesclagem['VALOR'].round(2)

        def verificar_nota(linha):
            if linha['BANDEIRA_x'] != linha['BANDEIRA_y']:
                return 'Bandeira Divergente'
            elif linha['PARCELA'] != linha['PARCELAS']:
                return 'Parcela Divergente'
            elif abs(linha['VALOR'] - linha['VALOR LIQUIDO']) > 3.59:
                return 'Valor Divergente'
            elif pd.isna(linha['NF']):
                return 'Sem NF'
            else:
                return 'OK'
            
        mesclagem['STATUS'] = mesclagem.apply(verificar_nota, axis=1)


        resultado = pd.merge(excel_notas, mesclagem)

        nome = f'{operadora}_{unidade}_{data}.xlsx'

        resultado.to_excel(nome, index=False)

        return resultado