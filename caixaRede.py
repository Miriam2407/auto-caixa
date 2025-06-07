import pandas as pd

global resultado

class Rede:
    def iniciar(planilha1, planilha2, operadora, unidade, data):

        df_notas = pd.read_excel(planilha1)
        df_vendas = pd.read_csv(planilha2, sep=';')


        df_notas1 = df_notas[['Num. NSU','No. Titulo', 'NF Eletr','Nome Cliente', 'Parcela', 'Vlr.Titulo']]

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


        exel_notas = df_notas1.groupby('NSU', as_index=False).agg({
            'VALOR': 'sum',
            'PARCELA': 'count',
            'BANDEIRA': 'first',
            'RPS': 'first',
            'NF': 'first'
        }, inplace=True)




        df_vendas = df_vendas[['NSU/CV', 'valor da venda atualizado','número de parcelas','bandeira', 'valor líquido']]

        df_vendas.rename(columns={
            'NSU/CV':'NSU', 
            'valor da venda atualizado':'VALOR BRUTO',
            'número de parcelas': 'PARCELAS',
            'bandeira': 'BANDEIRA1',
            'valor líquido':'VALOR LIQUIDO',
        },inplace=True)

        df_vendas['BANDEIRA'] = df_vendas['BANDEIRA1'].replace({
            'Visa': 'VISA',
            'Mastercard': 'MASTER',
            'Amex': 'AMEX',
            'Elo': 'ELO',
            'Hipercard': 'HIPER'
        })


        df_vendas['VALOR LIQUIDO'] = df_vendas['VALOR LIQUIDO'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)

        df_vendas['VALOR LIQUIDO'] = df_vendas['VALOR LIQUIDO'].astype(float)

        csv_vendas = df_vendas.drop(columns='BANDEIRA1')


        mesclagem = pd.merge(exel_notas, csv_vendas, on = 'NSU')


        mesclagem['VALOR'] = mesclagem['VALOR'].astype(float)

        mesclagem['VALOR'] = mesclagem['VALOR'].round(2)


        def verificar_nota(linha):
            if linha['BANDEIRA_x'] != linha['BANDEIRA_y']:
                return 'Bandeira Divergente'
            elif linha['PARCELA'] != linha['PARCELAS']:
                return 'Parcela Divergente'
            elif abs(linha['VALOR'] - linha['VALOR LIQUIDO']) > 0.05:
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
