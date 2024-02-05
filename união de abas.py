import pandas as pd

new_excel_file_path = r'C:\Users\moesios\Desktop\tratamento contorno python\ESTADOS BR (21).xlsx'

dfs = pd.read_excel(new_excel_file_path, sheet_name=None)

menu = dfs.get('0.1.MENU', pd.DataFrame())
formulario = dfs.get('1.FORMULARIO', pd.DataFrame())
veiculo = dfs.get('1.1.DADOS GERAIS DO VEÍCULO', pd.DataFrame())

resultado_final = formulario.merge(menu, left_on='ID_MENU', right_on='ID_MENU', how='left')

colunas_para_renomear = {
    'P_1.1': 'TIPO DE VEÍCULO',
    'P_1.2.1': 'CLASSE DO VEÍCULO_P',
    'P_1.2.2': 'CLASSE DO VEÍCULO_M',
    'P_1.2.3': 'CLASSE DO VEÍCULO_C',
    'P_1.4': 'NÚMERO DE PASSAGEIROS',
    'P_1.5': 'NÚMERO DE FUNCIONÁRIOS OU AUXILIARES DO MOTORISTA',
    'P_1.6': 'QUAL O TIPO DA CARGA?',
    'P_1.7': 'QUAL O PESO DA CARGA? (EM TONELADAS)',
    'P_2.2': 'HORÁRIO DE INÍCIO DA VIAGEM',
    'P_2.3': 'MOTIVO DE ESTAR NO LOCAL DE ORIGEM',
    'P_2.3.1': 'OUTROS_MOTIVOS_ORIGEM',
    'P_2.4': 'MOTIVO DE SE DESLOCAR PARA O LOCAL DE DESTINO',
    'P_2.4.1': 'OUTROS_MOTIVOS_DESTINO',
    'P_2.5': 'FREQUÊNCIA DE REALIZAÇÃO DA VIAGEM',
    'P_2.6': 'ESTADO DE ORIGEM DA VIAGEM',
    'P_2.6.1': 'MUNICÍPIO DE ORIGEM DA VIAGEM',
    'P_2.6.1.1': 'BAIRRO DE ORIGEM DA VIAGEM',
    'P_2.6.1.1.1': 'LOGRADOURO_ORIGEM',
    'P_2.6.1.1.2': 'PONTO DE REFERÊNCIA_ORIGEM',
    'P_2.7': 'ESTADO DE DESTINO DA VIAGEM',
    'P_2.7.1': 'MUNICÍPIO DE DESTINO DA VIAGEM',
    'P_2.7.1.1': 'BAIRRO DE DESTINO DA VIAGEM',
    'P_2.7.1.1.1': 'LOGRADOURO_DESTINO',
    'P_2.7.1.1.2': 'PONTO DE REFERÊNCIA_DESTINO',
    'P_0.1': 'IDENTIFICAÇÃO DO PESQUISADOR',
    'P_0.2': 'IDENTIFICAÇÃO DO POSTO DE PESQUISA',
    'P_0.3': 'SENTIDO DO TRÁFEGO PESQUISADO'
}

resultado_final = resultado_final.rename(columns=colunas_para_renomear)

resultado_final['ENDEREÇO_ORIGEM_CONCATENADO'] = (
    resultado_final['PONTO DE REFERÊNCIA_ORIGEM'].astype(str) + ' - ' +
    resultado_final['LOGRADOURO_ORIGEM'].astype(str) + ', ' +
    resultado_final['BAIRRO DE ORIGEM DA VIAGEM'].astype(str) + ', ' +
    resultado_final['MUNICÍPIO DE ORIGEM DA VIAGEM'].astype(str) + ' - ' +
    resultado_final['ESTADO DE ORIGEM DA VIAGEM'].astype(str)
)

resultado_final['ENDEREÇO_DESTINO_CONCATENADO'] = (
    resultado_final['PONTO DE REFERÊNCIA_DESTINO'].astype(str) + ' - ' +
    resultado_final['LOGRADOURO_DESTINO'].astype(str) + ', ' +
    resultado_final['BAIRRO DE DESTINO DA VIAGEM'].astype(str) + ', ' +
    resultado_final['MUNICÍPIO DE DESTINO DA VIAGEM'].astype(str) + ' - ' +
    resultado_final['ESTADO DE DESTINO DA VIAGEM'].astype(str)
)

resultado_final['ENDEREÇO_ORIGEM_CONCATENADO'] = resultado_final.apply(
    lambda row: ' - '.join(filter(lambda x: pd.notna(x), row[['PONTO DE REFERÊNCIA_ORIGEM', 'LOGRADOURO_ORIGEM', 'BAIRRO DE ORIGEM DA VIAGEM', 'MUNICÍPIO DE ORIGEM DA VIAGEM', 'ESTADO DE ORIGEM DA VIAGEM']])),
    axis=1
)

resultado_final['ENDEREÇO_DESTINO_CONCATENADO'] = resultado_final.apply(
    lambda row: ' - '.join(filter(lambda x: pd.notna(x), row[['PONTO DE REFERÊNCIA_DESTINO', 'LOGRADOURO_DESTINO', 'BAIRRO DE DESTINO DA VIAGEM', 'MUNICÍPIO DE DESTINO DA VIAGEM', 'ESTADO DE DESTINO DA VIAGEM']])),
    axis=1
)

resultado_final['VALIDAÇÃO GEO_ORIGEM'] = None
resultado_final['VALIDAÇÃO GEO_DESTINO'] = None
resultado_final['ENDEREÇO ORIGEM GEOCODIFICADO'] = None
resultado_final['ENDEREÇO DESTINO GEOCODIFICADO'] = None
resultado_final['LAT ORIGEM'] = None
resultado_final['LONG ORIGEM'] = None
resultado_final['LAT DESTINO'] = None
resultado_final['LONG DESTINO'] = None
resultado_final['ZT ORIGEM'] = None
resultado_final['ZT DESTINO'] = None

Ordem_Colunas = [
    'ID_MENU', 'ID_DADOS', 'ID_FORMULARIO', 'IDENTIFICAÇÃO DO PESQUISADOR', 'IDENTIFICAÇÃO DO POSTO DE PESQUISA',
    'SENTIDO DO TRÁFEGO PESQUISADO', 'DATA_y', 'HORA_y', 'LATLONG_y', 'EMAIL_y', 'VERSÃO DO APP_y', 'TIPO DE VEÍCULO',
    'CLASSE DO VEÍCULO_P', 'CLASSE DO VEÍCULO_M', 'CLASSE DO VEÍCULO_C', 'NÚMERO DE PASSAGEIROS',
    'NÚMERO DE FUNCIONÁRIOS OU AUXILIARES DO MOTORISTA', 'QUAL O TIPO DA CARGA?', 'QUAL O PESO DA CARGA? (EM TONELADAS)',
    'HORÁRIO DE INÍCIO DA VIAGEM', 'MOTIVO DE ESTAR NO LOCAL DE ORIGEM', 'OUTROS_MOTIVOS_ORIGEM',
    'MOTIVO DE SE DESLOCAR PARA O LOCAL DE DESTINO', 'OUTROS_MOTIVOS_DESTINO', 'FREQUÊNCIA DE REALIZAÇÃO DA VIAGEM',
    'ESTADO DE ORIGEM DA VIAGEM', 'MUNICÍPIO DE ORIGEM DA VIAGEM', 'BAIRRO DE ORIGEM DA VIAGEM', 'LOGRADOURO_ORIGEM',
    'PONTO DE REFERÊNCIA_ORIGEM', 'ENDEREÇO_ORIGEM_CONCATENADO','ENDEREÇO ORIGEM GEOCODIFICADO', 'LAT ORIGEM', 'LONG ORIGEM', 'ZT ORIGEM','VALIDAÇÃO GEO_ORIGEM', 'ESTADO DE DESTINO DA VIAGEM', 'MUNICÍPIO DE DESTINO DA VIAGEM',
    'BAIRRO DE DESTINO DA VIAGEM', 'LOGRADOURO_DESTINO', 'PONTO DE REFERÊNCIA_DESTINO', 'ENDEREÇO_DESTINO_CONCATENADO', 'ENDEREÇO DESTINO GEOCODIFICADO', 'LAT DESTINO', 'LONG DESTINO', 'ZT DESTINO',
    'VALIDAÇÃO GEO_DESTINO',
]

resultado_final = resultado_final[Ordem_Colunas]

caminho_saida_excel = r'C:\Users\moesios\Desktop\tratamento contorno python\Resultado_final.xlsx'
with pd.ExcelWriter(caminho_saida_excel, engine='xlsxwriter') as writer:
    resultado_final.to_excel(writer, sheet_name='Resultado_Final', index=False)

df = pd.read_excel(caminho_saida_excel, sheet_name='Resultado_Final')

counts = df['TIPO DE VEÍCULO'].value_counts()

result_df = pd.DataFrame({'TIPO DE VEÍCULO': counts.index, 'NÚMERO DE PESQUISAS': counts.values})

with pd.ExcelWriter(caminho_saida_excel, mode='a', engine='openpyxl') as writer:
    result_df.to_excel(writer, sheet_name='Veículos por tipo', index=False)

