import pandas as pd
import os
import glob

# caminho para ler os arquivos
folder_path = 'src\\data\\raw'

# Lista de todos os arquivos excel
excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

if not excel_files:
    print('Ninhum aquivo encontrado no formato xlsx')
else:

    # dataFrame
    df = []

    # Percorrer cada arquivo xlsx | tratativa de erro
    for excel_file in excel_files:
        try:
            # ler o aquivo excel
            df_temp = pd.read_excel(excel_file)

            # pega o nome do arquivo
            file_name = os.path.basename(excel_file)
            
            # Criar coluna para identificar pais
            if 'brasil' in file_name.lower():
                df_temp['location'] = 'br'
            elif 'france' in file_name.lower():
                df_temp['location'] = 'fr'
            elif 'italian' in file_name.lower():
                df_temp['location'] = 'it'
           
                
            #criar nova coluna para extrar a campanha    
            df_temp['campanha'] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')
            df_temp['file_name'] = file_name
            
            #guarda dados tratados de um dataframe 
            df.append(df_temp)
            print(df_temp)

        except Exception as e:
            print(f'Erro ao ler o arquivo {excel_file} : {e}')


if df:
    #concatena todas as tabelas salvas na df em uma unica tabela
    resultado = pd.concat(df, ignore_index=True)
    
    # arquivo de saida (arquivo tratado)
    output_file = os.path.join('src','data','ready', 'clean.xlsx')
    
    #definindo o motor de escrita
    writer = pd.ExcelWriter(output_file,engine='xlsxwriter')
    resultado.to_excel(writer, index=False)
    writer._save()
else:
    print('nenhum dado a ser salvo')