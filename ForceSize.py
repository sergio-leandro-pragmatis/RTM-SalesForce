import streamlit as st
import folium
from streamlit_folium import folium_static
import pandas as pd
from math import ceil
import pandas as pd
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
from PIL import Image


import numpy as np
from python_tsp.distances import great_circle_distance_matrix
from python_tsp.distances import tsplib_distance_matrix
from python_tsp.heuristics import solve_tsp_local_search, solve_tsp_simulated_annealing

def main():
    st.title('Pragmatis Consultoria - Squad Analytics')


    foto = Image.open('logo_analytics.png')
    st.image(foto, caption='Squad Analytics', use_column_width=False)


    menu = ['Simulador de Força de Vendas', 'Sobre']
    choice = st.sidebar.selectbox("Menu", menu)

    if choice == menu[0]:
        st.header(menu[0])

        st.subheader('Descrição')
        st.write('Essa solução para o problema do dimensionamento da força de vendas utiliza como base as seguintes informações: latitude e longitude dos pontos de venda, tempo médio de permanência no ponto de venda e o tempo médio entre os pontos de venda daquela rota.')

        st.write('Para o dimensionamento da força de venda, é a utilizada a seguinte expressão:')

        r'''
        $$ForcaDeVendas = \frac{\sum_{}^{}Freq_(PDV_i)* (Tempo Médio Visita_i + Tempo Medio Entre PDVs)}{CargaHoráriaSemanal}$$
        '''

        st.write('Para o cálculo do tempo médio entre rotas uma rota contendo todos os PDV é construída e a distância média entre esses PDV é calculada.')

        st.subheader('Input')

        CHS = st.number_input('Carga Horária Semanal', value = 40)

        FA = st.number_input('Fator de Ajuste', value = 1.4)
        
        st.write('Para o input utilize insira uma planilha no padrão abaixo (na mesma ordem no arquivo excel):')
        
        PDV_lat = [19.53796400,
                    19.51838100,
                    19.49321000,
                    19.52762000,
                    19.56123300]

        PDV_long = [-99.189358,
                    -99.156774,
                    -99.184939,
                    -99.14794,
                    -99.247383]

        Freq = [0,1,3,2,1,3]
        Tempo_medio = [9.526991577,
                        4.512250739,
                        5.47204333,
                        11.67188591,
                        11.99083285]
        dict_ = {"PDV_lat": PDV_lat, "PDV_long": PDV_long, "Freq": Freq, "Tempo_medio_visita": Tempo_medio}
        example = pd.DataFrame(dict_)
        st.table(example)
        
        
        def to_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            format1 = workbook.add_format({'num_format': '0.00'})
            worksheet.set_column('A:A', None, format1)
            writer.save()
            processed_data = output.getvalue()
            return processed_data
            df_xlsx = to_excel(base_lat_long, pd.DataFrame(dict))
            
            df_xlsx = to_excel(example)

            st.download_button(label='📥 Clique aqui para baixar o template ',
                               data=df_xlsx,
                               file_name='Template_ForcaDeVendas.xlsx')
        
        
        
        
        
        
        
        busca = st.file_uploader('Insira a planilha para a busca por aqui')

        if busca:
            base_lat_long = pd.read_excel(busca)
            st.dataframe(base_lat_long)
            colunas = base_lat_long.columns

            #Salvando informações:
            lat_ = base_lat_long[colunas[0]]
            long_ = base_lat_long[colunas[1]]
            freq_ = list(base_lat_long[colunas[2]])
            tempv_ = list(base_lat_long[colunas[3]])

            #Salvando em latlong shape:
            latlong = []
            longlat = []

            for i in range(len(lat_)):
                latlong.append([lat_[i], long_[i]])
            for i in range(len(lat_)):
                longlat.append([long_[i], lat_[i]])

            #Constrói Matriz de distâncias:
            distance_matrix = great_circle_distance_matrix(latlong)*FA

            permutation, distance = solve_tsp_local_search(
                distance_matrix,
                x0=None,
                perturbation_scheme="two_opt",
                max_processing_time=None,
                log_file=None,
            )
            permutation2, distance2 = solve_tsp_simulated_annealing(distance_matrix)
            permutation3, distance3 = solve_tsp_local_search(distance_matrix, x0=permutation2)
            permutation4, distance4 = solve_tsp_local_search(distance_matrix, x0=permutation3,                                                           perturbation_scheme="ps3")

            st.subheader('Rota Geral Contendo todos os Pontos de Venda:')

            cidades_ordem = []

            for i in permutation4:
                cidades_ordem.append(latlong[i])


            m = folium.Map(location=cidades_ordem[0], zoom_start=11)

            points = cidades_ordem
            for i in points:
                folium.Marker(i, color = 'green').add_to(m)
            folium.PolyLine(points, color='blue').add_to(m)
            folium_static(m)

            qtd_rotas = len(permutation4) - 1

            distancia_media = distance4 / qtd_rotas
            tempo_medio = 60*((distancia_media / 1000) / 50)
            tempo_medio_40 = 60*((distancia_media / 1000) / 40)
            tempo_medio_60 = 60*((distancia_media / 1000) / 60)

            st.write(f'Distância média entre PDVs (KM): {distancia_media/1000}')
            st.write(f'Tempo Médio entre PDVs (min - v = 50km/h): {tempo_medio}')
            st.write(f'Tempo Médio entre PDVs (min - v = 40km/h): {tempo_medio_40}')
            st.write(f'Tempo Médio entre PDVs (min - v = 60km/h): {tempo_medio_60}')

            Tempo_medio_rota = []
            Tempo_medio_rota_40 = []
            Tempo_medio_rota_60 = []

            for i in range(len(permutation4)):
                Tempo_medio_rota.append(tempo_medio )
                Tempo_medio_rota_40.append(tempo_medio_40 )
                Tempo_medio_rota_60.append(tempo_medio_60 )

            freq_ = list(base_lat_long[colunas[2]])
            tempv_ = list(base_lat_long[colunas[3]])


            tempo_medio_total = []
            tempo_medio_total_40 = []
            tempo_medio_total_60 = []

            tempo_medio_total = list(map(lambda v1, v2: v1 + v2, tempv_, Tempo_medio_rota))
            tempo_medio_total_40 = list(map(lambda v1, v2: v1 + v2, tempv_, Tempo_medio_rota_40))
            tempo_medio_total_60 = list(map(lambda v1, v2: v1 + v2, tempv_, Tempo_medio_rota_60))

            tempo_total_semanal = list(map(lambda v1, v2: v1 * v2, tempo_medio_total, freq_))
            tempo_total_semanal_40 = list(map(lambda v1, v2: v1 * v2, tempo_medio_total_40, freq_))
            tempo_total_semanal_60 = list(map(lambda v1, v2: v1 * v2, tempo_medio_total_60, freq_))

            base_lat_long['Tempo Médio entre PDVs (min - v = 50km/h)'] = Tempo_medio_rota
            base_lat_long['Tempo Médio entre PDVs (min - v = 40km/h)'] = Tempo_medio_rota_40
            base_lat_long['Tempo Médio entre PDVs (min - v = 60km/h)'] = Tempo_medio_rota_60

            base_lat_long['Tempo Médio Total na Semana (min - v = 50km/h)'] = tempo_total_semanal
            base_lat_long['Tempo Médio Total na Semana (min - v = 40km/h)'] = tempo_total_semanal_40
            base_lat_long['Tempo Médio Total na Semana (min - v = 60km/h)'] = tempo_total_semanal_60

            st.subheader('Tempos por PDV:')

            st.dataframe(base_lat_long)

            Tam_Forca_Vendas = ceil(sum(base_lat_long['Tempo Médio Total na Semana (min - v = 50km/h)'] )/ (CHS * 60))
            Tam_Forca_Vendas_40 = ceil(sum(base_lat_long['Tempo Médio Total na Semana (min - v = 40km/h)'] )/ (CHS * 60))
            Tam_Forca_Vendas_60 = ceil(sum(base_lat_long['Tempo Médio Total na Semana (min - v = 60km/h)']) / (CHS * 60))


            #st.write(  sum(base_lat_long['Tempo Médio Total na Semana (min - v = 50km/h)']),
            #           sum(base_lat_long['Tempo Médio Total na Semana (min - v = 40km/h)']),
            #          sum(base_lat_long['Tempo Médio Total na Semana (min - v = 60km/h)']) )


            dict = {'Tam_Forca_Vendas_40': [Tam_Forca_Vendas_40],'Tam_Forca_Vendas_50':[Tam_Forca_Vendas], 'Tam_Forca_Vendas_60':[Tam_Forca_Vendas_60]}


            st.subheader('Tamanho da Força de Vendas')

            st.dataframe(pd.DataFrame(dict))

            def to_excel(df,df2):
                output = BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                df2.to_excel(writer, index= False, sheet_name = 'Sheet2')

                workbook = writer.book

                worksheet = writer.sheets['Sheet1']
                worksheet2 = writer.sheets['Sheet2']
                format1 = workbook.add_format({'num_format': '0.00'})

                worksheet.set_column('A:A', None, format1)
                worksheet2.set_column('A:A', None, format1)
                writer.save()
                processed_data = output.getvalue()
                return processed_data


            df_xlsx = to_excel(base_lat_long, pd.DataFrame(dict))

            st.download_button(label='📥 Clique aqui para baixar os resultados ',
                               data=df_xlsx,
                               file_name='Força de Vendas Resultados.xlsx')


    if choice == menu[1]:
        st.header(menu[1])
        st.write('Essa ferramenta está em fase de protótipo, em caso de bugs ou sugestões de melhoria, entre em contato com: sergio.campos@pragmatis.com.br')
if __name__ == '__main__':
    main()
