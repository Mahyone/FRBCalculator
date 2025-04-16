
import streamlit as st
from PIL import Image
import pandas as pd
import numpy as np
from itertools import combinations
import openpyxl
import io
import os
import matplotlib.pyplot as plt
import plotly.express as px
import seaborn as sns

# Definir a configuração da página no início
st.set_page_config(page_title="Calculadora FRB - Alocação", page_icon="📊", layout="wide")

# Carregar e exibir o logo
logo = Image.open("FRBConsulting_Logo.PNG")  
st.image(logo, use_container_width=False) 


# Função para o Upload de Arquivo (script original do Upload)
def upload_arquivo():    

    # Reiniciar os dados no session_state
    if "df_staffheadcount" in st.session_state:
        del st.session_state.df_staffheadcount
    if "df_staffoccupancy_trat" in st.session_state:
        del st.session_state.df_staffoccupancy_trat
    if "df_subgroupadjacenties" in st.session_state:
        del st.session_state.df_subgroupadjacenties
    if "df_building_trat" in st.session_state:
        del st.session_state.df_building_trat

    if "df_unido" in st.session_state:
        del st.session_state.df_unido
    if "df_enriquecido" in st.session_state:
        del st.session_state.df_enriquecido
    if "df_proportional" in st.session_state:
        del st.session_state.df_proportional

     
    # Título da aplicação
    st.write("### Leitura e Processamento de Abas do Excel")

    # Dividir a interface em abas
    tabs = st.tabs(["Importar Arquivo", "Automação", "Cenarios", "Dashboards"])


    ##### ABA IMPORTAÇÃO #####   
    with tabs[0]:
        st.header("Importar Arquivo")
        
        # Função para carregar e processar os dados do Excel
        def process_excel_data(file_path):
            try:
                # Carregar tabelas do Excel
                df_staffheadcount = pd.read_excel(file_path, sheet_name='2. Staff Headcount ', skiprows=3, usecols="A:I")
                df_staffoccupancy = pd.read_excel(file_path, sheet_name='3. Staff Occupancy', skiprows=3, usecols="A:F")
                df_subgroupadjacenties = pd.read_excel(file_path, sheet_name='4. SubGroup Adjacencies', skiprows=3, usecols="A:E")
                df_building = pd.read_excel(file_path, sheet_name='5. Building Space Summary', skiprows=6, usecols="A:AC")

                # Processar df_staffheadcount
                rename_dfstaffheadcount = {
                    0: 'Current Location',
                    1: 'Group',
                    2: 'SubGroup',
                    3: 'FTE',
                    4: 'CW',
                    5: 'Growth',
                    6: 'Total',
                    7: 'Exception (Y/N)',
                    8: 'Comments'
                }        
                df_staffheadcount.columns = [rename_dfstaffheadcount.get(i, col) for i, col in enumerate(df_staffheadcount.columns)]
                
                df_staffheadcount = df_staffheadcount[
                    (df_staffheadcount['Total'] > 0) & 
                    (df_staffheadcount['Group'].notna())
                ]
                df_staffheadcount['Group'] = df_staffheadcount['Group'].astype(str)
                df_staffheadcount['SubGroup']  = df_staffheadcount['SubGroup'].fillna("").astype(str)
                df_staffheadcount['FTE'] = df_staffheadcount['FTE'].fillna(0).replace([np.inf, -np.inf], 0).astype(int)
                df_staffheadcount['CW'] = df_staffheadcount['CW'].fillna(0).replace([np.inf, -np.inf], 0).astype(int)
                df_staffheadcount['Growth'] = df_staffheadcount['Growth'].fillna(0).replace([np.inf, -np.inf], 0).astype(int)


                # Processar df_staffoccupancy
                rename_dfstaffoccupancy = {
                    0: 'Group',
                    1: 'HeadCount',
                    2: 'Avg Peak',
                    3: 'Avg Occupancy',
                    4: 'Perc Avg Peak',
                    5: 'Perc Avg'
                }        
                df_staffoccupancy.columns = [rename_dfstaffoccupancy.get(i, col) for i, col in enumerate(df_staffoccupancy.columns)]
                df_staffoccupancy = df_staffoccupancy[df_staffoccupancy['Group'] != 0]
                df_staffoccupancy['Group'] = df_staffoccupancy['Group'].astype(str)

                def preencher_valores(df):
                    df['Perc Avg Peak'] = np.where(df['Perc Avg Peak'].isna(), df['Avg Peak'] / df['HeadCount'], df['Perc Avg Peak'])
                    df['Perc Avg'] = np.where(df['Perc Avg'].isna(), df['Avg Occupancy'] / df['HeadCount'], df['Perc Avg'])
                    df['Avg Peak'] = np.where(df['Avg Peak'].isna(), df['Perc Avg Peak'] * df['HeadCount'], df['Avg Peak'])
                    df['Avg Occupancy'] = np.where(df['Avg Occupancy'].isna(), df['Perc Avg'] * df['HeadCount'], df['Avg Occupancy'])

                    df['Perc Avg Peak'] = df['Perc Avg Peak'] * 100
                    df['Perc Avg Peak'] = df['Perc Avg Peak'].round(0).astype(int)
                    df['Perc Avg'] = df['Perc Avg'] * 100
                    df['Perc Avg'] = df['Perc Avg'].round(0).astype(int)

                    df['Avg Peak'] = df['Avg Peak'].round(0).astype(int)
                    df['Avg Occupancy'] = df['Avg Occupancy'].round(0).astype(int)

                    return df

                df_staffoccupancy_trat = preencher_valores(df_staffoccupancy)

                # Processar df_subgroupadjacenties
                rename_dfsubgroupsadjacencies = {
                    0: 'Group',
                    1: 'SubGroup',
                    2: 'Adjacency Priority 1',
                    3: 'Adjacency Priority 2',
                    4: 'Adjacency Priority 3'
                }        
                df_subgroupadjacenties.columns = [rename_dfsubgroupsadjacencies.get(i, col) for i, col in enumerate(df_subgroupadjacenties.columns)]
                df_subgroupadjacenties = df_subgroupadjacenties[(df_subgroupadjacenties['Group'] != 0) & (df_subgroupadjacenties['Group'].notna())]
                df_subgroupadjacenties['Group'] = df_subgroupadjacenties['Group'].fillna("").astype(str)
                df_subgroupadjacenties['SubGroup'] = df_subgroupadjacenties['SubGroup'].fillna("").astype(str)


                df_building.rename(columns={df_building.columns[0]: 'Building Name'}, inplace=True)
                df_building.rename(columns={df_building.columns[1]: 'Primary Work Seats'}, inplace=True)
                df_building.rename(columns={df_building.columns[27]: 'Primary Work Seats'}, inplace=True)

                for col in df_building.columns:
                    if col != 'Building Name':
                        df_building[col] = pd.to_numeric(df_building[col], errors='coerce').fillna(0).astype(int)

                if 'Primary Work Seats' not in df_building.columns:
                    st.warning("Coluna 'Primary Work Seats' não encontrada. Adicionando valores padrão.")
                    df_building['Primary Work Seats'] = 0

                df_building_trat = df_building[
                    (df_building['Primary Work Seats'] > 0) & 
                    (df_building['Building Name'].notna())
                ]

                return df_staffheadcount, df_staffoccupancy_trat, df_subgroupadjacenties, df_building_trat
            except Exception as e:
                st.error(f"Erro ao processar o arquivo: {e}")
                return None, None, None, None

            
        # Função para substituir valores nulos e exibir tabelas sem índice
        def process_and_display_table(df):
            # Substituir NaN, NAT ou nulos por vazios
            df = df.fillna("")  # Substitui valores nulos por células vazias
            # Ajustar índice para começar de 1
            df.index = df.index + 1
            # Exibir a tabela sem o índice
            st.table(df)    


        # Verificar se o arquivo foi carregado
        uploaded_file = st.file_uploader("Carregue o arquivo Excel (DataCollection.xlsx):", type=["xlsx"])
        if uploaded_file:
            st.write("### Processando o arquivo...")
            
            # Processar os dados
            df_staffheadcount, df_staffoccupancy_trat, df_subgroupadjacenties, df_building_trat = process_excel_data(uploaded_file)
            
            # Verificar se os DataFrames foram carregados corretamente
            if df_staffheadcount is not None and df_staffoccupancy_trat is not None and df_subgroupadjacenties is not None and df_building_trat is not None:
                st.session_state.df_staffheadcount = df_staffheadcount
                st.session_state.df_staffoccupancy_trat = df_staffoccupancy_trat
                st.session_state.df_subgroupadjacenties = df_subgroupadjacenties
                st.session_state.df_building_trat = df_building_trat
                
                st.success("Arquivo processado com sucesso!")
                with st.expander("### Tabela 'Staff Headcount'"):
                    st.write("### Tabela 'Staff Headcount' :")
                    process_and_display_table(df_staffheadcount)

                with st.expander("### Tabela 'Staff Occupancy"):
                    st.write("### Tabela 'Staff Occupancy':")
                    process_and_display_table(df_staffoccupancy_trat)

                with st.expander("### Tabela 'SubGroup Adjacencies'"):
                    st.write("### Tabela 'SubGroup Adjacencies':")
                    process_and_display_table(df_subgroupadjacenties)  

                with st.expander("### Tabela 'Building'"):
                    st.write("### Tabela 'Building':")
                    process_and_display_table(df_building_trat)            
                
                
                # Processar e armazenar os dados enriquecidos
                df_unido = pd.merge(df_staffheadcount, df_subgroupadjacenties, how='left', on=['Group', 'SubGroup'])
                df_unido['FTE'] = df_unido['FTE'].fillna(0).round(0).astype(int)
                df_unido['CW'] = df_unido['CW'].fillna(0).round(0).astype(int)                
                df_unido['Growth'] = df_unido['Growth'].fillna(0).round(0).astype(int)

                # Aplicando a distribuição proporcional para Peak e Occupancy (fechados por Grupo no Excel)
                df_proportional = pd.merge(df_unido, df_staffoccupancy_trat, how='left', on='Group')

                # Calcular a proporção de HeadCount
                df_proportional['Proportionhc'] = df_proportional['Total'] / df_proportional['HeadCount']

                # Calcular os valores proporcionais de Peak e Avg Occupancy
                df_proportional['Proportional Peak'] = df_proportional['Avg Peak'] * df_proportional['Proportionhc']
                df_proportional['Proportional Avg'] = df_proportional['Avg Occupancy'] * df_proportional['Proportionhc']
                df_proportional['Proportional Peak'] = df_proportional['Proportional Peak'].round(0).astype(int)
                df_proportional['Proportional Avg'] = df_proportional['Proportional Avg'].round(0).astype(int)
                df_proportional.drop('HeadCount', axis=1, inplace=True)
                df_proportional.rename(columns={'Total': 'HeadCount'}, inplace=True)

                #df_enriquecido = df_proportional.copy()
                #df_enriquecido = df_enriquecido[['Current Location', 'Group', 'SubGroup', 'FTE','CW', 'Growth', 'HeadCount', 'Exception (Y/N)', 'Comments', 'Avg Peak', 'Avg Occupancy','Adjacency Priority 1', 'Adjacency Priority 2', 'Adjacency Priority 3']]      

                df_proportional = df_proportional[['Current Location', 'Group', 'SubGroup', 'FTE','CW', 'Growth', 'HeadCount', 'Exception (Y/N)', 'Proportional Peak', 'Proportional Avg',
                                                   'Adjacency Priority 1', 'Adjacency Priority 2', 'Adjacency Priority 3']]             


                # Exibir a tabela resultante
                st.write("### Abas Consolidadas em uma única tabela':")
                st.write("Os campos 'Proportional' são calculados quando há mais de um SubGroup para o mesmo Group, pois a informação de Peak e Avg Occ é cadastrada por Group.")
                st.session_state.df_proportional = df_proportional
                process_and_display_table(df_proportional)
                                

                # Botão para exportar tabela "Building" para Excel
                if st.button("Exportar Tabela 'Building' para Excel", key="export_building"):
                    with io.BytesIO() as output:
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            df_building_trat = st.session_state.df_building_trat
                            df_building_trat.to_excel(writer, sheet_name="Building", index=False)
                        st.download_button(
                            label="Download do Excel - Building",
                            data=output.getvalue(),
                            file_name="config_building.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                # Botão para exportar tabela "Grupos, SubGrupos e Adjacentes" para Excel
                if st.button("Exportar Tabela 'Grupos, SubGrupos e Adjacentes' para Excel", key="export_enriquecido"):
                    with io.BytesIO() as output:
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            df_proportional = st.session_state.df_proportional
                            df_proportional.to_excel(writer, sheet_name="Grupos, SubGrupos e Adjacentes", index=False)
                        st.download_button(
                            label="Download do Excel - Grupos, SubGrupos e Adjacentes",
                            data=output.getvalue(),
                            file_name="grupos_subgrupos_adjacentes.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            else:
                st.warning("Tabela 'Grupos, SubGrupos e Adjacentes' não disponível.")



    ##### ABA AUTOMAÇÃO #####
    with tabs[1]:
        st.header("Automação")
        st.write("Para o cálculo de espaços está sendo considerado 'Primary Work Seats'.")

        # Inicializar df_proportional como um DataFrame vazio, se não houver dados na sessão
        if "df_building_trat" not in st.session_state and "df_proportional" not in st.session_state:
            df_building_trat = pd.DataFrame()     
            df_proportional = pd.DataFrame()  
        else:
            df_building_trat = st.session_state.df_building_trat
            df_proportional = st.session_state.df_proportional

        # Verificar se o df_proportional tem dados antes de continuar
        if not df_building_trat.empty and not df_proportional.empty:

            with st.expander("### Dados Cadastrados"):

                # Carregar os dados e realizar a alocação
                if "df_building_trat" in st.session_state and "df_proportional" in st.session_state:
                    df_building_trat = st.session_state.df_building_trat
                    df_proportional = st.session_state.df_proportional

                    # Exibindo a Tabela 'Building Space Summary' com a linha de total
                    st.write("#### Tabela 'Building Space Summary'")
                    df_building_trat_total = df_building_trat.copy()
                    numeric_columns_building = df_building_trat.select_dtypes(include=['number']).columns
                    df_building_trat_total.loc['Total', 'Building Name'] = 'Total'
                    df_building_trat_total.loc['Total', numeric_columns_building] = df_building_trat_total[numeric_columns_building].sum()

                    st.dataframe(df_building_trat_total.fillna(""), use_container_width=False, hide_index=True)

                    # Exibindo a Tabela 'Grupos, SubGrupos e Adjacentes' com a linha de total
                    st.write("#### Tabela 'Grupos, SubGrupos e Adjacentes'")
                    df_proportional_total = df_proportional.copy()
                    numeric_columns_proportional = df_proportional.select_dtypes(include=['number']).columns
                    df_proportional_total.loc['Total', 'Group'] = 'Total'
                    df_proportional_total.loc['Total', numeric_columns_proportional] = df_proportional_total[numeric_columns_proportional].sum()
                    st.dataframe(df_proportional_total.fillna(""), use_container_width=False, hide_index=True)




            with st.expander("### Automação considerando HeadCount"):
                primary_work_seats = df_building_trat_total['Primary Work Seats'].iloc[-1].astype(int)
                total_seats_on_floor = df_building_trat_total['Total seats on floor'].iloc[-1].astype(int)
                total_headcount = df_proportional["HeadCount"].sum()
                    
                st.write(f"**Primary Work Seats**: {primary_work_seats} || **Total seats on floor**: {total_seats_on_floor}")
                st.write(f"**Total HeadCount**: {total_headcount}")

                # Função de alocação dos grupos nos andares
                def allocate_groups(df_proportional, floors):
                    allocation = {}  # Armazenar a alocação de grupos por andar
                    remaining_groups = df_proportional.sort_values(by='HeadCount', ascending=False)  # Ordenar por HeadCount
                    floor_names = list(floors.keys())
                    
                    # Copiar df_proportional para adicionar a coluna 'Building Name'
                    df_allocation = df_proportional.copy()
                    df_allocation['Building Name'] = 'Não Alocado'  # Coluna inicializada com valor "Não Alocado"
                    
                    # Criar um valor único para grupos sem SubGrupo
                    df_allocation['SubGroup'] = df_allocation['SubGroup'].fillna('NoSubGroup')
                    
                    # Alocar os grupos nos andares disponíveis
                    for _, group in remaining_groups.iterrows():
                        group_name = group['Group']
                        subgroup_name = group['SubGroup']
                        headcount = group['HeadCount']
                        
                        allocated = False  # Flag para verificar se o grupo foi alocado
                        
                        # Tentar alocar o grupo nos andares disponíveis
                        for floor_name in floor_names:
                            if floors[floor_name] >= headcount:
                                # Se couber, aloca
                                df_allocation.loc[(df_allocation['Group'] == group_name) & (df_allocation['SubGroup'] == subgroup_name), 'Building Name'] = floor_name
                                floors[floor_name] -= headcount
                                allocated = True  # Grupo foi alocado
                                break
                        
                        # Se não alocou, marca como "Não Alocado"
                        if not allocated:
                            df_allocation.loc[(df_allocation['Group'] == group_name) & (df_allocation['SubGroup'] == subgroup_name), 'Building Name'] = 'Não Alocado'
                    
                    return df_allocation, floors

                # Função de exibição de alocação com as tabelas ajustadas
                def display_allocation(df_allocation, remaining_floors, df_building_trat):
                    # Reordenar as colunas: "Building Name" em 1ª posição e "Current Location" em última
                    cols = df_allocation.columns.tolist()
                    if "Building Name" in cols and "Current Location" in cols:
                        new_order = (
                            ["Building Name"] +
                            [col for col in cols if col not in ("Building Name", "Current Location")] +
                            ["Current Location"]
                        )
                        df_allocation = df_allocation[new_order]
                    
                    # Ordenar o DataFrame por "Building Name" se ainda não estiver ordenado
                    df_allocation = df_allocation.sort_values(by='Building Name')
                    
                    # Obter os Building Names únicos conforme a ordem do DataFrame (primeira ocorrência)
                    unique_buildings = df_allocation["Building Name"].drop_duplicates().tolist()
                    # Alternar entre cinza claro e sem fundo (transparente)
                    building_colors = {building: "#D3D3D3" if i % 2 == 0 else "" 
                                    for i, building in enumerate(unique_buildings)}
                    
                    def highlight_building(row):
                        color = building_colors.get(row["Building Name"], "")
                        return ['background-color: ' + color] * len(row)
                    
                    st.write("#### Resultado da Automação - HeadCount")
                    df_allocation_styled = df_allocation.style.apply(highlight_building, axis=1)
                    st.dataframe(df_allocation_styled, use_container_width=False)
                    
                    # Exibir a capacidade restante nos andares
                    st.write("#### Capacidade restante nos andares - HeadCount:")
                    remaining_floors_df = pd.DataFrame(list(remaining_floors.items()), 
                                                    columns=['Building Name', 'Remaining Seats'])
                    st.dataframe(remaining_floors_df, use_container_width=False)
                    
                    
                    return df_allocation, remaining_floors_df


                # Carregar os dados e realizar a alocação
                if "df_building_trat" in st.session_state and "df_proportional" in st.session_state:
                    df_building_trat = st.session_state.df_building_trat
                    df_proportional = st.session_state.df_proportional

                    # Exibir as tabelas para debug
                    #st.write("### Tabela 'Building Space Summary'")
                    #st.dataframe(df_building_trat, use_container_width=False)
                    
                    #st.write("### Tabela 'Grupos, SubGrupos e Adjacentes'")
                    #st.dataframe(df_proportional, use_container_width=False)

                    # Extração da capacidade dos andares do df_building_trat
                    floors = dict(zip(df_building_trat['Building Name'], df_building_trat['Primary Work Seats']))

                    # Chamar a função de alocação
                    df_allocation, remaining_floors = allocate_groups(df_proportional, floors.copy())

                    # Exibir os resultados de alocação
                    df_allocation_result, remaining_floors_df_result = display_allocation(df_allocation, remaining_floors, df_building_trat)
                    cols = df_allocation.columns.tolist()
                    if "Building Name" in cols and "Current Location" in cols:
                        new_order = (
                            ["Building Name"] +
                            [col for col in cols if col not in ("Building Name", "Current Location")] +
                            ["Current Location"]
                        )
                        df_allocation = df_allocation[new_order]
                    
                    # Ordenar o DataFrame por "Building Name" se ainda não estiver ordenado
                    df_allocation = df_allocation.sort_values(by='Building Name')
                    dfautomation_hc = df_allocation.copy()
                    st.session_state.dfautomation_hc = dfautomation_hc  # Salvando no session_state

                    st.write("#### Grupos Não Alocados:")
                    df_hc_nonallocated = df_allocation_result[df_allocation_result['Building Name'] == 'Não Alocado']
                    numeric_columns = df_hc_nonallocated.select_dtypes(include='number').columns
                    total_row = df_hc_nonallocated[numeric_columns].sum()
                    total_row['Group'] = 'Total' 
                    total_row_df = pd.DataFrame([total_row])
                    df_hc_nonallocated_with_total = pd.concat([df_hc_nonallocated, total_row_df], ignore_index=True)
                    st.dataframe(df_hc_nonallocated_with_total, use_container_width=False)


                # Botão para exportar tabela "Resultados das Simulações" para Excel
                if st.button("Exportar Tabela 'Resultados das Simulações' para Excel", key="export_unificado"):
                    if "dfautomation_hc" in st.session_state:
                        # Acessa o DataFrame salvo no session_state e substitui NaN por string vazia
                        df_allocation_export = st.session_state.dfautomation_hc.fillna("")
                        
                        # Cria o arquivo Excel em memória
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            df_allocation_export.to_excel(writer, sheet_name="Simulações HC", index=False)
                        output.seek(0)
                        
                        # Botão de download, utilizando output.getvalue() para retornar os bytes do arquivo
                        st.download_button(
                            label="Download do Excel - Resultados das Simulações HeadCount",
                            data=output.getvalue(),
                            file_name="resultados_simulacoes_hc.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("Data not found: 'dfautomation_hc' não está disponível no session_state.")

                    

            with st.expander("### Automação considerando Peak"):
                st.write("Para os Groups + SubGroups que são 'Exception = Y' o valor considerado é Headcount - 1:1.")

                primary_work_seats = df_building_trat_total['Primary Work Seats'].iloc[-1].astype(int)
                total_seats_on_floor = df_building_trat_total['Total seats on floor'].iloc[-1].astype(int)
                total_proppeak = df_proportional["Proportional Peak"].sum()
                    
                st.write(f"**Primary Work Seats**: {primary_work_seats} || **Total seats on floor**: {total_seats_on_floor}")

                # Calcular o "Proportional Peak Exception" (Exception = Y) diretamente no backend
                total_proportional_Peak_exception = df_proportional.apply(
                    lambda row: row['HeadCount'] if row['Exception (Y/N)'] == 'Y' else row['Proportional Peak'],
                    axis=1
                ).sum()

                # Exibir o valor total
                st.write(f"**Total Avg Peak**: {total_proppeak} || **Total Avg Peak with Exception**: {total_proportional_Peak_exception}")


                def allocate_groups_peak(df_proportional, floors):
                    allocation = {}  # Armazenar a alocação de grupos por andar
                    remaining_groups = df_proportional.sort_values(by='HeadCount', ascending=False)  # Ordenar por HeadCount
                    floor_names = list(floors.keys())
                    
                    # Copiar df_proportional para adicionar a coluna 'Building Name'
                    df_allocation = df_proportional.copy()
                    df_allocation['Building Name'] = 'Não Alocado'  # Coluna inicializada com valor "Não Alocado"
                    
                    # Criar um valor único para grupos sem SubGrupo
                    df_allocation['SubGroup'] = df_allocation['SubGroup'].fillna('NoSubGroup')
                    
                    # Alocar os grupos nos andares disponíveis
                    for _, group in remaining_groups.iterrows():
                        group_name = group['Group']
                        subgroup_name = group['SubGroup']
                        
                        # Verificar se há exceção (se a coluna 'Exception' é 'Y')
                        exception = group['Exception (Y/N)']  # Ajuste o nome da coluna conforme necessário
                        
                        # Se houver uma exceção (Exception = 'Y'), usar HeadCount; caso contrário, usar Proportional Peak
                        if exception == 'Y':
                            headcount = group['HeadCount']
                        else:
                            headcount = group['Proportional Peak']  # Use o valor de 'Proportional Peak' para o cálculo
                        
                        allocated = False  # Flag para verificar se o grupo foi alocado
                        
                        # Tentar alocar o grupo nos andares disponíveis
                        for floor_name in floor_names:
                            if floors[floor_name] >= headcount:
                                # Se couber, aloca
                                df_allocation.loc[(df_allocation['Group'] == group_name) & (df_allocation['SubGroup'] == subgroup_name), 'Building Name'] = floor_name
                                floors[floor_name] -= headcount
                                allocated = True  # Grupo foi alocado
                                break
                        
                        # Se não alocou, marca como "Não Alocado"
                        if not allocated:
                            df_allocation.loc[(df_allocation['Group'] == group_name) & (df_allocation['SubGroup'] == subgroup_name), 'Building Name'] = 'Não Alocado'
                    
                    return df_allocation, floors

                
                # Função de exibição de alocação com as tabelas ajustadas
                def display_allocation(df_allocation, remaining_floors, df_building_trat):
                    # Ordenar os dados por 'Building Name'
                    df_allocation = df_allocation.sort_values(by='Building Name')
                    st.write("#### Resultado da Automação - Peak")
                    
                    # Criar a coluna 'Peak with Exception' com base na condição
                    df_allocation['Peak with Exception'] = df_allocation.apply(
                        lambda row: row['HeadCount'] if row['Exception (Y/N)'] == 'Y' else row['Proportional Peak'], 
                        axis=1
                    )
                    # Criar a nova coluna que calcula o % do HeadCount (multiplicado por 100)
                    df_allocation['Peak % of HeadCount'] = ((df_allocation['Peak with Exception'] / df_allocation['HeadCount']) * 100).round(0).astype(int)
                                        
                    # Remover coluna que não será mais necessária e renomear
                    df_allocation.drop(columns=['Proportional Peak'], inplace=True)
                    df_allocation.rename(columns={'Proportional Avg': 'Avg Occ'}, inplace=True)
                    
                    # Reordenar as colunas para que "Building Name" seja a 1ª e "Current Location" a última,
                    # e inserir a nova coluna após "Peak with Exception"
                    df_allocation = df_allocation[['Building Name', 'Group', 'SubGroup', 'FTE', 'CW', 'Growth', 
                                                'HeadCount', 'Exception (Y/N)', 'Peak with Exception', 'Peak % of HeadCount',
                                                'Avg Occ', 'Adjacency Priority 1', 'Adjacency Priority 2', 'Adjacency Priority 3', 
                                                'Current Location']]
                    
                    # Obter os Building Names únicos na ordem de aparecimento (após o sort)
                    unique_buildings = df_allocation['Building Name'].drop_duplicates().tolist()
                    # Definir cores alternadas: cinza claro para índices pares e transparente para ímpares
                    building_colors = {building: "#D3D3D3" if i % 2 == 0 else "" 
                                    for i, building in enumerate(unique_buildings)}
                    
                    # Função para aplicar o estilo de fundo para cada linha, com base no Building Name
                    def highlight_building(row):
                        color = building_colors.get(row['Building Name'], '')
                        return ['background-color: ' + color] * len(row)
                    
                    # Aplica o estilo alternado nas linhas e mantém a formatação específica para as colunas de Peak
                    df_allocation_styled = (
                        df_allocation
                        .style.apply(highlight_building, axis=1)
                        .applymap(lambda x: 'background-color: #D3D3D3', subset=['Peak with Exception', 'Peak % of HeadCount'])
                    )
                    
                    st.dataframe(df_allocation_styled, use_container_width=False)

                    # Exibir a capacidade restante nos andares
                    st.write("#### Capacidade restante nos andares - Peak:")
                    remaining_floors_df = pd.DataFrame(
                        list(remaining_floors.items()), 
                        columns=['Building Name', 'Remaining Seats']
                    )
                    st.dataframe(remaining_floors_df, use_container_width=False)

                    # Retornar o DataFrame modificado
                    return df_allocation, remaining_floors_df


            

                # Carregar os dados e realizar a alocação
                if "df_building_trat" in st.session_state and "df_proportional" in st.session_state:
                    df_building_trat = st.session_state.df_building_trat
                    df_proportional = st.session_state.df_proportional

                    # Exibir as tabelas para debug
                    #st.write("### Tabela 'Building Space Summary'")
                    #st.dataframe(df_building_trat, use_container_width=False)
                    
                    #st.write("### Tabela 'Grupos, SubGrupos e Adjacentes'")
                    #st.dataframe(df_proportional, use_container_width=False)

                    # Extração da capacidade dos andares do df_building_trat
                    floors = dict(zip(df_building_trat['Building Name'], df_building_trat['Primary Work Seats']))

                    # Chamar a função de alocação
                    df_allocation, remaining_floors = allocate_groups_peak(df_proportional, floors.copy())

                    # Exibir os resultados de alocação
                    df_allocation, remaining_floors_df = display_allocation(df_allocation, remaining_floors, df_building_trat)
                    cols = df_allocation.columns.tolist()
                    if "Building Name" in cols and "Current Location" in cols:
                        new_order = (
                            ["Building Name"] +
                            [col for col in cols if col not in ("Building Name", "Current Location")] +
                            ["Current Location"]
                        )
                        df_allocation = df_allocation[new_order]
                    
                    # Ordenar o DataFrame por "Building Name" se ainda não estiver ordenado
                    df_allocation = df_allocation.sort_values(by='Building Name')
                    dfautomation_peak = df_allocation.copy()
                    st.session_state.dfautomation_peak = dfautomation_peak  # Salvando no session_state

                    st.write("### Grupos Não Alocados:")
                    df_peak_nonallocated = df_allocation[df_allocation['Building Name'] == 'Não Alocado']
                    numeric_columns = df_peak_nonallocated.select_dtypes(include='number').columns
                    total_row = df_peak_nonallocated[numeric_columns].sum()
                    total_row['Group'] = 'Total' 
                    total_row_df = pd.DataFrame([total_row])
                    df_peak_nonallocated_total = pd.concat([df_peak_nonallocated, total_row_df], ignore_index=True)
                    st.dataframe(df_peak_nonallocated_total, use_container_width=False)


                # Botão para exportar tabela "Resultados das Simulações" para Excel
                if st.button("Exportar Tabela 'Resultados das Simulações' para Excel", key="export_unificado_peak"):
                    if "dfautomation_peak" in st.session_state:
                        # Acessa o DataFrame salvo no session_state e substitui NaN por string vazia
                        df_allocation_export = st.session_state.dfautomation_peak.fillna("")
                        
                        # Cria o arquivo Excel em memória
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            df_allocation_export.to_excel(writer, sheet_name="Simulações PEAK", index=False)
                        output.seek(0)
                        
                        # Botão de download, utilizando output.getvalue() para retornar os bytes do arquivo
                        st.download_button(
                            label="Download do Excel - Resultados das Simulações Peak",
                            data=output.getvalue(),
                            file_name="resultados_simulacoes_peak.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("Data not found: 'dfautomation_peak' não está disponível no session_state.")



            with st.expander("### Automação considerando Avg Occ"):
                st.write("Para os Groups + SubGroups que são 'Exception = Y' o valor considerado é Headcount - 1:1.")

                primary_work_seats = df_building_trat_total['Primary Work Seats'].iloc[-1].astype(int)
                total_seats_on_floor = df_building_trat_total['Total seats on floor'].iloc[-1].astype(int)
                total_propavg = df_proportional["Proportional Avg"].sum()
                    
                st.write(f"**Primary Work Seats**: {primary_work_seats} || **Total seats on floor**: {total_seats_on_floor}")

                # Calcular o "Proportional Avg Exception" (Exception = Y) diretamente no backend
                total_proportional_Avg_exception = df_proportional.apply(
                    lambda row: row['HeadCount'] if row['Exception (Y/N)'] == 'Y' else row['Proportional Avg'],
                    axis=1
                ).sum()

                # Exibir o valor total
                st.write(f"**Total Avg**: {total_propavg} || **Total Avg with Exception**: {total_proportional_Avg_exception}")

                
                def allocate_groups_avg(df_proportional, floors):
                    allocation = {}  # Armazenar a alocação de grupos por andar
                    remaining_groups = df_proportional.sort_values(by='HeadCount', ascending=False)  # Ordenar por HeadCount
                    floor_names = list(floors.keys())
                    
                    # Copiar df_proportional para adicionar a coluna 'Building Name'
                    df_allocation = df_proportional.copy()
                    df_allocation['Building Name'] = 'Não Alocado'  # Coluna inicializada com valor "Não Alocado"
                    
                    # Criar um valor único para grupos sem SubGrupo
                    df_allocation['SubGroup'] = df_allocation['SubGroup'].fillna('NoSubGroup')
                    
                    # Alocar os grupos nos andares disponíveis
                    for _, group in remaining_groups.iterrows():
                        group_name = group['Group']
                        subgroup_name = group['SubGroup']
                        
                        # Verificar se há exceção (se a coluna 'Exception' é 'Y')
                        exception = group['Exception (Y/N)']  # Ajuste o nome da coluna conforme necessário
                        
                        # Se houver uma exceção (Exception = 'Y'), usar HeadCount; caso contrário, usar Proportional Peak
                        if exception == 'Y':
                            headcount = group['HeadCount']
                        else:
                            headcount = group['Proportional Avg']  # Use o valor de 'Proportional Peak' para o cálculo                        
                        allocated = False  # Flag para verificar se o grupo foi alocado
                        
                        # Tentar alocar o grupo nos andares disponíveis
                        for floor_name in floor_names:
                            if floors[floor_name] >= headcount:
                                # Se couber, aloca
                                df_allocation.loc[(df_allocation['Group'] == group_name) & (df_allocation['SubGroup'] == subgroup_name), 'Building Name'] = floor_name
                                floors[floor_name] -= headcount
                                allocated = True  # Grupo foi alocado
                                break
                        
                        # Se não alocou, marca como "Não Alocado"
                        if not allocated:
                            df_allocation.loc[(df_allocation['Group'] == group_name) & (df_allocation['SubGroup'] == subgroup_name), 'Building Name'] = 'Não Alocado'
                    
                    return df_allocation, floors

                # Função de exibição de alocação com as tabelas ajustadas
                def display_allocation(df_allocation, remaining_floors, df_building_trat):
                    # Ordenar os dados por 'Building Name'
                    df_allocation = df_allocation.sort_values(by='Building Name')
                    st.write("#### Resultado da Automação - Avg Occ")
                    
                    # Criar a coluna 'Avg Occ with Exception'
                    df_allocation['Avg Occ with Exception'] = df_allocation.apply(
                        lambda row: row['HeadCount'] if row['Exception (Y/N)'] == 'Y' else row['Proportional Avg'], 
                        axis=1
                    )
                    # Criar a nova coluna que calcula o % do HeadCount (multiplicado por 100)
                    df_allocation['Avg Occ % of HeadCount'] = ((df_allocation['Avg Occ with Exception'] / df_allocation['HeadCount']) * 100).round(0).astype(int)
                    
                    # Reordenar as colunas para que "Building Name" seja a 1ª e "Current Location" a última
                    df_allocation = df_allocation[['Building Name', 'Group', 'SubGroup', 'FTE', 'CW', 'Growth', 
                                                'HeadCount', 'Exception (Y/N)', 'Avg Occ with Exception', 'Avg Occ % of HeadCount',
                                                'Adjacency Priority 1', 'Adjacency Priority 2', 'Adjacency Priority 3', 
                                                'Current Location']]
                    
                    # Obter os Building Names únicos na ordem de aparição (após o sort)
                    unique_buildings = df_allocation['Building Name'].drop_duplicates().tolist()
                    # Definir cores alternadas: cinza claro (#D3D3D3) para índices pares e transparente para ímpares
                    building_colors = {building: "#D3D3D3" if i % 2 == 0 else "" 
                                    for i, building in enumerate(unique_buildings)}
                    
                    # Função para aplicar o estilo de fundo para cada linha, com base no Building Name
                    def highlight_building(row):
                        color = building_colors.get(row['Building Name'], '')
                        return ['background-color: ' + color] * len(row)
                    
                    # Aplica o estilo alternado nas linhas e formata a coluna "Avg Occ with Exception" com o fundo fixo
                    df_allocation_styled = (
                        df_allocation
                        .style.apply(highlight_building, axis=1)
                        .applymap(lambda x: 'background-color: #D3D3D3', subset=['Avg Occ with Exception', 'Avg Occ % of HeadCount'])
                    )
                    
                    st.dataframe(df_allocation_styled, use_container_width=False)

                    # Exibir a capacidade restante nos andares
                    st.write("#### Capacidade restante nos andares - Avg:")
                    remaining_floors_df = pd.DataFrame(
                        list(remaining_floors.items()), 
                        columns=['Building Name', 'Remaining Seats']
                    )
                    st.dataframe(remaining_floors_df, use_container_width=False)

                    return df_allocation, remaining_floors_df


                # Carregar os dados e realizar a alocação
                if "df_building_trat" in st.session_state and "df_proportional" in st.session_state:
                    df_building_trat = st.session_state.df_building_trat
                    df_proportional = st.session_state.df_proportional

                    # Exibir as tabelas para debug
                    #st.write("### Tabela 'Building Space Summary'")
                    #st.dataframe(df_building_trat, use_container_width=False)
                    
                    #st.write("### Tabela 'Grupos, SubGrupos e Adjacentes'")
                    #st.dataframe(df_proportional, use_container_width=False)

                    # Extração da capacidade dos andares do df_building_trat
                    floors = dict(zip(df_building_trat['Building Name'], df_building_trat['Primary Work Seats']))

                    # Chamar a função de alocação
                    df_allocation, remaining_floors = allocate_groups_avg(df_proportional, floors.copy())

                    # Exibir os resultados de alocação
                    df_allocation, remaining_floors_df = display_allocation(df_allocation, remaining_floors, df_building_trat)
                    cols = df_allocation.columns.tolist()
                    if "Building Name" in cols and "Current Location" in cols:
                        new_order = (
                            ["Building Name"] +
                            [col for col in cols if col not in ("Building Name", "Current Location")] +
                            ["Current Location"]
                        )
                        df_allocation = df_allocation[new_order]
                    
                    # Ordenar o DataFrame por "Building Name" se ainda não estiver ordenado
                    df_allocation = df_allocation.sort_values(by='Building Name')
                    dfautomation_avg = df_allocation.copy()
                    st.session_state.dfautomation_avg = dfautomation_avg  # Salvando no session_state

                    st.write("### Grupos Não Alocados:")
                    df_avg_nonallocated = df_allocation[df_allocation['Building Name'] == 'Não Alocado']
                    numeric_columns = df_avg_nonallocated.select_dtypes(include='number').columns
                    total_row = df_avg_nonallocated[numeric_columns].sum()
                    total_row['Group'] = 'Total' 
                    total_row_df = pd.DataFrame([total_row])
                    df_avg_nonallocated_total = pd.concat([df_avg_nonallocated, total_row_df], ignore_index=True)
                    st.dataframe(df_avg_nonallocated_total, use_container_width=False)

                # Botão para exportar tabela "Resultados das Simulações" para Excel
                if st.button("Exportar Tabela 'Resultados das Simulações' para Excel", key="export_unificado_avgocc"):
                    if "dfautomation_hc" in st.session_state:
                        # Acessa o DataFrame salvo no session_state e substitui NaN por string vazia
                        df_allocation_export = st.session_state.dfautomation_avg.fillna("")
                        
                        # Cria o arquivo Excel em memória
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            df_allocation_export.to_excel(writer, sheet_name="Simulações Avg OCC", index=False)
                        output.seek(0)
                        
                        # Botão de download, utilizando output.getvalue() para retornar os bytes do arquivo
                        st.download_button(
                            label="Download do Excel - Resultados das Simulações Avg OCC",
                            data=output.getvalue(),
                            file_name="resultados_simulacoes_avgocc.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("Data not found: 'dfautomation_hc' não está disponível no session_state.")
        else:
            st.write("Por favor, carregue o arquivo para prosseguir.") 




   ##### ABA CENÁRIOS #####
    with tabs[2]:
        st.write("### Cenários de Alocação")

        # Inicializar df_proportional como um DataFrame vazio, se não houver dados na sessão
        if "df_proportional" not in st.session_state:
            df_proportional = pd.DataFrame()  # DataFrame vazio
        else:
            df_proportional = st.session_state.df_proportional

        if not df_proportional.empty:
            # Criação da tabela de cenários para "Informações Cadastradas"
            df_proportional_cenarios = df_proportional.copy()
            df_proportional_cenarios = df_proportional_cenarios[[
                "Group", "SubGroup", "Exception (Y/N)", "HeadCount", "Proportional Peak", "Proportional Avg"
            ]]
            df_proportional_cenarios.rename(
                columns={
                    "HeadCount": "1:1", 
                    "Proportional Peak": "Peak", 
                    "Proportional Avg": "Avg Occ"
                },
                inplace=True
            )

            with st.expander("#### **Informações Cadastradas**"):
                # Selectbox dentro do expander para escolha da visualização
                view_option = st.selectbox(
                    "Selecione a visualização:",
                    options=["Meu cenário", "Automação HeadCount", "Automação Peak", "Automação Avg Occ"]
                )

                if view_option == "Meu cenário":
                    st.dataframe(df_proportional_cenarios, use_container_width=False, hide_index=True)
                elif view_option == "Automação HeadCount":
                    if "dfautomation_hc" in st.session_state:
                        st.dataframe(st.session_state.dfautomation_hc, use_container_width=False, hide_index=True)
                    else:
                        st.info("Tabela de Automação HeadCount não disponível.")
                elif view_option == "Automação Peak":
                    if "dfautomation_peak" in st.session_state:
                        st.dataframe(st.session_state.dfautomation_peak, use_container_width=False, hide_index=True)
                    else:
                        st.info("Tabela de Automação Peak não disponível.")
                elif view_option == "Automação Avg Occ":
                    if "dfautomation_avg" in st.session_state:
                        st.dataframe(st.session_state.dfautomation_avg, use_container_width=False, hide_index=True)
                    else:
                        st.info("Tabela de Automação Avg Occ não disponível.")
            


            # Criação da tabela de cenários
            df_proportional_cenarios = df_proportional.copy()
            df_proportional_cenarios = df_proportional_cenarios[["Group", "SubGroup", "Exception (Y/N)", "HeadCount", "Proportional Peak", "Proportional Avg"]]
            df_proportional_cenarios.rename(columns={"HeadCount": "1:1", "Proportional Peak": "Peak", "Proportional Avg": "Avg Occ"}, inplace=True)

            # Cálculo dos Lugares Ocupados 1:1 acumulado
            df_proportional_cenarios['Lugares Ocupados 1:1'] = df_proportional_cenarios['1:1'].cumsum()

            # Juntando as informações dos edifícios com as informações do cenário
            df_final = df_building_trat.merge(df_proportional_cenarios, how='cross')  # Merge sem chave para manter todos os dados em cruzamento

            # Adicionando a coluna de chave-valor para "Group" + "SubGroup"
            df_final['Chave'] = df_final.apply(lambda row: f"{row['Group']} - {row['SubGroup']}" if row['SubGroup'] else f"{row['Group']} - ", axis=1)

            # Agora, vamos calcular os 'Lugares Disponíveis 1:1' individualmente para cada andar
            df_final['Lugares Disponíveis 1:1'] = df_final.groupby('Building Name')['Primary Work Seats'].transform('first') - df_final['Lugares Ocupados 1:1']
            
            # Cálculos para Peak e Avg com exceção
            def calcular_lugares_ocupados(row, column_name, headcount_column):
                if row['Exception (Y/N)'] == 'Y':
                    return row[headcount_column]
                return row[column_name]

            # Cálculo para 'Lugares Ocupados Peak'
            df_final['Lugares Ocupados Peak'] = df_final.apply(lambda row: calcular_lugares_ocupados(row, 'Peak', '1:1'), axis=1)
            
            # Cálculo para 'Lugares Ocupados Avg'
            df_final['Lugares Ocupados Avg'] = df_final.apply(lambda row: calcular_lugares_ocupados(row, 'Avg Occ', '1:1'), axis=1)

            # Cálculos acumulados para Peak e Avg
            df_final['Lugares Ocupados Peak'] = df_final.groupby('Building Name')['Lugares Ocupados Peak'].cumsum()
            df_final['Lugares Ocupados Avg'] = df_final.groupby('Building Name')['Lugares Ocupados Avg'].cumsum()

            # Calculando 'Lugares Disponíveis Peak' e 'Lugares Disponíveis Avg'
            df_final['Lugares Disponíveis Peak'] = df_final.groupby('Building Name')['Primary Work Seats'].transform('first') - df_final['Lugares Ocupados Peak']
            df_final['Lugares Disponíveis Avg'] = df_final.groupby('Building Name')['Primary Work Seats'].transform('first') - df_final['Lugares Ocupados Avg']

            # Inicializando a lista de tabelas no session_state, caso não tenha sido inicializada
            if "tables_to_append_dict" not in st.session_state:
                st.session_state.tables_to_append_dict = {}
            if "final_consolidated_df" not in st.session_state:
                st.session_state.final_consolidated_df = pd.DataFrame()

            
            # Exibindo as informações com expanders para cada 'Building Name'
            for building in df_final['Building Name'].unique():
                with st.expander(f"#### **Informações do Andar: {building}**"):
                    st.write(f"**Informações do Andar: {building}**")
                    df_building_data = df_final[df_final['Building Name'] == building].copy()
                    primary_work_seats = df_building_data['Primary Work Seats'].iloc[0]
                    total_seats_on_floor = df_building_data['Total seats on floor'].iloc[0]
            
                    st.write(f"**Primary Work Seats**: {primary_work_seats}")
                    st.write(f"**Total seats on floor**: {total_seats_on_floor}")
            
                    # Cria a coluna de concatenação para filtro (não exibida na tabela)
                    df_building_data['Concat_G_SB_HC'] = (
                        df_building_data['Group'] + ' - ' +
                        df_building_data['SubGroup'].fillna('') + ' - ' +
                        df_building_data['1:1'].astype(str)
                    )
                    group_subgroup_options = df_building_data['Concat_G_SB_HC'].drop_duplicates().tolist()
            
                    # Chave exclusiva para as seleções desta seção
                    building_key = f"selected_options_{building}"
                    if building_key not in st.session_state:
                        st.session_state[building_key] = []  # Inicializa com lista vazia
            
                    # Calcula as opções já gravadas globalmente (de todas as seções)
                    global_recorded = set()
                    for key in st.session_state.keys():
                        if key.startswith("selected_options_"):
                            global_recorded.update(st.session_state[key])
            
                    # Se a seção já tiver sido gravada, usa a seleção gravada e desabilita o multiselect;
                    # caso contrário, as opções disponíveis são as que não foram gravadas em outras seções.
                    if st.session_state[building_key]:
                        available_options = st.session_state[building_key]
                        multiselect_disabled = True
                    else:
                        available_options = [opt for opt in group_subgroup_options if opt not in global_recorded]
                        multiselect_disabled = False
            
                    # Exibe o multiselect – inicialmente, todas as opções disponíveis são selecionadas
                    selected_options = st.multiselect(
                        "Selecione os Grupos e Subgrupos (incluindo 1:1)",
                        options=available_options,
                        default=available_options,
                        key=f"multiselect_{building}",
                        disabled=multiselect_disabled
                    )
            
                    # Filtra a tabela de acordo com a seleção feita
                    if selected_options:
                        df_building_data_filtered = df_building_data[df_building_data['Concat_G_SB_HC'].isin(selected_options)]
                    else:
                        df_building_data_filtered = df_building_data
            
                    # Cálculos dinâmicos
                    df_building_data_filtered['Lugares Ocupados 1:1'] = df_building_data_filtered['1:1'].cumsum()
                    df_building_data_filtered['Lugares Disponíveis 1:1'] = (
                        df_building_data_filtered.groupby('Building Name')['Primary Work Seats'].transform('first')
                        - df_building_data_filtered['Lugares Ocupados 1:1']
                    )
                    df_building_data_filtered['Lugares Ocupados Peak'] = df_building_data_filtered['Peak'].cumsum()
                    df_building_data_filtered['Lugares Disponíveis Peak'] = (
                        df_building_data_filtered.groupby('Building Name')['Primary Work Seats'].transform('first')
                        - df_building_data_filtered['Lugares Ocupados Peak']
                    )
                    df_building_data_filtered['Lugares Ocupados Avg'] = df_building_data_filtered['Avg Occ'].cumsum()
                    df_building_data_filtered['Lugares Disponíveis Avg'] = (
                        df_building_data_filtered.groupby('Building Name')['Primary Work Seats'].transform('first')
                        - df_building_data_filtered['Lugares Ocupados Avg']
                    )
            
                    # st.dataframe(df_building_data_filtered, use_container_width=True)
            
                    # Entrada para margem de Risk
                    risk_value = st.text_input(
                        f"Risk (numérico, sem '%') para {building}",
                        value="",
                        key=f"risk_input_{building}"
                    )
                    risk_value = int(risk_value) if risk_value else 0
            
                    # Cálculos relacionados ao Risk
                    df_building_data_filtered['Risk 1:1'] = df_building_data_filtered['1:1'] * (1 - risk_value / 100)
                    df_building_data_filtered['Saldo Risk 1:1'] = (
                        df_building_data_filtered['Primary Work Seats'] - df_building_data_filtered['Risk 1:1']
                    )
                    df_building_data_filtered['Risk 1:1'] = df_building_data_filtered['Risk 1:1'].round(0).astype(int)
                    df_building_data_filtered['Saldo Risk 1:1'] = df_building_data_filtered['Saldo Risk 1:1'].round(0).astype(int)
            
                    df_building_data_filtered['Risk Peak'] = df_building_data_filtered['Peak'] * (1 - risk_value / 100)
                    df_building_data_filtered['Saldo Risk Peak'] = (
                        df_building_data_filtered['Primary Work Seats'] - df_building_data_filtered['Risk Peak']
                    )
                    df_building_data_filtered['Risk Peak'] = df_building_data_filtered['Risk Peak'].round(0).astype(int)
                    df_building_data_filtered['Saldo Risk Peak'] = df_building_data_filtered['Saldo Risk Peak'].round(0).astype(int)
            
                    df_building_data_filtered['Risk Avg Occ'] = df_building_data_filtered['Avg Occ'] * (1 - risk_value / 100)
                    df_building_data_filtered['Saldo Risk Avg Occ'] = (
                        df_building_data_filtered['Primary Work Seats'] - df_building_data_filtered['Risk Avg Occ']
                    )
                    df_building_data_filtered['Risk Avg Occ'] = df_building_data_filtered['Risk Avg Occ'].round(0).astype(int)
                    df_building_data_filtered['Saldo Risk Avg Occ'] = df_building_data_filtered['Saldo Risk Avg Occ'].round(0).astype(int)
            
                    # Renomeia as colunas para exibição
                    df_building_data_filtered.rename(columns={
                        "Lugares Disponíveis 1:1": "Saldo 1:1",
                        "Lugares Ocupados 1:1": "Occupied 1:1",
                        "Lugares Ocupados Peak": "Occupied Peak",
                        "Lugares Disponíveis Peak": "Saldo Peak",
                        "Lugares Ocupados Avg": "Occupied Avg",
                        "Lugares Disponíveis Avg": "Saldo Avg Occ"
                    }, inplace=True)
            
                    columns_to_display_filter = [
                        'Building Name', 'Group', 'SubGroup', 'Exception (Y/N)',
                        '1:1', 'Saldo 1:1', 'Peak', 'Saldo Peak', 'Avg Occ', 'Saldo Avg Occ'
                    ]
                    columns_risk = ['Saldo Risk 1:1', 'Saldo Risk Peak', 'Saldo Risk Avg Occ']
                    if risk_value != 0:
                        columns_to_display_filter += columns_risk
            
                    def colorize(val):
                        if isinstance(val, (int, float)):
                            return 'background-color: white' if val >= 0 else 'background-color: #FFBDBD'
                        return 'background-color: white'
            
                    def colorize_risk(val):
                        if isinstance(val, (int, float)):
                            return 'background-color: #DDEBF7'
                        return 'background-color: white'
            
                    styled_df = df_building_data_filtered[columns_to_display_filter].style
                    if risk_value != 0:
                        styled_df = styled_df.applymap(colorize_risk, subset=columns_risk)
                    styled_df = styled_df.applymap(colorize, subset=['1:1', 'Saldo 1:1', 'Peak', 'Saldo Peak', 'Avg Occ', 'Saldo Avg Occ'])
                    st.dataframe(styled_df, use_container_width=True, hide_index=True)
            
                    # Botão para Gravar Dados nesta seção
                    if st.button(f"Gravar Dados para {building}"):
                        st.session_state[building_key] = selected_options  # Armazena a seleção desta seção
                        st.session_state.tables_to_append_dict[building] = styled_df.data.copy()
                        st.success(f"Dados do prédio **{building}** gravados com sucesso!")
            
                    # Botão para Resetar a Seção (apenas esta seção é resetada)
                    if st.button(f"Resetar Seção para {building}"):
                        st.session_state[building_key] = []  # Limpa as seleções desta seção
                        try:
                            st.experimental_rerun()
                        except Exception:
                            st.info("Por favor, clique novamente em 'Resetar Seção' para ver as alterações na seção.")


            
            with st.expander("### **Resultado de todos os Cenários:**"):
                if "tables_to_append_dict" in st.session_state and st.session_state.tables_to_append_dict:
                    st.write("### **Dados Gravados**")
                    # Concatena todos os DataFrames armazenados
                    final_consolidated_df = pd.concat(
                        st.session_state.tables_to_append_dict.values(), ignore_index=True
                    )
            
                    # Função para arredondar e tratar valores numéricos
                    def round_and_convert_to_int(df):
                        numeric_columns = df.select_dtypes(include=['number']).columns
                        df[numeric_columns] = df[numeric_columns].replace([np.inf, -np.inf, np.nan], 0)
                        df[numeric_columns] = df[numeric_columns].round(0).astype(int)
                        return df
            
                    final_consolidated_df = round_and_convert_to_int(final_consolidated_df)
            
                    # Aplica cor de destaque na tabela consolidada
                    def colorize(val):
                        if isinstance(val, (int, float)):
                            return 'background-color: white' if val >= 0 else 'background-color: #FFBDBD'
                        return 'background-color: white'
            
                    final_consolidated_df_colour = final_consolidated_df.style.applymap(
                        colorize, subset=['1:1', 'Saldo 1:1', 'Peak', 'Saldo Peak', 'Avg Occ', 'Saldo Avg Occ']
                    )
            
                    st.write("#### **Consolidado de todos os cenários:**")
                    st.dataframe(final_consolidated_df_colour, use_container_width=True, hide_index=True)
            
                    # Cria a chave de identificação para grupos e subgrupos no consolidado
                    final_consolidated_df["Chave"] = final_consolidated_df.apply(
                        lambda row: f"{row['Group']} - {row['SubGroup']}" if row['SubGroup'] else f"{row['Group']} - ",
                        axis=1
                    )
            
                    # Cria a chave de identificação em df_proportional_cenarios
                    proportional_groups_subgroups = df_proportional_cenarios.copy()
                    proportional_groups_subgroups["Chave"] = proportional_groups_subgroups.apply(
                        lambda row: f"{row['Group']} - {row['SubGroup']}" if row['SubGroup'] else f"{row['Group']} - ",
                        axis=1
                    )
            
                    consolidated_groups_subgroups = final_consolidated_df[['Chave']].drop_duplicates()
            
                    # Encontra os grupos/subgrupos não alocados
                    df_non_allocated = proportional_groups_subgroups.merge(
                        consolidated_groups_subgroups, on="Chave", how="left", indicator=True
                    ).query('_merge == "left_only"').drop('_merge', axis=1)
            
                    st.write("#### **Grupos e Subgrupos Não Alocados**")
                    st.dataframe(df_non_allocated, use_container_width=True, hide_index=True)
            
                    # Armazena o DataFrame consolidado no session_state
                    st.session_state["final_consolidated_df"] = final_consolidated_df
            
                    if st.button("Exportar 'Cenários' para Excel", key="export_cenarios_excel"):
                        with io.BytesIO() as output:
                            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                                if "tables_to_append_dict" in st.session_state and st.session_state.tables_to_append_dict:
                                    final_consolidated_df = pd.concat(
                                        st.session_state.tables_to_append_dict.values(), ignore_index=True
                                    )
                                else:
                                    final_consolidated_df = pd.DataFrame()
                                if "df_non_allocated" in st.session_state:
                                    df_non_allocated = st.session_state.df_non_allocated.copy()
                                else:
                                    df_non_allocated = pd.DataFrame()
                                final_consolidated_df = final_consolidated_df.fillna("")
                                df_non_allocated = df_non_allocated.fillna("")
                                final_consolidated_df.to_excel(writer, sheet_name="Cenarios", index=False)
                                df_non_allocated.to_excel(writer, sheet_name="Não Alocados", index=False)
                            output.seek(0)
                            st.download_button(
                                label="Download do Excel - Resultados dos Cenários",
                                data=output.getvalue(),
                                file_name="resultados_cenarios.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                else:
                    st.write("Nenhum dado foi gravado ainda.")




   ##### ABA DASHBOARDS #####
    with tabs[3]:
        st.write("### DASHBOARDS")

        # Inicializar dfautomation_hc como um DataFrame vazio, se não houver dados na sessão
        if "dfautomation_hc" not in st.session_state:
            dfautomation_hc = pd.DataFrame()  # DataFrame vazio
            final_consolidated_df = pd.DataFrame()
            styled_df = pd.DataFrame()
            dfautomation_peak_dash = pd.DataFrame() 
            dfautomation_avg_dash = pd.DataFrame() 
            df_building_trat = pd.DataFrame()

        else:
            dfautomation_hc = st.session_state.dfautomation_hc
            # Se necessário, acesse outros DataFrames salvos no session_state
            final_consolidated_df = st.session_state.get('final_consolidated_df', pd.DataFrame())
            styled_df = st.session_state.get('styled_df', pd.DataFrame())
            dfautomation_peak_dash = st.session_state.get('dfautomation_peak_dash', pd.DataFrame())
            dfautomation_avg_dash = st.session_state.get('dfautomation_avg_dash', pd.DataFrame())
            df_building_dash = st.session_state.get('df_building_trat', pd.DataFrame())

        # Verificar se os DataFrames têm dados antes de continuar
        if not dfautomation_hc.empty:
            # Criação das tabelas de cenários
            dfautomation_hc_dash = dfautomation_hc.copy()
            dfautomation_peak_dash = dfautomation_peak.copy()
            dfautomation_avg_dash = dfautomation_avg.copy()
            df_building_dash = df_building_trat.copy()


            # Adicionar um seletor para o usuário escolher qual tabela exibir
            visao_selecionada = st.selectbox(
                'Escolha a tabela para visualizar:',
                ('Cenarios', 'Automações')
            )

            # Exibir a tabela com base na escolha do usuário
            if visao_selecionada == 'Automações':
            
                
            # Merge para unir informações de ocupação por HEADCOUNT
                df_merged_hc = pd.merge(dfautomation_hc_dash, df_building_dash[['Building Name', 'Primary Work Seats', 'Total seats on floor']], on='Building Name', how='left')
                df_merged_sorted_hc = df_merged_hc.sort_values(by=['Building Name', 'Group', 'SubGroup'])
                df_merged_sorted_hc['CumSum HeadCount'] = df_merged_sorted_hc.groupby('Building Name')['HeadCount'].cumsum()
                df_merged_sorted_hc['AvailableCumSum'] = df_merged_sorted_hc['Primary Work Seats'] - df_merged_sorted_hc['CumSum HeadCount']

                # Lugares Disponíveis por Andar
                df_last_cumsum_hc = df_merged_sorted_hc.groupby('Building Name').last().reset_index()
                df_last_cumsum_hc['Avail Total Seats HC'] = df_last_cumsum_hc['Total seats on floor'] - df_last_cumsum_hc['CumSum HeadCount']
                df_last_cumsum_hc['Avail Primary HC'] = df_last_cumsum_hc['Primary Work Seats'] - df_last_cumsum_hc['CumSum HeadCount']
                df_availability_hc = df_last_cumsum_hc[['Building Name', 'CumSum HeadCount', 'Primary Work Seats', 'Avail Primary HC', 'Total seats on floor',  'Avail Total Seats HC']]
                total_row_hc = df_availability_hc[['CumSum HeadCount', 'Primary Work Seats', 'Avail Primary HC', 'Total seats on floor',  'Avail Total Seats HC']].sum()
                total_row_hc['Building Name'] = 'Total'  
                df_avail_row_hc = pd.DataFrame([total_row_hc])
                df_avail_hc = pd.concat([df_availability_hc, df_avail_row_hc], ignore_index=True)
                df_avail_hc.rename(columns={"CumSum HeadCoun" : "Total HC"})

            #### BIG NUMBERS - CABEÇALHO PAINEL
                # Calculando as métricas
                # 1) Qtde de Buildings
                num_buildings = df_building_dash['Building Name'].nunique()

                # 2) Qtde de Groups
                num_groups = df_merged_hc['Group'].nunique()

                # 3) Qtde de Groups + SubGroups
                num_groups_subgroups = df_merged_hc[['Group', 'SubGroup']].drop_duplicates().shape[0]

                # 4) Total de Lugares Disponíveis (Primary Work Seats)
                total_primary_seats = df_building_dash['Primary Work Seats'].sum()

                # 5) Total de Lugares Disponíveis (Total Seats on Floor)
                total_floor_seats = df_building_dash['Total seats on floor'].sum()


                # Organizando os Big Numbers em colunas
                col1, col2, col3, col4, col5  = st.columns(5)  

                with col1:
                    # Título com fonte e alinhamento
                    st.markdown(f"<h4 style='text-align: center; background-color: #4682B4; padding: 10px; font-size: 16px; color: white; '> Qtde de Andares</h4>", unsafe_allow_html=True)                
                    # Número com fonte personalizada e centralizado
                    st.markdown(f"<h1 style='text-align: center; font-size: 28px; font-family: Arial, sans-serif; font-weight: bold;'>{num_buildings}</h1>", unsafe_allow_html=True)
                with col2:
                    st.markdown(f"<h4 style='text-align: center; background-color: #ADD8E6; padding: 10px; font-size: 16px;'> Qtde de Grupos</h4>", unsafe_allow_html=True) 
                    st.markdown(f"<h1 style='text-align: center; font-size: 28px; font-family: Arial, sans-serif; font-weight: bold;'>{num_groups}</h1>", unsafe_allow_html=True)
                with col3:
                    st.markdown(f"<h4 style='text-align: center; background-color: #4682B4; padding: 10px; font-size: 14px; color: white; '> Qtde de Grupos + SubGrupos</h4>", unsafe_allow_html=True)                
                    st.markdown(f"<h1 style='text-align: center; font-size: 28px; font-family: Arial, sans-serif; font-weight: bold;'>{num_groups_subgroups}</h1>", unsafe_allow_html=True)
                with col4:
                    st.markdown(f"<h4 style='text-align: center; background-color: #ADD8E6; padding: 10px; font-size: 14px; color: white; '> Total Primary Work Seats</h4>", unsafe_allow_html=True)                
                    st.markdown(f"<h1 style='text-align: center; font-size: 28px; font-family: Arial, sans-serif; font-weight: bold;'>{total_primary_seats}</h1>", unsafe_allow_html=True)
                with col5:
                    st.markdown(f"<h4 style='text-align: center; background-color: #4682B4; padding: 10px; font-size: 14px; color: white; '> Total Seats</h4>", unsafe_allow_html=True)                
                    st.markdown(f"<h1 style='text-align: center; font-size: 28px; font-family: Arial, sans-serif; font-weight: bold;'>{total_floor_seats}</h1>", unsafe_allow_html=True)



                # Adiciona um espaço maior usando <br> no Markdown
                st.markdown("<br><br><br>", unsafe_allow_html=True)  # Adiciona 3 quebras de linha
                


            #### BIG NUMBERS - CABEÇALHO DONUTS
                #### 1. HEADCOUNT
                total_headcount = df_merged_sorted_hc['HeadCount'].sum()

                # 1.1) % Alocados HEADCOUNT
                allocated_headcount = df_merged_sorted_hc[df_merged_sorted_hc['Building Name'] != 'Não Alocado']['HeadCount'].sum()
                percent_allocated_hc = (allocated_headcount / total_headcount) * 100 if total_headcount > 0 else 0
                percent_allocated_hc = percent_allocated_hc.round(0).astype(int)

                # 1.2) Qtde Não Alocados HEADCOUNT
                non_allocated_headcount = df_merged_sorted_hc[df_merged_sorted_hc['Building Name'] == 'Não Alocado']['HeadCount'].sum()
                non_allocated_groups_hc = df_merged_sorted_hc[df_merged_sorted_hc['Building Name'] == 'Não Alocado']
                non_allocated_groups_hc = non_allocated_groups_hc[["Group","SubGroup","HeadCount"]].sort_values(by=["Group","SubGroup"])

                # 1.3 - GRÁFICO DE DONUT - HEADCOUNT
                df_merged_sorted_hc['Status'] = df_merged_sorted_hc['Building Name'].apply(lambda x: 'Alocado' if x != 'Não Alocado' else 'Não Alocado')
                # Agrupando o HeadCount por Status
                headcount_by_status = df_merged_sorted_hc.groupby('Status').agg({'HeadCount': 'sum'}).reset_index()
                # Criando o gráfico de Donut com Plotly
                fighc = px.pie(headcount_by_status, 
                            names='Status', 
                            values='HeadCount', 
                            title='',
                            hole=0.3,  # Cria o efeito de donut
                            color='Status',
                            color_discrete_map={"Alocado": "#00CFFF", "Não Alocado": "#FFAA33"},  # Definindo as cores
                            labels={"HeadCount": "Total HeadCount"})  # Renomeia o label no gráfico
                # Adicionando os valores absolutos e percentuais como rótulos
                fighc.update_traces(textinfo='percent+label', pull=[0.1, 0.1])  # Exibindo percentagem e label
                fighc.update_layout(
                    title='',  # Título do gráfico
                    title_x=0,  # Alinha o título à esquerda
                    title_xanchor='left',  # Alinha o título à esquerda
                    title_font=dict(size=16, color="black", family="Arial"), 
                    legend_title="Status",
                    legend=dict(
                        x=1.05,  # Posiciona a legenda à direita
                        y=0.5,   # Ajuste vertical para que a legenda não sobreponha o gráfico
                        traceorder="normal",  # Ordem de exibição das legendas
                        orientation="v",  # Define a legenda na vertical
                        title="Status",
                    ),
                    margin=dict(t=50, b=50, l=50, r=50),  # Ajustando as margens do gráfico
                    plot_bgcolor="white",  # Cor de fundo do gráfico
                    paper_bgcolor="white",  # Cor de fundo da área do gráfico
                    width=450,  # Largura reduzida do gráfico
                    height=450,  # Altura reduzida do gráfico
                    )

                
                #### 2 - AVG PEAK            
                # Merge para unir informações de ocupação por AVG PEAK
                df_merged_peak = pd.merge(dfautomation_peak_dash, df_building_dash[['Building Name', 'Primary Work Seats', 'Total seats on floor']], on='Building Name', how='left')
                df_merged_sorted_peak = df_merged_peak.sort_values(by=['Building Name', 'Group', 'SubGroup'])
                
                # Calcular o "Proportional Peak Exception" (Exception = Y) diretamente no backend
                df_merged_sorted_peak['CumSum Peak_Exc'] = df_merged_sorted_peak.groupby('Building Name')['Peak with Exception'].cumsum()
                df_merged_sorted_peak['AvailableCumSum'] = df_merged_sorted_peak['Primary Work Seats'] - df_merged_sorted_peak['CumSum Peak_Exc']

                # Lugares Disponíveis por Andar
                df_last_cumsum_peak = df_merged_sorted_peak.groupby('Building Name').last().reset_index()
                df_last_cumsum_peak['Avail Total Seats Peak'] = df_last_cumsum_peak['Total seats on floor'] - df_last_cumsum_peak['CumSum Peak_Exc']
                df_last_cumsum_peak['Avail Primary Peak'] = df_last_cumsum_peak['Primary Work Seats'] - df_last_cumsum_peak['CumSum Peak_Exc']
                df_availability_peak = df_last_cumsum_peak[['Building Name', 'CumSum Peak_Exc', 'Primary Work Seats', 'Avail Primary Peak', 'Total seats on floor', 'Avail Total Seats Peak']]
                total_row_peak = df_availability_peak[['CumSum Peak_Exc', 'Primary Work Seats', 'Avail Primary Peak', 'Total seats on floor', 'Avail Total Seats Peak']].sum()
                total_row_peak['Building Name'] = 'Total'  
                df_avail_row_peak = pd.DataFrame([total_row_peak])
                df_avail_peak = pd.concat([df_availability_peak, df_avail_row_peak], ignore_index=True)
                df_avail_peak.rename(columns={"CumSum Peak_Exc" : "Total Peak"}, inplace=True)

                total_avgpeak = df_merged_sorted_peak['Peak with Exception'].sum()

                # 2.1) % Alocados AVG PEAK
                allocated_peak = df_merged_sorted_peak[df_merged_sorted_peak['Building Name'] != 'Não Alocado']['Peak with Exception'].sum()
                percent_allocated_peak = (allocated_peak / total_avgpeak) * 100 if total_avgpeak > 0 else 0
                percent_allocated_peak = percent_allocated_peak.round(0).astype(int)

                # 2.2) Qtde Não Alocados AVG PEAK
                non_allocated_peak = df_merged_sorted_peak[df_merged_sorted_peak['Building Name'] == 'Não Alocado']['Peak with Exception'].sum()
                non_allocated_groups_peak = df_merged_sorted_peak[df_merged_sorted_peak['Building Name'] == 'Não Alocado']
                non_allocated_groups_peak = non_allocated_groups_peak[["Group", "SubGroup", "Peak with Exception"]].sort_values(by=["Group","SubGroup"])

                # 2.3 - GRÁFICO DE DONUT - AVG PEAK
                df_merged_sorted_peak['Status'] = df_merged_sorted_peak['Building Name'].apply(lambda x: 'Alocado' if x != 'Não Alocado' else 'Não Alocado')
                # Agrupando o HeadCount por Status
                peak_by_status = df_merged_sorted_peak.groupby('Status').agg({'Peak with Exception': 'sum'}).reset_index()
                # Criando o gráfico de Donut com Plotly
                figpeak = px.pie(peak_by_status, 
                            names='Status', 
                            values='Peak with Exception', 
                            title='',
                            hole=0.3,  # Cria o efeito de donut
                            color='Status',
                            color_discrete_map={"Alocado": "#00CFFF", "Não Alocado": "#FFAA33"},  # Definindo as cores
                            labels={"Peak with Exception": "Total Peak Exc"})  # Renomeia o label no gráfico
                # Adicionando os valores absolutos e percentuais como rótulos
                figpeak.update_traces(textinfo='percent+label', pull=[0.1, 0.1])  # Exibindo percentagem e label
                figpeak.update_layout(
                    title='',  # Título do gráfico
                    title_x=0,  # Alinha o título à esquerda
                    title_xanchor='left',  # Alinha o título à esquerda
                    title_font=dict(size=16, color="black", family="Arial"), 
                    legend_title="Status",
                    legend=dict(
                        x=1.05,  # Posiciona a legenda à direita
                        y=0.5,   # Ajuste vertical para que a legenda não sobreponha o gráfico
                        traceorder="normal",  # Ordem de exibição das legendas
                        orientation="v",  # Define a legenda na vertical
                        title="Status",
                    ),
                    margin=dict(t=50, b=50, l=50, r=50),  # Ajustando as margens do gráfico
                    plot_bgcolor="white",  # Cor de fundo do gráfico
                    paper_bgcolor="white",  # Cor de fundo da área do gráfico
                    width=450,  # Largura reduzida do gráfico
                    height=450,  # Altura reduzida do gráfico
                    )
                

                #### 3 - AVG OCC
                # Merge para unir informações de ocupação por AVG OCC
                df_merged_avg = pd.merge(dfautomation_avg_dash, df_building_dash[['Building Name', 'Primary Work Seats','Total seats on floor']], on='Building Name', how='left')
                df_merged_sorted_avg = df_merged_avg.sort_values(by=['Building Name', 'Group', 'SubGroup'])
                
                # Calcular o "Proportional Peak Exception" (Exception = Y) diretamente no backend
                df_merged_sorted_avg['CumSum Avg_Exc'] = df_merged_sorted_avg.groupby('Building Name')['Avg Occ with Exception'].cumsum()
                df_merged_sorted_avg['AvailableCumSum'] = df_merged_sorted_avg['Primary Work Seats'] - df_merged_sorted_avg['CumSum Avg_Exc']

                # Lugares Disponíveis por Andar
                df_last_cumsum_avgocc = df_merged_sorted_avg.groupby('Building Name').last().reset_index()
                df_last_cumsum_avgocc['Avail Total Seats AvgOcc'] = df_last_cumsum_avgocc['Total seats on floor'] - df_last_cumsum_avgocc['CumSum Avg_Exc']
                df_last_cumsum_avgocc['Avail Primary AvgOcc'] = df_last_cumsum_avgocc['Primary Work Seats'] - df_last_cumsum_avgocc['CumSum Avg_Exc']
                df_availability_avgocc = df_last_cumsum_avgocc[['Building Name', 'CumSum Avg_Exc', 'Primary Work Seats', 'Avail Primary AvgOcc', 'Total seats on floor', 'Avail Total Seats AvgOcc']]
                total_row_avgocc = df_availability_avgocc[['CumSum Avg_Exc', 'Primary Work Seats', 'Avail Primary AvgOcc', 'Total seats on floor', 'Avail Total Seats AvgOcc']].sum()
                total_row_avgocc['Building Name'] = 'Total'  
                df_avail_row_avgocc = pd.DataFrame([total_row_avgocc])
                df_avail_avgocc = pd.concat([df_availability_avgocc, df_avail_row_avgocc], ignore_index=True)
                df_avail_avgocc.rename(columns={"CumSum Avg_Exc" : "Total Avg Occ"}, inplace=True)
                total_avgocc = df_merged_sorted_avg['Avg Occ with Exception'].sum()

                # 3.1) % Alocados AVG OCC
                allocated_avgocc = df_merged_sorted_avg[df_merged_sorted_avg['Building Name'] != 'Não Alocado']['Avg Occ with Exception'].sum()
                percent_allocated_avgocc = (allocated_avgocc / total_avgocc) * 100 if total_avgpeak > 0 else 0
                percent_allocated_avgocc = percent_allocated_avgocc.round(0).astype(int)

                # 3.2) Qtde Não Alocados AVG OCC
                non_allocated_avgocc = df_merged_sorted_avg[df_merged_sorted_avg['Building Name'] == 'Não Alocado']['Avg Occ with Exception'].sum()
                non_aloccated_groups_avgocc = df_merged_sorted_avg[df_merged_sorted_avg['Building Name'] == 'Não Alocado']
                non_aloccated_groups_avgocc = non_aloccated_groups_avgocc[["Group", "SubGroup", "Avg Occ with Exception"]].sort_values(by=["Group", "SubGroup"])


                # 3.3 - GRÁFICO DE DONUT - AVG OCC
                df_merged_sorted_avg['Status'] = df_merged_sorted_avg['Building Name'].apply(lambda x: 'Alocado' if x != 'Não Alocado' else 'Não Alocado')
                # Agrupando o HeadCount por Status
                avgocc_by_status = df_merged_sorted_avg.groupby('Status').agg({'Avg Occ with Exception': 'sum'}).reset_index()
                # Criando o gráfico de Donut com Plotly
                figavgocc = px.pie(avgocc_by_status, 
                            names='Status', 
                            values='Avg Occ with Exception', 
                            title='',
                            hole=0.3,  # Cria o efeito de donut
                            color='Status',
                            color_discrete_map={"Alocado": "#00CFFF", "Não Alocado": "#FFAA33"},  # Definindo as cores
                            labels={"Avg Occ with Exception": "Total Avg Exc"})  # Renomeia o label no gráfico
                # Adicionando os valores absolutos e percentuais como rótulos
                figavgocc.update_traces(textinfo='percent+label', pull=[0.1, 0.1])  # Exibindo percentagem e label
                figavgocc.update_layout(
                    title='',  # Título do gráfico
                    title_x=0,  # Alinha o título à esquerda
                    title_xanchor='left',  # Alinha o título à esquerda
                    title_font=dict(size=16, color="black", family="Arial"), 
                    legend_title="Status",
                    legend=dict(
                        x=1.05,  # Posiciona a legenda à direita
                        y=0.5,   # Ajuste vertical para que a legenda não sobreponha o gráfico
                        traceorder="normal",  # Ordem de exibição das legendas
                        orientation="v",  # Define a legenda na vertical
                        title="Status",
                    ),
                    margin=dict(t=50, b=50, l=50, r=50),  # Ajustando as margens do gráfico
                    plot_bgcolor="white",  # Cor de fundo do gráfico
                    paper_bgcolor="white",  # Cor de fundo da área do gráfico
                    width=450,  # Largura reduzida do gráfico
                    height=450,  # Altura reduzida do gráfico
                    )
                
            #### EXIBIÇÃO 
                col5, col6, col7,   = st.columns(3)

                with col5:
                    st.markdown(f"<h4 style='text-align: center; background-color: #707070; padding: 10px; font-size: 16px; color: white; '> Automação HeadCount</h4>", unsafe_allow_html=True)
                with col6:
                    st.markdown(f"<h4 style='text-align: center; background-color: #B0B0B0; padding: 10px; font-size: 16px; color: white; '> Automação Avg Peak</h4>", unsafe_allow_html=True) 
                with col7:
                    st.markdown(f"<h4 style='text-align: center; background-color: #707070; padding: 10px; font-size: 16px; color: white; '>  Automação Avg Occ</h4>", unsafe_allow_html=True) 


                # Adiciona um espaço maior usando <br> no Markdown
                st.markdown("<br>", unsafe_allow_html=True)  # Adiciona 3 quebras de linha

               
                col8, col9, col10,   = st.columns(3)
                with col8:
                    st.dataframe(df_avail_hc, use_container_width=False, hide_index=True)
                    st.plotly_chart(fighc)        
                    st.write("Groups e SubGroups não alocados")            
                    st.dataframe(non_allocated_groups_hc, use_container_width=False,hide_index=True)

                with col9:
                    st.dataframe(df_avail_peak, use_container_width=False, hide_index=True)
                    st.plotly_chart(figpeak)
                    st.write("Groups e SubGroups não alocados") 
                    st.dataframe(non_allocated_groups_peak, use_container_width=False, hide_index=True)

                with col10:
                    st.dataframe(df_avail_avgocc, use_container_width=False, hide_index=True)
                    st.plotly_chart(figavgocc)
                    st.write("Groups e SubGroups não alocados") 
                    st.dataframe(non_aloccated_groups_avgocc, use_container_width=False, hide_index=True)





                # Adiciona um espaço maior usando <br> no Markdown
                st.markdown("<br><br>", unsafe_allow_html=True)  # Adiciona 2 quebras de linha






            #### DROPBOX PARA DEEP DIVE
                # Adicionar um seletor para o usuário escolher qual tabela exibir
                tabela_selecionada = st.selectbox(
                    'Escolha a tabela para visualizar:',
                    ('Automação HeadCount', 'Automação Peak', 'Automação AvgOcc' )
                )


                #### HEADCOUNT
                if tabela_selecionada == "Automação HeadCount":

                #### GRÁFICO DE AVG E PEAK
                    #### GRÁFICO DE HEADCOUNT, PEAK e AVG OCC COM EXCEÇÃO

                    # 1. Criar as colunas com exceção no DataFrame original (caso ainda não existam)
                    df_merged_sorted_hc['Peak with Exception'] = df_merged_sorted_hc.apply(
                        lambda row: row['HeadCount'] if row['Exception (Y/N)'] == 'Y' else row['Proportional Peak'],
                        axis=1
                    )
                    df_merged_sorted_hc['Avg Occ with Exception'] = df_merged_sorted_hc.apply(
                        lambda row: row['HeadCount'] if row['Exception (Y/N)'] == 'Y' else row['Proportional Avg'],
                        axis=1
                    )

                    # 2. Agrupar por 'Group' somando os valores
                    df_grouped = df_merged_sorted_hc.groupby('Group').agg({
                        'HeadCount': 'sum',
                        'Peak with Exception': 'sum',
                        'Avg Occ with Exception': 'sum'
                    }).reset_index()
                    df_grouped = df_grouped.sort_values(by='HeadCount', ascending=False)


                    # 3. Reformular para o formato "long", para facilitar a criação do gráfico de barras
                    df_melted = df_grouped.melt(
                        id_vars=['Group'], 
                        value_vars=['HeadCount', 'Peak with Exception', 'Avg Occ with Exception'],
                        var_name='Metric', 
                        value_name='Value'
                    )

                    # 4. (Opcional) Criar coluna de texto para exibir os valores dentro das barras
                    df_melted['text'] = df_melted.apply(
                        lambda row: f"<b>{row['Metric']}:</b> {row['Value']}" if row['Value'] > 0 else "", 
                        axis=1
                    )

                    # 5. Criar o gráfico de barras com Plotly Express
                    fig = px.bar(
                        df_melted, 
                        x="Group", 
                        y="Value", 
                        color="Metric",               # Diferencia as barras pelo tipo de métrica
                        title="Distribuição de HeadCount, Peak e Avg Occ por Group",
                        labels={"Value": "Total", "Group": "Group", "Metric": "Métrica"},
                        text="text"                   # Exibe o texto configurado para cada barra
                    )

                    # 6. Ajustar o layout do gráfico
                    fig.update_layout(
                        barmode='group',             # Barras agrupadas
                        xaxis_title="Group",
                        yaxis_title="Total",
                        legend_title="Métrica",
                        margin=dict(t=50, b=50, l=50, r=50),
                        plot_bgcolor="white",
                        paper_bgcolor="white",
                        width=800,
                        height=450,
                    )

                    # 7. Exibir os valores dentro das barras
                    fig.update_traces(texttemplate='%{text}', textposition='inside')

                    # 8. Exibir o gráfico no Streamlit
                    st.plotly_chart(fig)
                    st.markdown("<br><br>", unsafe_allow_html=True)



                #### DATAFRAME HEADCOUNT + DONUT POR GROUP     
                    colhc1, colhc2 = st.columns(2)
                    with colhc1:
                        # Agrupando os dados e somando o HeadCount por Group
                        df_dash_hc = df_merged_sorted_hc[["Group", "SubGroup", "HeadCount","Proportional Peak", "Proportional Avg"]].copy()
                        df_grouped_hc = df_dash_hc.groupby('Group', as_index=False)['HeadCount'].sum()
                        figdonuthc = px.pie(df_grouped_hc, 
                                    names='Group', 
                                    values='HeadCount', 
                                    hole=0.3,  # Faz o gráfico ficar no formato de rosca
                                    title="Distribuição de HeadCount por Group")

                        # Exibindo o gráfico no Streamlit
                        st.plotly_chart(figdonuthc)

                    with colhc2:
                        df_dash_hc.rename(columns={"Proportional Peak" : "Peak Seats Required", "Proportional Avg" : "Avg Seats Required"}, inplace=True)
                        st.dataframe(df_dash_hc, use_container_width=False, hide_index=True)

                    
                    # Adiciona um espaço maior usando <br> no Markdown
                    st.markdown("<br><br>", unsafe_allow_html=True)  # Adiciona 2 quebras de linha
                        

                    #### GRÁFICO DE BARRAS - DISTRIBUIÇÃO DE HEADCOUNT POR ANDAR
                    colalochc1, colalochc2 = st.columns(2)

                    with colalochc1:
                        # Criando o gráfico de barras
                        df_merged_sorted_hc_bars = df_merged_sorted_hc.groupby(['Building Name', 'Group'], as_index=False)['HeadCount'].sum()
                        barshc = px.bar(df_merged_sorted_hc_bars, 
                                        x="Group", 
                                        y="HeadCount", 
                                        color="Building Name", 
                                        title="Distribuição de HeadCount por Alocação", 
                                        labels={"HeadCount": "Total HeadCount"},
                                        hover_data=["Building Name"],  # Exibe SubGroup ao passar o mouse
                                        category_orders={"Building Name": sorted(df_merged_sorted_hc['Building Name'].unique())})  # Exibe todos os Buildings disponíveis

                        # Adicionando os valores nas barras com texto
                        barshc.update_traces(text=df_merged_sorted_hc_bars['HeadCount'], textposition='inside', texttemplate='%{text}')

                        # Habilitando interatividade no gráfico
                        barshc.update_layout(
                            barmode='stack',  # Empilha as barras por SubGroup
                            xaxis_title="Group",
                            yaxis_title="Total HeadCount",
                            legend_title="Building Name",
                            legend=dict(
                                x=1.05,  # Posiciona a legenda à direita
                                y=0.5,   # Ajuste vertical para que a legenda não sobreponha o gráfico
                                traceorder="normal",  # Ordem de exibição das legendas
                                orientation="v",  # Define a legenda na vertical
                                title="Group"
                            ),
                            margin=dict(t=50, b=50, l=50, r=50),  # Ajustando as margens do gráfico
                            plot_bgcolor="white",  # Cor de fundo do gráfico
                            paper_bgcolor="white",  # Cor de fundo da área do gráfico
                            width=800,  # Largura reduzida do gráfico
                            height=550,  # Altura reduzida do gráfico
                        )

                        # Exibindo o gráfico no Streamlit
                        st.plotly_chart(barshc)


                    with colalochc2:
                        df_dash_aloc_hc = df_merged_sorted_hc[["Group", "SubGroup", "HeadCount","Proportional Peak", "Proportional Avg", "Building Name"]].copy()
                        df_dash_aloc_hc.rename(columns={"Proportional Peak" : "Peak Seats Required", "Proportional Avg" : "Avg Seats Required"}, inplace=True)
                        st.dataframe(df_dash_aloc_hc, use_container_width=False, hide_index=True)



            if visao_selecionada == 'Cenarios':
                if "final_consolidated_df" in st.session_state and not st.session_state.final_consolidated_df.empty:
                    final_consolidated_drop = st.session_state.final_consolidated_df.copy()
                    colunas_para_remover = ['Origem', 'Chave']
                    final_consolidated_dash = final_consolidated_drop.drop(columns=colunas_para_remover, errors='ignore')
                    st.session_state.final_consolidated_dash = final_consolidated_drop


                    # Merge para unir informações de ocupação por HEADCOUNT
                    df_merged_cenarios = pd.merge(final_consolidated_dash, df_building_dash[['Building Name', 'Primary Work Seats', 'Total seats on floor']], on='Building Name', how='left')
                    df_merged_cenarios = df_merged_cenarios.sort_values(by=['Building Name', 'Group', 'SubGroup'])
                    df_merged_cenarios['CumSum HeadCount'] = df_merged_cenarios.groupby('Building Name')['1:1'].cumsum()
                    df_merged_cenarios['AvailableCumSum'] = df_merged_cenarios['Primary Work Seats'] - df_merged_cenarios['CumSum HeadCount']

                    # Lugares Disponíveis por Andar
                    df_last_cumsum_cenarios = df_merged_cenarios.groupby('Building Name').last().reset_index()
                    df_last_cumsum_cenarios['Avail Total Seats HC'] = df_last_cumsum_cenarios['Total seats on floor'] - df_last_cumsum_cenarios['CumSum HeadCount']
                    df_last_cumsum_cenarios['Avail Primary HC'] = df_last_cumsum_cenarios['Primary Work Seats'] - df_last_cumsum_cenarios['CumSum HeadCount']
                    df_availability_cenarios = df_last_cumsum_cenarios[['Building Name', 'CumSum HeadCount', 'Primary Work Seats', 'Avail Primary HC', 'Total seats on floor',  'Avail Total Seats HC']]
                    total_row_cenarios = df_availability_cenarios[['CumSum HeadCount', 'Primary Work Seats', 'Avail Primary HC', 'Total seats on floor',  'Avail Total Seats HC']].sum()
                    total_row_cenarios['Building Name'] = 'Total'  
                    df_avail_row_cenarios = pd.DataFrame([total_row_cenarios])
                    df_avail_cenarios = pd.concat([df_availability_cenarios, df_avail_row_cenarios], ignore_index=True)
                    df_avail_row_cenarios.rename(columns={"CumSum HeadCoun" : "Total HC"})

                #### BIG NUMBERS - CABEÇALHO PAINEL
                    # Calculando as métricas
                    # 1) Qtde de Buildings
                    num_buildings = df_building_dash['Building Name'].nunique()

                    # 2) Qtde de Groups
                    num_groups = df_merged_cenarios['Group'].nunique()

                    # 3) Qtde de Groups + SubGroups
                    num_groups_subgroups = df_merged_cenarios[['Group', 'SubGroup']].drop_duplicates().shape[0]

                    # 4) Total de Lugares Disponíveis (Primary Work Seats)
                    total_primary_seats = df_building_dash['Primary Work Seats'].sum()

                    # 5) Total de Lugares Disponíveis (Total Seats on Floor)
                    total_floor_seats = df_building_dash['Total seats on floor'].sum()

                    # Organizando os Big Numbers em colunas
                    col1, col2, col3, col4, col5  = st.columns(5)  

                    with col1:
                        # Título com fonte e alinhamento
                        st.markdown(f"<h4 style='text-align: center; background-color: #4682B4; padding: 10px; font-size: 16px; color: white; '> Qtde de Andares</h4>", unsafe_allow_html=True)                
                        # Número com fonte personalizada e centralizado
                        st.markdown(f"<h1 style='text-align: center; font-size: 28px; font-family: Arial, sans-serif; font-weight: bold;'>{num_buildings}</h1>", unsafe_allow_html=True)
                    with col2:
                        st.markdown(f"<h4 style='text-align: center; background-color: #ADD8E6; padding: 10px; font-size: 16px;'> Qtde de Grupos</h4>", unsafe_allow_html=True) 
                        st.markdown(f"<h1 style='text-align: center; font-size: 28px; font-family: Arial, sans-serif; font-weight: bold;'>{num_groups}</h1>", unsafe_allow_html=True)
                    with col3:
                        st.markdown(f"<h4 style='text-align: center; background-color: #4682B4; padding: 10px; font-size: 14px; color: white; '> Qtde de Grupos + SubGrupos</h4>", unsafe_allow_html=True)                
                        st.markdown(f"<h1 style='text-align: center; font-size: 28px; font-family: Arial, sans-serif; font-weight: bold;'>{num_groups_subgroups}</h1>", unsafe_allow_html=True)
                    with col4:
                        st.markdown(f"<h4 style='text-align: center; background-color: #ADD8E6; padding: 10px; font-size: 14px; color: white; '> Total Primary Work Seats</h4>", unsafe_allow_html=True)                
                        st.markdown(f"<h1 style='text-align: center; font-size: 28px; font-family: Arial, sans-serif; font-weight: bold;'>{total_primary_seats}</h1>", unsafe_allow_html=True)
                    with col5:
                        st.markdown(f"<h4 style='text-align: center; background-color: #4682B4; padding: 10px; font-size: 14px; color: white; '> Total Seats</h4>", unsafe_allow_html=True)                
                        st.markdown(f"<h1 style='text-align: center; font-size: 28px; font-family: Arial, sans-serif; font-weight: bold;'>{total_floor_seats}</h1>", unsafe_allow_html=True)



                    # Adiciona um espaço maior usando <br> no Markdown
                    st.markdown("<br><br><br>", unsafe_allow_html=True)  # Adiciona 3 quebras de linha

                


                #### GRÁFICO DE AVG E PEAK
                    # Calculando os totais para cada grupo
                    df_grouped = final_consolidated_dash.groupby('Group').agg({
                        '1:1': 'sum',             # Soma o total de HeadCount por Group
                        'Peak': 'sum',            # Soma o Proportional Peak por Group
                        'Avg Occ': 'sum'          # Soma o Proportional Avg por Group
                    }).reset_index()

                    # Calculando os percentuais
                    df_grouped['Total Peak'] = (df_grouped['Peak'] / df_grouped['1:1']) * 100
                    df_grouped['Total Avg'] = (df_grouped['Avg Occ'] / df_grouped['1:1']) * 100

                    # Reformulando para ter os cálculos em linhas
                    df_melted = df_grouped.melt(id_vars=['Group'], value_vars=['Total Peak', 'Total Avg'], 
                                                var_name='CalculationType', value_name='Percentage')

                    # Criando as colunas de texto específicas para cada tipo de cálculo
                    df_melted['text'] = df_melted.apply(
                        lambda row: f"<b>{row['CalculationType']}:</b> {row['Percentage']:.1f}%" if row['Percentage'] > 0 else "", axis=1
                    )

                    # Criando o gráfico de barras
                    fig = px.bar(df_melted, 
                                x="Group", 
                                y="Percentage", 
                                color="CalculationType",  # Diferencia as barras pelo tipo de cálculo
                                title="Distribuição do Percentual de HeadCount por Group",
                                labels={"Percentage": "Percentual (%)", "Group": "Group", "CalculationType": "Cálculo"},
                                color_discrete_map={"Total Peak": "#006400", "Total Avg": "#32CD32"},  # Verde escuro e verde claro
                                text="text"  # Usando a coluna de texto específica para cada barra
                                )

                    # Habilitando interatividade no gráfico
                    fig.update_layout(
                        barmode='group',  # Barra agrupada (2 barras por grupo)
                        xaxis_title="Group",
                        yaxis_title="Percentual (%)",
                        yaxis=dict(range=[0, 100]),  # Fixando o limite do eixo Y em 100%
                        legend_title="Tipo de Percentual",
                        legend=dict(
                            x=1.05,  # Posiciona a legenda à direita
                            y=0.5,   # Ajuste vertical para que a legenda não sobreponha o gráfico
                            traceorder="normal",  # Ordem de exibição das legendas
                            orientation="v",  # Define a legenda na vertical
                            title="Group"
                        ),
                        margin=dict(t=50, b=50, l=50, r=50),  # Ajustando as margens do gráfico
                        plot_bgcolor="white",  # Cor de fundo do gráfico
                        paper_bgcolor="white",  # Cor de fundo da área do gráfico
                        width=800,  # Largura do gráfico
                        height=450,  # Altura do gráfico
                    )

                    # Adicionando os valores nas barras com texto HTML
                    fig.update_traces(texttemplate='%{text}', textposition='inside')  # Exibe o texto dentro das barras

                    # Exibe o gráfico no Streamlit
                    st.plotly_chart(fig)

                    # Adiciona um espaço maior usando <br> no Markdown
                    st.markdown("<br><br>", unsafe_allow_html=True)  # Adiciona 2 quebras de linha

                    
                #### GRÁFICO DE BARRAS - DISTRIBUIÇÃO DE HEADCOUNT POR ANDAR
                    colalochc1, colalochc2 = st.columns(2)

                    with colalochc1:
                        # Criando o gráfico de barras
                        final_consolidated_dash_bars = final_consolidated_dash.groupby(['Building Name', 'Group'], as_index=False)['1:1'].sum()
                        barscenarios = px.bar(final_consolidated_dash_bars, 
                                        x="Group", 
                                        y="1:1", 
                                        color="Building Name", 
                                        title="Distribuição de HeadCount por Alocação", 
                                        labels={"1:1": "Total HeadCount"},
                                        hover_data=["Building Name"],  # Exibe SubGroup ao passar o mouse
                                        category_orders={"Building Name": final_consolidated_dash_bars['Building Name'].unique()})  # Exibe todos os Buildings disponíveis

                        # Adicionando os valores nas barras com texto
                        barscenarios.update_traces(text=final_consolidated_dash_bars['1:1'], textposition='inside', texttemplate='%{text}')

                        # Habilitando interatividade no gráfico
                        barscenarios.update_layout(
                            barmode='stack',  # Empilha as barras por SubGroup
                            xaxis_title="Group",
                            yaxis_title="Total HeadCount",
                            legend_title="Building Name",
                            legend=dict(
                                x=1.05,  # Posiciona a legenda à direita
                                y=0.5,   # Ajuste vertical para que a legenda não sobreponha o gráfico
                                traceorder="normal",  # Ordem de exibição das legendas
                                orientation="v",  # Define a legenda na vertical
                                title="Group"
                            ),
                            margin=dict(t=50, b=50, l=50, r=50),  # Ajustando as margens do gráfico
                            plot_bgcolor="white",  # Cor de fundo do gráfico
                            paper_bgcolor="white",  # Cor de fundo da área do gráfico
                            width=800,  # Largura reduzida do gráfico
                            height=550,  # Altura reduzida do gráfico
                        )

                        # Exibindo o gráfico no Streamlit
                        st.plotly_chart(barscenarios)


                    with colalochc2:
                        final_consolidated_dash_cenarios = final_consolidated_dash[["Group", "SubGroup", "1:1","Peak", "Avg Occ", "Building Name"]].copy()
                        final_consolidated_dash_cenarios.rename(columns={"Peak" : "Peak Seats Required", "Avg Occ" : "Avg Seats Required"}, inplace=True)
                        st.dataframe(final_consolidated_dash_cenarios, use_container_width=False, hide_index=True)


                    
                    # Adiciona um espaço maior usando <br> no Markdown
                    st.markdown("<br><br>", unsafe_allow_html=True)  # Adiciona 2 quebras de linha

                    st.write("**Tabela Completa de Cenários:**")
                    st.dataframe(st.session_state.final_consolidated_dash, use_container_width=True, hide_index=True)


                else:
                    st.warning("Nenhum dado consolidado disponível. Grave os dados primeiro.")                            

        else:
            st.write("Nenhuma tabela disponível para exibição.")
 

# Tela Inicial com Seleção
st.title("Calculadora FRB - Alocação")
st.write("""
Aqui você escolhe a opção se realizar as alocações por Upload de Excel ou para Input das informações diretamente aqui pela Web.
""")
opcao = st.selectbox("Escolha uma opção", ["Selecione", "Upload de Arquivo"])

if opcao == "Upload de Arquivo":
    upload_arquivo()
