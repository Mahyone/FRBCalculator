
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
                df_building.rename(columns={df_building.columns[27]: 'Total seats on floor'}, inplace=True)

                for col in df_building.columns:
                    if col != 'Building Name':
                        df_building[col] = pd.to_numeric(df_building[col], errors='coerce').fillna(0).astype(int)

                if 'Primary Work Seats' not in df_building.columns:
                    st.warning("Coluna 'Primary Work Seats' não encontrada. Adicionando valores padrão.")
                    df_building['Primary Work Seats'] = 0

                df_building_trat = df_building[
                    (df_building['Total seats on floor'] > 0) & 
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
                df_proportional = df_proportional.drop_duplicates()       


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
                    # Reordena colunas e ordena
                    cols = df_allocation.columns.tolist()
                    if "Building Name" in cols and "Current Location" in cols:
                        new_order = (["Building Name"] +
                                    [c for c in cols if c not in ("Building Name", "Current Location")] +
                                    ["Current Location"])
                        df_allocation = df_allocation[new_order]
                    df_allocation = df_allocation.sort_values("Building Name")

                        # Calcula subtotais (Alocados vs Não Alocados)
                    num_cols = df_allocation.select_dtypes(include="number").columns.tolist()

                    # Calcula subtotais
                    alocados      = df_allocation[df_allocation["Building Name"] != "Não Alocado"]
                    nao_alocados  = df_allocation[df_allocation["Building Name"] == "Não Alocado"]

                    soma_alocados = alocados[num_cols].sum()
                    soma_alocados["Building Name"] = "Alocados"

                    soma_nao      = nao_alocados[num_cols].sum()
                    soma_nao["Building Name"] = "Não Alocados"

                    # **Novo: calcula Total Geral**
                    soma_geral    = df_allocation[num_cols].sum()
                    soma_geral["Building Name"] = "Total Geral"

                    # Constrói DataFrame de subtotais + total geral
                    df_subtotais = pd.DataFrame(
                        [soma_alocados, soma_nao, soma_geral],
                        columns=df_allocation.columns
                    )

                    # Concatena original + subtotais
                    df_tot = pd.concat([df_allocation, df_subtotais], ignore_index=True)

                        # Cria map de cores alternadas por prédio
                    unique_buildings = df_allocation["Building Name"].drop_duplicates().tolist()
                    building_colors = {
                        b: "#D3D3D3" if i % 2 == 0 else ""
                        for i, b in enumerate(unique_buildings)
                    }

                        # Função única de highlight (inclui subtotais em cinza médio + negrito
                    def highlight_rows(row):
                        name = row['Building Name']
                        if name == 'Total Geral':
                            # Azul petróleo escuro + texto em branco
                            return ['background-color: #004E64; color: #FFFFFF'] * len(row)
                        if name in ('Alocados', 'Não Alocados'):
                            # Tom mais claro do azul petróleo + texto em branco
                            return ['background-color: #357A91; color: #FFFFFF'] * len(row)
                        # linhas originais continuam com cinza claro alternado
                        color = building_colors.get(name, '')
                        return [f'background-color: {color}'] * len(row)

                    st.write("#### Resultado da Automação - HeadCount")
                    styled = df_tot.style.apply(highlight_rows, axis=1)
                    st.dataframe(styled, use_container_width=False)

                        # Tabela de capacidade + ocupados + restante
                    cap = (
                        df_building_trat[["Building Name", "Primary Work Seats"]]
                        .rename(columns={"Primary Work Seats": "Capacity"})
                    )
                    rem_df = (
                        pd.DataFrame(list(remaining_floors.items()), columns=["Building Name", "Remaining"])
                        .merge(cap, on="Building Name", how="left")
                    )
                    rem_df["Occupied"] = rem_df["Capacity"] - rem_df["Remaining"]
                    rem_df = rem_df[["Building Name", "Capacity", "Occupied", "Remaining"]]

                    st.write("#### Capacidade restante nos andares - HeadCount:")
                    st.dataframe(rem_df, use_container_width=False)

                    return df_tot, rem_df


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
                    total_row['Building Name'] = 'Total' 
                    total_row_df = pd.DataFrame([total_row])
                    df_hc_nonallocated_with_total = pd.concat([df_hc_nonallocated, total_row_df], ignore_index=True)
                    
                    def highlight_nonalloc_total(row):
                        if row["Building Name"] == "Total":
                            return ['background-color: #004E64; color: #FFFFFF'] * len(row)
                        return [""] * len(row)

                    styled_non = df_hc_nonallocated_with_total.style.apply(highlight_nonalloc_total, axis=1)
                    st.dataframe(styled_non, use_container_width=False)

                    ## st.dataframe(df_hc_nonallocated_with_total, use_container_width=False)


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
                def display_allocation_peak(df_allocation, remaining_floors, df_building_trat):
                        # Ordena e cria as colunas de Peak
                    df_allocation = df_allocation.sort_values(by='Building Name')
                    df_allocation['Peak with Exception'] = df_allocation.apply(
                        lambda r: r['HeadCount'] if r['Exception (Y/N)']=='Y' else r['Proportional Peak'], axis=1
                    )
                    df_allocation['Peak % of HeadCount'] = (
                        (df_allocation['Peak with Exception'] / df_allocation['HeadCount'])*100
                    ).round(0).astype(int)
                    df_allocation.drop(columns=['Proportional Peak'], inplace=True)
                    df_allocation.rename(columns={'Proportional Avg':'Avg Occ'}, inplace=True)

                        # Reordena colunas (inserindo as novas nos lugares certos)
                    cols = [
                        'Building Name','Group','SubGroup','FTE','CW','Growth',
                        'HeadCount','Exception (Y/N)','Peak with Exception','Peak % of HeadCount',
                        'Avg Occ','Adjacency Priority 1','Adjacency Priority 2',
                        'Adjacency Priority 3','Current Location'
                    ]
                    df_allocation = df_allocation[cols]

                        # Calcula subtotais e total geral
                    num_cols = df_allocation.select_dtypes(include='number').columns.tolist()
                    alocados     = df_allocation[df_allocation['Building Name']!='Não Alocado']
                    nao_alocados = df_allocation[df_allocation['Building Name']=='Não Alocado']

                    soma_alocados = alocados[num_cols].sum()
                    soma_alocados['Building Name'] = 'Alocados'

                    soma_nao = nao_alocados[num_cols].sum()
                    soma_nao['Building Name'] = 'Não Alocados'

                    soma_geral = df_allocation[num_cols].sum()
                    soma_geral['Building Name'] = 'Total Geral'

                    df_subtotais = pd.DataFrame(
                        [soma_alocados, soma_nao, soma_geral],
                        columns=df_allocation.columns
                    )
                    df_tot = pd.concat([df_allocation, df_subtotais], ignore_index=True)
                    dfautomation_peak = df_tot.copy()

                        # Cores alternadas por prédio (apenas para as linhas originais)
                    unique_buildings = df_allocation['Building Name'].drop_duplicates().tolist()
                    building_colors = {
                        b: '#D3D3D3' if i % 2 == 0 else ''
                        for i, b in enumerate(unique_buildings)
                    }

                    def highlight_rows(row):
                        name = row['Building Name']
                        if name == 'Total Geral':
                            # Azul petróleo escuro + texto em branco
                            return ['background-color: #004E64; color: #FFFFFF'] * len(row)
                        if name in ('Alocados', 'Não Alocados'):
                            # Tom mais claro do azul petróleo + texto em branco
                            return ['background-color: #357A91; color: #FFFFFF'] * len(row)
                        # linhas originais continuam com cinza claro alternado
                        color = building_colors.get(name, '')
                        return [f'background-color: {color}'] * len(row)
                    

                    st.write("#### Resultado da Automação - Peak")
                    st.dataframe(dfautomation_peak.style.apply(highlight_rows, axis=1),
                                use_container_width=False)

                        # Tabela de capacidade x ocupado x restante
                    cap = (
                        df_building_trat[['Building Name','Primary Work Seats']]
                        .rename(columns={'Primary Work Seats':'Capacity'})
                    )
                    rem_df = (
                        pd.DataFrame(list(remaining_floors.items()),
                                    columns=['Building Name','Remaining'])
                        .merge(cap, on='Building Name', how='left')
                    )
                    rem_df['Occupied'] = rem_df['Capacity'] - rem_df['Remaining']
                    rem_df = rem_df[['Building Name','Capacity','Occupied','Remaining']]

                    st.write("#### Capacidade restante nos andares - Peak:")
                    st.dataframe(rem_df, use_container_width=False)

                    return dfautomation_peak, rem_df


                # Carregar os dados e realizar a alocação
                if "df_building_trat" in st.session_state and "df_proportional" in st.session_state:
                    df_building_trat = st.session_state.df_building_trat
                    df_proportional  = st.session_state.df_proportional
                    floors = dict(zip(df_building_trat['Building Name'], df_building_trat['Primary Work Seats']))

                    # Função de alocação específica de Peak (já existente no seu código)
                    df_alloc_peak, remaining = allocate_groups_peak(df_proportional, floors.copy())

                    # Exibe e captura os DataFrames estilizados
                    dfautomation_peak, rem_peak_df = display_allocation_peak(
                        df_alloc_peak, remaining, df_building_trat
                    )
                    st.session_state.dfautomation_peak = dfautomation_peak

                        # Tabela de Não Alocados com Total em destaque
                    df_peak_non = dfautomation_peak[dfautomation_peak['Building Name']=='Não Alocado']
                    num_cols = df_peak_non.select_dtypes(include='number').columns
                    total_row = df_peak_non[num_cols].sum()
                    total_row['Building Name'] = 'Total'
                    df_peak_non_total = pd.concat(
                        [df_peak_non, pd.DataFrame([total_row])], ignore_index=True
                    )

                    def highlight_nonalloc_total(r):
                        return (['background-color: #004E64; color: #FFFFFF']
                                * len(r)) if r['Building Name']=='Total' else ['']*len(r)

                    st.write("#### Grupos Não Alocados - Peak:")
                    styled_non = df_peak_non_total.style.apply(highlight_nonalloc_total, axis=1)
                    st.dataframe(styled_non, use_container_width=False)


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
                def display_allocation_avg(df_allocation, remaining_floors, df_building_trat):
                        # Ordena por prédio
                    df_allocation = df_allocation.sort_values(by='Building Name')

                        # Cálculo de Avg Occ with Exception e %
                    df_allocation['Avg Occ with Exception'] = df_allocation.apply(
                        lambda r: r['HeadCount'] if r['Exception (Y/N)']=='Y' else r['Proportional Avg'],
                        axis=1
                    )
                    df_allocation['Avg Occ % of HeadCount'] = (
                        (df_allocation['Avg Occ with Exception'] / df_allocation['HeadCount']) * 100
                    ).round(0).astype(int)

                        # Reordena colunas
                    cols = [
                        'Building Name','Group','SubGroup','FTE','CW','Growth',
                        'HeadCount','Exception (Y/N)','Avg Occ with Exception','Avg Occ % of HeadCount',
                        'Adjacency Priority 1','Adjacency Priority 2','Adjacency Priority 3','Current Location'
                    ]
                    df_allocation = df_allocation[cols]

                        # Calcula subtotais e total geral
                    num_cols = df_allocation.select_dtypes(include='number').columns.tolist()

                    alocados     = df_allocation[df_allocation['Building Name']!='Não Alocado']
                    nao_alocados = df_allocation[df_allocation['Building Name']=='Não Alocado']

                    soma_alocados = alocados[num_cols].sum()
                    soma_alocados['Building Name'] = 'Alocados'

                    soma_nao = nao_alocados[num_cols].sum()
                    soma_nao['Building Name'] = 'Não Alocados'

                    soma_geral = df_allocation[num_cols].sum()
                    soma_geral['Building Name'] = 'Total Geral'

                    df_subtotais = pd.DataFrame(
                        [soma_alocados, soma_nao, soma_geral],
                        columns=df_allocation.columns
                    )
                    df_tot = pd.concat([df_allocation, df_subtotais], ignore_index=True)

                        # Renomeia para o nome de cálculo
                    dfautomation_avg = df_tot.copy()

                        # Prepara cores alternadas por prédio
                    unique_buildings = df_allocation['Building Name'].drop_duplicates().tolist()
                    building_colors = {
                        b: '#D3D3D3' if i % 2 == 0 else ''
                        for i, b in enumerate(unique_buildings)
                    }

                    def highlight_rows(row):
                        name = row['Building Name']
                        if name == 'Total Geral':
                            # Azul petróleo escuro + texto em branco
                            return ['background-color: #004E64; color: #FFFFFF'] * len(row)
                        if name in ('Alocados', 'Não Alocados'):
                            # Tom mais claro do azul petróleo + texto em branco
                            return ['background-color: #357A91; color: #FFFFFF'] * len(row)
                        # linhas originais continuam com cinza claro alternado
                        color = building_colors.get(name, '')
                        return [f'background-color: {color}'] * len(row)

                        # Exibição estilizada
                    st.write("#### Resultado da Automação - Avg Occ")
                    st.dataframe(
                        dfautomation_avg
                        .style.apply(highlight_rows, axis=1),
                        use_container_width=False
                    )

                        # Tabela de capacidade x ocupado x restante
                    cap = (
                        df_building_trat[['Building Name','Primary Work Seats']]
                        .rename(columns={'Primary Work Seats':'Capacity'})
                    )
                    rem_df = (
                        pd.DataFrame(list(remaining_floors.items()),
                                    columns=['Building Name','Remaining'])
                        .merge(cap, on='Building Name', how='left')
                    )
                    rem_df['Occupied'] = rem_df['Capacity'] - rem_df['Remaining']
                    rem_df = rem_df[['Building Name','Capacity','Occupied','Remaining']]

                    st.write("#### Capacidade restante nos andares - Avg Occ:")
                    st.dataframe(rem_df, use_container_width=False)

                    return dfautomation_avg, rem_df


                # Carregar os dados e realizar a alocação
                if "df_building_trat" in st.session_state and "df_proportional" in st.session_state:
                    df_building_trat = st.session_state.df_building_trat
                    df_proportional  = st.session_state.df_proportional
                    floors = dict(zip(df_building_trat['Building Name'], df_building_trat['Primary Work Seats']))

                    # chamada à sua função de alocação específica p/ Avg Occ
                    df_alloc_avg, remaining = allocate_groups_avg(df_proportional, floors.copy())

                    # exibe e captura como dfautomation_avg
                    dfautomation_avg, rem_avg_df = display_allocation_avg(
                        df_alloc_avg, remaining, df_building_trat
                    )
                    st.session_state.dfautomation_avg = dfautomation_avg

                        # Tabela “Não Alocados” com total destacado
                    df_avg_non = dfautomation_avg[dfautomation_avg['Building Name']=='Não Alocado']
                    tot = df_avg_non.select_dtypes(include='number').sum()
                    tot['Building Name'] = 'Total'
                    df_avg_non_total = pd.concat(
                        [df_avg_non, pd.DataFrame([tot])], ignore_index=True
                    )

                    def highlight_nonalloc_total(r):
                        return (
                            ['background-color: #004E64; color: #FFFFFF'] * len(r)
                            if r['Building Name']=='Total' else ['']*len(r)
                        )

                    st.write("#### Grupos Não Alocados - Avg Occ:")
                    st.dataframe(
                        df_avg_non_total.style.apply(highlight_nonalloc_total, axis=1),
                        use_container_width=False
                    )

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

        # Carrega os DataFrames do session_state
        df_hc       = st.session_state.get('dfautomation_hc', pd.DataFrame())
        df_peak     = st.session_state.get('dfautomation_peak', pd.DataFrame())
        df_avg      = st.session_state.get('dfautomation_avg', pd.DataFrame())
        df_building = st.session_state.get('df_building_trat', pd.DataFrame())

        if df_hc.empty or df_building.empty:
            st.info("Nenhum dado de alocação disponível. Execute a automação primeiro.")
            st.stop()


        # === BIG NUMBERS ===
        st.markdown("<h4 style='text-align:center'>Visão Consolidada</h4>", unsafe_allow_html=True)
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.markdown("<div style='background-color:#004E64;color:white;padding:10px;text-align:center'><b># Andares</b><br>" + str(df_building['Building Name'].nunique()) + "</div>", unsafe_allow_html=True)
        col2.markdown("<div style='background-color:#B0B0B0;color:black;padding:10px;text-align:center'><b># Groups</b><br>" + str(df_hc['Group'].nunique()) + "</div>", unsafe_allow_html=True)
        col3.markdown("<div style='background-color:#004E64;color:white;padding:10px;text-align:center'><b># Groups+SubGroups</b><br>" + str(df_hc[['Group','SubGroup']].drop_duplicates().shape[0]) + "</div>", unsafe_allow_html=True)
        col4.markdown("<div style='background-color:#B0B0B0;color:black;padding:10px;text-align:center'><b>Total Primary Seats</b><br>" + str(df_building['Primary Work Seats'].sum()) + "</div>", unsafe_allow_html=True)
        col5.markdown("<div style='background-color:#004E64;color:white;padding:10px;text-align:center'><b>Total Floor Seats</b><br>" + str(df_building['Total seats on floor'].sum()) + "</div>", unsafe_allow_html=True)


        # === Funções auxiliares ===
        def prepare_avail(df_base, key_col):
            df = pd.merge(df_base,
                        df_building[['Building Name','Primary Work Seats','Total seats on floor']],
                        on='Building Name', how='left')
            df = df.sort_values(['Building Name','Group','SubGroup'])

            # Cria coluna auxiliar para classificação
            df['Status'] = df['Building Name'].apply(lambda x: 'Não Alocado' if x == 'Não Alocado' else 'Alocado')

            # Totais
            total_geral = df[key_col].sum()
            alocado = df[df['Status'] == 'Alocado'][key_col].sum()
            nao_alocado = df[df['Status'] == 'Não Alocado'][key_col].sum()

            df_summary = pd.DataFrame([
                {'Status': 'Alocado', key_col: alocado},
                {'Status': 'Não Alocado', key_col: nao_alocado},
                {'Status': 'Total Geral', key_col: total_geral}
            ])

            return df, df_summary

       

        def plot_donut(df_base, key_col, title):
            df_base['Status'] = df_base['Building Name'].apply(lambda x: 'Alocado' if x!='Não Alocado' else 'Não Alocado')
            by_status = df_base.groupby('Status').agg({key_col:'sum'}).reset_index()
            fig = px.pie(by_status, names='Status', values=key_col, hole=0.3,
                        color='Status',
                        color_discrete_map={'Alocado':'#357A91','Não Alocado':'#FFAA33'},
                        title=f"% Alloc {title}")
            return fig, by_status
        

        def style_summary_table(df):
            def highlight(row):
                if row['Status'] == 'Total Geral':
                    return ['background-color: #004E64; color: white'] * len(row)
                elif row['Status'] == 'Alocado':
                    return ['background-color: #357A91; color: black'] * len(row)
                elif row['Status'] == 'Não Alocado':
                    return ['background-color: #FFAA33; color: black'] * len(row)
                else:
                    return [''] * len(row)
            return df.style.apply(highlight, axis=1)
        
        # Adiciona um espaço maior usando <br> no Markdown
        st.markdown("<br><br>", unsafe_allow_html=True)  # Adiciona 2 quebras de linha


        # === VISÃO COMPARATIVA SUPERIOR ===
        st.markdown("<h5 style='text-align:center'>Comparativo Consolidado</h5>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)

        for df_data, key_col, title, container in zip(
            [df_hc, df_peak, df_avg],
            ['HeadCount', 'Peak with Exception', 'Avg Occ with Exception'],
            ['HeadCount', 'Peak', 'Avg Occ'],
            [col1, col2, col3]
        ):
            with container:
                # Parte superior: consolida os totais
                df_full, resumo = prepare_avail(df_data, key_col)
                st.dataframe(style_summary_table(resumo), use_container_width=True, hide_index=True)

                # Gráfico de donut
                fig, _ = plot_donut(df_data, key_col, title)
                st.plotly_chart(fig, use_container_width=True, key=f"donut_{title}")

                # Parte inferior: apenas dados reais
                #df_detail = df_full[~df_full['Building Name'].isin(['Alocado', 'Não Alocado', 'Total Geral'])]
                df_detail = df_full.copy()

                # Identifica campo respectivo por cenário
                if title == 'HeadCount':
                    col_raw = 'HeadCount'
                    # GroupBy apenas com o campo absoluto
                    df_final = (
                        df_detail
                        .groupby(['Building Name', 'Group', 'SubGroup'], as_index=False)[[col_raw]]
                        .sum()
                    )
                else:
                    if title == 'Peak':
                        col_raw = 'Peak with Exception'
                        col_pct = 'Peak %'
                    elif title == 'Avg Occ':
                        col_raw = 'Avg Occ with Exception'
                        col_pct = 'Avg %'

                    # Cálculo de % com base no total real
                    total_value = df_detail[col_raw].sum()
                    df_detail[col_pct] = ((df_detail[col_raw] / total_value) * 100).round(0).astype(int)

                    # GroupBy com campo absoluto + %
                    df_final = (
                        df_detail
                        .groupby(['Building Name', 'Group', 'SubGroup'], as_index=False)[[col_raw, col_pct]]
                        .sum()
                    )


                # Exibe resultado
                st.dataframe(df_final, use_container_width=True, hide_index=True)


        # Adiciona um espaço maior usando <br> no Markdown
        st.markdown("<br><br>", unsafe_allow_html=True)  # Adiciona 2 quebras de linha
        st.markdown("---")

        # === DEEP DIVE ===
        tabela_selecionada = st.selectbox(
            'Escolha cenário para Deep Dive:',
            ('HeadCount','Peak','Avg Occ')
        )

        def render_deep(df_base, key_col, title):
            st.write(f"### Deep Dive: {title}")

            colx, coly = st.columns([2, 1])
            with colx:
                df_grp = df_base.groupby('Group').agg({key_col:'sum'}).reset_index()
                fig1 = px.pie(df_grp, names='Group', values=key_col, hole=0.3,
                            title=f"Distribuição {title} por Group")
                st.plotly_chart(fig1, use_container_width=True)
            with coly:
                st.dataframe(df_grp, use_container_width=True, hide_index=True)

            # Adiciona um espaço maior usando <br> no Markdown
            st.markdown("<br><br>", unsafe_allow_html=True)  # Adiciona 2 quebras de linha

            df_counts = df_base.groupby(['Building Name','Group'], as_index=False)[key_col].sum()
            pivot = df_counts.pivot(index='Building Name',columns='Group',values=key_col).fillna(0)
            pct   = pivot.div(pivot.sum(axis=1),axis=0)*100
            df_long = pct.reset_index().melt('Building Name',var_name='Group',value_name='Percent')
            df_long = df_long.merge(df_counts,on=['Building Name','Group'])
            ord_b = sorted(df_long['Building Name'].unique())
            ord_g = sorted(df_long['Group'].unique())
            fig2 = px.bar(df_long, x='Percent', y='Building Name', color='Group', orientation='h',
                        category_orders={'Building Name':ord_b,'Group':ord_g},
                        labels={'Percent':f'% {title}','Building Name':'Andar'},
                        title=f'% {title} por Grupo e Andar',
                        hover_data={'Percent':':.1f', key_col:True})
            fig2.update_layout(barmode='stack', xaxis=dict(ticksuffix='%'), margin=dict(l=80,r=20,t=30,b=30))
            st.plotly_chart(fig2, use_container_width=True)
            st.dataframe(df_counts, use_container_width=False, hide_index=True)

        if tabela_selecionada=='HeadCount':
            render_deep(df_hc,'HeadCount','HeadCount')
        elif tabela_selecionada=='Peak':
            render_deep(df_peak,'Peak with Exception','Peak')
        else:
            render_deep(df_avg,'Avg Occ with Exception','Avg Occ')




# Tela Inicial com Seleção
st.title("Calculadora FRB - Alocação")
st.write("""
Aqui você escolhe a opção se realizar as alocações por Upload de Excel ou para Input das informações diretamente aqui pela Web.
""")
opcao = st.selectbox("Escolha uma opção", ["Selecione", "Upload de Arquivo"])

if opcao == "Upload de Arquivo":
    upload_arquivo()
