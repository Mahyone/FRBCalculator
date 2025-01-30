
import streamlit as st
from PIL import Image
import pandas as pd
import numpy as np
from itertools import combinations
import io
import os

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
    tabs = st.tabs(["Glossário", "Importar Arquivo", "Automação", "Cenarios", "Dashboards"])

    with tabs[0]:
        st.header("Glossário")
        st.write("### **Para Importação do arquivo em Excel, é necessário que as abas estejam nesse padrão:**")

        st.write("#### Aba 'Staff HeadCount:")
        staffheadcountimage = Image.open("staffheadcountimage.PNG") 
        st.image(staffheadcountimage, use_container_width=False) 
        st.write(""" Assegurar que as informações estejam dispostas como na imagem:   
                 - Cabeçalho estar na linha 4   
                 - Preenchimento de informações entre colunas A à I, seguindo a ordem de preenchimento da imagem   
            """)
        
        st.write("#### Aba 'Staff Occupancy:")
        staffoccupancyimage = Image.open("staffoccupancyimage.PNG") 
        st.image(staffoccupancyimage, use_container_width=False) 
        st.write(""" Assegurar que as informações estejam dispostas como na imagem:   
                 - Cabeçalho estar na linha 4   
                 - Preenchimento de informações entre colunas A à F, seguindo a ordem de preenchimento da imagem   
            """)
        
        st.write("#### Aba 'SubGroup Adjacencies:")
        subgroupadjacenciesimage = Image.open("subgroupadjacenciesimage.PNG") 
        st.image(subgroupadjacenciesimage, use_container_width=False) 
        st.write(""" Assegurar que as informações estejam dispostas como na imagem:   
                 - Cabeçalho estar na linha 4   
                 - Preenchimento de informações entre colunas A à E, seguindo a ordem de preenchimento da imagem   
            """)

        st.write("#### Aba 'Building Space Summary:")
        buildingspaceimage = Image.open("buildingspaceimage.PNG") 
        st.image(buildingspaceimage, use_container_width=False) 
        st.write(""" Assegurar que as informações estejam dispostas como na imagem:   
                 - Cabeçalho estar na linha 7   
                 - Preenchimento de informações entre colunas A à U, seguindo a ordem de preenchimento da imagem   
            """)


    ##### ABA IMPORTAÇÃO #####   
    with tabs[1]:
        st.header("Importar Arquivo")
        
        # Função para carregar e processar os dados do Excel
        def process_excel_data(file_path):
            try:
                # Carregar tabelas do Excel
                df_staffheadcount = pd.read_excel(file_path, sheet_name='2. Staff Headcount ', skiprows=3, usecols="A:I")
                df_staffoccupancy = pd.read_excel(file_path, sheet_name='3. Staff Occupancy', skiprows=3, usecols="A:F")
                df_subgroupadjacenties = pd.read_excel(file_path, sheet_name='4. SubGroup Adjacencies', skiprows=3, usecols="A:E")
                df_building = pd.read_excel(file_path, sheet_name='5. Building Space Summary', skiprows=6, usecols="A:V")

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

                df_proportional = df_proportional[['Current Location', 'Group', 'SubGroup', 'FTE','CW', 'Growth', 'HeadCount', 'Exception (Y/N)', 'Comments', 'Proportional Peak', 'Proportional Avg',
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
    with tabs[2]:
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
                    # Ordenar os dados por 'Building Name'
                    df_allocation = df_allocation.sort_values(by='Building Name')
                    st.write("#### Resultado da Automação - HeadCount")
                    st.dataframe(df_allocation.fillna(""), use_container_width=False)

                    # Exibir a capacidade restante nos andares
                    st.write("#### Capacidade restante nos andares - HeadCount:")
                    remaining_floors_df = pd.DataFrame(list(remaining_floors.items()), columns=['Building Name', 'Remaining Seats'])
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
                    df_allocation, remaining_floors = allocate_groups(df_proportional, floors.copy())

                    # Exibir os resultados de alocação
                    df_allocation_result, remaining_floors_df_result = display_allocation(df_allocation, remaining_floors, df_building_trat)
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


                # Botão para exportar tabela "Building" para Excel
                if st.button("Exportar 'Automação HeadCount' para Excel", key="export_automacao_hc"):
                    with io.BytesIO() as output:
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            df_allocation = st.session_state.df_allocation
                            df_allocation_export = df_allocation.fillna("")                        
                            df_allocation_export.to_excel(writer, sheet_name="Automacao_HeadCount", index=False)
                        st.download_button(
                            label="Download do Excel - Automação HeadCount",
                            data=output.getvalue(),
                            file_name="automacao_headcount.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    

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
                st.write(f"**Total Avg Peak**: {total_proppeak} || **Total Avg Peak Exception**: {total_proportional_Peak_exception}")


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
                    st.dataframe(df_allocation.fillna(""), use_container_width=False)

                    # Exibir a capacidade restante nos andares
                    st.write("#### Capacidade restante nos andares - Peak:")
                    remaining_floors_df = pd.DataFrame(list(remaining_floors.items()), columns=['Building Name', 'Remaining Seats'])
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
                    dfautomation_peak = df_allocation.copy()
                    st.session_state.dfautomation_peak = dfautomation_peak  # Salvando no session_state

                    st.write("### Grupos Não Alocados:")
                    df_peak_nonallocated = df_allocation[df_allocation['Building Name'] == 'Não Alocado']
                    df_peak_nonallocated['Peak_Exception'] = df_peak_nonallocated.apply(
                    lambda row: row['HeadCount'] if row['Exception (Y/N)'] == 'Y' else row['Proportional Peak'],
                    axis=1
                    )

                    numeric_columns = df_peak_nonallocated.select_dtypes(include='number').columns
                    total_row = df_peak_nonallocated[numeric_columns].sum()
                    total_row['Group'] = 'Total' 
                    total_row_df = pd.DataFrame([total_row])
                    df_peak_nonallocated_total = pd.concat([df_peak_nonallocated, total_row_df], ignore_index=True)
                    st.dataframe(df_peak_nonallocated_total, use_container_width=False)



                # Botão para exportar tabela "Building" para Excel
                if st.button("Exportar 'Automação Peak' para Excel", key="export_automacao_peak"):
                    with io.BytesIO() as output:
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            df_allocation = st.session_state.df_allocation
                            df_allocation_export = df_allocation.fillna("")                        
                            df_allocation_export.to_excel(writer, sheet_name="Automacao_Peak", index=False)
                        st.download_button(
                            label="Download do Excel - Automação Peak",
                            data=output.getvalue(),
                            file_name="Automacao_Peak.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )



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
                st.write(f"**Total Avg**: {total_propavg} || **Total Avg Exception**: {total_proportional_Avg_exception}")

                
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
                    st.dataframe(df_allocation.fillna(""), use_container_width=False)

                    # Exibir a capacidade restante nos andares
                    st.write("#### Capacidade restante nos andares - Avg:")
                    remaining_floors_df = pd.DataFrame(list(remaining_floors.items()), columns=['Building Name', 'Remaining Seats'])
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
                    dfautomation_avg = df_allocation.copy()
                    st.session_state.dfautomation_avg = dfautomation_avg  # Salvando no session_state

                    st.write("### Grupos Não Alocados:")
                    df_avg_nonallocated = df_allocation[df_allocation['Building Name'] == 'Não Alocado']
                    df_avg_nonallocated['Avg_Exception'] = df_avg_nonallocated.apply(
                    lambda row: row['HeadCount'] if row['Exception (Y/N)'] == 'Y' else row['Proportional Avg'],
                    axis=1
                    )

                    numeric_columns = df_avg_nonallocated.select_dtypes(include='number').columns
                    total_row = df_avg_nonallocated[numeric_columns].sum()
                    total_row['Group'] = 'Total' 
                    total_row_df = pd.DataFrame([total_row])
                    df_avg_nonallocated_total = pd.concat([df_avg_nonallocated, total_row_df], ignore_index=True)
                    st.dataframe(df_avg_nonallocated_total, use_container_width=False)

                # Botão para exportar tabela "Building" para Excel
                if st.button("Exportar 'Automação Avg' para Excel", key="export_automacao_avg"):
                    with io.BytesIO() as output:
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            df_allocation = st.session_state.df_allocation
                            df_allocation_export = df_allocation.fillna("")                        
                            df_allocation_export.to_excel(writer, sheet_name="Automacao_Avg", index=False)
                        st.download_button(
                            label="Download do Excel - Automação Avg",
                            data=output.getvalue(),
                            file_name="Automacao_Avg.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
        else:
            st.write("Por favor, carregue o arquivo para prosseguir.") 




   ##### ABA CENÁRIOS #####
    with tabs[3]:
        st.write("### Cenários de Alocação")

        # Inicializar df_proportional como um DataFrame vazio, se não houver dados na sessão
        if "df_proportional" not in st.session_state:
            df_proportional = pd.DataFrame()  # DataFrame vazio
        else:
            df_proportional = st.session_state.df_proportional
            

        # Verificar se o df_proportional tem dados antes de continuar
        if not df_proportional.empty:
            # Criação da tabela de cenários
            df_proportional_cenarios = df_proportional.copy()
            df_proportional_cenarios = df_proportional_cenarios[["Group", "SubGroup", "Exception (Y/N)", "HeadCount", "Proportional Peak", "Proportional Avg"]]
            df_proportional_cenarios.rename(columns={"HeadCount": "1:1", "Proportional Peak": "Peak", "Proportional Avg": "Avg Occ"}, inplace=True)
            with st.expander(f"#### **Informações Cadastradas**"):
                st.dataframe(df_proportional_cenarios, use_container_width=False, hide_index=True)
            
        

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
            if "tables_to_append" not in st.session_state:
                st.session_state.tables_to_append = []  # Lista vazia inicialmente
            
            # Exibindo as informações com expanders para cada 'Building Name'
            for building in df_final['Building Name'].unique():
                with st.expander(f"#### **Informações do Andar: {building}**"):
                    st.write(f"**Informações do Andar: {building}**")
                    df_building_data = df_final[df_final['Building Name'] == building]
                    primary_work_seats = df_building_data['Primary Work Seats'].iloc[0]
                    total_seats_on_floor = df_building_data['Total seats on floor'].iloc[0]
                    
                    st.write(f"**Primary Work Seats**: {primary_work_seats}")
                    st.write(f"**Total seats on floor**: {total_seats_on_floor}")

                    # Dropdown de multiseleção para 'Group' e 'SubGroup' dentro de cada seção
                    groups_subgroups = df_building_data[['Group', 'SubGroup']].drop_duplicates()
                    group_subgroup_options = []

                    # Criando as combinações de 'Group' e 'SubGroup' para cada andar
                    for index, row in groups_subgroups.iterrows():
                        group_subgroup_options.append(f"{row['Group']} - {row['SubGroup']}" if row['SubGroup'] else f"{row['Group']} - ")

                    # Passando uma chave única para o multiselect usando o 'building' (nome do andar)
                    selected_options = st.multiselect(
                        "Selecione os Grupos e Subgrupos",
                        options=group_subgroup_options,
                        default=None,
                        key=f"multiselect_{building}"  # Usando o nome do andar como chave única
                    )

                    # Filtrando a tabela com base na seleção dentro do expander do building
                    if selected_options:
                        # Filtra os dados para que apenas as combinações selecionadas sejam mostradas
                        df_building_data_filtered = df_building_data[df_building_data['Chave'].isin(selected_options)]
                    else:
                        # Se não houver seleção, mostra todos os dados
                        df_building_data_filtered = df_building_data
                        

                    # Caixa de Texto para input de Margem de Growth e Risk          
                    growth_value = st.text_input(f"Growth (numérico, não digitar o símbolo '%') para {building}", value="", key=f"growth_input_{building}")
                    risk_value = st.text_input(f"Risk (numérico, não digitar o símbolo '%') para {building}", value="", key=f"risk_input_{building}")

                    # Converte os valores para inteiros, ou 0 se estiverem vazios
                    growth_value = int(growth_value) if growth_value else 0
                    risk_value = int(risk_value) if risk_value else 0

                    # Calcular as novas colunas, se Growth ou Risk forem preenchidos
                    df_building_data_filtered['Growth 1:1'] = df_building_data_filtered['1:1'] * (1 + growth_value / 100)
                    df_building_data_filtered['Saldo Growth 1:1'] = df_building_data_filtered['Primary Work Seats'] - df_building_data_filtered['Growth 1:1']
                    df_building_data_filtered['Growth 1:1'] = df_building_data_filtered['Growth 1:1'].round(0).astype(int)
                    df_building_data_filtered['Saldo Growth 1:1'] = df_building_data_filtered['Saldo Growth 1:1'].round(0).astype(int)

                    df_building_data_filtered['Growth Peak'] = df_building_data_filtered['Peak'] * (1 + growth_value / 100)
                    df_building_data_filtered['Saldo Growth Peak'] = df_building_data_filtered['Primary Work Seats'] - df_building_data_filtered['Growth Peak']
                    df_building_data_filtered['Growth Peak'] = df_building_data_filtered['Growth Peak'].round(0).astype(int)
                    df_building_data_filtered['Saldo Growth Peak'] = df_building_data_filtered['Saldo Growth Peak'].round(0).astype(int)

                    df_building_data_filtered['Growth Avg Occ'] = df_building_data_filtered['Avg Occ'] * (1 + growth_value / 100)
                    df_building_data_filtered['Saldo Growth Avg Occ'] = df_building_data_filtered['Primary Work Seats'] - df_building_data_filtered['Growth Avg Occ']
                    df_building_data_filtered['Growth Avg Occ'] = df_building_data_filtered['Growth Avg Occ'].round(0).astype(int)
                    df_building_data_filtered['Saldo Growth Avg Occ'] = df_building_data_filtered['Saldo Growth Avg Occ'].round(0).astype(int)

                    df_building_data_filtered['Risk 1:1'] = df_building_data_filtered['1:1'] * (1 - risk_value / 100)
                    df_building_data_filtered['Saldo Risk 1:1'] = df_building_data_filtered['Primary Work Seats'] - df_building_data_filtered['Risk 1:1']
                    df_building_data_filtered['Risk 1:1'] = df_building_data_filtered['Risk 1:1'].round(0).astype(int)
                    df_building_data_filtered['Saldo Risk 1:1'] = df_building_data_filtered['Saldo Risk 1:1'].round(0).astype(int)

                    df_building_data_filtered['Risk Peak'] = df_building_data_filtered['Peak'] * (1 - risk_value / 100)
                    df_building_data_filtered['Saldo Risk Peak'] = df_building_data_filtered['Primary Work Seats'] - df_building_data_filtered['Risk Peak']
                    df_building_data_filtered['Risk Peak'] = df_building_data_filtered['Risk Peak'].round(0).astype(int)
                    df_building_data_filtered['Saldo Risk Peak'] = df_building_data_filtered['Saldo Risk Peak'].round(0).astype(int)

                    df_building_data_filtered['Risk Avg Occ'] = df_building_data_filtered['Avg Occ'] * (1 - risk_value / 100)
                    df_building_data_filtered['Saldo Risk Avg Occ'] = df_building_data_filtered['Primary Work Seats'] - df_building_data_filtered['Risk Avg Occ']
                    df_building_data_filtered['Risk Avg Occ'] = df_building_data_filtered['Risk Avg Occ'].round(0).astype(int)
                    df_building_data_filtered['Saldo Risk Avg Occ'] = df_building_data_filtered['Saldo Risk Avg Occ'].round(0).astype(int)

                    df_building_data_filtered.rename(columns={"Lugares Disponíveis 1:1" : "Saldo 1:1", "Lugares Ocupados 1:1" : "Occupied 1:1", 
                                                              "Lugares Ocupados Peak" : "Occupied Peak", "Lugares Disponíveis Peak" : "Saldo Peak" , 
                                                              "Lugares Ocupados Avg" : "Occupied Avg", "Lugares Disponíveis Avg" : "Saldo Avg Occ"}, inplace=True)

                                    
                    ## Para exibir a tabela com os calculos
                    #columns_to_display = ['Building Name', 'Group', 'SubGroup', 'Exception (Y/N)', '1:1', 'Occupied 1:1' ,'Saldo 1:1', 'Peak', 'Occupied Peak', 'Saldo Peak',
                    #                      'Avg Occ', 'Occupied Avg', 'Saldo Avg Occ']
                    #if growth_value != 0:
                    #    columns_to_display += ['Growth 1:1', 'Saldo Growth 1:1', 'Growth Peak', 'Saldo Growth Peak', 'Growth Avg Occ', 'Saldo Growth Avg Occ']                
                    #if risk_value != 0:
                    #    columns_to_display += ['Risk 1:1', 'Saldo Risk 1:1', 'Risk Peak', 'Saldo Risk Peak', 'Risk Avg Occ', 'Saldo Risk Avg Occ']
                    # Salvando o DataFrame filtrado em uma variável
                    #df_filtered_output = df_building_data_filtered[columns_to_display]
                    #st.dataframe(df_filtered_output, use_container_width=False)


                    #### COLOROÇÃO DO PLANO DE FUNDO APENAS PARA DESTACAR QUE ESTAMOS ADICIONANDO COLUNAS DE GROWTH E RISK
                    columns_to_display_filter = ['Building Name', 'Group', 'SubGroup', 'Exception (Y/N)', '1:1', 'Saldo 1:1', 'Peak', 'Saldo Peak', 'Avg Occ', 'Saldo Avg Occ']
                    # Inicialização das colunas de Growth e Risk
                    columns_growth = ['Saldo Growth 1:1', 'Saldo Growth Peak', 'Saldo Growth Avg Occ']
                    columns_risk = ['Saldo Risk 1:1', 'Saldo Risk Peak', 'Saldo Risk Avg Occ']

                    # Condicional para adicionar as colunas de Growth e Risk
                    if growth_value != 0:
                        columns_to_display_filter += columns_growth

                    if risk_value != 0:
                        columns_to_display_filter += columns_risk

                    # Função para colorir as células com base no tipo de dado
                    def colorize(val):
                        if isinstance(val, (int, float)):
                            if val >= 0:
                                return 'background-color: white'  # Verde claro para valores positivos
                            elif val < 0:
                                return 'background-color: #FFBDBD'  # Coral para valores negativos
                        return 'background-color: white'  # Padrão branco para valores não numéricos

                    # Função para aplicar fundo cinza para Growth e Risk
                    def colorize_growth_risk(val, columns_type):
                        if isinstance(val, (int, float)):  # Verifica se o valor é numérico
                            if columns_type == 'Growth':  # Aplica para as colunas de Growth
                                if val >= 0 or val < 0:  # Valores maiores ou iguais a zero (para Growth)
                                    return 'background-color: #EDEDED'  # Fundo cinza claro
                            elif columns_type == 'Risk':  # Aplica para as colunas de Risk
                                if val >= 0 or val < 0:  # Valores menores ou iguais a zero (para Risk)
                                    return 'background-color: #DDEBF7'  # Fundo azul claro
                        return 'background-color: white'  # Para valores não numéricos ou outras condições

                    # Aplicando o estilo para as colunas de Growth, Risk e outras
                    styled_df = df_building_data_filtered[columns_to_display_filter].style

                    # Aplicando o estilo condicional para Growth
                    if growth_value != 0:
                        styled_df = styled_df.applymap(lambda val: colorize_growth_risk(val, 'Growth'), subset=columns_growth)

                    # Aplicando o estilo condicional para Risk
                    if risk_value != 0:
                        styled_df = styled_df.applymap(lambda val: colorize_growth_risk(val, 'Risk'), subset=columns_risk)

                    # Aplicando o estilo para as outras colunas (não Growth e Risk)
                    styled_df = styled_df.applymap(colorize, subset=['1:1', 'Saldo 1:1', 'Peak', 'Saldo Peak', 'Avg Occ', 'Saldo Avg Occ'])

                    # Exibindo a tabela com o estilo aplicado
                    st.dataframe(styled_df, use_container_width=False, hide_index=True)

                    # Botão de Gravar Dados
                    if st.button(f"Gravar Dados para {building}"):
                        # Verificar se o Building Name já existe na lista
                        building_exists = False
                        for idx, df in enumerate(st.session_state.tables_to_append):
                            if df['Building Name'].iloc[0] == building:
                                # Se o 'Building Name' já existir, substituímos os dados dessa seção
                                st.session_state.tables_to_append[idx] = styled_df.data  # Sobrescreve os dados
                                building_exists = True
                                break
                        
                        # Se o 'Building Name' não existir, adicionamos os novos dados à lista
                        if not building_exists:
                            st.session_state.tables_to_append.append(styled_df.data) 

        
        else:
            st.write("Por favor, carregue o arquivo para prosseguir.")

        # Ao final, exibir o expander para mostrar a tabela consolidada
        with st.expander("### **Resultado de todos os Cenários:**"):

            # Inicializar df_proportional como um DataFrame vazio, se não houver dados na sessão
            if "df_proportional_cenarios" not in st.session_state:
                df_proportional_cenarios = pd.DataFrame()  # DataFrame vazio
            else:
                df_proportional_cenarios = st.session_state.df_proportional_cenarios
                

            # Verificar se o df_proportional tem dados antes de continuar
            if not df_proportional_cenarios.empty:
                # Criação da tabela de cenários
                df_proportional_cenarios = df_proportional_cenarios.copy()

                # Se houver tabelas para consolidar
                if st.session_state.tables_to_append:
                    # Faz o append de todas as tabelas na lista
                    final_consolidated_df = pd.concat(st.session_state.tables_to_append, ignore_index=True)
                    
                    def round_and_convert_to_int(df):
                        # Seleciona apenas as colunas numéricas
                        numeric_columns = df.select_dtypes(include=['number']).columns
                        df[numeric_columns] = df[numeric_columns].replace([np.inf, -np.inf, np.nan], 0)
                        df[numeric_columns] = df[numeric_columns].round(0).astype(int)                
                        return df
                    final_consolidated_df = round_and_convert_to_int(final_consolidated_df)

                    # Aplicando o estilo para as colunas
                    final_consolidated_df_colour = final_consolidated_df.style
                    def colorize(val):
                        if isinstance(val, (int, float)):
                            if val >= 0:
                                return 'background-color: white'  # Verde claro para valores positivos
                            elif val < 0:
                                return 'background-color: #FFBDBD'  # Coral para valores negativos
                        return 'background-color: white'  # Padrão branco para valores não numéricos
                    final_consolidated_df_colour = final_consolidated_df_colour.applymap(colorize, subset=['1:1', 'Saldo 1:1', 'Peak', 'Saldo Peak', 'Avg Occ', 'Saldo Avg Occ'])
                    

                    # Exibe a tabela consolidada
                    st.write("#### **Consolidado de todos os cenários:**")
                    st.dataframe(final_consolidated_df_colour, use_container_width=True,  hide_index=True)

                    # Realizando o DISTINCT para obter os Grupos e Subgrupos não alocados
                    # Criando uma tabela de chave (Group - SubGroup) em final_consolidated_df
                    consolidated_groups_subgroups = final_consolidated_df[['Group', 'SubGroup']].drop_duplicates()

                    # Criando uma tabela de chave (Group - SubGroup) em df_proportional_cenarios
                    proportional_groups_subgroups = df_proportional_cenarios.copy()
                    proportional_groups_subgroups = proportional_groups_subgroups.drop('Lugares Ocupados 1:1', axis=1)

                    # Realizando a diferença entre os Grupos/Subgrupos
                    # Usamos 'merge' para identificar quais não estão em final_consolidated_df
                    df_non_allocated = proportional_groups_subgroups.merge(consolidated_groups_subgroups,
                                                                            how='left', 
                                                                            indicator=True).query('_merge == "left_only"').drop('_merge', axis=1)

                    # Exibindo os Grupos e Subgrupos não alocados
                    st.write("#### Grupos e Subgrupos Não Alocados")
                    st.dataframe(df_non_allocated, use_container_width=False, hide_index=True)

                    # Botão para exportar tabela "Cenários" para Excel
                    if st.button("Exportar 'Cenários' para Excel", key="export_cenarios_excel"):
                        with io.BytesIO() as output:
                            # Criação do ExcelWriter
                            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                                # Exportando final_consolidated_df para a aba "Cenarios"
                                final_consolidated_df = st.session_state.final_consolidated_df
                                final_consolidated_df = final_consolidated_df.fillna("")  # Substituindo NaN por ""
                                final_consolidated_df.to_excel(writer, sheet_name="Cenarios", index=False)

                                # Exportando df_non_allocated para a aba "Não Alocados"
                                df_non_allocated = st.session_state.df_non_allocated
                                df_non_allocated = df_non_allocated.fillna("")  # Substituindo NaN por ""
                                df_non_allocated.to_excel(writer, sheet_name="Não Alocados", index=False)

                            # Exibindo o botão de download
                            st.download_button(
                                label="Download do Excel - Cenários e Não Alocados",
                                data=output.getvalue(),
                                file_name="Cenarios_e_Nao_Alocados.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    

            else:
                st.write("Nenhum dado foi gravado ainda.")


        if st.button("Resetar Cenários"):
            # Limpando todos os dados armazenados no session_state
            st.session_state.clear()

            # Mensagem de confirmação
            st.success("Simulação resetada com sucesso!")

            # Opcional: Exibir algo ou recarregar a página, se necessário
            # Aqui você pode adicionar qualquer ação extra ou recarregar a interface
            st.rerun()  # Força a página a reiniciar
        



   ##### ABA DASHBOARDS #####
    with tabs[4]:
        st.write("### DASHBOARDS")
 

# Tela Inicial com Seleção
st.title("Calculadora FRB - Alocação")
st.write("""
Aqui você escolhe a opção se realizar as alocações por Upload de Excel ou para Input das informações diretamente aqui pela Web.
""")
opcao = st.selectbox("Escolha uma opção", ["Selecione", "Upload de Arquivo"])

if opcao == "Upload de Arquivo":
    upload_arquivo()



