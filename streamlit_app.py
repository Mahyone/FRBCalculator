
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
                df_building.rename(columns={df_building.columns[7]: 'Alternative Work Seats'}, inplace=True)
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

        # Carrega os dados
        df_building_trat = st.session_state.get('df_building_trat', pd.DataFrame())
        df_proportional  = st.session_state.get('df_proportional', pd.DataFrame())
        if df_building_trat.empty or df_proportional.empty:
            st.info("Carregue primeiro os dados em 'Building' e 'Proportional'.")
            st.stop()

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



        # Prepara dicionário de capacidade
        floors0 = dict(zip(df_building_trat['Building Name'], df_building_trat['Primary Work Seats']))

        # Função genérica de alocação com adjacentes
        def allocate_with_adj(df, floors, eff_col):
            df2 = df.copy().reset_index(drop=True)
            df2['Building Name'] = 'Não Alocado'
            df2['SubGroup'] = df2['SubGroup'].fillna('NoSubGroup')
            for idx, row in df2.sort_values(eff_col, ascending=False).iterrows():
                if df2.at[idx,'Building Name']!='Não Alocado': continue
                # monta lote principal + adjacentes
                lote = [idx]
                for adj_col in ['Adjacency Priority 1','Adjacency Priority 2','Adjacency Priority 3']:
                    adj = row.get(adj_col)
                    if pd.notna(adj):
                        mask = (
                            (df2['Group']==row['Group']) &
                            (df2['SubGroup']==adj) &
                            (df2['Building Name']=='Não Alocado')
                        )
                        lote += df2.index[mask].tolist()
                soma = df2.loc[lote, eff_col].sum()
                # tenta em bloco
                coloc = False
                for fl, cap in floors.items():
                    if cap >= soma:
                        df2.loc[lote,'Building Name'] = fl
                        floors[fl] -= soma
                        coloc = True
                        break
                # fallback só principal
                if not coloc:
                    hc = int(row[eff_col])
                    for fl, cap in floors.items():
                        if cap >= hc:
                            df2.at[idx,'Building Name'] = fl
                            floors[fl] -= hc
                            break
            return df2, floors

        # Função genérica de exibição
        def display_generic(df_alloc, floors_rem, dfb, key_col):
            # -- subtotais e estilo igual ao HeadCount --
            cols = df_alloc.columns.tolist()
            if "Current Location" in cols:
                df_alloc = df_alloc[["Building Name"] + [c for c in cols if c not in ("Building Name","Current Location")] + ["Current Location"]]
            df_alloc = df_alloc.sort_values("Building Name")
            num_cols = df_alloc.select_dtypes("number").columns
            a = df_alloc[df_alloc["Building Name"]!="Não Alocado"]
            na = df_alloc[df_alloc["Building Name"]=="Não Alocado"]
            soma_a = a[num_cols].sum();      soma_a["Building Name"]="Alocados"
            soma_na= na[num_cols].sum();     soma_na["Building Name"]="Não Alocados"
            soma_t = df_alloc[num_cols].sum();soma_t["Building Name"]="Total Geral"
            df_tot = pd.concat([df_alloc, pd.DataFrame([soma_a,soma_na,soma_t])], ignore_index=True)

            # cabeçalho
            st.write(f"#### Resultado da Automação - {key_col}")
            def hl(r):
                n = r["Building Name"]
                if n=="Total Geral":   return ["background:#004E64;color:white"]*len(r)
                if n in ("Alocados","Não Alocados"): return ["background:#357A91;color:white"]*len(r)
                return [""]*len(r)
            st.dataframe(df_tot.style.apply(hl,axis=1), use_container_width=False)

            # -- capacidade restante pelo Primary Work Seats --
            cap = dfb[["Building Name","Primary Work Seats"]].rename(columns={"Primary Work Seats":"Capacity"})
            occ = (
                df_alloc[df_alloc["Building Name"]!="Não Alocado"]
                .groupby("Building Name")[key_col].sum()
                .reset_index(name="Occupied")
            )
            rem = cap.merge(occ, on="Building Name", how="left").fillna(0)
            rem["Remaining"] = (rem["Capacity"] - rem["Occupied"]).clip(lower=0)
            tot = rem[["Capacity","Occupied","Remaining"]].sum(); tot["Building Name"]="Total Geral"
            rem = pd.concat([rem, pd.DataFrame([tot])], ignore_index=True)

            st.write(f"#### Capacidade restante nos andares - {key_col}:")
            st.dataframe(
                rem.style.apply(lambda r: ['background:#004E64;color:white']*len(r) if r["Building Name"]=="Total Geral" else ['']*len(r),
                                axis=1),
                use_container_width=False
            )
            return df_tot, rem

        # Define os 3 cenários
        scenarios = [
            {"title":"HeadCount", "eff_col":"HeadCount"},
            {"title":"Peak",      "eff_col":"Peak with Exception"},
            {"title":"Avg Occ",   "eff_col":"Avg Occ with Exception"}
        ]

        # Prepara coluna efetiva em df_proportional para cada cenário
        dfp = df_proportional.copy()
        # Peak exception
        dfp["Peak with Exception"] = dfp.apply(lambda r: r["HeadCount"] if r["Exception (Y/N)"]=="Y" else r["Proportional Peak"], axis=1)
        # Avg Occ exception
        dfp["Avg Occ with Exception"] = dfp.apply(lambda r: r["HeadCount"] if r["Exception (Y/N)"]=="Y" else r["Proportional Avg"], axis=1)

        # Loop pelos cenários
        cols_map = {
            "HeadCount": [
                "Building Name","Group","SubGroup","FTE","CW","Growth",
                "HeadCount","Exception (Y/N)",
                "Proportional Peak","Proportional Avg",
                "Adjacency Priority 1","Adjacency Priority 2","Adjacency Priority 3",
                "Current Location"
            ],
            "Peak": [
                "Building Name","Group","SubGroup","FTE","CW","Growth",
                "HeadCount","Exception (Y/N)",
                "Peak with Exception","Peak % of HeadCount",
                "Proportional Avg",
                "Adjacency Priority 1","Adjacency Priority 2","Adjacency Priority 3",
                "Current Location"
            ],
            "Avg Occ": [
                "Building Name","Group","SubGroup","FTE","CW","Growth",
                "HeadCount","Exception (Y/N)",
                "Avg Occ with Exception","Avg Occ % of HeadCount",
                "Adjacency Priority 1","Adjacency Priority 2","Adjacency Priority 3",
                "Current Location"
            ]
        }

        
        primary_work_seats     = int(df_building_trat['Primary Work Seats'].sum())
        alternative_work_seats = int(df_building_trat['Alternative Work Seats'].sum())
        total_seats_on_floor   = int(df_building_trat['Total seats on floor'].sum())

        for sc in scenarios:
            title   = sc["title"]
            eff_col = sc["eff_col"]

            with st.expander(f"### Automação considerando {title}"):
                # 0) cabeçalho de capacidades
                st.write(
                    f"**Primary Work Seats**: {primary_work_seats}  ||  "
                    f"**Alternative Work Seats**: {alternative_work_seats}  ||  "
                    f"**Total seats on floor**: {total_seats_on_floor}"
                )

                # 0.1) totais específicos de Peak/Avg Occ
                if title == "Peak":
                    total_pp = int(df_proportional["Proportional Peak"].sum())
                    total_pp_exc = int(dfp["Peak with Exception"].sum())
                    st.write(f"**Peak**: {total_pp}  ||  **Peak with Exception**: {total_pp_exc}")
                elif title == "Avg Occ":
                    total_pa = int(df_proportional["Proportional Avg"].sum())
                    total_pa_exc = int(dfp["Avg Occ with Exception"].sum())
                    st.write(f"**Avg Occ**: {total_pa}  ||  **Avg Occ with Exception**: {total_pa_exc}")
                else:  # HeadCount
                    st.write(f"**HeadCount**: {int(dfp['HeadCount'].sum())}")

                # 1) alocar
                floors = floors0.copy()
                df_alloc, floors_rem = allocate_with_adj(dfp, floors, eff_col)

                # 2) criar as colunas de exception e percentuais (já existia)
                df_alloc["Peak with Exception"]     = df_alloc.apply(
                    lambda r: r["HeadCount"] if r["Exception (Y/N)"]=="Y" else r["Proportional Peak"],
                    axis=1
                )
                df_alloc["Peak % of HeadCount"]     = (
                    df_alloc["Peak with Exception"] / df_alloc["HeadCount"] * 100
                ).round(0).astype(int)
                df_alloc["Avg Occ with Exception"]  = df_alloc.apply(
                    lambda r: r["HeadCount"] if r["Exception (Y/N)"]=="Y" else r["Proportional Avg"],
                    axis=1
                )
                df_alloc["Avg Occ % of HeadCount"]  = (
                    df_alloc["Avg Occ with Exception"] / df_alloc["HeadCount"] * 100
                ).round(0).astype(int)


                # 3) Extrair só as colunas desejadas
                cols = cols_map[title]
                df_display = df_alloc[cols]

                # === 3.1) Ordena e arredonda df_display ===
                df_display = df_display.sort_values(
                    ["Building Name","Group","SubGroup"], ascending=True
                )
                num_cols_disp = df_display.select_dtypes("number").columns
                df_display[num_cols_disp] = df_display[num_cols_disp].round(0).astype(int)

                # 4) subtotais + total geral
                num_cols = df_display.select_dtypes("number").columns
                grp_a    = df_display[df_display["Building Name"]!="Não Alocado"]
                grp_na   = df_display[df_display["Building Name"]=="Não Alocado"]
                soma_a   = grp_a[num_cols].sum();   soma_a["Building Name"]="Alocados"
                soma_na  = grp_na[num_cols].sum();  soma_na["Building Name"]="Não Alocados"
                soma_t   = df_display[num_cols].sum();soma_t["Building Name"]="Total Geral"
                df_tot = pd.concat([df_display, pd.DataFrame([soma_a, soma_na, soma_t])], ignore_index=True)
                num_cols_tot = df_tot.select_dtypes("number").columns
                df_tot[num_cols_tot] = df_tot[num_cols_tot].round(0).astype(int)

                # 5) Estilização única
                unique_b = df_display["Building Name"].drop_duplicates().tolist()
                building_colors = {b: "#D3D3D3" if i%2==0 else "" for i,b in enumerate(unique_b)}

                def highlight_rows(row):
                    name = row["Building Name"]
                    if name == "Total Geral":
                        return ["background-color: #004E64; color: white"] * len(row)
                    if name in ("Alocados","Não Alocados"):
                        return ["background-color: #357A91; color: white"] * len(row)
                    return [f"background-color: {building_colors.get(name,'')};"] * len(row)

                # **Adicione isto antes de usar highlight_total**
                def highlight_total(row):
                    if row["Building Name"] == "Total Geral":
                        return ["background-color: #004E64; color: white"] * len(row)
                    return [""] * len(row)

                # 6) exibir alocação completa
                st.write(f"#### Resultado da Automação - {title}")
                st.dataframe(df_tot.style.apply(highlight_rows,axis=1), hide_index=True)

                # 7) capacidade restante
                st.write(f"#### Capacidade restante nos andares - {title}:")
                cap = (
                    df_building_trat[["Building Name","Primary Work Seats"]]
                    .rename(columns={"Primary Work Seats":"Capacity"})
                )
                occ = (
                    df_alloc[df_alloc["Building Name"]!="Não Alocado"]
                    .groupby("Building Name")[eff_col]
                    .sum()
                    .reset_index(name="Occupied")
                )

                rem = cap.merge(occ,on="Building Name",how="left").fillna(0)
                rem["Remaining"] = (rem["Capacity"]-rem["Occupied"]).clip(lower=0)
                tot = rem[["Capacity","Occupied","Remaining"]].sum(); tot["Building Name"]="Total Geral"
                rem = pd.concat([rem,pd.DataFrame([tot])],ignore_index=True)
                rem = rem.round(0).astype({"Capacity":int,"Occupied":int,"Remaining":int})
                rem = rem.sort_values("Building Name", ascending=True)
                st.dataframe(rem.style.apply(highlight_total,axis=1), hide_index=True)

                # 8) tabela de Grupos Não Alocados
                st.write(f"#### Grupos e SubGrupos não alocados - {title}:")
                nao_df = df_tot[df_tot["Building Name"]=="Não Alocado"].copy()
                nao_df = nao_df.sort_values(["Building Name","Group","SubGroup"], ascending=True)
                st.dataframe(nao_df.style.apply(highlight_rows,axis=1), hide_index=True)

                # 9) exportar para Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    df_tot.to_excel(writer, sheet_name=title, index=False)
                output.seek(0)
                st.download_button(
                    label=f"Download Excel - {title}",
                    data=output.getvalue(),
                    file_name=f"automacao_{title.lower().replace(' ','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{title}"
                )

                # 10) salvar no session_state
                key = f"dfautomation_{title.lower().replace(' ','_')}"
                st.session_state[key] = df_tot

                


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
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.markdown(f"<div style='background:#004E64;color:white;padding:10px;text-align:center'><b># Andares</b><br>{df_building['Building Name'].nunique()}</div>", unsafe_allow_html=True)
        c2.markdown(f"<div style='background:#B0B0B0;color:black;padding:10px;text-align:center'><b># Groups</b><br>{df_hc['Group'].nunique()}</div>", unsafe_allow_html=True)
        c3.markdown(f"<div style='background:#004E64;color:white;padding:10px;text-align:center'><b># Groups+SubGroups</b><br>{df_hc[['Group','SubGroup']].drop_duplicates().shape[0]}</div>", unsafe_allow_html=True)
        c4.markdown(f"<div style='background:#B0B0B0;color:black;padding:10px;text-align:center'><b>Total Primary Seats</b><br>{df_building['Primary Work Seats'].sum()}</div>", unsafe_allow_html=True)
        c5.markdown(f"<div style='background:#004E64;color:white;padding:10px;text-align:center'><b>Total Floor Seats</b><br>{df_building['Total seats on floor'].sum()}</div>", unsafe_allow_html=True)

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
        

        def plot_donut(df_base, key_col, title):
            df2 = df_base.copy()
            df2['Status'] = df2['Building Name'].apply(lambda x:'Alocado' if x!='Não Alocado' else 'Não Alocado')
            by_st = df2.groupby('Status')[key_col].sum().reset_index()
            fig = px.pie(
                by_st, names='Status', values=key_col, hole=0.3,
                color='Status',
                color_discrete_map={'Alocado':'#357A91','Não Alocado':'#FFAA33'},
                title=f"% Alloc {title}"
            )
            return fig

        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("<h5 style='text-align:center'>Comparativo Consolidado</h5>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)

        for df_data, key_col, title, cont in zip(
            [df_hc, df_peak, df_avg],
            ['HeadCount','Peak with Exception','Avg Occ with Exception'],
            ['HeadCount','Peak','Avg Occ'],
            [col1, col2, col3]
        ):
            with cont:
                # totais + donut
                df_full, resumo = prepare_avail(df_data, key_col)
                st.dataframe(style_summary_table(resumo), use_container_width=True, hide_index=True)
                fig = plot_donut(df_data, key_col, title)
                st.plotly_chart(fig, use_container_width=True, key=f"donut_{title}")

                # 1) Tabela detalhada de grupos
                df_detail = df_full.copy()
                if title=='HeadCount':
                    raw = 'HeadCount'
                    df_final = df_detail.groupby(['Building Name','Group','SubGroup'], as_index=False)[[raw]].sum()
                else:
                    raw = key_col
                    pct = f"{title} %"
                    tot = df_detail[raw].sum()
                    df_detail[pct] = ((df_detail[raw]/tot)*100).round(0).astype(int)
                    df_final = df_detail.groupby(['Building Name','Group','SubGroup'], as_index=False)[[raw,pct]].sum()

                st.dataframe(df_final, use_container_width=True, hide_index=True)


                # 2) Andares com capacidade disponível
                cap_df = df_building[['Building Name','Primary Work Seats']]
                used  = df_full.groupby('Building Name')[key_col].sum().reset_index(name='Total Occupied')
                cap   = cap_df.merge(used, on='Building Name', how='left').fillna(0)
                cap['Total Available'] = cap['Primary Work Seats'] - cap['Total Occupied']
                cap_rem = cap[cap['Total Available']>0]
                st.markdown(f"##### Andares com capacidade disponível - {title}", unsafe_allow_html=True)
                st.dataframe(cap_rem[['Building Name','Primary Work Seats','Total Occupied','Total Available']], use_container_width=True, hide_index=True)                
                

        st.markdown("<br><br>---", unsafe_allow_html=True)




        # === DEEP DIVE ===
        tabela = st.selectbox('Escolha cenário para Deep Dive:', ['HeadCount','Peak','Avg Occ'])

        def render_deep(df_base, key_col, title):
            st.write(f"### Deep Dive: {title}")

            # --- 1) Donut com label, valor e % no hover e pull para fatias pequenas ---
            grp = df_base.groupby('Group')[key_col].sum().reset_index()
            total = grp[key_col].sum()
            grp['percent'] = (grp[key_col] / total * 100).round(0).astype(int)

            # computa pull: fatias <5% puxam 0.1
            pulls = [0.1 if p < 5 else 0 for p in grp['percent']]

            fig1 = px.pie(
                grp,
                names='Group',
                values=key_col,
                hole=0.3,
                title=f"Distribuição {title} por Group",
                hover_data=[key_col, 'percent'],                # mostra valor e %
                labels={key_col: 'Valor', 'percent': '%'},
            )
            fig1.update_traces(
                texttemplate='%{label}<br>%{value} || %{percent}%',
                textposition='inside',
                pull=pulls,
                hovertemplate=(
                    "<b>%{label}</b><br>" +
                    key_col + ": %{value}<br>" +
                    "%{percent}%<extra></extra>"
                ),
                textfont_size=10
            )
            cA, cB = st.columns([3,1])
            cA.plotly_chart(fig1, use_container_width=True, key=f"deep_donut_{title}")

            # tabela ao lado
            grp_table = grp.rename(columns={key_col: 'Total', 'percent': f'{title} %'})
            grp_table = grp_table[['Group', 'Total', f'{title} %']].sort_values('Group')
            cB.dataframe(grp_table, use_container_width=True, hide_index=True)

            st.markdown("<br><br>", unsafe_allow_html=True)

            # --- 2) Gráfico stacked horizontal com anotação de "Occupied || Primary || Total seats" ---
            st.markdown(
                "_Os números ao final de cada barra correspondem a:_  \n"
                "Occupied** ― total ocupado  \n"
                "Primary Work Seats** ― capacidade primária  \n"
                "Total seats on floor** ― total de assentos no andar  \n"
                "- Lembre-se que a alocação respeita o Primary Work Seats._"
            )

            df_counts = df_base.groupby(['Building Name','Group'], as_index=False)[key_col].sum()
            pivot = df_counts.pivot(index='Building Name', columns='Group', values=key_col).fillna(0)

            # garante todos os andares
            for b in df_building['Building Name'].unique():
                if b not in pivot.index:
                    pivot.loc[b] = [0]*pivot.shape[1]
            pivot = pivot.sort_index()

            pct = pivot.div(pivot.sum(axis=1), axis=0)*100
            long = pct.reset_index().melt('Building Name', var_name='Group', value_name='Percent')
            long = long.merge(df_counts, on=['Building Name','Group'])

            ord_b = list(pivot.index)
            ord_g = sorted(df_counts['Group'].unique())

            # mapeia capacidades
            cap_map = df_building.set_index('Building Name')['Total seats on floor'].to_dict()
            prim_map = df_building.set_index('Building Name')['Primary Work Seats'].to_dict()

            fig2 = px.bar(
                long,
                x='Percent',
                y='Building Name',
                color='Group',
                orientation='h',
                category_orders={'Building Name': ord_b, 'Group': ord_g},
                text=key_col,
                title=f'% {title} por Grupo e Andar'
            )
            # labels dentro das divisões
            fig2.update_traces(texttemplate='%{text}', textposition='inside')
            # anotações fora das barras
            for y_val in ord_b:
                used = int(df_counts.loc[df_counts['Building Name']==y_val, key_col].sum())
                prim = prim_map.get(y_val, 0)
                total_fl = cap_map.get(y_val, 0)
                fig2.add_annotation(
                    x=100, y=y_val,
                    text=f"{used} || {prim} || {total_fl}",
                    showarrow=False,
                    xanchor='left',
                    font=dict(size=9)
                )
            fig2.update_layout(
                barmode='stack',
                xaxis=dict(ticksuffix='%'),
                margin=dict(l=100, r=100, t=40, b=40),
                height=600
            )

            bar_col, tab_col = st.columns([3,1])
            bar_col.plotly_chart(fig2, use_container_width=True, key=f"deep_bar_{title}")

            # --- 3) Tabela lateral: todos os campos solicitados ---
            grp_cnt = (
                df_base[['Building Name','Group','SubGroup']]
                .drop_duplicates()
                .groupby('Building Name')
                .size()
                .reset_index(name='Total Groups+SubGroups')
            )
            occ = df_base.groupby('Building Name')[key_col].sum().reset_index(name='Total Occupied')
            prim = df_building[['Building Name','Primary Work Seats']].rename(
                columns={'Primary Work Seats':'Primary Work Seats'}
            )
            total_fl = df_building[['Building Name','Total seats on floor']]
            table = (
                grp_cnt
                .merge(occ, on='Building Name')
                .merge(prim, on='Building Name')
                .merge(total_fl, on='Building Name')
            )
            table['Total Available'] = table['Total seats on floor'] - table['Total Occupied']
            table = table[
                ['Building Name','Total Groups+SubGroups','Total Occupied',
                'Primary Work Seats','Total seats on floor','Total Available']
            ].sort_values(['Building Name'])
            tab_col.dataframe(table, use_container_width=True, hide_index=True)

        # chama o deep-dive
        if tabela=='HeadCount':
            render_deep(df_hc, 'HeadCount', 'HeadCount')
        elif tabela=='Peak':
            render_deep(df_peak, 'Peak with Exception', 'Peak')
        else:
            render_deep(df_avg, 'Avg Occ with Exception', 'Avg Occ')


        # === GRÁFICOS DE CAPACIDADE ===
        st.markdown("<br><br><h5 style='text-align:center'>Capacidade por Tipo de Assentos</h5>", unsafe_allow_html=True)
        df_seats = df_building.melt(
            id_vars='Building Name',
            value_vars=[
                'Primary Work Seats','Alternative Work Seats',
                'Total Enclosed Collab Seats','Total Open Collab Seats'
            ],
            var_name='Seat Type',
            value_name='Available Seats'
        )

        fig3 = px.bar(
            df_seats,
            x='Available Seats',
            y='Building Name',
            color='Seat Type',
            orientation='h',
            title='Disponibilidade por Tipo de Assentos',
            labels={'Available Seats':'Total Available Seats'}
        )

        # 1) Esconde os textos de cada segmento
        fig3.update_traces(
            texttemplate='%{x}',
            textposition='inside'
        )

        # 2) Calcula o total por andar
        totals = df_seats.groupby('Building Name')['Available Seats'].sum()

        # 3) Adiciona annotation com o total no final de cada barra
        for bld, tot in totals.items():
            fig3.add_annotation(
                x=tot, 
                y=bld,
                text=str(int(tot)),
                showarrow=False,
                xanchor='left',
                font=dict(size=10)
            )

        # 4) Ajustes de layout
        fig3.update_layout(
            height=500,
            margin=dict(l=100, r=50, t=40, b=40)
        )

        st.plotly_chart(fig3, use_container_width=True, key='cap_tipo_assentos')


        st.markdown("<br><br><h5 style='text-align:center'>Total de Assentos por Andar</h5>", unsafe_allow_html=True)
        df_floor = df_building[[
            'Building Name','Total Individual seats on floor','Total Collab seats on floor'
        ]].melt(
            id_vars='Building Name',
            var_name='Seat Category',
            value_name='Seats'
        )
        fig4 = px.bar(
            df_floor,
            x='Building Name',
            y='Seats',
            color='Seat Category',
            barmode='group',
            title='Assentos Individuais vs Colaborativos por Andar'
        )
        # exibe labels de cada barra
        fig4.update_traces(texttemplate='%{y}', textposition='auto')
        fig4.update_layout(height=500, margin=dict(l=80,r=50,t=40,b=40))
        st.plotly_chart(fig4, use_container_width=True, key='cap_andar_tipo')


# Tela Inicial com Seleção
st.title("Calculadora FRB - Alocação")
st.write("""
Aqui você escolhe a opção se realizar as alocações por Upload de Excel ou para Input das informações diretamente aqui pela Web.
""")
opcao = st.selectbox("Escolha uma opção", ["Selecione", "Upload de Arquivo"])

if opcao == "Upload de Arquivo":
    upload_arquivo()
