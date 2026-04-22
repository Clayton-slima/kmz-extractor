import streamlit as st
import zipfile
import pandas as pd
from bs4 import BeautifulSoup
import io

def processar_kmz(arquivo_kmz):
    # O KMZ é um arquivo ZIP. Vamos abri-lo e procurar o KML lá dentro.
    with zipfile.ZipFile(arquivo_kmz, 'r') as kmz:
        nomes_arquivos = kmz.namelist()
        kml_filename = [nome for nome in nomes_arquivos if nome.endswith('.kml')][0]
        
        with kmz.open(kml_filename, 'r') as kml_file:
            conteudo_kml = kml_file.read()

    soup = BeautifulSoup(conteudo_kml, 'xml')
    dados = []

    # Procura por todas as marcações (Placemarks)
    for placemark in soup.find_all('Placemark'):
        ponto = placemark.find('Point')
        if ponto:
            # --- NOVA LÓGICA: IDENTIFICANDO A PASTA ---
            # Busca se este ponto está dentro de uma tag <Folder>
            pasta_pai = placemark.find_parent('Folder')
            
            if pasta_pai:
                # Pega o nome da pasta (a primeira tag <name> logo abaixo de <Folder>)
                tag_nome_pasta = pasta_pai.find('name')
                nome_pasta = tag_nome_pasta.text.strip() if tag_nome_pasta else 'Pasta sem nome'
            else:
                nome_pasta = 'Geral (Raiz)'
            # ------------------------------------------

            # Extrai o nome da unidade
            nome_tag = placemark.find('name')
            nome = nome_tag.text.strip() if nome_tag else 'Ponto sem nome'
            
            # Extrai as coordenadas
            coords_tag = ponto.find('coordinates')
            if coords_tag:
                coords_texto = coords_tag.text.strip()
                partes = coords_texto.split(',')
                if len(partes) >= 2:
                    longitude = float(partes[0].strip())
                    latitude = float(partes[1].strip())
                    
                    dados.append({
                        'Pasta': nome_pasta, # Nova coluna adicionada aqui!
                        'Nome da Unidade': nome,
                        'Latitude': latitude,
                        'Longitude': longitude
                    })

    return pd.DataFrame(dados)

# --- INTERFACE DO STREAMLIT ---
st.set_page_config(page_title="Extrator KMZ para Excel", page_icon="🌍")

st.title("🌍 Extrator de Pontos: KMZ para Excel")
st.write("Faça o upload do seu arquivo `.kmz` do Google Earth para extrair as pastas, latitudes, longitudes e nomes em uma planilha.")

arquivo_upload = st.file_uploader("Selecione o arquivo .kmz", type=['kmz'])

if arquivo_upload is not None:
    try:
        df = processar_kmz(arquivo_upload)
        
        if df.empty:
            st.warning("Nenhum ponto (Point) foi encontrado neste arquivo KMZ.")
        else:
            st.success(f"{len(df)} pontos extraídos com sucesso!")
            
            st.dataframe(df, use_container_width=True)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Pontos Extraídos')
            
            st.download_button(
                label="📥 Baixar Planilha Excel",
                data=buffer.getvalue(),
                file_name="pontos_google_earth.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")