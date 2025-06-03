import streamlit as st
import pandas as pd
import requests
from io import BytesIO

# 🔐 Configurações dos clientes
WAREHOUSES_POR_CLIENTE = {
    "dunorte": ["1"],
    "cardeal": ["1"],
    "multigiro": ["4", "2"],
    "unilider": ["1001", "1004", "1003"],
    "supergiro": ["1"]
}

TOKENS_POR_CLIENTE = {
    "cardeal": "hrIGQKdoLbKoC3fWKSAKikTXDI9fdaSF",
    "dunorte": "9dPBMtw04jjlUtLa9VMB6CntAZJkt8gk",
    "multigiro": "89qv9YTfGBFMPZVSxm7O5kGgi3B3blQA",
    "supergiro": "XYfUuOrvKhMnyH2mEmDIaHjiNZwx1OjS",
    "unilider": "hhVluD2bw7MyOo6sdS6mgWE8ny6lCmNJ"
}

# 🔎 Detecta automaticamente a coluna que contém CNPJ
def detectar_coluna_cnpj(df):
    for col in df.columns:
        if 'cnpj' in col.lower():
            return col
    raise ValueError("❌ Nenhuma coluna com 'cnpj' foi encontrada no arquivo.")

# 🔗 Consulta API
def consultar_api(cnpj, warehouse, cliente, token):
    url = f"https://services.b2list.com/{cliente}/buyers/query"
    headers = {
        "X-TOKEN": token,
        "Content-Type": "application/json"
    }

    page = 0
    size = 100
    dados = []

    while True:
        payload = {
            "conditions": [
                {"key": "cnpj", "operator": "EQ", "value": cnpj},
                {"key": "warehouse", "operator": "IN", "value": warehouse}
            ],
            "page": page,
            "size": size,
            "sort": False,
        }

        try:
            response = requests.post(url, headers=headers, json=payload)

            if response.status_code == 401:
                st.error(f"❌ Token inválido ou expirado para o cliente '{cliente}'.")
                break

            response.raise_for_status()
            resultado = response.json()

            registros = resultado.get("results") or resultado.get("content") or []

            if not registros:
                break

            for item in registros:
                item["cnpj_consultado"] = cnpj
                item["warehouse_consultada"] = warehouse
                dados.append(item)

            total_pages = resultado.get("totalPages", 1)
            if page >= total_pages - 1:
                break

            page += 1

        except Exception as e:
            st.error(f"❌ Erro na consulta do CNPJ {cnpj} na warehouse {warehouse}: {e}")
            break

    return dados

# 🏗️ Processa a lista de CNPJs e retorna dataframe
def processar(cnpjs, cliente, token):
    warehouses = WAREHOUSES_POR_CLIENTE.get(cliente, [])
    
    if not warehouses:
        st.error(f"❌ Cliente '{cliente}' não possui warehouses cadastradas.")
        st.stop()

    resultados = []

    total_consultas = len(cnpjs) * len(warehouses)
    progresso = st.progress(0)
    status_text = st.empty()
    cont = 0

    for cnpj in cnpjs:
        for warehouse in warehouses:
            status_text.text(f"🔍 Consultando {cont + 1}/{total_consultas} — CNPJ: {cnpj}, Warehouse: {warehouse}")
            dados = consultar_api(cnpj, warehouse, cliente, token)
            resultados.extend(dados)
            cont += 1
            progresso.progress(cont / total_consultas)

    df_resultado = pd.DataFrame(resultados)
    return df_resultado


# 🎨 Interface Streamlit
st.set_page_config(page_title="Validador CNPJ x Warehouse", layout="wide")
st.title("🔍 Validação de CNPJ nas Warehouses")

uploaded_file = st.file_uploader("📤 Faça upload do arquivo Excel contendo os CNPJs", type=["xlsx", "xls"])
cliente = st.selectbox("🏢 Selecione o cliente", list(WAREHOUSES_POR_CLIENTE.keys()))

if uploaded_file and cliente:
    token = TOKENS_POR_CLIENTE.get(cliente)

    try:
        # 🚀 Leitura garantindo que o CNPJ seja string
        df_input = pd.read_excel(uploaded_file, dtype=str)
        
        st.subheader("📄 Prévia dos dados carregados:")
        st.dataframe(df_input)

        coluna_cnpj = detectar_coluna_cnpj(df_input)

        # 🔧 Normaliza o CNPJ (remove pontos, traços, barras)
        df_input[coluna_cnpj] = df_input[coluna_cnpj].str.replace(r'\D', '', regex=True)

        lista_cnpjs = df_input[coluna_cnpj].dropna().unique().tolist()

        if st.button("🚀 Iniciar Consulta"):
            with st.spinner("🔄 Processando... Aguarde!"):
                df_saida = processar(lista_cnpjs, cliente, token)

                if df_saida.empty:
                    st.warning("⚠️ Nenhum dado encontrado nas consultas.")
                else:
                    st.success("✅ Processamento concluído com sucesso!")
                    st.subheader("📜 Resultado da Consulta:")
                    st.dataframe(df_saida)

                    df_download = df_saida[['warehouse_consultada', 'cnpj_consultado']].drop_duplicates()

                    # 🔥 Geração do Excel corretamente formatado
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_download.to_excel(writer, index=False, sheet_name='Consulta')
                        workbook = writer.book
                        worksheet = writer.sheets['Consulta']

                        # Formatação texto para CNPJ
                        formato_texto = workbook.add_format({'num_format': '@'})

                        # Define larguras e aplica formatação
                        worksheet.set_column('A:A', 20)  # warehouse_consultada
                        worksheet.set_column('B:B', 30, formato_texto)  # cnpj_consultado como texto

                    output.seek(0)

                    st.download_button(
                        label="⬇️ Baixar Excel com Warehouse e CNPJ",
                        data=output,
                        file_name=f"warehouse_cnpj_{cliente}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"❌ Erro: {e}")

else:
    st.info("🚩 Por favor, envie um arquivo e selecione o cliente.")
