import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
from fpdf import FPDF
import gspread
from google.oauth2.service_account import Credentials

# Configuração da Página
st.set_page_config(page_title="Sistema de Confecção Pro", layout="wide", page_icon="🧵")

# Nome da Planilha no Google Sheets
GOOGLE_SHEET_NAME = "pedidos_confeccao"

# Função para conectar ao Google Sheets (corrigida sem duplicação de cache)
@st.cache_resource
def conectar_google_sheets():
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    if os.path.exists("credentials.json"):
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    else:
        # Puxa os dados dos secrets e corrige o formato da chave privada para aceitar as quebras de linha
        gcp_secrets = dict(st.secrets["gcp_service_account"])
        if "private_key" in gcp_secrets:
            gcp_secrets["private_key"] = gcp_secrets["private_key"].replace("\\n", "\n")
            
        creds = Credentials.from_service_account_info(gcp_secrets, scopes=SCOPES)
    
    client = gspread.authorize(creds)
    sheet = client.open(GOOGLE_SHEET_NAME).sheet1 
    return sheet

# Estrutura base das colunas
COLUNAS_EXCEL = [
    "ID", "Categoria", "Nome", "Telefone", "Endereco", "Responsavel", 
    "Itens_Detalhados", "Tecidos_Usados", "Qtde_Total",
    "Valor_Pecas", "Valor_Bordados", "Total", "Adiantamento", "Restante", 
    "Forma_Pagamento", "Info_Boletos", "Status", "Data_Pedido", "Prev_Entrega"
]

def ler_dados_sheets():
    sheet = conectar_google_sheets()
    dados = sheet.get_all_records()
    if not dados:
        df = pd.DataFrame(columns=COLUNAS_EXCEL)
    else:
        df = pd.DataFrame(dados)
        for col in COLUNAS_EXCEL:
            if col not in df.columns:
                df[col] = ""
    return df

def salvar_dados_sheets(df):
    sheet = conectar_google_sheets()
    sheet.clear()
    dados_para_salvar = [df.columns.tolist()] + df.fillna("").values.tolist()
    sheet.update(dados_para_salvar)

# Validação inicial do Google Sheets
try:
    df_temp = ler_dados_sheets()
    if df_temp.empty:
        salvar_dados_sheets(pd.DataFrame(columns=COLUNAS_EXCEL))
except Exception as e:
    st.error(f"Erro ao conectar com o Google Sheets. Verifique suas credenciais e o nome da planilha. Detalhe: {e}")
    st.stop()

# Lista padrão de tecidos
if "lista_tecidos" not in st.session_state:
    st.session_state.lista_tecidos = ["Brim", "Oxford", "Jeans", "Sarja", "Malha", "Moletom", "Helanca", "Outro"]

# Função para sanitizar textos para o FPDF
def limpa_texto(texto):
    if texto is None:
        return ""
    if not isinstance(texto, str):
        texto = str(texto)
    reemplazos = {
        '•': '-', '–': '-', '—': '-',
        '“': '"', '”': '"', '’': "'", '‘': "'",
        '…': '...', '–': '-'
    }
    for orig, dest in reemplazos.items():
        texto = texto.replace(orig, dest)
    return texto.encode('latin-1', 'ignore').decode('latin-1')

# Função para calcular dias úteis
def somar_dias_uteis(data_inicial, dias_uteis):
    data_atual = data_inicial
    dias_adicionados = 0
    while dias_adicionados < dias_uteis:
        data_atual += timedelta(days=1)
        if data_atual.weekday() < 5:
            dias_adicionados += 1
    return data_atual

# Classe FPDF customizada
class PDF_FichaProducao(FPDF):
    def footer(self):
        self.set_y(-12)
        self.set_font("Arial", 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, limpa_texto(f"Página {self.page_no()}"), align='C')

# Função para gerar PDF
def gerar_pdf_pedidos(pedidos_df):
    pdf = PDF_FichaProducao()
    pdf.set_margins(10, 10, 10)
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()

    ALTURA_APROX_CARD = 95

    for idx, (_, ped) in enumerate(pedidos_df.iterrows()):
        pdf.set_text_color(0, 0, 0)

        if idx > 0 and (pdf.get_y() + ALTURA_APROX_CARD > 280):
            pdf.add_page()

        pdf.set_fill_color(30, 41, 59)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(190, 7, limpa_texto("SISTEMA DE CONFECÇÃO PRO - FICHA DE PRODUÇÃO"), border=0, fill=True, align='C', ln=True)

        pdf.set_text_color(0, 0, 0)

        pdf.set_font("Arial", 'B', 10)
        pdf.set_fill_color(255, 255, 255)
        pdf.cell(120, 6, limpa_texto(f" PEDIDO #{ped['ID']} - Categoria: {ped['Categoria']}"), border='TLB', fill=True)
        pdf.cell(70, 6, limpa_texto(f"STATUS: {ped['Status']} "), border='TRB', fill=True, align='R', ln=True)

        pdf.set_font("Arial", '', 8)
        resp = str(ped['Responsavel']) if pd.notna(ped['Responsavel']) and str(ped['Responsavel']).strip() != "" else "N/A"
        end_txt = str(ped['Endereco']).strip() if pd.notna(ped['Endereco']) else ""
        forma_pag = str(ped['Forma_Pagamento']) if pd.notna(ped['Forma_Pagamento']) else "N/A"
        info_bol = f" | Venc. Boletos: {ped['Info_Boletos']}" if pd.notna(ped['Info_Boletos']) and str(ped['Info_Boletos']).strip() != "" else ""
        
        label_entrega = "Data de Entrega:" if ped['Status'] == "Entregue" else "Previsão de Entrega:"
        
        pdf.cell(95, 4.5, limpa_texto(f"Cliente: {ped['Nome']}"), border='TLR')
        pdf.cell(95, 4.5, limpa_texto(f"Data do Pedido: {ped['Data_Pedido']}"), border='TLR', ln=True)
        
        pdf.cell(95, 4.5, limpa_texto(f"Telefone: {ped['Telefone']}"), border='LR')
        pdf.cell(95, 4.5, limpa_texto(f"{label_entrega} {ped['Prev_Entrega']}"), border='LR', ln=True)
        
        if end_txt:
            pdf.cell(190, 4.5, limpa_texto(f"Endereço: {end_txt}"), border='LRT', ln=True)
            
        pdf.cell(190, 4.5, limpa_texto(f"Responsável: {resp} | Pagamento: {forma_pag}{info_bol}"), border='LR', ln=True)
        pdf.cell(190, 4.5, limpa_texto(f"Tecidos Utilizados: {ped['Tecidos_Usados']}"), border='LRB', ln=True)

        pdf.set_font("Arial", 'B', 8)
        pdf.set_fill_color(225, 228, 232)
        pdf.cell(190, 5, limpa_texto(" DETALHAMENTO DOS ITENS PARA CORTE / PRODUÇÃO"), ln=True, fill=True, border=1)
        
        pdf.set_font("Arial", '', 8)
        itens_split = str(ped['Itens_Detalhados']).split(" | ")
        for item_txt in itens_split:
            if item_txt.strip():
                pdf.multi_cell(190, 4.5, limpa_texto(f"- {item_txt.strip()}"), border='LRB')

        pdf.set_font("Arial", 'B', 8)
        pdf.set_fill_color(245, 245, 245)
        pdf.cell(47.5, 5, limpa_texto(f"Total Peças: R$ {float(ped['Valor_Pecas']):.2f}"), border=1, fill=True)
        pdf.cell(47.5, 5, limpa_texto(f"Total Bordados: R$ {float(ped['Valor_Bordados']):.2f}"), border=1, fill=True)
        pdf.cell(47.5, 5, limpa_texto(f"TOTAL: R$ {float(ped['Total']):.2f}"), border=1, fill=True)
        pdf.cell(47.5, 5, limpa_texto(f"Restante: R$ {float(ped['Restante']):.2f}"), border=1, fill=True, ln=True)

        pdf.ln(8)

    try:
        out = pdf.output(dest='S')
    except TypeError:
        out = pdf.output()

    if isinstance(out, str):
        return out.encode('latin1')
    return bytes(out)

# Estilização CSS
st.markdown("""
    <style>
    .stButton>button { 
        width: 100%; 
        border-radius: 8px; 
        font-weight: bold; 
        transition: all 0.3s ease;
    }
    div[data-testid="stMetricValue"] {
        font-size: 26px;
        color: #0284c7;
    }
    </style>
    """, unsafe_allow_html=True)

# --- NAVEGAÇÃO SIDEBAR ---
st.sidebar.title("🧵 Confecção Pro")
menu = st.sidebar.radio("Navegação", ["📝 Novo Pedido", "✏️ Editar Pedido", "🔄 Fluxo de Produção", "📋 Todos os Pedidos"])

# --- PÁGINA 1: NOVO PEDIDO ---
if menu == "📝 Novo Pedido":
    st.title("📝 Lançar Novo Pedido")
    
    if "categoria" not in st.session_state:
        st.session_state.categoria = "INDIVIDUAL"
    if "itens_pedido" not in st.session_state:
        st.session_state.itens_pedido = [{
            "tecido": "Brim", "descricao": "", "qtde": 1, "cores": "",
            "th": "", "sexo": "Masc", "servico": "", "bordado": "Não",
            "desc_bordado": "", "valor_peca": 0.0, "valor_bordado": 0.0
        }]

    st.write("### Selecione a Categoria:")
    col_cat1, col_cat2, col_cat3, col_cat4 = st.columns(4)
    if col_cat1.button("👤 INDIVIDUAL", use_container_width=True): 
        st.session_state.categoria = "INDIVIDUAL"
    if col_cat2.button("🏢 EMPRESARIAL", use_container_width=True): 
        st.session_state.categoria = "EMPRESARIAL"
    if col_cat3.button("🪡 BORDADO", use_container_width=True): 
        st.session_state.categoria = "BORDADO"
    if col_cat4.button("🛠️ CONSERTO", use_container_width=True): 
        st.session_state.categoria = "CONSERTO"
    
    st.info(f"📌 Categoria Selecionada: **{st.session_state.categoria}**")

    with st.container(border=True):
        st.subheader("👤 Dados do Cliente")
        c1, c2, c3 = st.columns([3, 2, 2])
        nome = c1.text_input("Nome / Razão Social:")
        telefone = c2.text_input("Telefone / WhatsApp:")
        responsavel = c3.text_input("Responsável / Atendente:")

        if st.session_state.categoria == "EMPRESARIAL":
            st.markdown("🏢 **Endereço Empresarial**")
            endereco = st.text_input("Endereço Empresarial Completo (Rua, Nº, Bairro, Cidade - UF, CEP):")
        else:
            endereco = ""

        st.subheader("📅 Datas & Prazos")
        c4, c5, c6 = st.columns([2, 2, 3])
        data_atual_dt = datetime.now()
        data_ped = c4.date_input("Data do Pedido:", data_atual_dt, format="DD/MM/YYYY")
        opcao_prazo = c5.selectbox("Prazo de Entrega:", ["15 Dias Úteis", "30 Dias Úteis", "Personalizado"])
        
        if opcao_prazo == "15 Dias Úteis":
            prev_ent_calc = somar_dias_uteis(data_ped, 15)
            prev_ent = c6.date_input("Prev. Entrega:", prev_ent_calc, format="DD/MM/YYYY")
        elif opcao_prazo == "30 Dias Úteis":
            prev_ent_calc = somar_dias_uteis(data_ped, 30)
            prev_ent = c6.date_input("Prev. Entrega:", prev_ent_calc, format="DD/MM/YYYY")
        else:
            prev_ent = c6.date_input("Prev. Entrega:", data_ped, format="DD/MM/YYYY")

    st.subheader("👕 Itens do Pedido")
    
    total_pecas_val = 0.0
    total_bordados_val = 0.0
    tecidos_do_pedido = set()
    item_para_remover = None

    for idx, item in enumerate(st.session_state.itens_pedido):
        with st.container(border=True):
            st.markdown(f"#### Item #{idx+1}")
            ci1, ci2, ci3, ci4 = st.columns([2, 3, 1, 1])
            
            opcoes_tecidos = st.session_state.lista_tecidos + ["➕ Adicionar novo..."]
            tec_sel = ci1.selectbox(f"Tecido (Item {idx+1}):", opcoes_tecidos, key=f"tec_{idx}")
            
            if tec_sel == "➕ Adicionar novo...":
                tec_custom = ci1.text_input(f"Especifique o Tecido (Item {idx+1}):", key=f"custom_tec_{idx}")
                item["tecido"] = tec_custom if tec_custom else "Outro"
            else:
                item["tecido"] = tec_sel

            if item["tecido"]:
                tecidos_do_pedido.add(item["tecido"])

            item["descricao"] = ci2.text_input(f"Descrição (Item {idx+1}):", value=item["descricao"], key=f"desc_{idx}")
            item["qtde"] = ci3.number_input(f"Qtde (Item {idx+1}):", min_value=1, value=item["qtde"], step=1, key=f"qtde_{idx}")
            item["th"] = ci4.text_input(f"Tam/TH (Item {idx+1}):", value=item["th"], key=f"th_{idx}")

            ci5, ci6, ci7, ci8 = st.columns([2, 1, 2, 2])
            item["cores"] = ci5.text_input(f"Cores (Item {idx+1}):", value=item["cores"], key=f"cor_{idx}")
            item["sexo"] = ci6.selectbox(f"Sexo (Item {idx+1}):", ["Masc", "Fem", "Uni"], key=f"sexo_{idx}")
            item["servico"] = ci7.text_input(f"Serviço (Item {idx+1}):", value=item["servico"], key=f"serv_{idx}")
            item["bordado"] = ci8.selectbox(f"Bordado (Item {idx+1}):", ["Não", "Sim"], key=f"bord_{idx}")

            if item["bordado"] == "Sim":
                cb1, cb2 = st.columns([3, 1])
                item["desc_bordado"] = cb1.text_input(f"Detalhes do Bordado (Item {idx+1}):", value=item["desc_bordado"], key=f"dbord_{idx}")
                item["valor_bordado"] = cb2.number_input(f"Val. Bordado Un. (R$):", min_value=0.0, value=item["valor_bordado"], format="%.2f", key=f"vbord_{idx}")
            else:
                item["desc_bordado"] = ""
                item["valor_bordado"] = 0.0

            cv1, _ = st.columns([2, 2])
            item["valor_peca"] = cv1.number_input(f"Valor Unitário da Peça (R$):", min_value=0.0, value=item["valor_peca"], format="%.2f", key=f"vpeca_{idx}")

            subtotal_item = (item["valor_peca"] + item["valor_bordado"]) * item["qtde"]
            st.caption(f"Subtotal deste item: **R$ {subtotal_item:.2f}**")

            total_pecas_val += item["valor_peca"] * item["qtde"]
            total_bordados_val += item["valor_bordado"] * item["qtde"]

            if len(st.session_state.itens_pedido) > 1:
                if st.button(f"🗑️ Remover Item #{idx+1}", key=f"del_{idx}"):
                    item_para_remover = idx

    if item_para_remover is not None:
        st.session_state.itens_pedido.pop(item_para_remover)
        st.rerun()

    if st.button("➕ Adicionar Outro Item ao Pedido"):
        st.session_state.itens_pedido.append({
            "tecido": "Brim", "descricao": "", "qtde": 1, "cores": "",
            "th": "", "sexo": "Masc", "servico": "", "bordado": "Não",
            "desc_bordado": "", "valor_peca": 0.0, "valor_bordado": 0.0
        })
        st.rerun()

    st.divider()
    st.subheader("💰 Resumo Financeiro & Forma de Pagamento")
    
    v_total_geral = total_pecas_val + total_bordados_val
    
    cv1, cv2, cv3, cv4 = st.columns(4)
    cv1.metric("Total Peças", f"R$ {total_pecas_val:.2f}")
    cv2.metric("Total Bordados", f"R$ {total_bordados_val:.2f}")
    cv3.metric("VALOR TOTAL", f"R$ {v_total_geral:.2f}")
    
    v_adiant = cv4.number_input("Adiantamento (R$):", format="%.2f", min_value=0.0)
    v_restante = v_total_geral - v_adiant
    st.write(f"### Restante a Pagar: **R$ {v_restante:.2f}**")

    st.markdown("#### 💳 Forma de Pagamento")
    cp1, _ = st.columns([2, 2])
    forma_pagamento = cp1.selectbox("Selecione a Forma de Pagamento:", ["PIX", "Dinheiro", "Cartão de Crédito", "Cartão de Débito", "Boleto", "Faturado / A Prazo", "Outro"])
    
    info_boletos_str = ""
    if forma_pagamento == "Boleto":
        with st.container(border=True):
            st.markdown("📄 **Vencimento dos Boletos / Parcelamento**")
            num_boletos = st.number_input("Quantidade de Boletos / Parcelas:", min_value=1, max_value=12, value=1, step=1)
            
            datas_boletos = []
            cols_bol = st.columns(min(int(num_boletos), 4))
            for b in range(int(num_boletos)):
                c_idx = b % 4
                data_padrao = data_ped + timedelta(days=30*(b+1))
                dt_b = cols_bol[c_idx].date_input(f"Vencimento Boleto #{b+1}:", data_padrao, format="DD/MM/YYYY", key=f"bol_dt_{b}")
                datas_boletos.append(f"Parc #{b+1}: {dt_b.strftime('%d/%m/%Y')}")
            
            info_boletos_str = " | ".join(datas_boletos)

    st.write("")
    if st.button("💾 SALVAR PEDIDO", use_container_width=True, type="primary"):
        if not nome:
            st.error("❌ Por favor, preencha pelo menos o Nome do cliente.")
        else:
            df_existente = ler_dados_sheets()
            novo_id_num = len(df_existente) + 1
            novo_id_str = f"{novo_id_num}"

            itens_resumo_str = ""
            qtde_total_pedido = 0
            for idx, it in enumerate(st.session_state.itens_pedido):
                itens_resumo_str += f"[{idx+1}] {it['qtde']}x {it['descricao']} (Tecido: {it['tecido']}, Tam: {it['th']}, Cor: {it['cores']}, Bordado: {it['bordado']}) | "
                qtde_total_pedido += it['qtde']

            tecidos_str = ", ".join(list(tecidos_do_pedido))

            novo_pedido = {
                "ID": novo_id_str,
                "Categoria": st.session_state.categoria,
                "Nome": nome,
                "Telefone": telefone,
                "Endereco": endereco,
                "Responsavel": responsavel,
                "Itens_Detalhados": itens_resumo_str,
                "Tecidos_Usados": tecidos_str,
                "Qtde_Total": qtde_total_pedido,
                "Valor_Pecas": total_pecas_val,
                "Valor_Bordados": total_bordados_val,
                "Total": v_total_geral,
                "Adiantamento": v_adiant,
                "Restante": v_restante,
                "Forma_Pagamento": forma_pagamento,
                "Info_Boletos": info_boletos_str,
                "Status": "Pedido Recebido",
                "Data_Pedido": data_ped.strftime("%d/%m/%Y"),
                "Prev_Entrega": prev_ent.strftime("%d/%m/%Y")
            }

            df_novo = pd.concat([df_existente, pd.DataFrame([novo_pedido])], ignore_index=True)
            salvar_dados_sheets(df_novo)
            
            st.success(f"✅ Pedido **#{novo_id_str}** salvo com sucesso no Google Sheets!")
            st.session_state.itens_pedido = [{
                "tecido": "Brim", "descricao": "", "qtde": 1, "cores": "",
                "th": "", "sexo": "Masc", "servico": "", "bordado": "Não",
                "desc_bordado": "", "valor_peca": 0.0, "valor_bordado": 0.0
            }]

# --- PÁGINA 2: EDITAR PEDIDO ---
elif menu == "✏️ Editar Pedido":
    st.title("✏️ Editar / Alterar Informações do Pedido")
    df = ler_dados_sheets()
    
    if df.empty:
        st.info("Nenhum pedido cadastrado para editar.")
    else:
        ped_id_sel = st.selectbox("Selecione o ID do Pedido que deseja alterar:", df['ID'].astype(str))
        ped_idx = df[df['ID'].astype(str) == ped_id_sel].index[0]
        ped = df.loc[ped_idx]

        with st.form("form_editar_pedido"):
            st.subheader(f"Alterando Pedido #{ped['ID']}")
            
            ce1, ce2, ce3 = st.columns(3)
            cat_edit = ce1.selectbox("Categoria:", ["INDIVIDUAL", "EMPRESARIAL", "BORDADO", "CONSERTO"], index=["INDIVIDUAL", "EMPRESARIAL", "BORDADO", "CONSERTO"].index(ped['Categoria']) if ped['Categoria'] in ["INDIVIDUAL", "EMPRESARIAL", "BORDADO", "CONSERTO"] else 0)
            status_edit = ce2.selectbox("Status:", ["Pedido Recebido", "Corte", "Costura", "Bordado", "Pronto", "Entregue"], index=["Pedido Recebido", "Corte", "Costura", "Bordado", "Pronto", "Entregue"].index(ped['Status']) if ped['Status'] in ["Pedido Recebido", "Corte", "Costura", "Bordado", "Pronto", "Entregue"] else 0)
            resp_edit = ce3.text_input("Responsável:", value=str(ped['Responsavel']) if pd.notna(ped['Responsavel']) else "")

            st.markdown("#### Cliente & Endereço")
            c_ed1, c_ed2 = st.columns(2)
            nome_edit = c_ed1.text_input("Nome / Razão Social:", value=str(ped['Nome']) if pd.notna(ped['Nome']) else "")
            tel_edit = c_ed2.text_input("Telefone:", value=str(ped['Telefone']) if pd.notna(ped['Telefone']) else "")
            end_edit = st.text_input("Endereço Empresarial:", value=str(ped['Endereco']) if pd.notna(ped['Endereco']) else "")

            st.markdown("#### Detalhamento dos Itens")
            itens_edit = st.text_area("Descrição dos Itens:", value=str(ped['Itens_Detalhados']) if pd.notna(ped['Itens_Detalhados']) else "", help="Separados por |")
            tecidos_edit = st.text_input("Tecidos Utilizados:", value=str(ped['Tecidos_Usados']) if pd.notna(ped['Tecidos_Usados']) else "")

            st.markdown("#### Financeiro & Forma de Pagamento")
            f1, f2, f3, f4 = st.columns(4)
            v_pecas_edit = f1.number_input("Valor Peças (R$):", value=float(ped['Valor_Pecas']) if pd.notna(ped['Valor_Pecas']) and ped['Valor_Pecas'] != '' else 0.0, format="%.2f")
            v_bord_edit = f2.number_input("Valor Bordados (R$):", value=float(ped['Valor_Bordados']) if pd.notna(ped['Valor_Bordados']) and ped['Valor_Bordados'] != '' else 0.0, format="%.2f")
            v_total_edit = f3.number_input("Total Geral (R$):", value=float(ped['Total']) if pd.notna(ped['Total']) and ped['Total'] != '' else 0.0, format="%.2f")
            v_adiant_edit = f4.number_input("Adiantamento (R$):", value=float(ped['Adiantamento']) if pd.notna(ped['Adiantamento']) and ped['Adiantamento'] != '' else 0.0, format="%.2f")
            
            restante_edit = v_total_edit - v_adiant_edit
            st.caption(f"Novo Valor Restante a Pagar: **R$ {restante_edit:.2f}**")

            fp_edit = st.selectbox("Forma de Pagamento:", ["PIX", "Dinheiro", "Cartão de Crédito", "Cartão de Débito", "Boleto", "Faturado / A Prazo", "Outro"], index=["PIX", "Dinheiro", "Cartão de Crédito", "Cartão de Débito", "Boleto", "Faturado / A Prazo", "Outro"].index(ped['Forma_Pagamento']) if ped['Forma_Pagamento'] in ["PIX", "Dinheiro", "Cartão de Crédito", "Cartão de Débito", "Boleto", "Faturado / A Prazo", "Outro"] else 0)
            boletos_edit = st.text_input("Info Boletos / Parcelas (Vencimentos):", value=str(ped['Info_Boletos']) if pd.notna(ped['Info_Boletos']) else "")

            btn_salvar_edicao = st.form_submit_button("💾 SALVAR ALTERAÇÕES", type="primary", use_container_width=True)

            if btn_salvar_edicao:
                df.loc[ped_idx, 'Categoria'] = cat_edit
                df.loc[ped_idx, 'Status'] = status_edit
                df.loc[ped_idx, 'Responsavel'] = resp_edit
                df.loc[ped_idx, 'Nome'] = nome_edit
                df.loc[ped_idx, 'Telefone'] = tel_edit
                df.loc[ped_idx, 'Endereco'] = end_edit
                df.loc[ped_idx, 'Itens_Detalhados'] = itens_edit
                df.loc[ped_idx, 'Tecidos_Usados'] = tecidos_edit
                df.loc[ped_idx, 'Valor_Pecas'] = v_pecas_edit
                df.loc[ped_idx, 'Valor_Bordados'] = v_bord_edit
                df.loc[ped_idx, 'Total'] = v_total_edit
                df.loc[ped_idx, 'Adiantamento'] = v_adiant_edit
                df.loc[ped_idx, 'Restante'] = restante_edit
                df.loc[ped_idx, 'Forma_Pagamento'] = fp_edit
                df.loc[ped_idx, 'Info_Boletos'] = boletos_edit
                
                salvar_dados_sheets(df)
                st.success(f"✅ Pedido #{ped['ID']} atualizado com sucesso no Google Sheets!")
                st.rerun()

# --- PÁGINA 3: FLUXO DE PRODUÇÃO & IMPRESSÃO ---
elif menu == "🔄 Fluxo de Produção":
    st.title("🔄 Fluxo de Produção & Fila de Trabalho")
    
    df = ler_dados_sheets()
    
    if df.empty:
        st.info("Nenhum pedido cadastrado no momento.")
    else:
        col1, col2, col3 = st.columns(3)
        
        filtro_categoria = col1.selectbox("Filtrar por Categoria:", ["Todas", "INDIVIDUAL", "EMPRESARIAL", "BORDADO", "CONSERTO"])
        
        todos_tecidos = set()
        for tecs in df['Tecidos_Usados'].dropna():
            for t in str(tecs).split(", "):
                if t.strip():
                    todos_tecidos.add(t.strip())
        
        filtro_tecido = col2.selectbox("Filtrar por Tecido:", ["Todos"] + list(todos_tecidos))
        status_opcoes = ["Todos", "Pedido Recebido", "Corte", "Costura", "Bordado", "Pronto", "Entregue"]
        filtro_status = col3.selectbox("Filtrar por Status do Fluxo:", status_opcoes)
        
        query = df.copy()
        
        if filtro_categoria != "Todas":
            query = query[query['Categoria'] == filtro_categoria]

        if filtro_status != "Todos":
            query = query[query['Status'] == filtro_status]
            
        if filtro_tecido != "Todos":
            query = query[query['Tecidos_Usados'].astype(str).str.contains(filtro_tecido, case=False, na=False)]
        
        st.write(f"Exibindo **{len(query)}** pedido(s)")
        st.dataframe(query, use_container_width=True)

        if not query.empty:
            st.divider()
            
            st.subheader("⚙️ Atualizar Status do Pedido")
            c_id, c_status = st.columns(2)
            id_marcar = c_id.selectbox("Selecione o ID do pedido:", query['ID'])
            novo_status = c_status.selectbox("Novo Status do Fluxo:", ["Pedido Recebido", "Corte", "Costura", "Bordado", "Pronto", "Entregue"])
            
            if st.button("Atualizar Status", type="primary"):
                df.loc[df['ID'].astype(str) == str(id_marcar), 'Status'] = novo_status
                
                if novo_status == "Entregue":
                    data_hoje_str = datetime.now().strftime("%d/%m/%Y")
                    df.loc[df['ID'].astype(str) == str(id_marcar), 'Prev_Entrega'] = data_hoje_str
                    st.success(f"Status do Pedido #{id_marcar} atualizado para 'Entregue'!")
                else:
                    st.success(f"Status do Pedido #{id_marcar} atualizado para '{novo_status}'!")
                    
                salvar_dados_sheets(df)
                st.rerun()

            st.divider()

            st.subheader("🖨️ Impressão de Fichas para Corte / Produção (PDF)")
            
            ids_selecionados = st.multiselect(
                "Selecione um ou mais pedidos para imprimir a ficha de produção:", 
                options=query['ID'].tolist(),
                default=query['ID'].tolist() if len(query) <= 5 else []
            )

            if ids_selecionados:
                pedidos_para_pdf = query[query['ID'].isin(ids_selecionados)]
                pdf_bytes = gerar_pdf_pedidos(pedidos_para_pdf)
                
                nome_arquivo_pdf = f"fichas_producao_{len(ids_selecionados)}_pedidos.pdf" if len(ids_selecionados) > 1 else f"ficha_producao_pedido_{ids_selecionados[0]}.pdf"
                
                st.download_button(
                    label=f"📄 Baixar PDF para Impressão ({len(ids_selecionados)} Ficha(s))",
                    data=pdf_bytes,
                    file_name=nome_arquivo_pdf,
                    mime="application/pdf",
                    use_container_width=True
                )
            else:
                st.info("Selecione ao menos um pedido acima para gerar o PDF.")

# --- PÁGINA 4: TODOS OS PEDIDOS ---
elif menu == "📋 Todos os Pedidos":
    st.title("📋 Banco de Dados Completo (Google Sheets)")
    df = ler_dados_sheets()
    
    if df.empty:
        st.info("O banco de dados está vazio.")
    else:
        st.dataframe(df, use_container_width=True)
        
        @st.cache_data
        def convert_df_to_excel(df_to_convert):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_to_convert.to_excel(writer, index=False, sheet_name='Pedidos')
            return output.getvalue()

        import io
        excel_data = convert_df_to_excel(df)
        st.download_button(
            label="📥 Baixar Planilha em Excel (.xlsx)",
            data=excel_data,
            file_name="relatorio_pedidos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )