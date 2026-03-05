import streamlit as st
import zipfile
import io
import os
import re
import pandas as pd
import gc
import shutil

# --- CONFIGURAÇÃO E ESTILO (CLONE ABSOLUTO DO DIAMOND TAX) ---
st.set_page_config(page_title="GARIMPEIRO", layout="wide", page_icon="⛏️")

def aplicar_estilo_premium():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;800&family=Plus+Jakarta+Sans:wght@400;700&display=swap');

        header, [data-testid="stHeader"] { display: none !important; }
        .stApp { 
            background: radial-gradient(circle at top right, #FFDEEF 0%, #F8F9FA 100%) !important; 
        }

        [data-testid="stSidebar"] {
            background-color: #FFFFFF !important;
            border-right: 1px solid #FFDEEF !important;
            min-width: 400px !important;
            max-width: 400px !important;
        }

        div.stButton > button {
            color: #6C757D !important; 
            background-color: #FFFFFF !important; 
            border: 1px solid #DEE2E6 !important;
            border-radius: 15px !important;
            font-family: 'Montserrat', sans-serif !important;
            font-weight: 800 !important;
            height: 60px !important;
            text-transform: uppercase;
            transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275) !important;
            width: 100% !important;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05) !important;
        }

        div.stButton > button:hover {
            transform: translateY(-5px) !important;
            box-shadow: 0 10px 20px rgba(255,105,180,0.2) !important;
            border-color: #FF69B4 !important;
            color: #FF69B4 !important;
        }

        [data-testid="stFileUploader"] { 
            border: 2px dashed #FF69B4 !important; 
            border-radius: 20px !important;
            background: #FFFFFF !important;
            padding: 20px !important;
        }

        div.stDownloadButton > button {
            background-color: #FF69B4 !important; 
            color: white !important; 
            border: 2px solid #FFFFFF !important;
            font-weight: 700 !important;
            border-radius: 15px !important;
            box-shadow: 0 0 15px rgba(255, 105, 180, 0.3) !important;
            text-transform: uppercase;
            width: 100% !important;
        }

        h1, h2, h3 {
            font-family: 'Montserrat', sans-serif;
            font-weight: 800;
            color: #FF69B4 !important;
            text-align: center;
        }

        .instrucoes-card {
            background-color: rgba(255, 255, 255, 0.7);
            border-radius: 15px;
            padding: 20px;
            border-left: 5px solid #FF69B4;
            margin-bottom: 20px;
            min-height: 280px;
        }

        [data-testid="stMetric"] {
            background: white !important;
            border-radius: 20px !important;
            border: 1px solid #FFDEEF !important;
            padding: 15px !important;
        }
        </style>
    """, unsafe_allow_html=True)

aplicar_estilo_premium()

# --- VARIÁVEIS DE SISTEMA DE ARQUIVOS ---
TEMP_EXTRACT_DIR = "temp_garimpo_zips"
TEMP_UPLOADS_DIR = "temp_garimpo_uploads"

# --- MOTOR DE IDENTIFICAÇÃO MELHORADO ---
def identify_xml_info(content_bytes, client_cnpj, file_name):
    client_cnpj_clean = "".join(filter(str.isdigit, str(client_cnpj))) if client_cnpj else ""
    nome_puro = os.path.basename(file_name)
    if nome_puro.startswith('.') or nome_puro.startswith('~') or not nome_puro.lower().endswith('.xml'):
        return None, False
    
    resumo = {
        "Arquivo": nome_puro, 
        "Chave": "", 
        "Tipo": "Outros", 
        "Série": "0",
        "Número": 0, 
        "Status": "NORMAIS", 
        "Pasta": "",
        "Valor": 0.0, 
        "Conteúdo": b"", 
        "Ano": "0000", 
        "Mes": "00",
        "Operacao": "SAIDA", 
        "Data_Emissao": "",
        "CNPJ_Emit": "", 
        "Nome_Emit": "", 
        "Doc_Dest": "", 
        "Nome_Dest": ""
    }
    
    try:
        content_str = content_bytes[:45000].decode('utf-8', errors='ignore')
        tag_l = content_str.lower()
        if '<?xml' not in tag_l and '<inf' not in tag_l and '<inut' not in tag_l and '<retinut' not in tag_l: 
            return None, False
        
        tp_nf_match = re.search(r'<tpnf>([01])</tpnf>', tag_l)
        if tp_nf_match:
            if tp_nf_match.group(1) == "0":
                resumo["Operacao"] = "ENTRADA"
            else:
                resumo["Operacao"] = "SAIDA"

        resumo["CNPJ_Emit"] = re.search(r'<emit>.*?<cnpj>(\d+)</cnpj>', tag_l, re.S).group(1) if re.search(r'<emit>.*?<cnpj>(\d+)</cnpj>', tag_l, re.S) else ""
        resumo["Nome_Emit"] = re.search(r'<emit>.*?<xnome>(.*?)</xnome>', tag_l, re.S).group(1).upper() if re.search(r'<emit>.*?<xnome>(.*?)</xnome>', tag_l, re.S) else ""
        resumo["Doc_Dest"] = re.search(r'<dest>.*?<(?:cnpj|cpf)>(.*?)</(?:cnpj|cpf)>', tag_l, re.S).group(1) if re.search(r'<dest>.*?<(?:cnpj|cpf)>(.*?)</(?:cnpj|cpf)>', tag_l, re.S) else ""
        resumo["Nome_Dest"] = re.search(r'<dest>.*?<xnome>(.*?)</xnome>', tag_l, re.S).group(1).upper() if re.search(r'<dest>.*?<xnome>(.*?)</xnome>', tag_l, re.S) else ""

        data_match = re.search(r'<(?:dhemi|demi|dhregevento|dhrecbto)>(\d{4})-(\d{2})-(\d{2})', tag_l)
        if data_match: 
            resumo["Data_Emissao"] = f"{data_match.group(1)}-{data_match.group(2)}-{data_match.group(3)}"
            resumo["Ano"] = data_match.group(1)
            resumo["Mes"] = data_match.group(2)

        if '<inutnfe' in tag_l or '<retinutnfe' in tag_l or '<procinut' in tag_l:
            resumo["Status"] = "INUTILIZADOS"
            resumo["Tipo"] = "NF-e"
            
            if '<mod>65</mod>' in tag_l: resumo["Tipo"] = "NFC-e"
            elif '<mod>57</mod>' in tag_l: resumo["Tipo"] = "CT-e"
            
            resumo["Série"] = re.search(r'<serie>(\d+)</', tag_l).group(1) if re.search(r'<serie>(\d+)</', tag_l) else "0"
            ini = re.search(r'<nnfini>(\d+)</', tag_l).group(1) if re.search(r'<nnfini>(\d+)</', tag_l) else "0"
            fin = re.search(r'<nnffin>(\d+)</', tag_l).group(1) if re.search(r'<nnffin>(\d+)</', tag_l) else ini
            
            resumo["Número"] = int(ini)
            resumo["Range"] = (int(ini), int(fin))
            
            if resumo["Ano"] == "0000":
                ano_match = re.search(r'<ano>(\d+)</', tag_l)
                if ano_match: resumo["Ano"] = "20" + ano_match.group(1)[-2:]
                    
            resumo["Chave"] = f"INUT_{resumo['Série']}_{ini}"

        else:
            match_ch = re.search(r'<(?:chnfe|chcte|chmdfe)>(\d{44})</', tag_l)
            if not match_ch:
                match_ch = re.search(r'id=["\'](?:nfe|cte|mdfe)?(\d{44})["\']', tag_l)
                if match_ch: resumo["Chave"] = match_ch.group(1)
                else: resumo["Chave"] = ""
            else:
                resumo["Chave"] = match_ch.group(1)

            # Busca exata da Série na Tag (Evita erros)
            serie_match = re.search(r'<serie>0*(\d+)</serie>', tag_l)
            if serie_match:
                resumo["Série"] = serie_match.group(1)
            elif resumo["Chave"] and len(resumo["Chave"]) == 44:
                resumo["Série"] = str(int(resumo["Chave"][22:25]))

            if resumo["Chave"] and len(resumo["Chave"]) == 44:
                resumo["Ano"] = "20" + resumo["Chave"][2:4]
                resumo["Mes"] = resumo["Chave"][4:6]
                resumo["Número"] = int(resumo["Chave"][25:34])
                if not resumo["Data_Emissao"]: resumo["Data_Emissao"] = f"{resumo['Ano']}-{resumo['Mes']}-01"

            # Busca exata do Número da Nota na Tag
            nnf_match = re.search(r'<nnf>0*(\d+)</nnf>', tag_l)
            if nnf_match:
                resumo["Número"] = int(nnf_match.group(1))

            tipo = "NF-e"
            if '<mod>65</mod>' in tag_l: tipo = "NFC-e"
            elif '<mod>57</mod>' in tag_l or '<infcte' in tag_l: tipo = "CT-e"
            elif '<mod>58</mod>' in tag_l or '<infmdfe' in tag_l: tipo = "MDF-e"
            
            status = "NORMAIS"
            if '110111' in tag_l or '<cstat>101</cstat>' in tag_l: status = "CANCELADOS"
            elif '110110' in tag_l: status = "CARTA_CORRECAO"
                
            resumo["Tipo"] = tipo
            resumo["Status"] = status

            if status == "NORMAIS":
                v_match = re.search(r'<(?:vnf|vtprest|vreceb)>([\d.]+)</', tag_l)
                if v_match: resumo["Valor"] = float(v_match.group(1))
                else: resumo["Valor"] = 0.0
            
        if not resumo["CNPJ_Emit"] and resumo["Chave"] and not resumo["Chave"].startswith("INUT_"): 
            resumo["CNPJ_Emit"] = resumo["Chave"][6:20]
        
        if resumo["Mes"] == "00": resumo["Mes"] = "01"
        if resumo["Ano"] == "0000": resumo["Ano"] = "2000"

        is_p = (resumo["CNPJ_Emit"] == client_cnpj_clean)
        
        if is_p:
            resumo["Pasta"] = f"EMITIDOS_CLIENTE/{resumo['Operacao']}/{resumo['Tipo']}/{resumo['Status']}/{resumo['Ano']}/{resumo['Mes']}/Serie_{resumo['Série']}"
        else:
            resumo["Pasta"] = f"RECEBIDOS_TERCEIROS/{resumo['Operacao']}/{resumo['Tipo']}/{resumo['Ano']}/{resumo['Mes']}"
            
        return resumo, is_p
        
    except Exception as e: 
        return None, False

# --- FUNÇÃO RECURSIVA OTIMIZADA PARA DISCO ---
def extrair_recursivo(conteudo_ou_file, nome_arquivo):
    if not os.path.exists(TEMP_EXTRACT_DIR): os.makedirs(TEMP_EXTRACT_DIR)
        
    if nome_arquivo.lower().endswith('.zip'):
        try:
            if hasattr(conteudo_ou_file, 'read'): file_obj = conteudo_ou_file
            else: file_obj = io.BytesIO(conteudo_ou_file)
                
            with zipfile.ZipFile(file_obj) as z:
                for sub_nome in z.namelist():
                    if sub_nome.startswith('__MACOSX') or os.path.basename(sub_nome).startswith('.'): continue
                        
                    if sub_nome.lower().endswith('.zip'):
                        temp_path = z.extract(sub_nome, path=TEMP_EXTRACT_DIR)
                        with open(temp_path, 'rb') as f_temp: yield from extrair_recursivo(f_temp, sub_nome)
                        try: os.remove(temp_path)
                        except: pass
                    elif sub_nome.lower().endswith('.xml'):
                        yield (os.path.basename(sub_nome), z.read(sub_nome))
        except: pass
            
    elif nome_arquivo.lower().endswith('.xml'):
        if hasattr(conteudo_ou_file, 'read'): yield (os.path.basename(nome_arquivo), conteudo_ou_file.read())
        else: yield (os.path.basename(nome_arquivo), conteudo_ou_file)

# --- LIMPEZA DE PASTAS TEMPORÁRIAS ---
def limpar_arquivos_temp():
    try:
        for f in os.listdir('.'):
            if f.endswith('.zip') and ('z_org_final' in f or 'z_todos_final' in f):
                try: os.remove(f)
                except: pass
            
        if os.path.exists(TEMP_EXTRACT_DIR): shutil.rmtree(TEMP_EXTRACT_DIR, ignore_errors=True)
        if os.path.exists(TEMP_UPLOADS_DIR): shutil.rmtree(TEMP_UPLOADS_DIR, ignore_errors=True)
    except: pass

# --- FUNÇÃO DE RECALCULO (Deixa o código muito mais leve e evita repetições) ---
def recalcular_tabelas(lista_relatorio):
    lote_recalc = {}
    for item in lista_relatorio:
        key = item["Chave"]
        is_p = "EMITIDOS_CLIENTE" in item["Pasta"]
        if key in lote_recalc:
            if item["Status"] in ["CANCELADOS", "INUTILIZADOS"]: lote_recalc[key] = (item, is_p)
        else: lote_recalc[key] = (item, is_p)

    audit_map, canc_list, inut_list, aut_list, geral_list = {}, [], [], [], []
    
    for k, (res, is_p) in lote_recalc.items():
        origem_label = f"EMISSÃO PRÓPRIA ({res['Operacao']})" if is_p else f"TERCEIROS ({res['Operacao']})"
        registro_detalhado = {
            "Origem": origem_label, "Operação": res["Operacao"], "Modelo": res["Tipo"], "Série": res["Série"], 
            "Nota": res["Número"], "Data Emissão": res["Data_Emissao"], "CNPJ Emitente": res["CNPJ_Emit"], 
            "Nome Emitente": res["Nome_Emit"], "Doc Destinatário": res["Doc_Dest"], "Nome Destinatário": res["Nome_Dest"], 
            "Chave": res["Chave"], "Status Final": res["Status"], "Valor": res["Valor"], "Ano": res["Ano"], "Mes": res["Mes"]
        }
        
        if res["Status"] == "INUTILIZADOS":
            r = res.get("Range", (res["Número"], res["Número"]))
            for n in range(r[0], r[1] + 1):
                item_inut = registro_detalhado.copy()
                item_inut.update({"Nota": n, "Status Final": "INUTILIZADA", "Valor": 0.0})
                geral_list.append(item_inut)
        else: geral_list.append(registro_detalhado)

        if is_p:
            sk = (res["Tipo"], res["Série"])
            if sk not in audit_map: audit_map[sk] = {"nums": set(), "valor": 0.0}
                
            if res["Status"] == "INUTILIZADOS":
                r = res.get("Range", (res["Número"], res["Número"]))
                for n in range(r[0], r[1] + 1): 
                    audit_map[sk]["nums"].add(n)
                    inut_list.append({"Modelo": res["Tipo"], "Série": res["Série"], "Nota": n})
            else:
                if res["Número"] > 0:
                    audit_map[sk]["nums"].add(res["Número"])
                    if res["Status"] == "CANCELADOS": canc_list.append(registro_detalhado)
                    elif res["Status"] == "NORMAIS": aut_list.append(registro_detalhado)
                    audit_map[sk]["valor"] += res["Valor"]

    res_final, fal_final = [], []
    for (t, s), dados in audit_map.items():
        ns = sorted(list(dados["nums"]))
        if ns:
            n_min, n_max = ns[0], ns[-1]
            res_final.append({
                "Documento": t, "Série": s, "Início": n_min, "Fim": n_max, 
                "Quantidade": len(ns), "Valor Contábil (R$)": round(dados["valor"], 2)
            })
            for b in sorted(list(set(range(n_min, n_max + 1)) - set(ns))): fal_final.append({"Tipo": t, "Série": s, "Nº Faltante": b})
                
    return pd.DataFrame(res_final), pd.DataFrame(fal_final), pd.DataFrame(canc_list), pd.DataFrame(inut_list), pd.DataFrame(aut_list), pd.DataFrame(geral_list), canc_list, inut_list, aut_list

# --- INTERFACE ---
st.markdown("<h1>⛏️ O GARIMPEIRO</h1>", unsafe_allow_html=True)

with st.container():
    m_col1, m_col2 = st.columns(2)
    with m_col1:
        st.markdown("""
        <div class="instrucoes-card">
            <h3>📖 Como usar o sistema</h3>
            <ol>
                <li><b>Identificar Empresa:</b> Coloque o CNPJ à esquerda.</li>
                <li><b>Enviar as Notas:</b> Arraste as pastas/zips.</li>
                <li><b>Analisar:</b> Inicie o Garimpo.</li>
                <li><b>Validar:</b> Autentique com Excel se necessário.</li>
                <li><b>Exportar:</b> Filtre e baixe ficheiros consolidados.</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
    with m_col2:
        st.markdown("""
        <div class="instrucoes-card">
            <h3>📊 O que o sistema faz</h3>
            <ul>
                <li>Acha Notas Perdidas (Buracos)</li>
                <li>Limpa Cancelamentos e Inutilizadas</li>
                <li>Filtros avançados por Série, Modelo, e Origem</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

keys_to_init = [
    'garimpo_ok', 'confirmado', 'relatorio', 'df_resumo', 'df_faltantes', 'df_canceladas', 
    'df_inutilizadas', 'df_autorizadas', 'df_geral', 'df_divergencias', 'st_counts', 
    'validation_done', 'export_ready', 'org_zip_file', 'todos_zip_file', 'excel_buffer'
]

for k in keys_to_init:
    if k not in st.session_state:
        if 'df' in k: st.session_state[k] = pd.DataFrame()
        elif k == 'relatorio': st.session_state[k] = []
        elif k == 'st_counts': st.session_state[k] = {"CANCELADOS": 0, "INUTILIZADOS": 0, "AUTORIZADAS": 0}
        elif 'zip_file' in k: st.session_state[k] = ""
        elif 'excel_buffer' in k: st.session_state[k] = None
        else: st.session_state[k] = False

with st.sidebar:
    st.markdown("### 🔍 Configuração")
    cnpj_input = st.text_input("CNPJ DO CLIENTE", placeholder="00.000.000/0001-00")
    cnpj_limpo = "".join(filter(str.isdigit, cnpj_input))
    
    if cnpj_input and len(cnpj_limpo) != 14: st.error("⚠️ CNPJ Inválido.")
    if len(cnpj_limpo) == 14:
        if st.button("✅ LIBERAR OPERAÇÃO"): st.session_state['confirmado'] = True
            
    st.divider()
    if st.button("🗑️ RESETAR SISTEMA"):
        limpar_arquivos_temp()
        st.session_state.clear()
        st.rerun()

if st.session_state['confirmado']:
    if not st.session_state['garimpo_ok']:
        uploaded_files = st.file_uploader("📂 ARQUIVOS XML/ZIP (Suporta grandes volumes):", accept_multiple_files=True)
        if uploaded_files and st.button("🚀 INICIAR GRANDE GARIMPO"):
            limpar_arquivos_temp() 
            os.makedirs(TEMP_UPLOADS_DIR, exist_ok=True)
            
            lote_dict = {}
            progresso_bar = st.progress(0)
            status_text = st.empty()
            total_arquivos = len(uploaded_files)
            
            with st.status("⛏️ Minerando e salvando fisicamente...", expanded=True) as status_box:
                for i, f in enumerate(uploaded_files):
                    with open(os.path.join(TEMP_UPLOADS_DIR, f.name), "wb") as out_f: out_f.write(f.read())
                
                lista_salvos = os.listdir(TEMP_UPLOADS_DIR)
                total_salvos = len(lista_salvos)
                
                for i, f_name in enumerate(lista_salvos):
                    if i % 50 == 0: gc.collect()
                    progresso_bar.progress((i + 1) / total_salvos)
                    status_text.text(f"⛏️ Lendo conteúdo: {f_name}")
                    
                    try:
                        with open(os.path.join(TEMP_UPLOADS_DIR, f_name), "rb") as file_obj:
                            for name, xml_data in extrair_recursivo(file_obj, f_name):
                                res, is_p = identify_xml_info(xml_data, cnpj_limpo, name)
                                if res:
                                    key = res["Chave"]
                                    if key in lote_dict:
                                        if res["Status"] in ["CANCELADOS", "INUTILIZADOS"]: lote_dict[key] = (res, is_p)
                                    else: lote_dict[key] = (res, is_p)
                                del xml_data 
                    except Exception as e: continue
                
                status_box.update(label="✅ Leitura Concluída!", state="complete", expanded=False)
                progresso_bar.empty()
                status_text.empty()

            rel_list = [res for key, (res, is_p) in lote_dict.items()]
            df_res, df_fal, df_canc, df_inut, df_aut, df_ger, l_canc, l_inut, l_aut = recalcular_tabelas(rel_list)

            st.session_state.update({
                'relatorio': rel_list, 'df_resumo': df_res, 'df_faltantes': df_fal, 'df_canceladas': df_canc, 
                'df_inutilizadas': df_inut, 'df_autorizadas': df_aut, 'df_geral': df_ger,
                'st_counts': {"CANCELADOS": len(l_canc), "INUTILIZADOS": len(l_inut), "AUTORIZADAS": len(l_aut)}, 
                'garimpo_ok': True, 'export_ready': False
            })
            st.rerun()
    else:
        # --- TELA INICIAL PÓS-PROCESSAMENTO ---
        sc = st.session_state['st_counts']
        c1, c2, c3 = st.columns(3)
        c1.metric("📦 AUTORIZADAS (PRÓPRIAS)", sc.get("AUTORIZADAS", 0))
        c2.metric("❌ CANCELADAS (PRÓPRIAS)", sc.get("CANCELADOS", 0))
        c3.metric("🚫 INUTILIZADAS (PRÓPRIAS)", sc.get("INUTILIZADOS", 0))
        
        st.markdown("### 📊 RESUMO POR SÉRIE")
        st.dataframe(st.session_state['df_resumo'], use_container_width=True, hide_index=True)
        
        st.markdown("---")
        col_audit, col_canc, col_inut = st.columns(3)
        with col_audit:
            qtd_buracos = len(st.session_state['df_faltantes']) if not st.session_state['df_faltantes'].empty else 0
            st.markdown(f"### ⚠️ BURACOS ({qtd_buracos})")
            if not st.session_state['df_faltantes'].empty: st.dataframe(st.session_state['df_faltantes'], use_container_width=True, hide_index=True)
            else: st.info("✅ Tudo em ordem.")
        with col_canc:
            st.markdown("### ❌ CANCELADAS")
            if not st.session_state['df_canceladas'].empty: st.dataframe(st.session_state['df_canceladas'], use_container_width=True, hide_index=True)
            else: st.info("ℹ️ Nenhuma nota.")
        with col_inut:
            st.markdown("### 🚫 INUTILIZADAS")
            if not st.session_state['df_inutilizadas'].empty: st.dataframe(st.session_state['df_inutilizadas'], use_container_width=True, hide_index=True)
            else: st.info("ℹ️ Nenhuma nota.")

        st.divider()

        # --- MÓDULO: INUTILIZADAS MANUAIS ---
        if not st.session_state['df_faltantes'].empty:
            st.markdown("### 🛠️ INFORMAR NOTAS INUTILIZADAS (SEM XML)")
            with st.expander("Consulte a Sefaz e selecione abaixo as notas inutilizadas."):
                opcoes_buracos = [f"{row['Tipo']} | Série {row['Série']} | Nota {row['Nº Faltante']}" for idx, row in st.session_state['df_faltantes'].iterrows()]
                buracos_selecionados = st.multiselect("Selecione as notas:", opcoes_buracos)
                
                if st.button("CONFIRMAR INUTILIZAÇÃO"):
                    if buracos_selecionados:
                        with st.spinner("Atualizando..."):
                            for selecao in buracos_selecionados:
                                partes = selecao.split(" | ")
                                tipo_man, serie_man, nota_man = partes[0].strip(), partes[1].replace("Série", "").strip(), int(partes[2].replace("Nota", "").strip())
                                res_manual = {
                                    "Arquivo": "REGISTRO_MANUAL", "Chave": f"MANUAL_INUT_{tipo_man}_{serie_man}_{nota_man}", "Tipo": tipo_man, "Série": serie_man, "Número": nota_man, 
                                    "Status": "INUTILIZADOS", "Pasta": f"EMITIDOS_CLIENTE/SAIDA/{tipo_man}/INUTILIZADOS/0000/01/Serie_{serie_man}", "Valor": 0.0, "Conteúdo": b"", 
                                    "Ano": "0000", "Mes": "01", "Operacao": "SAIDA", "Data_Emissao": "", "CNPJ_Emit": cnpj_limpo, "Nome_Emit": "INSERÇÃO MANUAL", "Doc_Dest": "", "Nome_Dest": ""
                                }
                                st.session_state['relatorio'].append(res_manual)
                            
                            df_res, df_fal, df_canc, df_inut, df_aut, df_ger, l_canc, l_inut, l_aut = recalcular_tabelas(st.session_state['relatorio'])
                            st.session_state.update({'df_resumo': df_res, 'df_faltantes': df_fal, 'df_canceladas': df_canc, 'df_inutilizadas': df_inut, 'df_autorizadas': df_aut, 'df_geral': df_ger, 'st_counts': {"CANCELADOS": len(l_canc), "INUTILIZADOS": len(l_inut), "AUTORIZADAS": len(l_aut)}})
                            st.rerun()

        # --- MÓDULO: DESFAZER INUTILIZAÇÃO MANUAL ---
        inut_manuais = [item for item in st.session_state['relatorio'] if item.get('Arquivo') == "REGISTRO_MANUAL"]
        if inut_manuais:
            with st.expander("🔙 DESFAZER INUTILIZAÇÃO MANUAL"):
                opcoes_desfazer = [f"{item['Tipo']} | Série {item['Série']} | Nota {item['Número']}" for item in inut_manuais]
                desfazer_selecionados = st.multiselect("Remover da lista de inutilizadas:", opcoes_desfazer)
                
                if st.button("DESFAZER E ATUALIZAR TABELAS"):
                    if desfazer_selecionados:
                        with st.spinner("Removendo..."):
                            chaves_removidas = [f"MANUAL_INUT_{s.split(' | ')[0].strip()}_{s.split(' | ')[1].replace('Série', '').strip()}_{int(s.split(' | ')[2].replace('Nota', '').strip())}" for s in desfazer_selecionados]
                            st.session_state['relatorio'] = [i for i in st.session_state['relatorio'] if i['Chave'] not in chaves_removidas]
                            
                            df_res, df_fal, df_canc, df_inut, df_aut, df_ger, l_canc, l_inut, l_aut = recalcular_tabelas(st.session_state['relatorio'])
                            st.session_state.update({'df_resumo': df_res, 'df_faltantes': df_fal, 'df_canceladas': df_canc, 'df_inutilizadas': df_inut, 'df_autorizadas': df_aut, 'df_geral': df_ger, 'st_counts': {"CANCELADOS": len(l_canc), "INUTILIZADOS": len(l_inut), "AUTORIZADAS": len(l_aut)}})
                            st.rerun()

        st.divider()
        
        # --- ETAPA 2: VALIDAR COM EXCEL ---
        st.markdown("### 🕵️ ETAPA 2: VALIDAR COM RELATÓRIO DE AUTENTICIDADE")
        if st.session_state.get('validation_done'):
            if len(st.session_state['df_divergencias']) > 0: st.warning("⚠️ Status atualizados. Verifique a aba Divergências no Excel.")
            else: st.success("✅ O status dos XMLs está 100% alinhado com a SEFAZ.")

        with st.expander("Subir Excel e atualizar status real"):
            auth_file = st.file_uploader("Suba o Excel (.xlsx) [Col A=Chave, Col F=Status]", type=["xlsx", "xls"], key="auth_up")
            if auth_file and st.button("🔄 VALIDAR E ATUALIZAR"):
                df_auth = pd.read_excel(auth_file)
                auth_dict = {str(row.iloc[0]).strip(): str(row.iloc[5]).strip().upper() for idx, row in df_auth.iterrows() if len(str(row.iloc[0]).strip()) == 44}
                
                div_list = []
                for item in st.session_state['relatorio']:
                    if item["Chave"] in auth_dict and "CANCEL" in auth_dict[item["Chave"]]:
                        if item["Status"] == "NORMAIS":
                            div_list.append({"Chave": item["Chave"], "Nota": item["Número"], "Status XML": "AUTORIZADA", "Status Real": "CANCELADA"})
                        item["Status"] = "CANCELADOS"
                
                df_res, df_fal, df_canc, df_inut, df_aut, df_ger, l_canc, l_inut, l_aut = recalcular_tabelas(st.session_state['relatorio'])
                st.session_state.update({'df_divergencias': pd.DataFrame(div_list), 'df_resumo': df_res, 'df_faltantes': df_fal, 'df_canceladas': df_canc, 'df_inutilizadas': df_inut, 'df_autorizadas': df_aut, 'df_geral': df_ger, 'st_counts': {"CANCELADOS": len(l_canc), "INUTILIZADOS": len(l_inut), "AUTORIZADAS": len(l_aut)}, 'validation_done': True})
                st.rerun()

        st.divider()

        # --- MÓDULO: ADICIONAR MAIS ARQUIVOS ---
        with st.expander("➕ ADICIONAR MAIS ARQUIVOS (SEM RESETAR)"):
            extra_files = st.file_uploader("Adicionar arquivos ao lote atual:", accept_multiple_files=True, key="extra_files")
            if extra_files and st.button("PROCESSAR E ATUALIZAR LISTA"):
                with st.spinner("Adicionando..."):
                    os.makedirs(TEMP_UPLOADS_DIR, exist_ok=True)
                    for f in extra_files:
                        with open(os.path.join(TEMP_UPLOADS_DIR, f.name), "wb") as out_f: out_f.write(f.read())
                        f.seek(0)
                        try:
                            for name, xml_data in extrair_recursivo(f, f.name):
                                res, is_p = identify_xml_info(xml_data, cnpj_limpo, name)
                                if res and not any(item['Chave'] == res['Chave'] for item in st.session_state['relatorio']):
                                    st.session_state['relatorio'].append(res)
                        except: pass
                    
                    st.session_state['export_ready'] = False
                    df_res, df_fal, df_canc, df_inut, df_aut, df_ger, l_canc, l_inut, l_aut = recalcular_tabelas(st.session_state['relatorio'])
                    st.session_state.update({'df_resumo': df_res, 'df_faltantes': df_fal, 'df_canceladas': df_canc, 'df_inutilizadas': df_inut, 'df_autorizadas': df_aut, 'df_geral': df_ger, 'st_counts': {"CANCELADOS": len(l_canc), "INUTILIZADOS": len(l_inut), "AUTORIZADAS": len(l_aut)}})
                    st.rerun()

        st.divider()

        # --- ETAPA 3: FILTROS E EXPORTAÇÃO ---
        st.markdown("### ⚙️ ETAPA 3: FILTROS AVANÇADOS E EXPORTAÇÃO")
        st.info("Deixe as caixas vazias para exportar TUDO, ou filtre os dados.")
        
        todas_origens = ["EMISSÃO PRÓPRIA", "TERCEIROS"]
        anos_meses = sorted(list(set([f"{r.get('Ano', '0000')}/{r.get('Mes', '00')}" for r in st.session_state['relatorio'] if r.get('Ano', '0000') != '0000'])))
        modelos = sorted(list(set([r.get('Tipo', '') for r in st.session_state['relatorio']])))
        series = sorted(list(set([str(r.get('Série', '0')) for r in st.session_state['relatorio']])))
        
        with st.container():
            f_col1, f_col2, f_col3, f_col4 = st.columns(4)
            with f_col1: filtro_origem = st.multiselect("📌 Origem:", todas_origens)
            with f_col2: 
                filtro_meses = st.multiselect("📅 Ano/Mês:", anos_meses)
                aplicar_mes_so_na_propria = st.checkbox("Mês APENAS na Própria?", value=True)
            with f_col3: filtro_modelos = st.multiselect("📄 Modelo:", modelos)
            with f_col4: filtro_series = st.multiselect("🔢 Série:", series)

        if st.button("🚀 PROCESSAR E GERAR ARQUIVOS FINAIS"):
            with st.spinner("Buscando no HD e montando pacotes..."):
                
                # CORREÇÃO APLICADA AQUI: Limpa apenas os zips finais antigos para não apagar os uploads!
                for f in os.listdir('.'):
                    if f.startswith('z_org_final') or f.startswith('z_todos_final'):
                        try: os.remove(f)
                        except: pass
                
                os.makedirs(TEMP_EXTRACT_DIR, exist_ok=True)
                
                # --- EXCEL ---
                df_geral_filtrado = st.session_state['df_geral'].copy()
                if not df_geral_filtrado.empty:
                    if filtro_origem:
                        condicoes = []
                        if "EMISSÃO PRÓPRIA" in filtro_origem: condicoes.append(df_geral_filtrado['Origem'].str.contains('PRÓPRIA'))
                        if "TERCEIROS" in filtro_origem: condicoes.append(df_geral_filtrado['Origem'].str.contains('TERCEIROS'))
                        df_geral_filtrado = df_geral_filtrado[pd.concat(condicoes, axis=1).any(axis=1)]
                    if filtro_meses and not df_geral_filtrado.empty:
                        df_geral_filtrado['Mes_Comp'] = df_geral_filtrado['Ano'] + "/" + df_geral_filtrado['Mes']
                        if aplicar_mes_so_na_propria: df_geral_filtrado = df_geral_filtrado[(df_geral_filtrado['Mes_Comp'].isin(filtro_meses)) | (df_geral_filtrado['Origem'].str.contains('TERCEIROS')) | (df_geral_filtrado['Status Final'] == 'INUTILIZADA')]
                        else: df_geral_filtrado = df_geral_filtrado[(df_geral_filtrado['Mes_Comp'].isin(filtro_meses)) | (df_geral_filtrado['Status Final'] == 'INUTILIZADA')]
                    if filtro_modelos and not df_geral_filtrado.empty: df_geral_filtrado = df_geral_filtrado[df_geral_filtrado['Modelo'].isin(filtro_modelos)]
                    if filtro_series and not df_geral_filtrado.empty: df_geral_filtrado = df_geral_filtrado[df_geral_filtrado['Série'].astype(str).isin(filtro_series)]

                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                    st.session_state['df_resumo'].to_excel(writer, sheet_name='Resumo_Auditoria', index=False)
                    df_geral_filtrado.to_excel(writer, sheet_name='Geral_Filtrado', index=False)
                    st.session_state['df_faltantes'].to_excel(writer, sheet_name='Buracos', index=False)
                    st.session_state['df_canceladas'].to_excel(writer, sheet_name='Canceladas', index=False)
                    st.session_state['df_inutilizadas'].to_excel(writer, sheet_name='Inutilizadas', index=False)
                    st.session_state['df_autorizadas'].to_excel(writer, sheet_name='Autorizadas', index=False)
                    if not st.session_state['df_divergencias'].empty: st.session_state['df_divergencias'].to_excel(writer, sheet_name='Divergencias', index=False)
                
                st.session_state['excel_buffer'] = buffer_excel.getvalue()

                # --- ZIPS (Físico sem quebra de blocos) ---
                org_name = 'z_org_final_completo.zip'
                todos_name = 'z_todos_final_completo.zip'
                
                z_org = zipfile.ZipFile(org_name, "w", zipfile.ZIP_DEFLATED)
                z_todos = zipfile.ZipFile(todos_name, "w", zipfile.ZIP_DEFLATED)
                
                if os.path.exists(TEMP_UPLOADS_DIR):
                    for f_name in os.listdir(TEMP_UPLOADS_DIR):
                        with open(os.path.join(TEMP_UPLOADS_DIR, f_name), "rb") as f_temp:
                            for name, xml_data in extrair_recursivo(f_temp, f_name):
                                res, is_p = identify_xml_info(xml_data, cnpj_limpo, name)
                                if res:
                                    manter = True
                                    if filtro_origem:
                                        if ("EMISSÃO PRÓPRIA" if is_p else "TERCEIROS") not in filtro_origem: manter = False
                                    if manter and filtro_meses:
                                        passar_mes = False
                                        if f"{res['Ano']}/{res['Mes']}" in filtro_meses: passar_mes = True
                                        elif res["Status"] == "INUTILIZADOS": passar_mes = True
                                        elif aplicar_mes_so_na_propria and not is_p: passar_mes = True
                                        if not passar_mes: manter = False
                                    if manter and filtro_modelos and res["Tipo"] not in filtro_modelos: manter = False
                                    if manter and filtro_series and str(res["Série"]) not in filtro_series: manter = False
                                    if manter:
                                        z_org.writestr(f"{res['Pasta']}/{name}", xml_data)
                                        z_todos.writestr(name, xml_data)
                                del xml_data
                
                if z_org: z_org.close()
                if z_todos: z_todos.close()

                st.session_state['org_zip_file'] = org_name
                st.session_state['todos_zip_file'] = todos_name
                st.session_state['export_ready'] = True
                st.rerun()

        # --- BOTÕES DE DOWNLOAD OTIMIZADOS ---
        if st.session_state.get('export_ready'):
            st.success("✅ Tudo empacotado, filtrado e pronto para baixar!")
            
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown("### 📂 ORGANIZADO (PASTAS)")
                if os.path.exists(st.session_state['org_zip_file']):
                    with open(st.session_state['org_zip_file'], 'rb') as f:
                        st.download_button("📥 BAIXAR ORGANIZADO", data=f, file_name="garimpo_filtrado_org.zip", mime="application/zip", use_container_width=True)

            with c2:
                st.markdown("### 📦 TODOS (SÓ XML)")
                if os.path.exists(st.session_state['todos_zip_file']):
                    with open(st.session_state['todos_zip_file'], 'rb') as f:
                        st.download_button("📥 BAIXAR TODOS", data=f, file_name="todos_xml_soltos.zip", mime="application/zip", use_container_width=True)

            with c3:
                st.markdown("### 📊 RELATÓRIO EXCEL")
                if st.session_state.get('excel_buffer'):
                    st.download_button("📊 BAIXAR EXCEL MASTER", st.session_state['excel_buffer'], "auditoria_detalhada_filtrada.xlsx", use_container_width=True, mime="application/vnd.ms-excel")

        st.divider()
        if st.button("⛏️ NOVO GARIMPO / LIMPAR TUDO"):
            limpar_arquivos_temp()
            st.session_state.clear()
            st.rerun()
else:
    st.warning("👈 Insira o CNPJ na barra lateral para começar.")
