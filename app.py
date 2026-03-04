import streamlit as st
import zipfile
import io
import os
import gc
import shutil

# --- CONFIGURAÇÃO E ESTILO ---
st.set_page_config(page_title="EXTRATOR TURBO XML", layout="wide", page_icon="⚡")

def aplicar_estilo_premium():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;800&family=Plus+Jakarta+Sans:wght@400;700&display=swap');

        header, [data-testid="stHeader"] { display: none !important; }
        .stApp { 
            background: radial-gradient(circle at top right, #E0F7FA 0%, #F8F9FA 100%) !important; 
        }

        [data-testid="stSidebar"] {
            background-color: #FFFFFF !important;
            border-right: 1px solid #E0F7FA !important;
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
            box-shadow: 0 10px 20px rgba(0,188,212,0.2) !important;
            border-color: #00BCD4 !important;
            color: #00BCD4 !important;
        }

        [data-testid="stFileUploader"] { 
            border: 2px dashed #00BCD4 !important; 
            border-radius: 20px !important;
            background: #FFFFFF !important;
            padding: 20px !important;
        }

        div.stDownloadButton > button {
            background-color: #00BCD4 !important; 
            color: white !important; 
            border: 2px solid #FFFFFF !important;
            font-weight: 700 !important;
            border-radius: 15px !important;
            box-shadow: 0 0 15px rgba(0, 188, 212, 0.3) !important;
            text-transform: uppercase;
            width: 100% !important;
        }

        h1, h2, h3 {
            font-family: 'Montserrat', sans-serif;
            font-weight: 800;
            color: #00BCD4 !important;
            text-align: center;
        }

        .instrucoes-card {
            background-color: rgba(255, 255, 255, 0.7);
            border-radius: 15px;
            padding: 20px;
            border-left: 5px solid #00BCD4;
            margin-bottom: 20px;
            min-height: 200px;
        }

        [data-testid="stMetric"] {
            background: white !important;
            border-radius: 20px !important;
            border: 1px solid #E0F7FA !important;
            padding: 15px !important;
        }
        </style>
    """, unsafe_allow_html=True)

aplicar_estilo_premium()

# --- VARIÁVEIS DE SISTEMA DE ARQUIVOS (PREVENÇÃO DE QUEDA DE MEMÓRIA) ---
TEMP_EXTRACT_DIR = "temp_extrator_zips"
TEMP_UPLOADS_DIR = "temp_extrator_uploads"
MAX_XML_PER_ZIP = 10000  # Limite elástico para evitar queda no download

# --- FUNÇÃO RECURSIVA OTIMIZADA PARA DISCO (BUSCA APENAS XMLS) ---
def extrair_recursivo(conteudo_ou_file, nome_arquivo):
    if not os.path.exists(TEMP_EXTRACT_DIR): 
        os.makedirs(TEMP_EXTRACT_DIR)
        
    if nome_arquivo.lower().endswith('.zip'):
        try:
            if hasattr(conteudo_ou_file, 'read'):
                file_obj = conteudo_ou_file
            else:
                file_obj = io.BytesIO(conteudo_ou_file)
                
            with zipfile.ZipFile(file_obj) as z:
                for sub_nome in z.namelist():
                    if sub_nome.startswith('__MACOSX') or os.path.basename(sub_nome).startswith('.'): 
                        continue
                        
                    if sub_nome.lower().endswith('.zip'):
                        # Extrai zip dentro de zip para o disco (Bypass na RAM)
                        temp_path = z.extract(sub_nome, path=TEMP_EXTRACT_DIR)
                        with open(temp_path, 'rb') as f_temp:
                            yield from extrair_recursivo(f_temp, sub_nome)
                        try: 
                            os.remove(temp_path)
                        except: 
                            pass
                    elif sub_nome.lower().endswith('.xml'):
                        yield (os.path.basename(sub_nome), z.read(sub_nome))
        except: 
            pass
            
    elif nome_arquivo.lower().endswith('.xml'):
        if hasattr(conteudo_ou_file, 'read'): 
            yield (os.path.basename(nome_arquivo), conteudo_ou_file.read())
        else: 
            yield (os.path.basename(nome_arquivo), conteudo_ou_file)

# --- LIMPEZA DE PASTAS TEMPORÁRIAS ---
def limpar_arquivos_temp():
    try:
        for f in os.listdir('.'):
            if f.endswith('.zip') and f.startswith('lote_xml_bruto'):
                try: os.remove(f)
                except: pass
            
        if os.path.exists(TEMP_EXTRACT_DIR): 
            shutil.rmtree(TEMP_EXTRACT_DIR, ignore_errors=True)
            
        if os.path.exists(TEMP_UPLOADS_DIR): 
            shutil.rmtree(TEMP_UPLOADS_DIR, ignore_errors=True)
    except: 
        pass

# --- DIVISOR DE LOTES HTML (Para deixar botões organizados) ---
def chunk_list(lst, n):
    for i in range(0, len(lst), n): 
        yield lst[i:i + n]

# --- INTERFACE PRINCIPAL ---
st.markdown("<h1>⚡ EXTRATOR TURBO XML</h1>", unsafe_allow_html=True)

with st.container():
    st.markdown("""
    <div class="instrucoes-card">
        <h3>🚀 Operação Força Bruta (Extração Direta)</h3>
        <p>Esse aplicativo foi desenhado com um único objetivo: <b>agrupar todos os arquivos XML soltos num único lugar, ignorando regras contábeis, pastas ou CNPJs.</b></p>
        <ul>
            <li>Arrastou um ZIP cheio de outras pastas e zips dentro? Ele vasculha tudo, acha os XMLs e planifica.</li>
            <li>Zero travamentos: O sistema lê de forma assíncrona usando o HD para não esgotar a memória do servidor.</li>
            <li>Se o volume for titânico, ele divide automaticamente os resultados em lotes de 10.000 notas (aprox. 25MB) para facilitar o download.</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# Controle de sessão
if 'processamento_concluido' not in st.session_state:
    st.session_state['processamento_concluido'] = False
if 'lotes_gerados' not in st.session_state:
    st.session_state['lotes_gerados'] = []
if 'total_xmls' not in st.session_state:
    st.session_state['total_xmls'] = 0

with st.sidebar:
    st.markdown("### ⚙️ Controles")
    if st.button("🗑️ RESETAR SISTEMA"):
        limpar_arquivos_temp()
        st.session_state.clear()
        st.rerun()

if not st.session_state['processamento_concluido']:
    uploaded_files = st.file_uploader("📂 ARRASTE AQUI SEUS ARQUIVOS (XML ou ZIP):", accept_multiple_files=True)
    
    if uploaded_files and st.button("⚡ EXTRAIR E PLANIFICAR TUDO"):
        limpar_arquivos_temp() 
        os.makedirs(TEMP_UPLOADS_DIR, exist_ok=True)
        
        progresso_bar = st.progress(0)
        status_text = st.empty()
        total_arquivos = len(uploaded_files)
        
        with st.status("⚡ Turbo Ativado! Escaneando e compactando...", expanded=True) as status_box:
            
            # 1. Salva uploads fisicamente no disco para evitar estouro de RAM
            for i, f in enumerate(uploaded_files):
                caminho_salvo = os.path.join(TEMP_UPLOADS_DIR, f.name)
                with open(caminho_salvo, "wb") as out_f:
                    out_f.write(f.read())
            
            lista_salvos = os.listdir(TEMP_UPLOADS_DIR)
            total_salvos = len(lista_salvos)
            
            nomes_unicos_vistos = set() # Evita sobreposição de arquivos com mesmo nome exato
            lotes_parts = []
            current_count = 0
            curr_part = 1
            
            zip_name = f'lote_xml_bruto_pt{curr_part}.zip'
            z_out = zipfile.ZipFile(zip_name, "w", zipfile.ZIP_DEFLATED)
            lotes_parts.append(zip_name)
            
            total_xmls_extraidos = 0
            
            # 2. Varredura cega por XMLs
            for i, f_name in enumerate(lista_salvos):
                if i % 20 == 0: 
                    gc.collect() # Faxineiro de RAM
                    
                progresso_bar.progress((i + 1) / total_salvos)
                status_text.text(f"⚡ Varrulhando arquivo {i+1}/{total_salvos}: {f_name}")
                
                caminho_leitura = os.path.join(TEMP_UPLOADS_DIR, f_name)
                try:
                    with open(caminho_leitura, "rb") as file_obj:
                        todos_xmls = extrair_recursivo(file_obj, f_name)
                        for name, xml_data in todos_xmls:
                            
                            # Proteção contra nomes duplicados na mesma pasta
                            nome_final = name
                            if nome_final in nomes_unicos_vistos:
                                name_parts = os.path.splitext(name)
                                random_sufix = str(random.randint(1000, 9999))
                                nome_final = f"{name_parts[0]}_{random_sufix}{name_parts[1]}"
                                
                            nomes_unicos_vistos.add(nome_final)
                            
                            # Lógica Anti-Crash (Fechamento de lotes se ficar muito pesado)
                            if current_count >= MAX_XML_PER_ZIP:
                                z_out.close()
                                curr_part += 1
                                zip_name = f'lote_xml_bruto_pt{curr_part}.zip'
                                z_out = zipfile.ZipFile(zip_name, "w", zipfile.ZIP_DEFLATED)
                                lotes_parts.append(zip_name)
                                current_count = 0
                                
                            # Escreve o arquivo "solto" no zip (sem pastas)
                            z_out.writestr(nome_final, xml_data)
                            current_count += 1
                            total_xmls_extraidos += 1
                            
                            del xml_data 
                except Exception as e: 
                    continue
            
            if z_out:
                z_out.close()
            
            status_box.update(label=f"✅ Sucesso! {total_xmls_extraidos} XMLs localizados e planificados.", state="complete", expanded=False)
            progresso_bar.empty()
            status_text.empty()

        st.session_state['total_xmls'] = total_xmls_extraidos
        st.session_state['lotes_gerados'] = lotes_parts
        st.session_state['processamento_concluido'] = True
        st.rerun()

else:
    # --- TELA DE RESULTADO (DOWNLOADS) ---
    st.success(f"🎉 Processamento concluído! Foram extraídos e planificados **{st.session_state['total_xmls']}** arquivos XML.")
    
    st.markdown("### 📥 DOWNLOADS DISPONÍVEIS")
    st.write("Os seus arquivos estão empacotados e divididos em lotes (se necessário) para um download rápido e seguro.")
    
    lista_lotes = st.session_state['lotes_gerados']
    
    for row in chunk_list(lista_lotes, 3):
        cols = st.columns(len(row))
        for idx, part_name in enumerate(row):
            if os.path.exists(part_name):
                part_num = re.search(r'pt(\d+)', part_name).group(1)
                tamanho_mb = os.path.getsize(part_name) / (1024 * 1024)
                
                label = f"📦 BAIXAR LOTE {part_num} ({tamanho_mb:.1f} MB)" if len(lista_lotes) > 1 else f"📦 BAIXAR TODOS OS XMLS ({tamanho_mb:.1f} MB)"
                with cols[idx]:
                    with open(part_name, 'rb') as f:
                        st.download_button(
                            label=label, 
                            data=f.read(), 
                            file_name=f"xml_extraidos_brutos_pt{part_num}.zip", 
                            mime="application/zip", 
                            use_container_width=True
                        )
                        
    st.divider()
    if st.button("⚡ FAZER NOVA EXTRAÇÃO"):
        limpar_arquivos_temp()
        st.session_state.clear()
        st.rerun()
