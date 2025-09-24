import streamlit as st
import pandas as pd
import webbrowser
import pyautogui
import time

st.set_page_config(
    page_title="Automação ServiceNow",
    page_icon="🤖",
    layout="wide"
)

st.title("-- Automação ServiceNow - Preenchimento de RITM --")
st.markdown("---")

# Configurações de segurança do PyAutoGUI
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3  # Reduzido para melhor performance

class ServiceNowAutomation:
    def __init__(self):
        self.dados = None
        self.linha_selecionada = None
        self.screen_width, self.screen_height = pyautogui.size()
    
    def focus_browser_window(self):
        """Foca na janela do navegador com métodos múltiplos"""
        try:
            st.info("Tentando focar na janela do ServiceNow...")
            for i in range(5, 0, -1):
                st.warning(f"⏳ Automação iniciará em {i} segundos - Certifique-se que a janela está ativa!")
                time.sleep(1)
           
        except Exception as e:
            st.error("Erro no focus automático!")
            st.warning("**OBRIGATÓRIO:** Clique MANUALMENTE na janela do ServiceNow!")
            for i in range(5, 0, -1):
                st.error(f"⏳ Clique na janela em {i} segundos...")
                time.sleep(1)
   
   # Função que carrega o excel
    def load_excel_data(self, uploaded_file, sheet_name: str):
        try:
            self.dados = pd.read_excel(uploaded_file, sheet_name=sheet_name, skipfooter=2, dtype=str)
            self.dados = self.dados.dropna(subset=['Colaborador']) # Apagando campos vazios
            self.dados = self.dados.fillna('')
           
            # Converter colunas para string
            colunas_importantes = ['Colaborador', 'Matricula', 'Alocação', 'Fornecedor', 'Perfil', 'VALOR', 'PREVISTO', 'Centro', 'Gestor']
            for coluna in colunas_importantes:
                if coluna in self.dados.columns:
                    self.dados[coluna] = self.dados[coluna].astype(str).str.strip()
            return True
        except Exception as e:
            st.error(f"❌ Erro ao carregar arquivo: {e}")
            return False

    def fill_snow_form(self):
        """Preenche o formulário do ServiceNow com coordenadas FIXAS"""
        try:
            linha = self.linha_selecionada
           
            time.sleep(2)
            pyautogui.click(x=671, y=756)
            time.sleep(4)
            pyautogui.scroll(-500)
            time.sleep(1)
            pyautogui.click(x=671, y=756)
            time.sleep(1)
            pyautogui.click(x=641, y=684)
            time.sleep(3)
           
            if linha['Alocação'] == 'Capex':
                pyautogui.click(x=146, y=863)
            elif linha['Alocação'] == 'Opex':
                pyautogui.click(x=286, y=861)
           
            time.sleep(2)
           
            pyautogui.scroll(-550)
            time.sleep(1.5)

            pyautogui.click(x=475, y=619)
            time.sleep(1.5)
            pyautogui.click(x=462, y=755)
            time.sleep(4)
           
            pyautogui.click(x=498, y=729)
            time.sleep(1)
           
            pyautogui.write(linha['Fornecedor'])
            pyautogui.press('enter')
           
            time.sleep(1)
           
            # Nome
            pyautogui.click(x=410, y=866)
            time.sleep(0.5)
            pyautogui.write(linha['Colaborador'])
            time.sleep(0.5)
            # Matricula
            pyautogui.click(x=867, y=841)
            time.sleep(0.5)
            pyautogui.write(str(int(float(linha['Matricula']))))
            time.sleep(0.5)
            # Cargo
            pyautogui.click(x=525, y=964)
            time.sleep(0.5)
            pyautogui.write(linha['Perfil'])
           
            time.sleep(2)
           
            # Descendo a tela
            pyautogui.scroll(-350)
            time.sleep(1.5)
           
           
            # Data inicio
            pyautogui.click(x=664, y=787)
            time.sleep(0.5)
            pyautogui.click(x=701, y=520)
            time.sleep(0.5)
            pyautogui.click(x=897, y=765)
            time.sleep(0.5)
            # Data fim
            pyautogui.click(x=1284,y=786)
            pyautogui.click(x=1595,y=436)
            pyautogui.click(x=1555,y=678)
            pyautogui.click(x=1525,y=775)
            time.sleep(2)

            # Valor mensal
            pyautogui.click(x=536, y=652)
            time.sleep(0.5)
            pyautogui.write(linha['VALOR'])
            time.sleep(0.5)
            # Valor total de meses
            pyautogui.click(x=1047, y=641)
            time.sleep(0.5)
            pyautogui.write(linha['PREVISTO'])
            pyautogui.press('enter')
            time.sleep(1)
           
            # Ordem estatística (apenas para OPEX)
            if linha['Alocação'] == "Opex":
                pyautogui.click(x=912, y=909)
                time.sleep(0.5)
                pyautogui.write(linha['Ordem estatisica'])
                time.sleep(0.5)
           
            # Descendo mais
            pyautogui.scroll(-250)
            time.sleep(1.5)
           
            # Centro de custo
            if linha['Alocação'] == "Opex":
                pyautogui.click(x=629, y=700)
                time.sleep(0.5)
                pyautogui.write(linha['PREVISTO'])
                time.sleep(0.5)
                pyautogui.click(x=824, y=695)
                time.sleep(0.5)
                pyautogui.write(linha['Centro'].strip())
            elif linha['Alocação'] == "Capex":
                pyautogui.click(x=431, y=714)
                time.sleep(0.5)
                pyautogui.write(linha['Centro'].strip())
           
            time.sleep(1)
           
            pyautogui.click(x=509, y=857)
            time.sleep(0.5)
            pyautogui.write(f"Abertura de chamado para {linha['Perfil']} no valor mensal de {linha['VALOR']} com o centro de custo {linha['Centro']}")
           
            time.sleep(1.5)
            pyautogui.scroll(-250)
            time.sleep(1.5)
           
            pyautogui.click(x=461, y=676)
            time.sleep(0.5)
            pyautogui.write(linha['Área solicitante'])
            time.sleep(1.5)
            pyautogui.press('enter')
            time.sleep(0.5)

            pyautogui.click(x=459, y=778)
            time.sleep(0.5)
            pyautogui.write(linha['Área destino'])
            time.sleep(1.5)
            pyautogui.press('enter')
            time.sleep(0.5)
           
            # Nome area
            pyautogui.click(x=502, y=949)
            time.sleep(0.5)
            pyautogui.write(linha['Diretoria'])
            time.sleep(1.5)
            pyautogui.click(x=500, y=907)
            time.sleep(0.5)
           
            pyautogui.scroll(-300)
            time.sleep(1.5)
           
            # Gerencia
            pyautogui.click(x=612, y=714)
            time.sleep(1)
            pyautogui.write(linha['Área'])
            pyautogui.press('enter')
            time.sleep(2)
           
            # Diretor
            pyautogui.click(x=628, y=831)
            time.sleep(1)
            pyautogui.write(linha['Diretor'].strip())
            time.sleep(1)
            pyautogui.press('enter')
           
            time.sleep(2)
           
            # Gestor responsavel
            pyautogui.click(x=264, y=950)
            time.sleep(1)
            pyautogui.write(linha['Gestor'].strip())
            pyautogui.press('enter')
            pyautogui.scroll(-300)
            time.sleep(1)
           
            # Abrir área de trabalho (final)
            pyautogui.click(x=1204, y=915)
            time.sleep(1)
            st.success("✅ Preenchimento concluído com sucesso!")
            return True
           
        except Exception as e:
            st.error(f"❌ Erro durante o preenchimento: {e}")
            return False

# Inicializar a automação
if 'snow_bot' not in st.session_state:
    st.session_state.snow_bot = ServiceNowAutomation()


snow_bot = st.session_state.snow_bot


# Sidebar para upload
with st.sidebar:
    st.header("📁 Configurações")
   
    st.info(f"🖥️ Resolução: {snow_bot.screen_width}x{snow_bot.screen_height}")
   
    uploaded_file = st.file_uploader("Carregar planilha Excel", type=['xlsx'])
   
    if uploaded_file is not None:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            selected_sheet = st.selectbox("Selecione a aba:", sheet_names)
           
            if st.button("Carregar Dados da Planilha"):
                with st.spinner("Carregando dados..."):
                    if snow_bot.load_excel_data(uploaded_file, selected_sheet):
                        st.success(f"Dados carregados! {len(snow_bot.dados)} registros encontrados.")
                        if 'selected_colaborador' in st.session_state:
                            del st.session_state.selected_colaborador
        except Exception as e:
            st.error(f"Erro ao ler arquivo: {e}")

# Área principal
if snow_bot.dados is not None:
    st.header("👥 Seleção do Colaborador")
   
    colaboradores = snow_bot.dados['Colaborador'].tolist()
   
    if 'selected_colaborador' not in st.session_state:
        st.session_state.selected_colaborador = colaboradores[0] if colaboradores else ""
   
    selected_colaborador = st.selectbox(
        "Selecione o colaborador:",
        options=colaboradores,
        index=colaboradores.index(st.session_state.selected_colaborador) if st.session_state.selected_colaborador in colaboradores else 0,
        key="colaborador_select"
    )
   
    st.session_state.selected_colaborador = selected_colaborador
   
    if selected_colaborador:
        linha_idx = snow_bot.dados[snow_bot.dados['Colaborador'] == selected_colaborador].index[0]
        snow_bot.linha_selecionada = snow_bot.dados.loc[linha_idx]
       
        st.subheader("Dados do Colaborador Selecionado")
       
        col1, col2, col3 ,col4 = st.columns(4)
       
        with col1:
            st.info(f"**Colaborador:** {snow_bot.linha_selecionada['Colaborador']}")
            st.info(f"**Matrícula:** {snow_bot.linha_selecionada['Matricula']}")
            st.info(f"**Perfil:** {snow_bot.linha_selecionada['Perfil']}")
       
        with col2:
            st.info(f"**Alocação:** {snow_bot.linha_selecionada['Alocação']}")
            st.info(f"**Fornecedor:** {snow_bot.linha_selecionada['Fornecedor']}")
            st.info(f"**Gestor:** {snow_bot.linha_selecionada['Gestor']}")
       
        with col3:
            st.info(f"**Valor Mensal:** R$ {snow_bot.linha_selecionada['VALOR']}")
            st.info(f"**Valor Previsto:** R$ {snow_bot.linha_selecionada['PREVISTO']}")
            st.info(f"**Centro de Custo:** {snow_bot.linha_selecionada['Centro']}")
    
        with col4:
            st.info(f"**Área destino:** {snow_bot.linha_selecionada['Área destino']}")
            st.info(f"**Área:** {snow_bot.linha_selecionada['Área']}")
            st.info(f"**Diretoria:** {snow_bot.linha_selecionada['Diretoria']}")
       
        st.markdown("---")
       
        st.header("🚀 Executar Automação")
       
        st.error("""
        **⚠️ INSTRUÇÕES CRÍTICAS:**
       
        1. **Certifique-se** que a resolução é **1920x1080**
        2. **AGUARDE 6 segundos** após abrir o link
        3. **NÃO MOVA O MOUSE** durante a execução
        4. **Certifique se os campos necessarios estão na planilha**
        """)
       
        with st.container():
            col1, col2, col3 = st.columns([1,2,1])
            with col2:
                if st.button("🎯 EXECUTAR AUTOMAÇÃO",
                           type="primary",
                           use_container_width=True,
                           key="run_automation"):
                    
                   
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                   
                    try:
                        # Etapa 1: Abrir ServiceNow
                        status_text.text("🔗 Abrindo ServiceNow...")
                        webbrowser.open("https://centrord.service-now.com/esc?id=sc_cat_item&sys_id=8fd94b438740d51064c5a8e80cbb35dd")
                        progress_bar.progress(20)
                       
                        # Tempo maior para carregamento
                        for i in range(12, 0, -1):
                            status_text.text(f"⏳ Aguardando carregamento da página... {i}s")
                            time.sleep(1)
                       
                        # Etapa 2: Focar na janela
                        status_text.text("🎯 Focando na janela...")
                        snow_bot.focus_browser_window()
                        progress_bar.progress(40)
                       
                        # Etapa 3: Verificação adicional
                        status_text.text("🔍 Verificação final...")
                        st.info("🚨 **ÚLTIMA VERIFICAÇÃO:**")
                        st.info("✅ A janela do ServiceNow está ativa?")
                        st.info("✅ A página está totalmente carregada?")
                        st.info("✅ Você pode ver o formulário RITM?")
                       
                        time.sleep(3)
                        progress_bar.progress(60)
                       
                        # Etapa 4: Preencher formulário
                        status_text.text("⌨️ Preenchendo formulário...")
                        success = snow_bot.fill_snow_form()
                        progress_bar.progress(100)
                       
                        if success:
                            status_text.text("✅ Automação concluída!")
                            st.balloons()
                            st.success(f"Processo finalizado! Colaborador {snow_bot.linha_selecionada['Colaborador']} Verifique o ServiceNow.")
                        else:
                            status_text.text("Erro na automação")
                            st.error("Houve um erro durante o preenchimento.")
                           
                    except Exception as e:
                        st.error(f"Erro inesperado: {e}")
else:
    st.info("👈 Faça upload de uma planilha Excel para começar.")

# Informações técnicas
with st.expander("🔧 INFORMAÇÕES TÉCNICAS"):
    st.markdown("""
    **MELHORIAS IMPLEMENTADAS:**
   
    ✅ **Coordenadas FIXAS**  \\
    ✅ **Verificar os campos necessários para a automação** -- Campos necessários na documentação  \\
    ✅ **Cliques mais precisos**  \\
    ✅ **Direcionamento de capex e opex**  \\
    ✅ **Processo automático** - apenas anexar o arquivo automático
   
    **REQUISITOS:**
    - Resolução: **1920x1080**
    - Sistema: **Windows**
    """)

st.markdown("---")
st.markdown("*Automação ServiceNow v5.0 - Coordenadas Fixas Otimizadas*")