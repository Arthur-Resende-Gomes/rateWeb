import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import os
import platform

URL_SERVICENOW = "https://centrord.service-now.com/esc?id=sc_cat_item&sys_id=8fd94b438740d51064c5a8e80cbb35dd"

def get_chrome_options():
    """Configura Chrome Options para diferentes ambientes"""
    options = Options()
    
    # Detecção do ambiente
    is_cloud = any([
        'streamlit' in platform.platform().lower(),
        '/home/appuser' in os.getcwd(),
        '/home/adminuser' in os.getcwd(),
        os.path.exists('/usr/bin/chromium'),
        os.path.exists('/usr/bin/chromium-browser')
    ])
    
    if is_cloud:
        # Configurações para Streamlit Cloud
        st.info("🌐 Detectado ambiente cloud - usando configuração headless")
        
        # Tenta encontrar o binário do Chromium
        chromium_paths = [
            "/usr/bin/chromium",
            "/usr/bin/chromium-browser",
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable"
        ]
        
        chromium_binary = None
        for path in chromium_paths:
            if os.path.exists(path):
                chromium_binary = path
                break
        
        if chromium_binary:
            options.binary_location = chromium_binary
            st.success(f"✅ Chromium encontrado: {chromium_binary}")
        else:
            st.warning("⚠️ Chromium não encontrado nos caminhos padrão")
        
        # Configurações headless obrigatórias
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-plugins")
        options.add_argument("--disable-images")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--remote-debugging-port=9222")
        options.add_argument("--disable-background-timer-throttling")
        options.add_argument("--disable-renderer-backgrounding")
        options.add_argument("--disable-backgrounding-occluded-windows")
        options.add_argument("--disable-web-security")
        options.add_argument("--disable-features=VizDisplayCompositor")
    else:
        # Configurações para ambiente local
        st.info("💻 Detectado ambiente local")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")
    
    return options, is_cloud

def get_chrome_driver():
    """Cria o driver Chrome com configurações apropriadas"""
    options, is_cloud = get_chrome_options()
    
    if is_cloud:
        # Para Streamlit Cloud - usa chromedriver do sistema
        chromedriver_paths = [
            "/usr/bin/chromedriver",
            "/usr/local/bin/chromedriver",
            "/app/.chromedriver/bin/chromedriver"
        ]
        
        chromedriver_path = None
        for path in chromedriver_paths:
            if os.path.exists(path):
                chromedriver_path = path
                break
        
        if chromedriver_path:
            service = Service(chromedriver_path)
            st.success(f"✅ ChromeDriver encontrado: {chromedriver_path}")
        else:
            st.error("❌ ChromeDriver não encontrado")
            service = Service()  # Tenta usar do PATH
    else:
        # Para ambiente local - usa webdriver-manager
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            service = Service(ChromeDriverManager().install())
        except ImportError:
            st.warning("⚠️ webdriver-manager não instalado, tentando usar driver do sistema")
            service = Service()
        except Exception as e:
            st.error(f"❌ Erro com webdriver-manager: {e}")
            service = Service()
    
    return webdriver.Chrome(service=service, options=options)

def wait_and_click(driver, xpath, timeout=15):
    element = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    element.click()

def wait_and_send_keys(driver, element_id, text, timeout=15):
    element = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.ID, element_id))
    )
    element.clear()
    element.send_keys(text)

def robust_select2_fill(driver, container_ids, search_text, max_attempts=3):    
    for container_id in container_ids:
        for attempt in range(max_attempts):
            try:
                print(f"Tentativa {attempt + 1} para {container_id} com texto '{search_text}'")
                
                try:
                    container = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, container_id))
                    )
                    if not container.is_displayed():
                        continue
                except:
                    continue
                
                try:
                    driver.execute_script(f"document.getElementById('{container_id}').click();")
                    time.sleep(1)
                    
                    # Procura campo de busca ativo
                    search_input = None
                    search_selectors = [
                        "input.select2-input:focus",
                        ".select2-dropdown-open input.select2-input",
                        ".select2-search input[type='text']",
                        "input[id*='autogen'][id*='search']:focus"
                    ]
                    
                    for selector in search_selectors:
                        try:
                            search_input = driver.find_element(By.CSS_SELECTOR, selector)
                            if search_input.is_displayed():
                                break
                        except:
                            continue
                    
                    if search_input:
                        search_input.clear()
                        search_input.send_keys(search_text)
                        time.sleep(2)
                        
                        # Clica no resultado
                        result = WebDriverWait(driver, 8).until(
                            EC.element_to_be_clickable((By.XPATH, f"//div[contains(@class,'select2-result-label') and contains(text(), '{search_text}')]"))
                        )
                        result.click()
                        time.sleep(1)
                        
                        # Verifica se foi selecionado
                        try:
                            selected = driver.find_element(By.CSS_SELECTOR, f"#{container_id} .select2-chosen")
                            if search_text.lower() in selected.text.lower():
                                print(f"✅ Sucesso com {container_id}")
                                return True
                        except:
                            pass
                
                except Exception as e:
                    print(f"Estratégia 1 falhou: {e}")
                
                # ESTRATÉGIA 2: JavaScript direto
                try:
                    # Fecha qualquer dropdown aberto
                    driver.execute_script("jQuery('.select2-drop').hide();")
                    time.sleep(0.5)
                    
                    # Abre o dropdown
                    driver.execute_script(f"""
                        var container = jQuery('#{container_id}');
                        container.select2('open');
                    """)
                    time.sleep(1)
                    
                    # Digita o texto
                    driver.execute_script(f"""
                        var searchField = jQuery('.select2-search input:visible');
                        if (searchField.length > 0) {{
                            searchField.val('{search_text}');
                            searchField.trigger('keyup');
                        }}
                    """)
                    time.sleep(2)
                    
                    # Clica no primeiro resultado visível
                    result = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, f"//div[contains(@class,'select2-result-label') and contains(text(), '{search_text}')]"))
                    )
                    result.click()
                    time.sleep(1)
                    
                    # Verifica se foi selecionado
                    try:
                        selected = driver.find_element(By.CSS_SELECTOR, f"#{container_id} .select2-chosen")
                        if search_text.lower() in selected.text.lower():
                            print(f"✅ Sucesso com JavaScript para {container_id}")
                            return True
                    except:
                        pass
                        
                except Exception as e:
                    print(f"Estratégia 2 falhou: {e}")
                
                # ESTRATÉGIA 3: Força bruta com Enter
                try:
                    container.click()
                    time.sleep(1)
                    
                    # Envia o texto + Enter
                    active_element = driver.switch_to.active_element
                    active_element.clear()
                    active_element.send_keys(search_text)
                    time.sleep(2)
                    active_element.send_keys(Keys.ENTER)
                    time.sleep(1)
                    
                    # Verifica se funcionou
                    try:
                        selected = driver.find_element(By.CSS_SELECTOR, f"#{container_id} .select2-chosen")
                        if search_text.lower() in selected.text.lower():
                            print(f"✅ Sucesso com Enter para {container_id}")
                            return True
                    except:
                        pass
                        
                except Exception as e:
                    print(f"Estratégia 3 falhou: {e}")
                
                # Pequena pausa antes da próxima tentativa
                time.sleep(1)
                
            except Exception as e:
                print(f"Erro geral na tentativa {attempt + 1} para {container_id}: {e}")
                time.sleep(1)
                continue
    
    print(f"❌ Falha em todos os containers para '{search_text}'")
    return False

def preencher_campo_dinamico(driver, field_identifiers, text, timeout=15):
    """Preenche campos que podem ter IDs dinâmicos"""
    for identifier in field_identifiers:
        try:
            if identifier.startswith("//"):
                element = WebDriverWait(driver, timeout).until(
                    EC.presence_of_element_located((By.XPATH, identifier))
                )
            else:
                element = WebDriverWait(driver, timeout).until(
                    EC.presence_of_element_located((By.ID, identifier))
                )
            element.clear()
            element.send_keys(text)
            return True
        except:
            continue
    return False

def debug_environment():
    """Função para debugar o ambiente"""
    st.subheader("🔍 Informações do Ambiente")
    
    env_info = {
        "Sistema": platform.system(),
        "Plataforma": platform.platform(),
        "Diretório atual": os.getcwd(),
        "HOME": os.environ.get('HOME', 'N/A'),
        "PATH": os.environ.get('PATH', 'N/A')[:200] + "...",
    }
    
    # Verifica binários
    binaries = {
        "Chromium": ["/usr/bin/chromium", "/usr/bin/chromium-browser"],
        "Chrome": ["/usr/bin/google-chrome", "/usr/bin/google-chrome-stable"],
        "ChromeDriver": ["/usr/bin/chromedriver", "/usr/local/bin/chromedriver"]
    }
    
    for name, paths in binaries.items():
        found = [path for path in paths if os.path.exists(path)]
        env_info[f"{name}"] = found if found else "❌ Não encontrado"
    
    for key, value in env_info.items():
        st.write(f"**{key}**: {value}")

def preencher_servicenow(dados):
    try:
        st.info("🚀 Iniciando configuração do driver...")
        
        # Cria o driver com configurações apropriadas
        driver = get_chrome_driver()
        st.success("✅ Driver Chrome inicializado com sucesso!")

        driver.get(URL_SERVICENOW)
        st.info("🌐 Navegando para ServiceNow...")

        st.warning("⚠️ **ATENÇÃO**: Faça login manual no ServiceNow e aguarde o carregamento do formulário...")
        st.info("⏳ Aguardando formulário carregar (timeout: 2 minutos)...")
        
        # Aguarda o formulário carregar
        try:
            WebDriverWait(driver, 120).until(
                EC.presence_of_element_located((By.ID, "s2id_sp_formfield_u_rd_selecione_o_fluxo_de_trabalho_novo"))
            )
            st.success("✅ Formulário carregado!")
        except Exception as e:
            st.error(f"❌ Timeout aguardando formulário: {e}")
            driver.quit()
            return

        st.info("📝 Iniciando preenchimento do formulário...")

        # --- Fluxo de trabalho ---
        driver.find_element(By.ID, "s2id_sp_formfield_u_rd_selecione_o_fluxo_de_trabalho_novo").click()
        wait_and_click(driver, "//div[contains(@class,'select2-result-label') and text()='Contratação de Rate Card']")

        # --- Radio buttons ---
        alocacao = dados["Alocação"].strip().lower()
        radios_spans = driver.find_elements(
            By.XPATH,
            "//input[@name='radio_button_across_4c12d8058717d21082f586270cbb351f_8fd94b438740d51064c5a8e80cbb35dd']/following-sibling::span"
        )
        if alocacao == "capex":
            radios_spans[0].click()
        elif alocacao == "opex":
            radios_spans[1].click()
        
        time.sleep(5)

        driver.find_element(By.ID, "s2id_sp_formfield_u_rd_tipo_contratacao_ratecard").click()
        wait_and_click(driver, "//div[contains(@class,'select2-result-label') and text()='Contratação Ratecard fornecedor Ativo']")

        driver.find_element(By.ID, "s2id_sp_formfield_u_rd_fornecedor").click()
        driver.find_element(By.ID, "s2id_autogen7_search").send_keys(dados['Fornecedor'])
        time.sleep(1)
        wait_and_click(driver, f"//div[contains(@class,'select2-result-label') and contains(., '{dados['Fornecedor']}')]")
        
        # --- Consultor ---
        wait_and_send_keys(driver, "sp_formfield_u_rd_matricula_pagto_rate", str(dados["Matricula"]))
        wait_and_send_keys(driver, "sp_formfield_u_rd_nome_completo_pagto_rate", dados["Colaborador"])
        wait_and_send_keys(driver, "sp_formfield_u_rd_cargo_pagto_rate", dados["Perfil"])
        
        wait_and_send_keys(driver, "sp_formfield_u_rd_valor_mensal_168hr", dados["VALOR"])
        wait_and_send_keys(driver, "sp_formfield_u_rd_valor_total_fxpr", dados["PREVISTO"])

        # --- Escopo ---
        wait_and_send_keys(
            driver,
            "sp_formfield_u_rd_descreva_escopo",
            f"Abertura de chamado para {dados['Perfil']} no valor mensal de {dados['VALOR']} com o centro de custo {dados['Centro']}"
        )
        
        codigo_identifiers = [
            "sp_formfield_u_rd_codigo_projeto",
            "sp_formfield_u_rd_codigo_orcamento",
            "//input[contains(@id, 'codigo_projeto')]",
            "//input[contains(@id, 'codigo_orcamento')]"
        ]
        preencher_campo_dinamico(driver, codigo_identifiers, dados["Centro"])
        
        if alocacao == "opex":
            # Ordem estatística
            ordem_identifiers = [
                "sp_formfield_u_rd_ordem_estatistica",
                "sp_formfield_u_rd_ordem_estatisica",
                "//input[contains(@id, 'ordem_estat')]"
            ]
            preencher_campo_dinamico(driver, ordem_identifiers, dados["Ordem estatisica"])
            
            # Valor do orçamento
            valor_orcamento_identifiers = [
                "sp_formfield_u_rd_valor_do_orcamento",
                "//input[contains(@id, 'valor_do_orcamento')]",
                "//input[contains(@id, 'valor_orcamento')]"
            ]
            if "Valor do orçamento" in dados:
                preencher_campo_dinamico(driver, valor_orcamento_identifiers, dados["Valor do orçamento"])
            elif "VALOR" in dados:
                preencher_campo_dinamico(driver, valor_orcamento_identifiers, dados["VALOR"])
        
        time.sleep(3)
        area_solicitante_containers = [
            "s2id_sp_formfield_u_rd_qual_centro_custo_area_solicitante",
            "s2id_sp_formfield_u_rd_centro_custo_area_solicitante",
            "s2id_sp_formfield_u_rd_area_solicitante"
        ]
        robust_select2_fill(driver, area_solicitante_containers, dados['Área solicitante'])
        
        time.sleep(3)
        area_destino_containers = [
            "s2id_sp_formfield_u_rd_centro_custo_destino",
            "s2id_sp_formfield_u_rd_area_destino",
            "s2id_sp_formfield_u_rd_destino"
        ]
        robust_select2_fill(driver, area_destino_containers, dados['Área destino'])

        time.sleep(3)
        diretoria_containers = [
            "s2id_sp_formfield_u_rd_diretoria",
            "s2id_sp_formfield_u_rd_diretor",
            "s2id_sp_formfield_u_rd_diretoria_responsavel"
        ]
        robust_select2_fill(driver, diretoria_containers, dados['Diretoria'])
        
        time.sleep(3)
        gerencia_containers = [
            "s2id_sp_formfield_u_rd_gerencia",
            "s2id_sp_formfield_u_rd_area",
            "s2id_sp_formfield_u_rd_gerencia_responsavel"
        ]
        robust_select2_fill(driver, gerencia_containers, dados['Área'])
        
        # --- Diretor ---
        time.sleep(3)
        try:
            diretor_span = driver.find_element(By.ID, "select2-chosen-18")
            diretor_span.click()
            time.sleep(2)
            
            diretor_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "s2id_autogen18_search"))
            )
            diretor_input.clear()
            diretor_input.send_keys(dados['Diretor'])
            time.sleep(3)
            
            try:
                diretor_result = WebDriverWait(driver, 8).until(
                    EC.element_to_be_clickable((By.XPATH, f"//div[contains(@class,'select2-result-label') and contains(text(), '{dados['Diretor']}')]"))
                )
                diretor_result.click()
                print("✅ Campo Diretor preenchido com sucesso!")
            except:
                try:
                    first_result = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".select2-result-label"))
                    )
                    first_result.click()
                    print("✅ Campo Diretor preenchido com primeiro resultado disponível!")
                except:
                    print("❌ Falha ao selecionar resultado para Diretor")
                    
        except Exception as e:
            print(f"❌ Erro ao preencher campo Diretor: {e}")
        
        # --- Gestor ---
        time.sleep(3)
        try:
            gestor_span = driver.find_element(By.ID, "select2-chosen-19")
            gestor_span.click()
            time.sleep(2)
            
            gestor_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "s2id_autogen19_search"))
            )
            gestor_input.clear()
            gestor_input.send_keys(dados['Gestor'])
            time.sleep(3)
            try:
                gestor_result = WebDriverWait(driver, 8).until(
                    EC.element_to_be_clickable((By.XPATH, f"//div[contains(@class,'select2-result-label') and contains(text(), '{dados['Gestor']}')]"))
                )
                gestor_result.click()
                print("✅ Campo Gestor preenchido com sucesso!")
            except:
                try:
                    first_result = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".select2-result-label"))
                    )
                    first_result.click()
                    print("✅ Campo Gestor preenchido com primeiro resultado disponível!")
                except:
                    print("❌ Falha ao selecionar resultado para Gestor")
                    
        except Exception as e:
            print(f"❌ Erro ao preencher campo Gestor: {e}")

        st.success("✅ Formulário preenchido com sucesso!")
        
        # Comportamento diferente para ambiente cloud vs local
        _, is_cloud = get_chrome_options()
        if is_cloud:
            st.info("🌐 Ambiente cloud: aguardando 60 segundos antes de fechar o navegador...")
            st.warning("⚠️ **IMPORTANTE**: Revise rapidamente o formulário e finalize se necessário!")
            time.sleep(60)
            driver.quit()
            st.info("🔒 Navegador fechado automaticamente (ambiente cloud)")
        else:
            st.info("💻 Ambiente local: navegador permanecerá aberto para revisão...")
            st.warning("⚠️ **REVISE** o formulário no navegador antes de submeter!")
            
            # Aguarda indefinidamente ou até erro
            try:
                WebDriverWait(driver, 3600).until(lambda d: False)
            except:
                pass

    except Exception as e:
        st.error(f"❌ Erro durante a automação: {e}")
        import traceback
        st.error(f"**Detalhes do erro:**\n```\n{traceback.format_exc()}\n```")
        
        # Tenta fechar o driver se existir
        try:
            driver.quit()
        except:
            pass

# --- Interface Streamlit ---
st.title("🚀 Automação ServiceNow - Preenchimento de Formulário")

# Mostrar informações do ambiente
_, is_cloud = get_chrome_options()
if is_cloud:
    st.success("🌐 **Executando em ambiente cloud** - Selenium configurado para modo headless")
    
    with st.expander("🔍 Debug do Ambiente"):
        debug_environment()
else:
    st.info("💻 **Executando em ambiente local** - Selenium abrirá navegador normalmente")

uploaded_file = st.file_uploader("📂 Carregue a planilha Excel", type=["xlsx", "xls"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        aba = st.selectbox("Escolha a aba da planilha:", xls.sheet_names)

        df = pd.read_excel(uploaded_file, sheet_name=aba)
        
        st.success(f"✅ Planilha carregada com sucesso! {len(df)} registros encontrados.")
        
        with st.expander("👀 Visualizar dados da planilha"):
            st.dataframe(df)

        if "Colaborador" in df.columns:
            pessoa = st.selectbox("Selecione a pessoa:", df["Colaborador"].tolist())
            dados_pessoa = df[df["Colaborador"] == pessoa].iloc[0].to_dict()

            st.subheader("📋 Dados que serão preenchidos:")
            
            col1, col2 = st.columns(2)
            with col1:
                st.write("**👤 Informações Pessoais:**")
                st.write(f"• **Colaborador**: {dados_pessoa.get('Colaborador', 'N/A')}")
                st.write(f"• **Matrícula**: {dados_pessoa.get('Matricula', 'N/A')}")
                st.write(f"• **Perfil**: {dados_pessoa.get('Perfil', 'N/A')}")
                
            with col2:
                st.write("**💰 Valores:**")
                st.write(f"• **Valor Mensal**: {dados_pessoa.get('VALOR', 'N/A')}")
                st.write(f"• **Valor Previsto**: {dados_pessoa.get('PREVISTO', 'N/A')}")
                st.write(f"• **Centro de Custo**: {dados_pessoa.get('Centro', 'N/A')}")

            if st.button("🚀 **EXECUTAR AUTOMAÇÃO**", type="primary"):
                st.info("🚀 Iniciando automação...")
                preencher_servicenow(dados_pessoa)
                
        else:
            st.error("❌ A coluna 'Colaborador' não foi encontrada na planilha.")
            st.info("**Colunas disponíveis**: " + ", ".join(df.columns.tolist()))
            
    except Exception as e:
        st.error(f"❌ Erro ao processar a planilha: {str(e)}")

# Footer
st.markdown("---")
st.markdown("""
### 📌 **Instruções Importantes:**

**Para ambiente Cloud:**
- O navegador roda em modo headless (sem interface gráfica)
- Você terá 60 segundos para revisar o formulário
- Faça o login rapidamente quando solicitado

**Para ambiente Local:**
- O navegador abre normalmente na tela
- Você pode revisar com calma antes de submeter
- O navegador fica aberto até você fechar manualmente

**⚠️ Lembre-se:** Sempre revise o formulário antes de submeter no ServiceNow!
""")
