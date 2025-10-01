import streamlit as st # Web
import pandas as pd # Pegar df
from selenium import webdriver # Pra ir pra web
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import time

URL_SERVICENOW = "https://centrord.service-now.com/esc?id=sc_cat_item&sys_id=8fd94b438740d51064c5a8e80cbb35dd" # URL service now

def wait_and_click(driver, xpath, timeout=15): #Espera do elemnto aparecer para ser clicavel , passando o driver* , xpath (caminho da class ou id) e o tempo
    element = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    element.click()

def wait_and_send_keys(driver, element_id, text, timeout=15): # Espera tambem mas com keys passando o texto
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
                                print(f"Sucesso com {container_id}")
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
                            print(f"Sucesso com JavaScript para {container_id}")
                            return True
                    except:
                        pass
                        
                except Exception as e:
                    print(f"Estratégia 2 falhou: {e}")
                
                # Força bruta com Enter
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
                            print(f"Sucesso com Enter para {container_id}")
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
    
    print(f"Falha em todos os containers para '{search_text}'")
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

def debug_select2_containers(driver):
    
    try:
        containers = driver.find_elements(By.CSS_SELECTOR, "[id*='s2id_']")
        visible_containers = []
        for container in containers:
            if container.is_displayed():
                container_id = container.get_attribute('id')
                try:
                    label_element = driver.find_element(By.CSS_SELECTOR, f"label[for='{container_id.replace('s2id_', '')}']")
                    label_text = label_element.text
                except:
                    label_text = "Sem label"
                
                visible_containers.append({
                    'id': container_id,
                    'label': label_text
                })
        
        print("=== CONTAINERS SELECT2 VISÍVEIS ===")
        for container in visible_containers:
            print(f"ID: {container['id']} | Label: {container['label']}")
        print("=== FIM DEBUG ===")
        
        return visible_containers
    except Exception as e:
        print(f"Erro no debug: {e}")
        return []

def preencher_servicenow(dados):
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    try:
        driver.get(URL_SERVICENOW)

        st.info("Faça login manual no ServiceNow e aguarde o carregamento do formulário...")
        WebDriverWait(driver, 120).until(
            EC.presence_of_element_located((By.ID, "s2id_sp_formfield_u_rd_selecione_o_fluxo_de_trabalho_novo"))
        )

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
        
        
        wait_and_send_keys(driver, "sp_formfield_u_rd_matricula_pagto_rate", str(dados["Matricula"]))
        wait_and_send_keys(driver, "sp_formfield_u_rd_nome_completo_pagto_rate", dados["Colaborador"])
        wait_and_send_keys(driver, "sp_formfield_u_rd_cargo_pagto_rate", dados["Perfil"])
        
        wait_and_send_keys(driver, "sp_formfield_u_rd_valor_mensal_168hr", dados["VALOR"])
        wait_and_send_keys(driver, "sp_formfield_u_rd_valor_total_fxpr", dados["PREVISTO"])

        # --- Escopo ---
        wait_and_send_keys(
            driver,
            "sp_formfield_u_rd_descreva_escopo",
            f"Abertura de chamado para {dados['Perfil']} no valor mensal de {dados['VALOR']} no {dados['Centro']}"
        )
        
        
        codigo_identifiers = [
            "sp_formfield_u_rd_codigo_projeto",
            "sp_formfield_u_rd_codigo_orcamento",
            "//input[contains(@id, 'codigo_projeto')]",
            "//input[contains(@id, 'codigo_orcamento')]"
        ]
        preencher_campo_dinamico(driver, codigo_identifiers, dados["Centro"].strip())
        
        if alocacao == "opex":
            # Ordem estatística
            ordem_identifiers = [
                "sp_formfield_u_rd_ordem_estatistica",
                "sp_formfield_u_rd_ordem_estatisica",
                "//input[contains(@id, 'ordem_estat')]"
            ]
            preencher_campo_dinamico(driver, ordem_identifiers, dados["Ordem estatisica"])
            
            # Valor do orçamento (NOVO CAMPO ADICIONADO)
            valor_orcamento_identifiers = [
                "sp_formfield_u_rd_valor_do_orcamento",
                "//input[contains(@id, 'valor_do_orcamento')]",
                "//input[contains(@id, 'valor_orcamento')]"
            ]
            # Verifica se existe a coluna no dados
            if "Valor do orçamento" in dados:
                preencher_campo_dinamico(driver, valor_orcamento_identifiers, dados["Valor do orçamento"])
            elif "VALOR" in dados:  # Fallback para usar o mesmo valor
                preencher_campo_dinamico(driver, valor_orcamento_identifiers, dados["VALOR"])
        

        time.sleep(2)
        area_solicitante_containers = [
            "s2id_sp_formfield_u_rd_qual_centro_custo_area_solicitante",
            "s2id_sp_formfield_u_rd_centro_custo_area_solicitante",
            "s2id_sp_formfield_u_rd_area_solicitante"
        ]
        robust_select2_fill(driver, area_solicitante_containers, dados['Área solicitante'])
        

        time.sleep(2)
        area_destino_containers = [
            "s2id_sp_formfield_u_rd_centro_custo_destino",
            "s2id_sp_formfield_u_rd_area_destino",
            "s2id_sp_formfield_u_rd_destino"
        ]
        robust_select2_fill(driver, area_destino_containers, dados['Área destino'])


        time.sleep(2)
        diretoria_containers = [
            "s2id_sp_formfield_u_rd_diretoria",
            "s2id_sp_formfield_u_rd_diretor",
            "s2id_sp_formfield_u_rd_diretoria_responsavel"
        ]
        robust_select2_fill(driver, diretoria_containers, dados['Diretoria'])
        
        # --- Gerência/Área (CORRIGIDO COM NOVA FUNÇÃO) ---
        time.sleep(2)
        gerencia_containers = [
            "s2id_sp_formfield_u_rd_gerencia",
            "s2id_sp_formfield_u_rd_area",
            "s2id_sp_formfield_u_rd_gerencia_responsavel"
        ]
        robust_select2_fill(driver, gerencia_containers, dados['Área'])
        
        # --- Diretor (MÉTODO DIRETO COM IDs CORRETOS) ---
        time.sleep(2)
        print("Tentando preencher campo Diretor...")
        
        try:
            # Clica no span para abrir o dropdown do Diretor
            diretor_span = driver.find_element(By.ID, "select2-chosen-18")
            diretor_span.click()
            time.sleep(2)
            
            # Digita no campo de busca específico do Diretor
            diretor_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "s2id_autogen18_search"))
            )
            diretor_input.clear()
            diretor_input.send_keys(dados['Diretor'])
            time.sleep(2)
            
            # Clica no resultado
            try:
                diretor_result = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, f"//div[contains(@class,'select2-result-label') and contains(text(), '{dados['Diretor']}')]"))
                )
                diretor_result.click()
                print("Campo Diretor preenchido com sucesso!")
            except:
                # Fallback: clica no primeiro resultado visível
                try:
                    first_result = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".select2-result-label"))
                    )
                    first_result.click()
                    print("Campo Diretor preenchido com primeiro resultado disponível!")
                except:
                    print("Falha ao selecionar resultado para Diretor")
                    
        except Exception as e:
            print(f"Erro ao preencher campo Diretor: {e}")
        
        # --- Gestor (MÉTODO DIRETO COM IDs CORRETOS) ---
        time.sleep(2)
        print("Tentando preencher campo Gestor...")
        
        try:
            # Clica no span para abrir o dropdown do Gestor
            gestor_span = driver.find_element(By.ID, "select2-chosen-19")
            gestor_span.click()
            time.sleep(2)
            
            gestor_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "s2id_autogen19_search"))
            )
            gestor_input.clear()
            gestor_input.send_keys(dados['Gestor'])
            time.sleep(3)
            try:
                gestor_result = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, f"//div[contains(@class,'select2-result-label') and contains(text(), '{dados['Gestor']}')]"))
                )
                gestor_result.click()
                print("Campo Gestor preenchido com sucesso!")
            except:
                try:
                    first_result = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".select2-result-label"))
                    )
                    first_result.click()
                    print("Campo Gestor preenchido com primeiro resultado disponível!")
                except:
                    print("Falha ao selecionar resultado para Gestor")
                    
        except Exception as e:
            print(f"Erro ao preencher campo Gestor: {e}")

        st.success("Formulário preenchido! Confira o navegador. Ele permanecerá aberto para revisão.")
        
        WebDriverWait(driver, 3600).until(lambda d: False)

    except Exception as e:
        st.error(f"Erro durante a automação: {e}")
        time.sleep(10)
    finally:
        pass

st.title("Automação ServiceNow - Preenchimento de Formulário")

uploaded_file = st.file_uploader("Carregue a planilha Excel", type=["xlsx", "xls"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    aba = st.selectbox("Escolha a aba da planilha:", xls.sheet_names)

    df = pd.read_excel(uploaded_file, sheet_name=aba, dtype={"Matricula": str, "Ordem estatisica": str})
    st.dataframe(df)

    if "Colaborador" in df.columns:
        pessoa = st.selectbox("Selecione a pessoa:", df["Colaborador"].tolist())
        dados_pessoa = df[df["Colaborador"] == pessoa].iloc[0].to_dict()

        st.write("Dados selecionados:")
        st.json(dados_pessoa)

        if st.button("Executar Automação"):
            st.success("Rodando automação... o navegador vai abrir")
            preencher_servicenow(dados_pessoa)
    else:
        st.error("A coluna 'Colaborador' não foi encontrada na planilha.")
