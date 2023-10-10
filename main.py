from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from pathlib import Path
import os
from RPA.Browser.Selenium import Selenium
from RPA.Email.ImapSmtp import ImapSmtp
from time import sleep
from RPA.Assistant import Assistant
import ctypes  
import sys
from datetime import date

def start_dialog():
    dialog = Assistant()
    # dialog.add_image(r"imagens\\iel.jpg", 250, 150)
    dialog.add_text_input('date_input', 'Data de faturamento no formato DD/MM/AAAA', required=True)
    dialog.add_submit_buttons(["Executar", "Cancelar"])
    user_input = dialog.run_dialog(90000, title="Assistente de Faturamento", height=185, width=450, location="Center")

    return user_input

def clear_directory(directory_1, directory_2):
    file = FileSystem()
    try:
        file.empty_directory(directory_1)
        file.empty_directory(directory_2)
        print("diretórios limpos!")
    except:
        print('Não foi possível limpar os diretórios')

def creat_new_service_note(browser: Selenium, faturamento_date):
    chrome_prefs = {
    "download.default_directory": f"{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\base",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True,
}
    user = "ielce.rpa"
    password = 'Ielce2@'
    url_login = 'https://gestaovagas.iel-ce.org.br/hlogin.aspx'
    url_faturamentos = 'https://gestaovagas.iel-ce.org.br/HResListaFaturamentos.aspx?hitrpaginainicial.aspx'
    id_user = 'vUSU_USUARIO'
    id_password = 'vUSU_SENHA'
    id_login_button = 'BTNLOGIN'
    id_usuarios_test = 'SearchType'
    id_novo_button = 'BTNNOVO'
    id_grupodefaturamento_field = 'RSGRF_ID'
    id_salvar_button = 'BTNSALVAR'
    id_imprimir_button = 'BTNIMPRIMIR'

    # TODO: criar verificação de login
    browser.open_chrome_browser(url=url_login, preferences=chrome_prefs)

    def check_login(browser: Selenium):
        test = browser.does_page_contain_element(id_user)
        print(test)

        return test

    count = 0
    while True:
        try:
            if count == 3:
                print('Foram efetuadas 3 tentativas de login sem êxito. Revise as credenciais')
                browser.close_browser()
                break
            test = check_login(browser)
            if test == False:
                break
            else:
                print('Fazendo login...')
                browser.press_keys(id_user, 'CTRL + A')
                browser.input_text_when_element_is_visible(id_user, user)
                browser.press_keys(id_password, 'CTRL + A')
                browser.input_text_when_element_is_visible(id_password, password)
                browser.click_element_when_visible(id_login_button)
                count +=1
        except:
            print('Login validado')

    browser.wait_until_page_contains_element(id_usuarios_test, 90)
    browser.go_to(url_faturamentos)
    browser.click_element_when_visible(id_novo_button)
    browser.select_from_list_by_value(id_grupodefaturamento_field, '46')
    browser.execute_javascript(f"document.getElementById('RSFAT_DATA').value = '{faturamento_date}'")
    browser.click_element_when_visible(id_salvar_button)

    # esperar faturamento
    browser.wait_until_element_is_visible(id_imprimir_button, 1000)

def download_service_note(browser: Selenium, file_path):
    id_imprimir_button = 'BTNIMPRIMIR'
    id_relatorio_de_conferencia = 'vTIPORELATORIO'
    id_imprimir_2_button = 'BTNPRINT'

    browser.click_element_when_visible(id_imprimir_button)
    browser.wait_until_element_is_visible(id_relatorio_de_conferencia)
    sleep(2)
    browser.select_from_list_by_value(id_relatorio_de_conferencia, 'X')
    browser.click_element_when_visible(id_imprimir_2_button)    

    def wait_until_download_finish(file_path):
        file = FileSystem()
        test = file.does_file_exist(file_path)
    
        return test
    
    while True:
        test = wait_until_download_finish(file_path)
        if test == True:
            print("download finalizado")
            break
        else:
            print("aguardando download finalizar")
        sleep(1)

def extract_pages_from_pdf(file_path):
    pdf = PDF()
    pdf.open_pdf(file_path)
    for page_num in range(1, pdf.get_number_of_pages() + 1):
        page_text = pdf.get_text_from_pdf(file_path, page_num)
        
        if "Razão Social:" in str(page_text):
            start_page = page_num
        if "Total Cliente:" in str(page_text) and start_page is not None:
            end_page = page_num    
        
        list_of_pages = list(range(start_page, end_page + 1))

        if len(list_of_pages) != 0:
            pdf.extract_pages_from_pdf(file_path, f"{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\output\\service_note_{start_page} - {end_page}.pdf", list_of_pages)

def get_company_name_and_rename(directory: str):
    pdf = PDF()
    file_ = FileSystem()
    list_of_files = os.listdir(directory)
    for file in list_of_files:
        try:
            pdf.open_pdf(f"{directory}\\{file}") 
            company_name = pdf.find_text("Razão Social:", 1) 
            company_name = company_name[0].neighbours[0]
            company_name = company_name.replace("/", "-")
            pdf.close_pdf()
        except:
            print(f'Não foi possível buscar o nome da empresa no arquivo {file}')

        try:
            old_name = f"{directory}\\{file}"
            new_name = f"{directory}\\{company_name}.pdf"

            file_.move_file(old_name, new_name)
        except:
            old_name = f"{directory}\\{file}"
            new_name = f"{directory}\\{company_name}_mesmo_nome.pdf"

            file_.move_file(old_name, new_name)   

def get_company_name(directory): # função redundante. Aprimorar em versões posteriores
    pdf = PDF()
    list_of_files = os.listdir(directory)
    errors = []
    list_of_names = []
    for file in list_of_files:
        try:
            pdf.open_pdf(f"{directory}\\{file}") 
            company_name = pdf.find_text("Razão Social:", 1) 
            company_name = company_name[0].neighbours[0]
            list_of_names.append(company_name) 
        except:
            errors.append(file)

    return list_of_names

def send_email(company_name: str, email_atribbute: list[str]):
    service_note_path = f"{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\output\\{company_name}.pdf"
    gmail_account = "ielce.rpa@sfiec.org.br"
    gmail_app_password = "rbew hgih nqco nszj"
    mail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)
    mail.authorize(account=gmail_account, password=gmail_app_password)
    mail.send_message(sender=gmail_account,
                        recipients= email_atribbute,
                        subject=f"(TESTE DO RPA. IGNORE) RELATÓRIO DE FATURAMENTO DE ESTAGIÁRIOS IEL - AGOSTO DE 2023", #automatizar mês e ano
                        body=f"""
Boa tarde!\n
Segue em anexo, relação referente ao serviço Gestão de Estágio - Mês de referência XXXX/2023.\n
Aguardamos a confirmação do relatório até o dia  XX/XX/2023 , pois iniciaremos a emissão das Notas Fiscais.
                        """,
                        attachments=[service_note_path])
    
def get_email_info(browser: Selenium, directory: str):
    list_of_emails = [] #para criar um relatório na próxima versão
    list_of_errors = []
    count = 0

    list_of_company_names = get_company_name(directory)
    id_pesquisa_field = 'vPESQUISA'
    url_intranet = 'https://gestaovagas.iel-ce.org.br/HComListaPessoas.aspx?hitrpaginainicial.aspx'
    id_primeirocamporazaosocial_field = 'span_vPES_RAZAO_SOCIAL_0001'
    id_contato_button = 'Tab_TAB1Containerpanel2'
    id_email_field = 'span_W0324CNT_EMAIL_0001'
    js_code = '''// Pegar linhas
const linhas = document.querySelectorAll('#W0324GridContainerTbl .WorkWithOdd')

// Declarar array onde serão colocados os emails
let email_cobranca = []

// Iterar Linhas
linhas.forEach(linha=>{
    // Para cada linha, pegar checkbox da terceira coluna 
    const check_cobranca = linha.querySelector('[data-colindex="3"] input[type="checkbox"]')
    
    // Para cada linha, pegar conteudo da 6° coluna (email) 
    const email = linha.querySelector('[data-colindex="6"] span').innerText
    
    // Se o checkbox tem o valor='S' (está marcado), pegue o email e coloque no array email_cobranca
    if(check_cobranca.value=='S'){
        email_cobranca.push(email)
    }
})

// Printar email_cobranca
return(email_cobranca)'''

    for company_name in list_of_company_names:
        try:
            browser.go_to(url_intranet)
            browser.input_text_when_element_is_visible(id_pesquisa_field, company_name)
            sleep(5)
            company_name_verification = browser.get_text(id_primeirocamporazaosocial_field)
            print(f'{company_name_verification} é igual a {company_name}')
            if company_name_verification == company_name:
                browser.click_element_when_visible(id_primeirocamporazaosocial_field)
                browser.click_element_when_visible(id_contato_button)
                sleep(2)
                # email_atribbute = browser.get_text(id_email_field)
                email_atribbute = browser.execute_javascript(js_code)
                print(email_atribbute)
                list_of_emails.append(email_atribbute)
                email_atribbute = ['rfcosta@sfiec.org.br', 'gpmonteiro@sfiec.org.br'] 
                sleep(1)
                send_email(company_name, email_atribbute)
                sleep(1)
                count +=1
            else:
                print('não foi possível verificar o nome da empresa')

        except:
            print(f'Não foi possível buscar o email da empresa {company_name}')
            list_of_errors.append(company_name)

    sleep(3)
    browser.close_browser()

    return count, list_of_errors, list_of_emails

if __name__ == "__main__":
    browser = Selenium()
    browser.auto_close = False

    user_input = start_dialog()
    if user_input.get("submit") != "Executar":
        sys.exit()
    else:
        file_path = f"{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\base\\arresrelfatextratoestagios.pdf"
        directory = f"{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\output"

        clear_directory(f"{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\output", 
                        f"{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\base")
        creat_new_service_note(browser, user_input.get('date_input')) #criar tratativa para o caso de já existir um grupo igual
        download_service_note(browser, file_path)
        extract_pages_from_pdf(file_path)
        get_company_name_and_rename(directory)
        number_of_emails, list_of_errors, list_of_emails = get_email_info(browser, directory)
        ctypes.windll.user32.MessageBoxW(0, f"Total de extratos enviados: {number_of_emails}", "Aviso", 1)
        ctypes.windll.user32.MessageBoxW(0, f'Lista de empresas que foram feitos os envios: {list_of_emails}', 'Aviso', 1)
        if len(list_of_errors) != 0:
            ctypes.windll.user32.MessageBoxW(0, f'lista de empresas que não foi possível fazer o envio: {list_of_errors}', 'Aviso', 1)

