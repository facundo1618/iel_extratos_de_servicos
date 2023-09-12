from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from pathlib import Path
import os
from RPA.Browser.Selenium import Selenium
from RPA.Email.ImapSmtp import ImapSmtp
from time import sleep




def creat_new_service_note(browser: Selenium):
    chrome_prefs = {
    "download.default_directory": f"{Path().home()}\\Desktop\\Palácio\\Códigos\\iel_extratos_de_serviço\\base",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True,
}
    user = "rfcosta"
    password = 'Iel23@'
    url_login = 'https://gestaovagas.iel-ce.org.br/hlogin.aspx'
    url_faturamentos = 'https://gestaovagas.iel-ce.org.br/HResListaFaturamentos.aspx?hitrpaginainicial.aspx'
    id_user = 'vUSU_USUARIO'
    id_password = 'vUSU_SENHA'
    id_login_button = 'BTNLOGIN'
    id_usuarios_test = 'SearchType'
    id_novo_button = 'BTNNOVO'
    id_grupodefaturamento_field = 'RSGRF_ID'
    id_datadefaturamento_field = 'RSFAT_DATA'
    id_salvar_button = 'BTNSALVAR'
    id_imprimir_button = 'BTNIMPRIMIR'
    javascript_code = "document.getElementById('RSFAT_DATA').textContent = '04092023'"

    data_de_faturamento = '30/09/2023'

    # TODO: criar verificação de login
    browser.open_chrome_browser(url=url_login, preferences=chrome_prefs)
    browser.input_text_when_element_is_visible(id_user, user)
    browser.input_text_when_element_is_visible(id_password, password)
    browser.click_element_when_visible(id_login_button)

    browser.wait_until_page_contains_element(id_usuarios_test, 90)
    browser.go_to(url_faturamentos)
    browser.click_element_when_visible(id_novo_button)
    browser.select_from_list_by_value(id_grupodefaturamento_field, '45')
    browser.execute_javascript(f"document.getElementById('RSFAT_DATA').value = '{data_de_faturamento}'")
    # TODO: verificar se download está correto
    browser.click_element_when_visible(id_salvar_button)

    # esperar faturamento
    browser.wait_until_element_is_visible(id_imprimir_button, 1000)
    

def download_service_note(browser: Selenium):
    id_imprimir_button = 'BTNIMPRIMIR'
    id_relatorio_de_conferencia = 'vTIPORELATORIO'
    id_imprimir_2_button = 'BTNPRINT'

    browser.click_element_when_visible(id_imprimir_button)
    browser.wait_until_element_is_visible(id_relatorio_de_conferencia)
    sleep(2)
    browser.select_from_list_by_value(id_relatorio_de_conferencia, 'X')
    browser.click_element_when_visible(id_imprimir_2_button)         
        
def extract_pages_from_pdf(directory):
    pdf = PDF()
    pdf.open_pdf(directory)
    for page_num in range(1, pdf.get_number_of_pages() + 1):
        page_text = pdf.get_text_from_pdf(directory, page_num)
        
        if "Razão Social:" in str(page_text):
            start_page = page_num
        if "Total Cliente:" in str(page_text) and start_page is not None:
            end_page = page_num    
        
        list_of_pages = list(range(start_page, end_page + 1))

        if len(list_of_pages) != 0:
            pdf.extract_pages_from_pdf(directory,f"{Path().home()}\\Desktop\\Palácio\\Códigos\\iel_extratos_de_serviço\\output\\service_note_{start_page} - {end_page}.pdf", list_of_pages)

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

def send_email(company_name: str, email_atribbute: str):
    service_note_path = f"{Path().home()}\\Desktop\\Palácio\\Códigos\\iel_extratos_de_serviço\\output\\{company_name}.pdf"
    gmail_account = "rfcosta@sfiec.org.br"
    gmail_app_password = "japvpatyzugkbruc"
    mail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)
    mail.authorize(account=gmail_account, password=gmail_app_password)
    mail.send_message(sender=gmail_account,
                        recipients= email_atribbute,
                        subject=f"NOTA DE SERVIÇO - {company_name}",
                        body=f"""
Olá!\n
Este é um email teste referente ao RPA de faturamento. Ignore-o.
                        """,
                        attachments=[service_note_path])
    
def get_email_info(browser: Selenium, directory: str):
    list_of_emails = []

    list_of_company_names = get_company_name(directory)
    id_pesquisa_field = 'vPESQUISA'
    url_intranet = 'https://gestaovagas.iel-ce.org.br/HComListaPessoas.aspx?hitrpaginainicial.aspx'
    id_primeirocamporazaosocial_field = 'span_PES_NOME_FANTASIA_0001' 
    id_contato_button = 'Tab_TAB1Containerpanel2'
    id_email_field = 'span_W0324CNT_EMAIL_0001'

    for company_name in list_of_company_names:
        try:
            browser.go_to(url_intranet)
            browser.input_text_when_element_is_visible(id_pesquisa_field, company_name)
            # TODO: criar verificação para nome da empresa
            sleep(1.5)
            browser.click_element_when_visible(id_primeirocamporazaosocial_field)
            browser.click_element_when_visible(id_contato_button)
            email_atribbute = browser.get_text(id_email_field)
            # email_atribbute = 'rfcosta@sfiec.org.br'
            list_of_emails.append(email_atribbute)

            send_email(company_name, email_atribbute)
        except:
            print(f'Não foi possível buscar o email da empresa {company_name}')

if __name__ == "__main__":
    browser = Selenium()
    browser.auto_close = False

    # file_directory = r"base\\grupo_teste_valorizza.pdf"
    creat_new_service_note(browser)
    download_service_note(browser)
    # extract_pages_from_pdf(file_directory)
    # get_company_name_and_rename(r'C:\\Users\\rfcosta\\Desktop\\Palácio\\Códigos\\iel_extratos_de_serviço\\output')
    # enviar e-mails
    pass