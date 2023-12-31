from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from pathlib import Path
import os
from RPA.Browser.Selenium import Selenium
from RPA.Email.ImapSmtp import ImapSmtp
from time import sleep
from RPA.Assistant import Assistant

def start_dialog():
    dialog = Assistant()
    # dialog.add_image(r"imagens\\iel.jpg", 250, 150)
    dialog.add_file_input("add_pdf_file", 'Selecione o arquivo de faturamento', file_type="pdf")
    dialog.add_submit_buttons(["Executar", "Cancelar"])
    user_input = dialog.run_dialog(90000, title="Assistente de Faturamento", height=165, width=400, location="Center")

    return user_input

def clear_directory(directory):
    file = FileSystem()
    try:
        file.empty_directory(directory)
        print("diretório limpo!")
    except:
        print('Não foi possível limpar o diretório')

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
    chrome_prefs = {
    "download.default_directory": f"{Path().home()}\\Desktop\\Palácio\\Códigos\\iel_extratos_de_serviço\\base",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True,
}
    user = "rfcosta"
    password = 'Iel23@'
    url_login = 'https://gestaovagas.iel-ce.org.br/hlogin.aspx'
    id_user = 'vUSU_USUARIO'
    id_password = 'vUSU_SENHA'
    id_login_button = 'BTNLOGIN'
    id_usuarios_test = 'SearchType'

    # TODO: criar verificação de login
    browser.open_chrome_browser(url=url_login, preferences=chrome_prefs)
    browser.input_text_when_element_is_visible(id_user, user)
    browser.input_text_when_element_is_visible(id_password, password)
    browser.click_element_when_visible(id_login_button)

    browser.wait_until_page_contains_element(id_usuarios_test, 90)

# --------------------------------------------------------------------------------------------------------------------------------------------------- #

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
            sleep(1)
            email_atribbute = browser.get_text(id_email_field)
            # email_atribbute = 'rfcosta@sfiec.org.br'
            list_of_emails.append(email_atribbute)

            send_email(company_name, email_atribbute)
        except:
            print(f'Não foi possível buscar o email da empresa {company_name}')

    sleep(3)
    browser.close_browser()

if __name__ == '__main__':
    browser = Selenium()
    browser.auto_close = False

    output_path = r'C:\\Users\\rfcosta\\Desktop\\Palácio\\Códigos\\iel_extratos_de_serviço\\output'

    file_path = start_dialog()
    clear_directory(output_path)
    extract_pages_from_pdf(file_path.get('add_pdf_file')[0])
    get_company_name_and_rename(output_path)
    get_email_info(browser, output_path)



