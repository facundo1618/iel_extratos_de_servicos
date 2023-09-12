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
    browser.select_from_list_by_value(id_grupodefaturamento_field, '45')
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

# mudar nome do parâmetro para file_path
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

def send_email(company_name: str, email_atribbute: str):
    service_note_path = f"{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\output\\{company_name}.pdf"
    gmail_account = "rfcosta@sfiec.org.br"
    gmail_app_password = "japvpatyzugkbruc"
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
    count = 0

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
            # ******TODO: criar verificação para nome da empresa******
            sleep(5) # criando a verificação, tira o sleep()
            company_name_verification = browser.get_text(id_primeirocamporazaosocial_field)
            print(f'{company_name_verification} é igual a {company_name}')
            if company_name_verification == company_name:
                browser.click_element_when_visible(id_primeirocamporazaosocial_field)
                browser.click_element_when_visible(id_contato_button)
                sleep(1.5)
                email_atribbute = browser.get_text(id_email_field)
                # email_atribbute = 'rfcosta@sfiec.org.br'
                list_of_emails.append(email_atribbute)

                send_email(company_name, email_atribbute)
                count +=1
            else:
                print('não foi possível verificar se o nome da empresa')

        except:
            print(f'Não foi possível buscar o email da empresa {company_name}')

    sleep(3)
    browser.close_browser()

    return count

def move_files_to_dedicated_folder(path_of_dedicated_folder):
    def get_month():
        today = date.today()
        today = today.strftime("%m")

        if today == "01":
            today = "Dezembro"

        elif today == "02":
            today = "Janeiro"

        elif today == "03":
            today = "Fevereiro"

        elif today == "04":
            today = "Março"

        elif today == "05":
            today = "Abril"

        elif today == "06":
            today = "Maio"

        elif today == "07":
            today = "Junho"

        elif today == "08":
            today = "Julho"

        elif today == "09":
            today = "Agosto"

        elif today == "10":
            today = "Setembro"

        elif today == "11":
            today = "Outubro"

        elif today == "12":
            today = "Novembro"

        else:
            print("Mês inválido")

        return today
    
    file_ = FileSystem()
    list_of_files = os.listdir(path_of_dedicated_folder)
    for file in list_of_files:
        old_directory = f'{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\output\\{file}'
        new_directory = f'{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\faturamentos passados\\{get_month()}-2023\\{file}'
        file_.move_file(old_directory, new_directory)

if __name__ == "__main__":
    browser = Selenium()
    browser.auto_close = False

    # #pôr imagem do IEL na interface
    user_input = start_dialog() # a data precisa ser 30/09/2023 para pegar as 30 empresas
    if user_input.get("submit") != "Executar":
        sys.exit()
    else:
    #     # clear_directory(f"{Path().home()}\\Desktop\Palácio\Códigos\iel_extratos_de_serviço\output", 
    #     #                 f"{Path().home()}\\Desktop\\Palácio\\Códigos\\iel_extratos_de_serviço\\base")
        clear_directory(f"{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\output", 
                        f"{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\base")
        creat_new_service_note(browser, user_input.get('date_input'))
        file_path = f"{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\base\\arresrelfatextratoestagios.pdf"
        download_service_note(browser, file_path)
        extract_pages_from_pdf(file_path)
        directory = f"{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\output"
        get_company_name_and_rename(directory)
        number_of_emails = get_email_info(browser, directory)
        ctypes.windll.user32.MessageBoxW(0, f"Total de extratos enviados: {number_of_emails}", "Aviso", 1)
        # dedicated_folder = f'{Path().home()}\\Documents\\RPA_IEL_FATURAMENTO\\output'
        # move_files_to_dedicated_folder(dedicated_folder)

    #     # no final, é necessário criar alguma tratativa para a exclusão dos dados. (mover para uma pasta externa ainda na pasta "documentos")

