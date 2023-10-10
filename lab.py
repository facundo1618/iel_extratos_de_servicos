from RPA.Browser.Selenium import Selenium
from time import sleep

url_login = 'https://gestaovagas.iel-ce.org.br/hlogin.aspx'
url_intranet = 'https://gestaovagas.iel-ce.org.br/HComListaPessoas.aspx?hitrpaginainicial.aspx'
id_user = 'vUSU_USUARIO'
id_password = 'vUSU_SENHA'
user = "ielce.rpa"
password = 'Ielce2@'
id_pesquisa_field = 'vPESQUISA'
id_login_button = 'BTNLOGIN'
id_primeirocamporazaosocial_field = 'span_vPES_RAZAO_SOCIAL_0001'
id_contato_button = 'Tab_TAB1Containerpanel2'
id_table = 'W0324GridContainerTbl'
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

browser = Selenium()
browser.auto_close = False
browser.open_chrome_browser(url_login)
browser.input_text_when_element_is_visible(id_user, user)
browser.input_text_when_element_is_visible(id_password, password)
browser.click_element_when_visible(id_login_button)
sleep(9)
browser.go_to(url_intranet)
browser.input_text_when_element_is_visible(id_pesquisa_field, 'FAE SISTEMAS DE MEDICAO S/A')
sleep(3)
browser.click_element_when_visible(id_primeirocamporazaosocial_field)
browser.click_element_when_visible(id_contato_button)
email_cobranca = browser.execute_javascript(js_code)
print(email_cobranca)



