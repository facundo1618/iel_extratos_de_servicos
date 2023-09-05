# from pathlib import Path
# company_name = 'TESTE 1'
# service_note_path = f"{Path().home()}\\Desktop\\Palácio\\Códigos\\iel_extratos_de_serviço\\output\\{company_name}.pdf"

# print(service_note_path)
# # def creat_new_service_note_pure(browser: Chrome):
# #     user = "rfcosta"
# #     password = 'Iel23@'
# #     url_login = 'https://gestaovagas.iel-ce.org.br/hlogin.aspx'
# #     url_faturamentos = 'https://gestaovagas.iel-ce.org.br/HResListaFaturamentos.aspx?hitrpaginainicial.aspx'
# #     id_user = 'vUSU_USUARIO'
# #     id_password = 'vUSU_SENHA'
# #     id_login_button = 'BTNLOGIN'

# #     browser.get(url_login)

# #     count = 0
# #     while True:
# #         try:
# #             if count == 3:
# #                 print('Quantidade de tentativas excedidas!')
# #                 break
# #             WebDriverWait(browser, 1).until(EC.invisibility_of_element((By.ID, id_user)))
# #             break
            
# #         except:
# #             WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, id_user))).send_keys(Keys.CONTROL + 'a')
# #             WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, id_user))).send_keys(user)
# #             WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, id_password))).send_keys(Keys.CONTROL + 'a')
# #             WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, id_password))).send_keys(password)
# #             WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, id_login_button))).click()
# #             count += 1

# #     browser.get(url_faturamentos)
# #     # clicar no botão "novo"
# #     # selecionar grupo e data
# #     # clicar no botão "imprimir"
# #     # selecionar "extrato de estágio"
# #     # clicar no botão imprimir
# #     # pode abrir outra aba antes do download