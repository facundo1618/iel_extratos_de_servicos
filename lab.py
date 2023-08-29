from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from pathlib import Path
import os

def get_company_name_and_rename(directory: str):
    pdf = PDF()
    file_ = FileSystem()
    list_of_files = os.listdir(directory)
    for file in list_of_files:
        # try:
        pdf.open_pdf(f"{directory}\\{file}") 
        company_name = pdf.find_text("Razão Social:", 1) 
        company_name = company_name[0].neighbours[0]
        company_name = company_name.replace("/", "-")
        pdf.close_pdf()

        try:
            old_name = f"{directory}\\{file}"
            new_name = f"{directory}\\{company_name}.pdf"

            file_.move_file(old_name, new_name)
        except:
            old_name = f"{directory}\\{file}"
            new_name = f"{directory}\\{company_name}_mesmo_nome.pdf"

            file_.move_file(old_name, new_name)            
        # except:
        #     pass

def extract_pages_from_pdf(directory):
    pdf = PDF()
    file = FileSystem() 
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

file_directory = r"base\\Valorizza - Geral.pdf"
extract_pages_from_pdf(file_directory)
get_company_name_and_rename(r'C:\\Users\\rfcosta\\Desktop\\Palácio\\Códigos\\iel_extratos_de_serviço\\output')
