from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from pathlib import Path
import os

def open_pdf_and_save_pages(directory: str):
    pdf = PDF()
    file = FileSystem() 
    pdf.open_pdf(directory)
    start_page = None
    for page_num in range(1, pdf.get_number_of_pages() + 1):
        page_text = pdf.get_text_from_pdf(directory, page_num)
        if "Extrato de Serviços" in page_text:
            start_page = page_num
        elif "Total Cliente:" in page_text and start_page is not None:
            end_page = page_num

        pdf.extract_pages_from_pdf(directory,f"{Path().home()}\\Desktop\\Palácio\\Códigos\\iel_extratos_de_serviço\\output\\service_note_{start_page} - {end_page}.pdf", page_num)
    
    pdf.close_pdf()

def extract_service_extracts(pdf_path, output_directory):
    pdf = PDF()
    pdf.open_pdf(pdf_path)
    
    start_page = None
    for page_num in range(1, pdf.get_number_of_pages() + 1):
        page_text = pdf.get_text_from_pdf(pdf_path, page_num)
        
        if "Extrato de Serviços" in page_text:
            start_page = page_num
        elif "Total Cliente:" in page_text and start_page is not None:
            end_page = page_num
            extract_pdf = PDF()
            for i in range(start_page, end_page + 1):
                extract_pdf.add_page(pdf.get_page(i))
            
            output_path = f"{output_directory}/extract_{start_page}-{end_page}.pdf"
            extract_pdf.save(output_path)
            extract_pdf.close_pdf()
            
            start_page = None
    
    pdf.close_pdf()

def get_company_name_and_rename(directory: str):
    pdf = PDF()
    list_of_files = os.listdir(directory)
    for file in list_of_files:
        try:
            pdf.open_pdf(f"C:\\Users\\rfcosta\\Desktop\\teste\\{file}") 
            company_name = pdf.find_text("Razão Social:", 1) 
            company_name = company_name[0].neighbours[0]
            print(company_name)
        except:
            pass

if __name__ == "__main__":
    file_directory = r"base\\Valorizza - Geral.pdf"
    open_pdf_and_save_pages(file_directory)
    # get_company_name_and_rename(r"C:\Users\rfcosta\Desktop\Palácio\Códigos\iel_extratos_de_serviço\output")
    # extract_service_extracts(file_directory, r"C:\\Users\\rfcosta\\Desktop\\Palácio\\Códigos\\iel_extratos_de_serviço\\output")


