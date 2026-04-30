import os
import time
import glob
import calendar
import traceback
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import pandas as pd

load_dotenv()

LOGIN = os.getenv("LOGIN_FEN")
SENHA = os.getenv("SENHA_FEN")

# --- CONFIGURAÇÕES ---
ANO_PESQUISA = "2026"       # Ano que será selecionado no site (ex: 2025, 2026)
MES_PESQUISA = "Abril"       # Mês que será selecionado no site (ex: Janeiro, Fevereiro, Março)
PASTA_COMPETENCIA = "042026" # Nome da pasta onde os arquivos serão salvos
# ---------------------

def format_excel_file(file_path):
    print(f"Formatting {file_path} to standard structure...")
    try:
        # Read all as string to preserve precision for CNPJ
        df = pd.read_excel(file_path, header=1, dtype=str)
        if len(df.columns) >= 10:
            new_cols = [
                df.columns[0], # A -> A (Data)
                df.columns[5], # F -> B (Chassis)
                df.columns[7], # H -> C (Fabricante)
                df.columns[1], # B -> D (Modelo)
                df.columns[9], # J -> E (Municipio)
                df.columns[8], # I -> F (UF)
                df.columns[3], # D -> CNPJ
                df.columns[6]  # G -> Placa
            ]
            df_formatted = df[new_cols].copy()
            df_formatted.columns = ['Data', 'Chassis', 'Fabricante', 'Modelo', 'Municipio', 'UF', 'CNPJ', 'Placa']
            
            # Insert two blank columns at G and H (indices 6 and 7) with unique space names
            df_formatted.insert(6, ' ', 1)
            df_formatted.insert(7, '  ', '')
            
            df_formatted.to_excel(file_path, index=False)
            print(f"Successfully formatted {file_path}")
        else:
            print(f"Skipped formatting {file_path}: Does not match expected column structure.")
    except Exception as e:
        print(f"Failed to format {file_path}: {e}")

def combine_spreadsheets(outros_qtd=0):
    print("Combining all formatted spreadsheets...")
    try:
        target_dir = f"downloads/{PASTA_COMPETENCIA}"
        all_files = glob.glob(f"{target_dir}/*_relatorio.xlsx")
        if not all_files:
            print("No files found to combine.")
            return

        df_list = []
        for file in all_files:
            df = pd.read_excel(file, dtype=str)
            df_list.append(df)

        if df_list:
            combined_df = pd.concat(df_list, ignore_index=True)
            
            if outros_qtd > 0:
                try:
                    month = int(PASTA_COMPETENCIA[:2])
                    year = int(PASTA_COMPETENCIA[2:])
                    last_day = calendar.monthrange(year, month)[1]
                    data_str = f"{last_day:02d}/{month:02d}/{year}"
                except Exception:
                    data_str = "31/12/2099"
                
                outros_row = pd.DataFrame([{
                    'Data': data_str,
                    'Chassis': 'outros',
                    'Fabricante': 'outros',
                    'Modelo': 'outros',
                    'Municipio': 'outros',
                    'UF': 'outros',
                    ' ': outros_qtd,
                    '  ': 'outros',
                    'CNPJ': 'outros',
                    'Placa': 'outros'
                }])
                combined_df = pd.concat([combined_df, outros_row], ignore_index=True)
                
            out_path = f"{target_dir}/consolidado.xlsx"
            combined_df.to_excel(out_path, index=False)
            print(f"Successfully created {out_path} with {len(all_files)} files combined.")
    except Exception as e:
        print(f"Failed to combine spreadsheets: {e}")

def run():
    print(f"Loaded credentials for user: {LOGIN}")
    playwright = None
    browser = None
    context = None
    page = None
    try:
        playwright = sync_playwright().start()
        print("Launching browser...")
        # headless=False to make execution visible
        browser = playwright.chromium.launch(headless=False, slow_mo=500)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        print("Navigating to login page...")
        page.goto("https://www.tela.com.br/inteligencia/Home/Index?ReturnUrl=%2finteligencia%2fConcessionaria")

        print("Applying credentials...")
        page.fill("input#Usuario", str(LOGIN))
        page.fill("input#Senha", str(SENHA))

        print("Submitting login form...")
        page.click("button.submitButton")

        print("Waiting for authenticated session...")
        try:
            page.wait_for_url("**/Concessionaria**", timeout=30000)
        except Exception:
            page.wait_for_load_state("domcontentloaded")
            page.wait_for_timeout(3000)

        print("Navigating to 'Meu Negócio' page...")
        page.goto(
            "https://www.tela.com.br/inteligencia/Concessionaria/Emplacamento/MeuNegocio",
            wait_until="domcontentloaded"
        )
        page.wait_for_selector("span#select2-cmbAno-container", timeout=30000)
        page.wait_for_selector("span#select2-cmbMes-container", timeout=30000)

        # Pesquisa pelo ano
        print(f"Selecting month '{ANO_PESQUISA}'...")
        selected_year = page.locator("span#select2-cmbAno-container").inner_text().strip()
        if selected_year == ANO_PESQUISA:
            print("Target month already selected.")
        else:
            # Clicking the span that opens the select2 dropdown
            page.click("span#select2-cmbAno-container")
            
            # Playwright exact text locator for the option list
            try:
                # We wait for the dropdown to appear and select the specific month
                list_item = page.locator("li", has_text=ANO_PESQUISA).first
                list_item.click(timeout=5000)
            except Exception as e:
                print(f"Could not select {ANO_PESQUISA} using select2 container: {e}")
                # Fallback: try to select directly if it's a native select
                try:
                    page.locator("select#cmbAno").select_option(label=ANO_PESQUISA, force=True, timeout=5000)
                except Exception as select_e:
                    print(f"Fallback native select failed: {select_e}")

        # Pesquisa pelo mês
        print(f"Selecting month '{MES_PESQUISA}'...")
        selected_month = page.locator("span#select2-cmbMes-container").inner_text().strip()
        if selected_month == MES_PESQUISA:
            print("Target month already selected.")
        else:
            # Clicking the span that opens the select2 dropdown
            page.click("span#select2-cmbMes-container")
            
            # Playwright exact text locator for the option list
            try:
                # We wait for the dropdown to appear and select the specific month
                list_item = page.locator("li", has_text=MES_PESQUISA).first
                list_item.click(timeout=5000)
            except Exception as e:
                print(f"Could not select {MES_PESQUISA} using select2 container: {e}")
                # Fallback: try to select directly if it's a native select
                try:
                    page.locator("select#cmbMes").select_option(label=MES_PESQUISA, force=True, timeout=5000)
                except Exception as select_e:
                    print(f"Fallback native select failed: {select_e}")

        print("Clicking search...")
        page.click("a#btnPesquisar")
        
        print("Waiting for table to load...")
        # Waiting for the specific table class or id
        # <table class="table table-condensed table-bordered" ...
        try:
            page.wait_for_selector("table.table-condensed", timeout=15000)
        except Exception:
            print("Table didn't load in time, but proceeding to see if elements exist.")

        # Wait for AJAX data to settle
        print("Waiting for AJAX data to settle...")
        try:
            page.wait_for_load_state("networkidle", timeout=5000)
        except Exception:
            pass
        time.sleep(8)

        print("Extracting brands list...")
        # Elements like: <a href="#" id="marcaarea" class="clickMarcaArea" data-id="JEEP">480</a>
        brand_links = page.locator("a.clickMarcaArea").all()
        
        print(f"Found {len(brand_links)} brand targets to process.")
        
        target_dir = f"downloads/{PASTA_COMPETENCIA}"
        os.makedirs(target_dir, exist_ok=True)

        print("Extracting 'Outros' quantity...")
        outros_qtd = 0
        try:
            outros_row = page.locator("xpath=//tr[td[normalize-space(text())='Outros']]").first
            outros_qtd_text = outros_row.locator("td").nth(2).inner_text()
            outros_qtd = int(outros_qtd_text.strip())
            print(f"'Outros' quantity found: {outros_qtd}")
        except Exception as e:
            print(f"Could not extract 'Outros' quantity: {e}")

        for i in range(len(brand_links)):
            # Re-read elements directly to avoid stale element reference if the DOM refreshes
            link = page.locator("a.clickMarcaArea").nth(i)
            brand_name = link.get_attribute("data-id")
            if not brand_name:
                brand_name = f"marca_{i}"
                
            print(f"Processing brand: {brand_name}")
            
            # Click the brand to trigger whatever action (e.g., opening a modal)
            link.click()
            time.sleep(2)

            # Look for Excel export button and wait for download
            # <span>Excel</span>
            print(f"Downloading Excel spreadsheet for {brand_name}...")
            try:
                with page.expect_download(timeout=10000) as download_info:
                    # Depending on how it's structured, we might search for text 'Excel'
                    page.locator("span", has_text="Excel").first.click()
                
                download = download_info.value
                file_path = f"{target_dir}/{brand_name}_relatorio.xlsx"
                download.save_as(file_path)
                print(f"Successfully saved {file_path}")
                
                # Format the excel file automatically after saving
                format_excel_file(file_path)
            except Exception as e:
                print(f"Failed to download Excel for {brand_name}: {e}")
            
            # In some systems we need to explicitly close the modal or wait before the next click
            time.sleep(1)

        print("All downloads completed! Closing browser.")
        combine_spreadsheets(outros_qtd)
    except Exception:
        print("Scraper failed. Full traceback:")
        traceback.print_exc()
        raise
    finally:
        if page is not None:
            try:
                page.close()
            except Exception as e:
                print(f"Failed to close page cleanly: {e}")
        if context is not None:
            try:
                context.close()
            except Exception as e:
                print(f"Failed to close context cleanly: {e}")
        if browser is not None:
            try:
                browser.close()
            except Exception as e:
                print(f"Failed to close browser cleanly: {e}")
        if playwright is not None:
            try:
                playwright.stop()
            except Exception as e:
                print(f"Failed to stop Playwright cleanly: {e}")

if __name__ == "__main__":
    run()






