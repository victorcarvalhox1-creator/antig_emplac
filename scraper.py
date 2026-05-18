import os
import time
import glob
import traceback
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import pandas as pd

load_dotenv()

LOGIN = os.getenv("LOGIN_FEN")
SENHA = os.getenv("SENHA_FEN")

# --- COMPETÊNCIAS A PROCESSAR ---
# Adicione ou remova competências conforme necessário.
# Deixe apenas as que deseja rodar marcadas (não comentadas).
COMPETENCIAS = [
    {"ano": "2025", "mes": "Janeiro",   "pasta": "012025"},
    {"ano": "2025", "mes": "Fevereiro",  "pasta": "022025"},
    {"ano": "2025", "mes": "Março",      "pasta": "032025"},
    {"ano": "2025", "mes": "Abril",      "pasta": "042025"},
    {"ano": "2025", "mes": "Maio",       "pasta": "052025"},
    {"ano": "2025", "mes": "Junho",      "pasta": "062025"},
    #{"ano": "2025", "mes": "Julho",      "pasta": "072025"},
    # {"ano": "2025", "mes": "Agosto",     "pasta": "082025"},
    # {"ano": "2025", "mes": "Setembro",   "pasta": "092025"},
    # {"ano": "2025", "mes": "Outubro",    "pasta": "102025"},
    # {"ano": "2025", "mes": "Novembro",   "pasta": "112025"},
    # {"ano": "2025", "mes": "Dezembro",   "pasta": "122025"},
]
# --------------------------------

def format_excel_file(file_path):
    print(f"Formatting {file_path} to standard structure...")
    try:
        # header=1 pula a linha de título "Dados de Mercado Personalizados"
        # Colunas do site: [0]CNPJ [1]Razão Social [2]Emplacamento [3]Chassi
        #                  [4]Placa [5]Fabricante [6]Modelo [7]Ano Fabricação
        #                  [8]Estado [9]Município
        df = pd.read_excel(file_path, header=1, dtype=str)
        if len(df.columns) < 10:
            print(f"Skipped formatting {file_path}: only {len(df.columns)} columns.")
            return

        result = pd.DataFrame()
        result['Data']        = df.iloc[:, 2]   # Emplacamento
        result['Chassis']     = df.iloc[:, 3]   # Chassi
        result['Fabricante']  = df.iloc[:, 5]   # Fabricante
        result['Modelo']      = df.iloc[:, 6]   # Modelo
        result['Municipio']   = df.iloc[:, 9]   # Município
        result['UF']          = df.iloc[:, 8]   # Estado
        result['Segmento']    = '1'

        # Sub Segmento = Fabricante + primeiro token do modelo após "/"
        # Ex: "FIAT/DOBLO ADV 1.8 FLEX" -> "FIAT DOBLO"
        #     "I/BYD DOLPHIN MINI GS EV" com Fabricante=BYD -> "BYD BYD"
        def sub_segmento(row):
            fab   = str(row['Fabricante']).strip()
            model = str(row['Modelo']).strip()
            token = model.split('/', 1)[1].split()[0] if '/' in model else model.split()[0]
            return f"{fab} {token}"

        result['Sub Segmento'] = result.apply(sub_segmento, axis=1)
        result['DN ou CNPJ']   = df.iloc[:, 0]  # CNPJ
        result['Ano Fab.']     = df.iloc[:, 7]  # Ano Fabricação
        result['Tipo']         = df.iloc[:, 3].astype(str) + '/' + df.iloc[:, 4].astype(str)  # Chassi/Placa

        result.to_excel(file_path, index=False)
        print(f"Successfully formatted {file_path}")
    except Exception as e:
        print(f"Failed to format {file_path}: {e}")

def combine_spreadsheets(pasta_competencia):
    print("Combining all formatted spreadsheets...")
    try:
        target_dir = f"downloads/{pasta_competencia}"
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
            out_path = f"{target_dir}/consolidado_{pasta_competencia}.xlsx"
            combined_df.to_excel(out_path, index=False)
            print(f"Successfully created {out_path} with {len(all_files)} files combined.")
    except Exception as e:
        print(f"Failed to combine spreadsheets: {e}")

def processar_competencia(page, competencia):
    ano = competencia["ano"]
    mes = competencia["mes"]
    pasta = competencia["pasta"]

    print(f"\n{'='*50}")
    print(f"Iniciando competência: {mes}/{ano} -> pasta: {pasta}")
    print(f"{'='*50}")

    print("Navigating to 'Meu Negócio' page...")
    page.goto(
        "https://www.tela.com.br/inteligencia/Concessionaria/Emplacamento/ShareRanking",
        wait_until="domcontentloaded"
    )
    page.wait_for_selector("select#cmbAno", timeout=30000)

    print(f"Selecting year '{ano}'...")
    page.locator("select#cmbAno").select_option(label=ano)
    page.wait_for_function(
        "document.getElementById('cmbMes').value === '0'",
        timeout=10000
    )

    print(f"Selecting month '{mes}'...")
    page.locator("select#cmbMes").select_option(label=mes)
    page.wait_for_function(
        "document.getElementById('cmbDia').value === '0'",
        timeout=10000
    )

    print("Selecting day 'Todos'...")
    page.locator("select#cmbDia").select_option(value="")

    print("Clicking search...")
    page.click("a#btnPesquisar")

    print("Waiting for brands table to load...")
    try:
        page.wait_for_selector("#divTab1 table tbody tr", timeout=15000)
    except Exception:
        print("Table didn't load in time, but proceeding to see if elements exist.")

    print("Waiting for AJAX data to settle...")
    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass
    time.sleep(8)

    print("Extracting brands list...")
    brand_elements = page.locator("#divTab1 table tbody tr td a").all()
    brand_names = []
    for elem in brand_elements:
        href = elem.get_attribute("href") or ""
        if "chamaDetalhe" in href:
            name = elem.inner_text().strip()
            if name and name.lower() != "total":
                brand_names.append(name)

    print(f"Found {len(brand_names)} brands to process.")

    target_dir = f"downloads/{pasta}"
    os.makedirs(target_dir, exist_ok=True)

    for brand_name in brand_names:
        print(f"Processing brand: {brand_name}")
        try:
            brand_link = page.locator(f"#divTab1 table tbody tr td a[href*=\"chamaDetalhe\"]").filter(has_text=brand_name).first
            brand_link.click()

            total_selector = f"a[href*=\"chamaDescritivo('Total', '{brand_name}')\"]"
            page.wait_for_selector(total_selector, timeout=15000)

            page.locator(total_selector).first.click()
            time.sleep(2)

            page.wait_for_selector("button.buttons-excel", timeout=20000)
            time.sleep(1)

            print(f"Downloading Excel for {brand_name}...")
            with page.expect_download(timeout=30000) as download_info:
                page.locator("button.buttons-excel").first.click()

            download = download_info.value
            file_path = f"{target_dir}/{brand_name}_relatorio.xlsx"
            download.save_as(file_path)
            print(f"Successfully saved {file_path}")
            format_excel_file(file_path)

            page.evaluate("setTimeout(function(){ voltaTabelas(); }, 300)")
            page.wait_for_selector("#divTab1 table tbody tr", state="visible", timeout=15000)
            time.sleep(1)

        except Exception as e:
            print(f"Failed processing {brand_name}: {e}")
            try:
                page.evaluate("setTimeout(function(){ if(typeof voltaTabelas==='function') voltaTabelas(); }, 300)")
                page.wait_for_selector("#divTab1 table tbody tr", state="visible", timeout=10000)
                time.sleep(1)
            except Exception:
                pass

    print(f"Competência {mes}/{ano} concluída. Gerando consolidado...")
    combine_spreadsheets(pasta)


def run():
    if not COMPETENCIAS:
        print("Nenhuma competência marcada para processar. Edite a lista COMPETENCIAS no topo do arquivo.")
        return

    print(f"Loaded credentials for user: {LOGIN}")
    print(f"Competências a processar: {len(COMPETENCIAS)}")
    for c in COMPETENCIAS:
        print(f"  - {c['mes']}/{c['ano']} -> {c['pasta']}")

    playwright = None
    browser = None
    context = None
    page = None
    try:
        playwright = sync_playwright().start()
        print("Launching browser...")
        browser = playwright.chromium.launch(headless=True)
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

        for i, competencia in enumerate(COMPETENCIAS, 1):
            print(f"\nProcessando competência {i} de {len(COMPETENCIAS)}...")
            try:
                processar_competencia(page, competencia)
            except Exception as e:
                print(f"Erro na competência {competencia['mes']}/{competencia['ano']}: {e}")
                traceback.print_exc()
                print("Tentando continuar para a próxima competência...")

        print("\nTodas as competências processadas! Fechando browser.")
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
    