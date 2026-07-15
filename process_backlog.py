import os
import re
import time
import csv
import requests
from bs4 import BeautifulSoup

# Configurações
USERNAME = "gabriel.sojo"
PASSWORD = "Dorinh@001"
CSV_PATH = "backlog.csv"

def clean(val):
    if val is None:
        return ""
    s = str(val).strip()
    s = s.replace('\ufeff', '')
    s = s.replace('\xef\xbb\xbf', '')
    s = s.replace('"', '').replace("'", '')
    return s.strip()

def main():
    global CSV_PATH
    # Check if backlog.csv exists
    if not os.path.exists(CSV_PATH):
        print(f"Aviso: Arquivo '{CSV_PATH}' padrão não encontrado na pasta.")
        # Try to search for other CSV files
        csv_files = [f for f in os.listdir('.') if f.endswith('.csv') and f != 'backlog.csv']
        if csv_files:
            print(f"Arquivos CSV encontrados: {csv_files}")
            default_csv = csv_files[0]
            choice = input(f"Deseja usar '{default_csv}'? (S/N): ").strip().lower()
            if choice == 's':
                CSV_PATH = default_csv
            else:
                CSV_PATH = input("Digite o nome do arquivo CSV: ").strip()
        else:
            CSV_PATH = input("Digite o nome do arquivo CSV: ").strip()
            
        if not os.path.exists(CSV_PATH):
            print(f"Erro: Arquivo '{CSV_PATH}' não encontrado. Encerrando.")
            return

    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    })

    # 1. Obter página de login para capturar o token CSRF e os nomes de campo dinâmicos
    print("Passo 1: Acessando a página de login para obter tokens...")
    login_url = "https://lcdesk.lowcost.com.br/index.php"
    try:
        response = session.get(login_url, timeout=15)
    except Exception as e:
        print(f"Erro ao conectar ao GLPI: {e}")
        return
    
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Encontrar token CSRF
    csrf_input = soup.find('input', {'name': '_glpi_csrf_token'})
    if not csrf_input:
        print("Erro: Não foi possível localizar o token CSRF.")
        return
    csrf_token = csrf_input['value']

    # Encontrar os campos dinâmicos de usuário e senha pelos IDs correspondentes
    user_input = soup.find('input', {'id': 'login_name'})
    pass_input = soup.find('input', {'id': 'login_password'})
    remember_input = soup.find('input', {'id': 'login_remember'})

    if not user_input or not pass_input:
        print("Erro: Não foi possível identificar os nomes de campo dinâmicos para login.")
        return
        
    user_field_name = user_input['name']
    pass_field_name = pass_input['name']
    remember_field_name = remember_input['name'] if remember_input else None

    # 2. Realizar o Login
    print("Passo 2: Realizando a autenticação...")
    login_data = {
        "_glpi_csrf_token": csrf_token,
        "noAUTO": "0",
        "redirect": "",
        user_field_name: USERNAME,
        pass_field_name: PASSWORD,
        "submit": "submit"
    }
    if remember_field_name:
        login_data[remember_field_name] = "1"

    submit_url = "https://lcdesk.lowcost.com.br/front/login.php"
    login_response = session.post(submit_url, data=login_data)

    # Validar se o login foi bem-sucedido
    test_url = "https://lcdesk.lowcost.com.br/ajax/common.tabs.php?_glpi_tab=Ticket%24main&_target=%2Ffront%2Fticket.form.php&_itemtype=Ticket&id=2606240015"
    test_response = session.get(test_url)
    if "actor-type" not in test_response.text:
        print("Erro: Falha na autenticação. Verifique suas credenciais de usuário/senha.")
        return
    print("Autenticação efetuada com sucesso!")

    # 3. Processar CSV
    print(f"Passo 3: Lendo o arquivo CSV '{CSV_PATH}'...")
    
    # Read rows and detect encoding
    encodings = ['utf-8-sig', 'iso-8859-1', 'utf-8']
    rows = []
    headers = []
    success_read = False
    
    for encoding in encodings:
        try:
            with open(CSV_PATH, mode='r', encoding=encoding, errors='replace') as f:
                reader = csv.reader(f, delimiter=';')
                headers = next(reader)
                rows = list(reader)
                success_read = True
                print(f"Lido com sucesso usando codificação: {encoding}")
                break
        except Exception:
            continue
            
    if not success_read:
        print("Erro: Não foi possível ler o arquivo CSV. Verifique a codificação.")
        return

    # Encontrar coluna de ID e de atribuição técnica
    id_col = -1
    tech_col = -1
    
    for i, h in enumerate(headers):
        h_clean = clean(h).lower()
        if h_clean == 'id':
            id_col = i
        elif 'atribu' in h_clean or 'atend' in h_clean:
            tech_col = i
            
    if id_col == -1:
        print("Erro: Coluna 'ID' não localizada no cabeçalho do CSV.")
        return
        
    if tech_col == -1:
        # Se não encontrar a coluna do técnico, vamos criá-la
        headers.append("Atribuído - Técnico")
        tech_col = len(headers) - 1
        for r in rows:
            r.append("")
        print("Coluna de técnico não encontrada. Criando nova coluna 'Atribuído - Técnico'.")
    else:
        print(f"Coluna de técnico identificada: '{headers[tech_col]}' (índice {tech_col})")

    # Percorre as linhas
    updated_count = 0
    total_processed = 0
    
    print("Iniciando a busca de atribuições em branco no GLPI...")
    
    for idx, r in enumerate(rows):
        # Asegurar que o tamanho da linha corresponda ao cabeçalho (caso tenha criado coluna)
        while len(r) < len(headers):
            r.append("")
            
        ticket_id = clean(r[id_col])
        if not ticket_id.isdigit():
            continue
            
        tech_val = clean(r[tech_col])
        if tech_val != "":
            # Já tem técnico
            continue
            
        print(f"Chamado {ticket_id} (Linha {idx+2}): Buscando no GLPI...")
        total_processed += 1
        
        ajax_url = f"https://lcdesk.lowcost.com.br/ajax/common.tabs.php?_glpi_tab=Ticket%24main&_target=%2Ffront%2Fticket.form.php&_itemtype=Ticket&id={ticket_id}"
        
        try:
            res = session.get(ajax_url, timeout=10)
            ticket_soup = BeautifulSoup(res.text, 'html.parser')
            
            select_elem = ticket_soup.find('select', {'data-actor-type': 'assign'})
            assignments = []
            if select_elem:
                selected_options = select_elem.find_all('option', {'selected': 'true'})
                for opt in selected_options:
                    assignments.append(opt.text.strip())
            
            assignment_text = ", ".join(assignments)
            if assignment_text != "":
                r[tech_col] = assignment_text
                print(f"-> Atribuído a: '{assignment_text}'")
                updated_count += 1
            else:
                print("-> Nenhuma atribuição encontrada no GLPI.")
                
        except Exception as ex:
            print(f"-> Erro ao buscar chamado {ticket_id}: {ex}")
            r[tech_col] = "ERROR"
            
        time.sleep(0.150)

    print(f"\nSincronização concluída! Total de chamados com técnico atualizado: {updated_count} de {total_processed} pesquisados.")

    # 4. Gravar arquivo CSV atualizado
    print("Passo 4: Gravando arquivo CSV atualizado...")
    try:
        # Save back to CSV using the same coding (utf-8-sig keeps compatibility with Excel)
        with open(CSV_PATH, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(headers)
            writer.writerows(rows)
        print(f"Arquivo '{CSV_PATH}' salvo com sucesso!")
    except Exception as e:
        print(f"Erro ao gravar o arquivo CSV: {e}")

if __name__ == "__main__":
    main()
