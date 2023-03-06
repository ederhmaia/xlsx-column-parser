from colorama import init, Fore
import os
import pandas as pd
from typing import List, Union

init(autoreset=True)

def banner() -> None:
    print(f'''
    {Fore.CYAN}            ,    _             
    {Fore.CYAN}           /|   | |         {Fore.WHITE}    -.. .. .- -. .-
    {Fore.CYAN}         _/_\_  >_<         {Fore.MAGENTA}    cpf xsx extractor    
    {Fore.CYAN}        .-\-/.   |          {Fore.MAGENTA}     by @ederhmaia
    {Fore.CYAN}       /  | | \_ |          {Fore.WHITE}    -.. .. .- -. .-
    {Fore.CYAN}       \ \| |\__(/                  
    {Fore.CYAN}       /(`---')  |          {Fore.RED}    the responsability
    {Fore.CYAN}      / /     \  |          {Fore.RED}    of using this tool
    {Fore.CYAN}   _.'  \\'-'  /  |          {Fore.RED}    is yours, not mine.
    {Fore.CYAN}   `----'`=-='   '    
    
    ''')

def clear() -> None:
    os.system('cls' if os.name == 'nt' else 'clear')
    banner()

def find_cpf_in_xlsx(file_path) -> List[str]:
    df = pd.read_excel(file_path)
    
    cpf_column: None | str = None
    for col in df.columns:
        if str(col).lower() == "cpf":
            cpf_column = col
            break

    if cpf_column is None:
        print(f"{Fore.RED}✘ {Fore.WHITE}Coluna CPF não encontrada no arquivo {Fore.RED}✘")
        return []

    cpf_list = [str(cpf).strip() for cpf in df[cpf_column]]
    return cpf_list


def extract_cpf_values(output_file_path, cpf_list) -> None:
    clear()
    with open(output_file_path, "w") as f:
        for cpf in cpf_list:
            print(f"{Fore.GREEN}↪ {Fore.WHITE}CPF {Fore.CYAN}{cpf}{Fore.WHITE} extraído {Fore.GREEN}")
            f.write(cpf + "\n")
    
    clear()
    print(f"{Fore.GREEN}✔ {Fore.WHITE}Arquivo de saída {Fore.CYAN}{output_file_path}{Fore.WHITE} gerado com sucesso {Fore.GREEN}✔")


def main() -> None:
    xlsx_files: List[str] = [f for f in os.listdir('./') if f.endswith(".xlsx")]
    
    if not xlsx_files:
        print(f"{Fore.RED}✘ {Fore.WHITE}Nenhum arquivo XLSX encontrado no diretório atual {Fore.RED}✘")
        return
    
    print(f"{Fore.GREEN}↓ {Fore.WHITE}Arquivos XLSX disponíveis {Fore.GREEN}↓")
    
    for i, xlsx_file in enumerate(xlsx_files):
        print(f"{Fore.CYAN}{i+1}.{Fore.WHITE} {xlsx_file}")

    xlsx_index: None | int = None

    print()
    print(f"{Fore.GREEN}↓ {Fore.WHITE}Selecione o arquivo XLSX {Fore.GREEN}↓")

    while xlsx_index is None:
        try:
            xlsx_index: int = int(input(f"{Fore.GREEN}>{Fore.WHITE} "))
            
            if not 1 <= xlsx_index <= len(xlsx_files):
                raise ValueError()

        except ValueError:
            print(f"{Fore.RED}✘ {Fore.WHITE}Opção inválida {Fore.RED}✘")
            print()
            xlsx_index = None
    
    xlsx_file_path: str = os.path.join('./', xlsx_files[xlsx_index-1])

    clear()

    print(f"{Fore.GREEN}↓ {Fore.WHITE}Nome do arquivo de saída {Fore.GREEN}↓")
    while True:
        try:
            output_file_path: str = input(f"{Fore.GREEN}>{Fore.WHITE} ")

            if not output_file_path.endswith(".txt"):
                output_file_path += ".txt"

            if os.path.exists(output_file_path):
                raise ValueError()

            break

        except ValueError:
            print(f"{Fore.RED}✘ {Fore.WHITE}Arquivo já existe {Fore.RED}✘")
            print()
            continue

    cpf_lists: List[str] = find_cpf_in_xlsx(xlsx_file_path)
    print()
    extract_cpf_values(output_file_path, cpf_lists)
    
    
if __name__ == "__main__":
    clear()
    main()
