
"""CÓDIGO DESENVOLVIDO POR FLAVIO CORDEIRO PARA PROCESSAR OS ARQUIVOS DE CTE
   ORIGINÁRIOS DO MERCADO LIVRE E TRANSFORMÁ-LOS EM UMA PLANILHA PARA MELHOR VISUALIZAÇÃO.
   SERÃO GERADOS DOIS ARQUIVOS. UM COM TODOS OS CAMPOS DO XML ORGANIZADOS EM COLUNAS
   E UM SEGUNDO COM OS POSSÍVEIS ARQUIVOS NÃO PROCESSADOS.
   ÚLTIMO REVIEW DO CÓDIGO EM 21/07/2023."""

import os
import tkinter as tk
from tkinter import filedialog
import xml.etree.ElementTree as ET
import openpyxl as px
from datetime import datetime

# Função para selecionar a pasta com os arquivos XML
def select_xml_folder():
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Selecione a pasta com os arquivos XML")
    return folder_path

# Função para selecionar a pasta onde salvar o arquivo XLSX
def select_xlsx_folder():
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Selecione a pasta onde salvar o arquivo XLSX")
    return folder_path

# Função para verificar se o XML tem a estrutura esperada
def is_valid_xml(xml_filepath):
    try:
        tree = ET.parse(xml_filepath)
        root_element = tree.getroot()
        ns = {"cte": "http://www.portalfiscal.inf.br/cte"}
        root_element.find("cte:CTe", ns).find("cte:infCte", ns).find("cte:emit", ns)
        return True
    except:
        return False

# Função para extrair todos os campos do XML
def extract_xml_data(xml_folder):
    data = []
    ignored_files = []

    for root, _, files in os.walk(xml_folder):
        for filename in files:
            if filename.endswith(".xml"):
                xml_filepath = os.path.join(root, filename)

                if is_valid_xml(xml_filepath):
                    tree = ET.parse(xml_filepath)
                    root_element = tree.getroot()
                    ns = {"cte": "http://www.portalfiscal.inf.br/cte"}
                    cte_element = root_element.find("cte:CTe", ns)
                    infcte_element = cte_element.find("cte:infCte", ns)
                    emit_element = infcte_element.find("cte:emit", ns)
                    rem_element = infcte_element.find("cte:rem", ns)
                    dest_element = infcte_element.find("cte:dest", ns)
                    vprest_element = infcte_element.find("cte:vPrest", ns)

                    # Extrair os campos do XML (adapte com os campos que você deseja extrair)
                    cUF = infcte_element.findtext("cte:ide/cte:cUF", namespaces=ns)
                    cCT = infcte_element.findtext("cte:ide/cte:cCT", namespaces=ns)
                    CFOP = infcte_element.findtext("cte:ide/cte:CFOP", namespaces=ns)
                    natOp = infcte_element.findtext("cte:ide/cte:natOp", namespaces=ns)
                    mod = infcte_element.findtext("cte:ide/cte:mod", namespaces=ns)
                    serie = infcte_element.findtext("cte:ide/cte:serie", namespaces=ns)
                    nCT = infcte_element.findtext("cte:ide/cte:nCT", namespaces=ns)
                    dhEmi = infcte_element.findtext("cte:ide/cte:dhEmi", namespaces=ns)
                    tpImp = infcte_element.findtext("cte:ide/cte:tpImp", namespaces=ns)
                    tpEmis = infcte_element.findtext("cte:ide/cte:tpEmis", namespaces=ns)
                    cDV = infcte_element.findtext("cte:ide/cte:cDV", namespaces=ns)
                    tpAmb = infcte_element.findtext("cte:ide/cte:tpAmb", namespaces=ns)
                    tpCTe = infcte_element.findtext("cte:ide/cte:tpCTe", namespaces=ns)
                    procEmi = infcte_element.findtext("cte:ide/cte:procEmi", namespaces=ns)
                    verProc = infcte_element.findtext("cte:ide/cte:verProc", namespaces=ns)
                    cMunEnv = infcte_element.findtext("cte:ide/cte:cMunEnv", namespaces=ns)
                    xMunEnv = infcte_element.findtext("cte:ide/cte:xMunEnv", namespaces=ns)
                    UFEnv = infcte_element.findtext("cte:ide/cte:UFEnv", namespaces=ns)
                    modal = infcte_element.findtext("cte:ide/cte:modal", namespaces=ns)
                    tpServ = infcte_element.findtext("cte:ide/cte:tpServ", namespaces=ns)
                    cMunIni = infcte_element.findtext("cte:ide/cte:cMunIni", namespaces=ns)
                    xMunIni = infcte_element.findtext("cte:ide/cte:xMunIni", namespaces=ns)
                    UFIni = infcte_element.findtext("cte:ide/cte:UFIni", namespaces=ns)
                    cMunFim = infcte_element.findtext("cte:ide/cte:cMunFim", namespaces=ns)
                    xMunFim = infcte_element.findtext("cte:ide/cte:xMunFim", namespaces=ns)
                    UFFim = infcte_element.findtext("cte:ide/cte:UFFim", namespaces=ns)
                    retira = infcte_element.findtext("cte:ide/cte:retira", namespaces=ns)
                    indIEToma = infcte_element.findtext("cte:ide/cte:indIEToma", namespaces=ns)
                    CNPJToma = infcte_element.findtext("cte:ide/cte:toma4/cte:CNPJ", namespaces=ns)
                    IEToma = infcte_element.findtext("cte:ide/cte:toma4/cte:IE", namespaces=ns)
                    xNomeToma = infcte_element.findtext("cte:ide/cte:toma4/cte:xNome", namespaces=ns)
                    xLgrToma = infcte_element.findtext("cte:ide/cte:toma4/cte:enderToma/cte:xLgr", namespaces=ns)
                    nroToma = infcte_element.findtext("cte:ide/cte:toma4/cte:enderToma/cte:nro", namespaces=ns)
                    xBairroToma = infcte_element.findtext("cte:ide/cte:toma4/cte:enderToma/cte:xBairro", namespaces=ns)
                    cMunToma = infcte_element.findtext("cte:ide/cte:toma4/cte:enderToma/cte:cMun", namespaces=ns)
                    xMunToma = infcte_element.findtext("cte:ide/cte:toma4/cte:enderToma/cte:xMun", namespaces=ns)
                    CEPToma = infcte_element.findtext("cte:ide/cte:toma4/cte:enderToma/cte:CEP", namespaces=ns)
                    UFToma = infcte_element.findtext("cte:ide/cte:toma4/cte:enderToma/cte:UF", namespaces=ns)
                    xObs = infcte_element.findtext("cte:compl/cte:xObs", namespaces=ns)
                    CNPJEmit = infcte_element.findtext("cte:emit/cte:CNPJ", namespaces=ns)
                    IEEmit = infcte_element.findtext("cte:emit/cte:IE", namespaces=ns)
                    xNomeEmit = infcte_element.findtext("cte:emit/cte:xNome", namespaces=ns)
                    xLgrEmit = infcte_element.findtext("cte:emit/cte:enderEmit/cte:xLgr", namespaces=ns)
                    nroEmit = infcte_element.findtext("cte:emit/cte:enderEmit/cte:nro", namespaces=ns)
                    xBairroEmit = infcte_element.findtext("cte:emit/cte:enderEmit/cte:xBairro", namespaces=ns)
                    cMunEmit = infcte_element.findtext("cte:emit/cte:enderEmit/cte:cMun", namespaces=ns)
                    xMunEmit = infcte_element.findtext("cte:emit/cte:enderEmit/cte:xMun", namespaces=ns)
                    CEPEmit = infcte_element.findtext("cte:emit/cte:enderEmit/cte:CEP", namespaces=ns)
                    UFEmit = infcte_element.findtext("cte:emit/cte:enderEmit/cte:UF", namespaces=ns)
                    CNPJRem = infcte_element.findtext("cte:rem/cte:CNPJ", namespaces=ns)
                    IERem = infcte_element.findtext("cte:rem/cte:IE", namespaces=ns)
                    xNomeRem = infcte_element.findtext("cte:rem/cte:xNome", namespaces=ns)
                    xLgrRem = infcte_element.findtext("cte:rem/cte:enderReme/cte:xLgr", namespaces=ns)
                    nroRem = infcte_element.findtext("cte:rem/cte:enderReme/cte:nro", namespaces=ns)
                    xBairroRem = infcte_element.findtext("cte:rem/cte:enderReme/cte:xBairro", namespaces=ns)
                    cMunRem = infcte_element.findtext("cte:rem/cte:enderReme/cte:cMun", namespaces=ns)
                    xMunRem = infcte_element.findtext("cte:rem/cte:enderReme/cte:xMun", namespaces=ns)
                    CEPRem = infcte_element.findtext("cte:rem/cte:enderReme/cte:CEP", namespaces=ns)
                    UFRem = infcte_element.findtext("cte:rem/cte:enderReme/cte:UF", namespaces=ns)
                    CPFDes = infcte_element.findtext("cte:dest/cte:CPF", namespaces=ns)
                    xNomeDes = infcte_element.findtext("cte:dest/cte:xNome", namespaces=ns)
                    xLgrDes = infcte_element.findtext("cte:dest/cte:enderDest/cte:xLgr", namespaces=ns)
                    nroDes = infcte_element.findtext("cte:dest/cte:enderDest/cte:nro", namespaces=ns)
                    xCplDes = infcte_element.findtext("cte:dest/cte:enderDest/cte:xCpl", namespaces=ns)
                    xBairroDes = infcte_element.findtext("cte:dest/cte:enderDest/cte:xBairro", namespaces=ns)
                    cMunDes = infcte_element.findtext("cte:dest/cte:enderDest/cte:cMun", namespaces=ns)
                    xMunDes = infcte_element.findtext("cte:dest/cte:enderDest/cte:xMun", namespaces=ns)
                    CEPDes = infcte_element.findtext("cte:dest/cte:enderDest/cte:CEP", namespaces=ns)
                    UFDes = infcte_element.findtext("cte:dest/cte:enderDest/cte:UF", namespaces=ns)
                    vTPrest = float(infcte_element.findtext("cte:vPrest/cte:vTPrest", namespaces=ns))
                    vRec = float(infcte_element.findtext("cte:vPrest/cte:vRec", namespaces=ns))
                    CSTICMS = infcte_element.findtext("cte:imp/cte:ICMS/cte:ICMS00/cte:CST", namespaces=ns)
                    vBCICMS = float(infcte_element.findtext("cte:imp/cte:ICMS/cte:ICMS00/cte:vBC", namespaces=ns))
                    pICMS = float(infcte_element.findtext("cte:imp/cte:ICMS/cte:ICMS00/cte:pICMS", namespaces=ns))
                    vICMS = float(infcte_element.findtext("cte:imp/cte:ICMS/cte:ICMS00/cte:vICMS", namespaces=ns))
                    vTotTrib = float(infcte_element.findtext("cte:imp/cte:vTotTrib", namespaces=ns))
                    infAdFisco = infcte_element.findtext("cte:imp/cte:infAdFisco", namespaces=ns)
                    chaveNFe = infcte_element.findtext("cte:infCTeNorm/cte:infDoc/cte:infNFe/cte:chave", namespaces=ns)
                    dPrevNFe = infcte_element.findtext("cte:infCTeNorm/cte:infDoc/cte:infNFe/cte:dPrev", namespaces=ns)
                    RNTRC = infcte_element.findtext("cte:infCTeNorm/cte:infModal/cte:rodo/cte:RNTRC", namespaces=ns)


                    # INSERE OS DADOS EXTRAÍDOS NA LISTA
                    data.append([
                        cUF,
                        cCT, 
                        CFOP, 
                        natOp, 
                        mod, 
                        serie, 
                        nCT, 
                        dhEmi, 
                        tpImp, 
                        tpEmis,
                        cDV, 
                        tpAmb, 
                        tpCTe, 
                        procEmi, 
                        verProc, 
                        cMunEnv, 
                        xMunEnv, 
                        UFEnv,
                        modal, 
                        tpServ, 
                        cMunIni, 
                        xMunIni, 
                        UFIni, 
                        cMunFim, 
                        xMunFim, 
                        UFFim,
                        retira, 
                        indIEToma, 
                        CNPJToma, 
                        IEToma, 
                        xNomeToma, 
                        xLgrToma, 
                        nroToma,
                        xBairroToma, 
                        cMunToma, 
                        xMunToma, 
                        CEPToma, 
                        UFToma, 
                        xObs,
                        CNPJEmit, 
                        IEEmit, 
                        xNomeEmit, 
                        xLgrEmit, 
                        nroEmit, 
                        xBairroEmit,
                        cMunEmit, 
                        xMunEmit, 
                        CEPEmit, 
                        UFEmit, 
                        CNPJRem, 
                        IERem, 
                        xNomeRem,
                        xLgrRem, 
                        nroRem, 
                        xBairroRem, 
                        cMunRem, 
                        xMunRem, 
                        CEPRem, 
                        UFRem,
                        CPFDes, 
                        xNomeDes, 
                        xLgrDes, 
                        nroDes, 
                        xCplDes, 
                        xBairroDes,
                        cMunDes, 
                        xMunDes, 
                        CEPDes, 
                        UFDes, 
                        vTPrest, 
                        vRec, 
                        CSTICMS, 
                        vBCICMS, 
                        pICMS, 
                        vICMS,
                        vTotTrib, 
                        infAdFisco, 
                        chaveNFe, 
                        dPrevNFe, 
                        RNTRC
                    ])

                else:
                    ignored_files.append(filename)

    return data, ignored_files

# Função para salvar os dados em um arquivo XLSX
def save_to_xlsx(data, xlsx_folder):
    if not xlsx_folder.endswith(os.path.sep):
        xlsx_folder += os.path.sep

    now = datetime.now()
    formatted_date = now.strftime("%Y-%m-%d_%H-%M")
    output_filename = f"CTEs_{formatted_date}.xlsx"

    workbook = px.Workbook()
    sheet = workbook.active
    sheet.title = "CTEs Processados"

    for row in data:
        sheet.append(row)

    output_filepath = os.path.join(xlsx_folder, output_filename)
    workbook.save(output_filepath)
    print(f"Arquivo XLSX criado com sucesso em: {output_filepath}")

def main():
    xml_folder = select_xml_folder()

    data, ignored_files = extract_xml_data(xml_folder)

    xlsx_folder = select_xlsx_folder()

    # Títulos das colunas (adaptar de acordo com os campos extraídos do XML)
    headers = [
        "cUF", 
        "cCT", 
        "CFOP", 
        "natOp", 
        "mod", 
        "serie", 
        "nCT", 
        "dhEmi", 
        "tpImp", 
        "tpEmis",
        "cDV", 
        "tpAmb", 
        "tpCTe", 
        "procEmi", 
        "verProc", 
        "cMunEnv", 
        "xMunEnv", 
        "UFEnv",
        "modal", 
        "tpServ", 
        "cMunIni", 
        "xMunIni", 
        "UFIni", 
        "cMunFim", 
        "xMunFim", 
        "UFFim",
        "retira", 
        "indIEToma", 
        "CNPJToma", 
        "IEToma", 
        "xNomeToma", 
        "xLgrToma", 
        "nroToma",
        "xBairroToma", 
        "cMunToma", 
        "xMunToma", 
        "CEPToma", 
        "UFToma", 
        "xObs",
        "CNPJEmit", 
        "IEEmit", 
        "xNomeEmit", 
        "xLgrEmit", 
        "nroEmit", 
        "xBairroEmit",
        "cMunEmit", 
        "xMunEmit", 
        "CEPEmit", 
        "UFEmit", 
        "CNPJRem", 
        "IERem", 
        "xNomeRem",
        "xLgrRem", 
        "nroRem", 
        "xBairroRem", 
        "cMunRem", 
        "xMunRem", 
        "CEPRem", 
        "UFRem",
        "CPFDes", 
        "xNomeDes", 
        "xLgrDes", 
        "nroDes", 
        "xCplDes", 
        "xBairroDes",
        "cMunDes", 
        "xMunDes", 
        "CEPDes", 
        "UFDes", 
        "vTPrestacao", 
        "vRecebido",
        "CSTICMS", 
        "vBCICMS", 
        "pICMS", 
        "vICMS", 
        "vTotTrib", 
        "infAdFisco",
        "chaveNFe", 
        "dPrevNFe", 
        "RNTRC"
    ]

    # Adiciona os títulos como primeira linha do arquivo XLSX
    data.insert(0, headers)

    # Criação da relação dos nomes dos arquivos que não foram processados
    if ignored_files:
        now = datetime.now()
        formatted_date = now.strftime("%Y-%m-%d_%H-%M")
        ignored_filename = f"ArquivosInvalidos_{formatted_date}.xlsx"

        workbook = px.Workbook()
        sheet = workbook.active
        sheet.title = "Arquivos Invalidos"

        for ignored_file in ignored_files:
            sheet.append([ignored_file])

        ignored_filepath = os.path.join(xlsx_folder, ignored_filename)
        workbook.save(ignored_filepath)
        print(f"Arquivos invalidos salvos em: {ignored_filepath}")

    save_to_xlsx(data, xlsx_folder)

if __name__ == "__main__":
    main()
