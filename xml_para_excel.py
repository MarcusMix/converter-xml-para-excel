import xmltodict
import os
import pandas as pd
from datetime import datetime
import json

def pegar_infos(nome_arquivo, valores, contador_sucesso, contador_erros):

    if nome_arquivo.lower() == 'desktop.ini':
        print(f"Ignorando o arquivo {nome_arquivo}")
        return contador_sucesso, contador_erros
    
    print(f"Pegou as informações {nome_arquivo}")
    caminho_completo = os.path.join(pasta_xml_absoluta, nome_arquivo)

    with open(caminho_completo, 'rb') as arquivo_xml:

        try:
            dict_arquivo = xmltodict.parse(arquivo_xml)
                
            if "nfeProc" in dict_arquivo:
                info_nf = dict_arquivo["nfeProc"]["NFe"]["infNFe"]
            elif "NFe" in dict_arquivo:
                info_nf = dict_arquivo["NFe"]["infNFe"]
            else:
                print("Arquivo de erro " , nome_arquivo)
                contador_erros += 1

                return contador_sucesso, contador_erros

            numero_nota = info_nf["@Id"]
            cnpj_fornecedor = info_nf["emit"]["CNPJ"]
            fornecedor = info_nf["emit"]["xFant"]
            nfe_numero = info_nf["ide"]["nNF"]

            if "transp" in info_nf:

                if "vol" in info_nf["transp"]:

                    if "pesoB" in info_nf["transp"]["vol"]:
                        peso_bruto = info_nf["transp"]["vol"]["pesoB"]
                    else:
                        peso_bruto = "Não informado"
                else:
                    peso_bruto = "Não informado"
            else:
                peso_bruto = "Não informado"
        
            if "ide" in info_nf:

                if "emit" in info_nf["ide"]:    
                    valor_frete = info_nf["ide"]["emit"]["ICMSTot"]["VFrete"]
                else:
                    valor_frete = 0

                if "dhSaiEnt" in info_nf["ide"]:
                    data = info_nf["ide"]["dhSaiEnt"]
                else:
                    data = info_nf["ide"]["dhEmi"]
                
            else:
                data = "Data não informada"

            cnpj_loja = info_nf["dest"]["CNPJ"]
            loja = info_nf["dest"]["xNome"]

            produtos = info_nf.get("det", [])

            if "cobr" in info_nf:
                valor_nota = info_nf["cobr"]["fat"]["vLiq"]
            else: 
                valor_nota = "Não informado"
            
            try:
                if isinstance(produtos, list):
                    for produto in produtos:
                        produto_data = {
                            "codigo_unico": (nfe_numero + produto["prod"]["cProd"]),
                            "codigo_produto" : produto["prod"]["cProd"],
                            "numero_nota": numero_nota,
                            "nota": nfe_numero,
                            "cnpj_fornecedor": cnpj_fornecedor,
                            "fornecedor": fornecedor,
                            "data": datetime.strptime(data[:10] , "%Y-%m-%d").strftime("%d/%m/%Y"),
                            "cnpj_loja": cnpj_loja,
                            "loja": loja,
                            "valor_nota": valor_nota,
                            "valor_frete" : valor_frete,
                            "peso_bruto" : peso_bruto,
                            "produto": produto["prod"]["xProd"],
                            "quantidade": round(float(produto["prod"]["qCom"]), 2), 
                            "valor": round(float(produto["prod"]["vUnCom"]), 2)
                        }
                        valores.append(produto_data)
                        contador_sucesso += 1
                elif isinstance(produtos, dict):
                    
                    produto_data = {
                        "codigo_unico": (nfe_numero + produto["prod"]["cProd"]),
                        "codigo_produto" : produto["prod"]["cProd"],
                        "numero_nota": numero_nota,
                        "nota": nfe_numero,
                        "cnpj_fornecedor": cnpj_fornecedor,
                        "fornecedor": fornecedor,
                        "data": datetime.strptime(data[:10], "%Y-%m-%d").strftime("%d/%m/%Y"),
                        "cnpj_loja": cnpj_loja,
                        "loja": loja,
                        "valor_nota": valor_nota,
                        "valor_frete" : valor_frete,
                        "peso_bruto" : peso_bruto,
                        "produto": produto["prod"]["xProd"],
                        "quantidade": round(float(produto["prod"]["qCom"]), 2),
                        "valor": round(float(produto["prod"]["vUnCom"]), 2)
                    }
                    valores.append(produto_data)
                    contador_sucesso += 1
            except Exception as e:
                print(json.dumps(dict_arquivo, indent=4))
                print(f"Erro no arquivo {nome_arquivo}: {e}")
                contador_erros += 1

        except xmltodict.expat.ExpatError as e:
            print(f"Arquivo {nome_arquivo} NÃO É XML: {e}")
            contador_erros +=1
 
            
    return contador_sucesso, contador_erros

colunas = ["codigo_unico", "codigo_produto", "numero_nota", "nota", "cnpj_fornecedor", "fornecedor", "data", "cnpj_loja", "loja", "valor_nota", "valor_frete", "peso_bruto", "produto", "quantidade", "valor"]
valores = []


pasta_xml_absoluta = "C:\\Users\\adm\\Área de Trabalho\\python"

lista_arquivos = os.listdir(pasta_xml_absoluta)

contador_sucesso = 0
contador_erros = 0

for arquivo in lista_arquivos:
    contador_sucesso, contador_erros = pegar_infos(arquivo, valores, contador_sucesso, contador_erros)

tabela = pd.DataFrame(columns=colunas, data=valores)

print("Processamento concluído!")
print("Arquivos processados com sucesso:", contador_sucesso)
print("Arquivos com erros:", contador_erros)

tabela.to_excel("XML_NOVO.xlsx", index=False)
