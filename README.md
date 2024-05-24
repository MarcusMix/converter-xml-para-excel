# Extrair arquivos XML para XLSX (Excel)

Este repositório contém um script Python para ler uma pasta com arquivos XML e convertê-los para o formato XLSX (Excel).

## Instalação

Para usar o código, siga os passos abaixo:

1. Clone o repositório:

    ```bash
    git clone https://github.com/MarcusMix/converter-xml-para-excel
    ```

2. Navegue até o diretório do projeto:

    ```bash
    cd converter-xml-para-excel
    ```

3. Instale as dependências necessárias (requer `pip` e `virtualenv`):

    ```bash
    virtualenv venv
    source venv/bin/activate
    pip install -r requirements.txt
    ```

## Uso

1. Coloque os arquivos XML na pasta indicada no código.
2. Execute o script para processar os arquivos e gerar um arquivo XLSX (Excel).

    ```bash
    python converter.py
    ```

## Fluxo

O script segue o fluxo abaixo:

1. **Leitura da Pasta**: O script acessa a pasta indicada no código que contém os arquivos XML.
2. **Processamento dos Arquivos XML**: Cada arquivo XML é lido e processado.
3. **Criação do Arquivo XLSX**: Após o processamento, um arquivo XLSX é criado contendo os dados dos arquivos XML.

![Fluxo do Processo](https://i.imgur.com/y8PIg61.png)

## Exemplos de Uso

Converter arquivos XML para Excel de forma automatizada.

### Entrada (Arquivos XML)

Coloque seus arquivos XML na pasta especificada no código.

### Saída (Arquivo XLSX)

O arquivo XLSX gerado estará na pasta de saída no código.


## Licença

Este projeto está licenciado sob a licença MIT. Veja o arquivo [LICENSE](https://choosealicense.com/licenses/mit/) para mais detalhes.




### Autor

Desenvolvido por MarcusMix

---

### Referências

- [Documentação do Python](https://docs.python.org/3/)
- [Pandas Library](https://pandas.pydata.org/)
- [openpyxl Library](https://openpyxl.readthedocs.io/en/stable/)

---

