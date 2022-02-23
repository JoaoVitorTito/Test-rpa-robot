*** Settings ***
Documentation     Test RPA form web
Library           RPA.Browser.Selenium    auto_close=${FALSE}
Library           RPA.Desktop.Windows
Library           RPA.HTTP
Library           RPA.Excel.Files
Library           RPA.PDF
Library           RPA.Tables
Library           RPA.JSON
Library           RPA.FileSystem
*** Tasks ***
Test RPA form web
    Open the intranet website
    Enter Register adm
    Register adm
    EnterProduct
    Fill the products using the data from the Excel file
    List Products
    Export the table as a Excel

*** Keywords ***
Open the intranet website
    Open Available Browser    https://front.serverest.dev/login
    Sleep    1

Enter Register adm
    Wait Until Page Contains Element    css=[data-testid="cadastrar"]
    Click Link    css=[data-testid="cadastrar"]

    Sleep    1

Register adm
    Sleep     1
    Input Text    nome    random
    Input Text    email    sidera@gmail.com
    Input password    password    12345
    Click Button    administrador
    Click Button    css=[data-testid="cadastrar"]

EnterProduct
    Wait Until Page Contains Element    css=[data-testid="cadastrarProdutos"]
    Click Link    css=[data-testid="cadastrarProdutos"]
    
Register Product
    Sleep    1
    [Arguments]    ${product_sell}
    Click Link       css=[data-testid="cadastrar-produtos"]
    Input Text    nome    ${product_sell}[Nome]
    Input Text    price    ${product_sell}[Preço]
    Input Text    css=[data-testid="descricao"]    ${product_sell}[Descrição]
    Input Text    quantity    ${product_sell}[Quantidade]
    Http Get    https://api.thecatapi.com/v1/images/search      overwrite=True 
    Download    https://cdn2.thecatapi.com/images/ebv.jpg    ${OUTPUT_DIR}${/}imagem.jpg
    Choose File    imagem    D:\\teste-rpa-robot\\output\\imagem.jpg
    Click Button    Cadastrar
    Wait Until Page Contains Element    css=[data-testid="cadastrar-produtos"]
    Click Link       css=[data-testid="cadastrar-produtos"]

Fill the products using the data from the Excel file
    Open Workbook    Tabela de Produtos.xlsx
    ${product_sells}=    Read Worksheet As Table    header=True
    Close Workbook
    FOR    ${product_sell}    IN    @{product_sells}
        Register Product    ${product_sell}
    END

List Products
    Wait Until Page Contains Element    css=[data-testid="listar-produtos"]
    Click Link    css=[data-testid="listar-produtos"]

Export the table as a Excel
     Create Workbook    produtos.xlsx
    Set Worksheet Value    1    1    Nome
    Set Worksheet Value    1    2    Preço
    Set Worksheet Value    1    3    Descrição
    Set Worksheet Value    1    4    Quantidade
    Set Worksheet Value    1    5    _id
    Set Worksheet Value    1    6    imagem
    ${response}=    Http Get    https://serverest.dev/produtos
    Wait Until Page Contains Element    css=[data-testid="listar-produtos"]  
    Append Rows To Worksheet    ${response.json()}[produtos]
    Save Workbook
    Close Workbook
    Close All Browsers