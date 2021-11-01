import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep

# Importando a planilha modelo
Path = input("Insira o caminho da planilha.: ")
BaseDado = pd.read_excel(rf'{Path}')

# Criando colunas vazias no DataFrame
col_names = ["ValorFornecedor", "Pos 1",
             "Pos 2", "Pos 3", "Pos 4", "Pos 5", "Status"]

# ANOTAÇÃO: Necessito modificar isso
for col in col_names:
    BaseDado.loc[0, col] = None

# Informando a licitação de interesse na coleta
Pregao = str(input("Digite o número do pregão(Ex:012021): "))

# Armazenando o link para ser aberto no navegador posteriormente
url = "http://comprasnet.gov.br/acesso.asp?url=/livre/Pregao/lista_pregao_filtro.asp?Opc=2"

# Informando o Path para o arquivo "chromedriver.exe" e armazenando em "driver"
driver = webdriver.Chrome(
    r"E:\Programacao\Python Scripts\chromedriver.exe")

# Abrindo o google chrome no site informado
driver.get(url)

# Navegador em tela cheia
driver.maximize_window()

# Configurando o tempo de espera em segundos até que o elemento html seja encontrado
wdw = WebDriverWait(driver, 2000)
wdw2 = WebDriverWait(driver, 2000)

# Mudando o main frame do site que será navegado
driver.switch_to.frame("main2")


CodUasg = wdw.until(
    EC.presence_of_element_located((By.XPATH, "//input[@id='co_uasg']"))
)
CodUasg.send_keys("150182")

NumPreg = wdw.until(EC.presence_of_element_located(
    (By.XPATH, "//input[@id='numprp']")))
NumPreg.send_keys(Pregao)

driver.execute_script("ValidaForm();")
sleep(1)

PrClick = wdw.until(
    EC.presence_of_element_located(
        (By.XPATH, "//a[contains(text(), '{}')]".format(Pregao))
    )
)
PrClick.click()

print("Preencha o CAPTCHA")

wdw.until(
    EC.presence_of_element_located(
        (By.XPATH, "//center[contains(text(), 'MINISTÉRIO DA EDUCAÇÃO')]")
    )
)
print("Portal de compras encontrado")

fail = ["Item deserto", "Cancelado no julgamento"]

############################## Alterei aqui ##############################
item, collecting, page, n = [1, True, 0, len(BaseDado.iloc[:, 0])]

while collecting:
    actual_item = (page)*100+item
    print(f'\n\n\n\n \U0001F40D Buscando pelo item: {actual_item}\n\n')

    Vencedores = wdw2.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "//tr[contains(@class, 'tex3a')][{}]/td[6]".format(
                    item)
            ),
        )
    )

    if Vencedores.text in fail:
        print("Status do item ", actual_item, ":", Vencedores.text)
        BaseDado.iloc[actual_item-1, 11] = Vencedores.text
        if actual_item == n:
            collecting = False
        elif item == 100:
            item -= 99
            page += 1
            driver.execute_script("javascript:PaginarItens('Proxima');")
        else:
            item += 1
    else:
        Vencedores = wdw2.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//tr[./td/a[contains(@href, 'javascript:void(0)')]][{}]/td[6]/a".format(
                        item),
                )
            )
        )

        print("Status do item ", actual_item, ":", Vencedores.text, "\n")
        BaseDado.iloc[actual_item-1, 11] = Vencedores.text
        Vencedores.click()
        driver.switch_to.window(driver.window_handles[1])
        Captcha = wdw2.until(
            EC.presence_of_element_located((By.XPATH, "//h2")))
        if Captcha.text == "ACOMPANHAMENTO DE PREGÃO":
            print("Preencha o CAPTCHA")
            wdw.until(EC.presence_of_element_located(
                (By.CLASS_NAME, "pregao")))

        vencedor = 1
        for colocado in range(1, 6):
            try:
                InfoEmpresa = []

                Cnpj = driver.find_element_by_xpath(
                    "//tr[@class = 'tex3'][{}]/td[1]".format(colocado)
                )

                Nome = driver.find_element_by_xpath(
                    "//tr[@class = 'tex3'][{}]/td[2]".format(colocado)
                )

                ValorEmpresa = driver.find_element_by_xpath(
                    "//tr[@class = 'tex3'][{}]/td[4]".format(colocado)
                ).text

                StatusEmpresa = driver.find_element_by_xpath(
                    "//tr[@class='tex3'][{}]/td[@class='tex3b']".format(
                        colocado)
                ).text

                ProdutoInfo = driver.find_element_by_xpath(
                    "//tr[@class='tex5a'][{}]/td[@colspan='6']".format(
                        colocado)
                )
                ProdutoInfo = ProdutoInfo.text.split()

                Marca = []
                Fabricante = []
                Versão = []
                i = 0
                while ProdutoInfo[i] != "Descrição":
                    if ProdutoInfo[i] == "Marca:":
                        PosMarca = i
                    elif ProdutoInfo[i] == "Fabricante:":
                        PosFabricante = i
                    elif ProdutoInfo[i] == "Modelo":
                        PosVersão = i
                    i += 1

                for j in range(PosMarca + 1, PosFabricante):
                    Marca.append(ProdutoInfo[j])
                for j in range(PosFabricante + 1, PosVersão):
                    Fabricante.append(ProdutoInfo[j])
                for j in range(PosVersão + 3, i):
                    Versão.append(ProdutoInfo[j])

                Marca = " ".join(Marca)
                Fabricante = " ".join(Fabricante)
                Versão = " ".join(Versão)

                InfoEmpresa = [
                    ("CNPJ: " + Cnpj.text),
                    ("NOME: " + Nome.text),
                    ("R$ " + ValorEmpresa),
                    ("Marca: " + Marca),
                    ("Fabricante: " + Fabricante),
                    ("Modelo / Versão: " + Versão),
                ]

                BaseDado.iloc[actual_item-1, (5 + colocado)] = (
                    InfoEmpresa[0]
                    + "\n"
                    + InfoEmpresa[1]
                    + "\n"
                    + InfoEmpresa[2]
                    + "\n"
                    + InfoEmpresa[3]
                    + "\n"
                    + InfoEmpresa[4]
                    + "\n"
                    + InfoEmpresa[5]
                )

                print(InfoEmpresa)

                if StatusEmpresa == "Recusado":
                    vencedor += 1
                    BaseDado.iloc[actual_item-1, (5 + colocado)] = "Recusado \U0001F6AB \n" + \
                        BaseDado.iloc[actual_item-1, (5 + colocado)]
                    print(f"{StatusEmpresa} '\U0001F6AB'")
                elif StatusEmpresa == "Aceito":
                    BaseDado.iloc[actual_item-1, (5 + colocado)] = "Aceito \U00002705 \n" + \
                        BaseDado.iloc[actual_item-1, (5 + colocado)]
                    print(f"{StatusEmpresa} '\u2705'")

                if colocado == vencedor:
                    try:
                        ValorNegociado = driver.find_element_by_xpath(
                            "//tr[@class='tex3'][{}]/td[not(node())]".format(colocado))
                        BaseDado.iloc[actual_item-1, 5] = float(
                            ValorEmpresa.replace(".", "").replace(",", "."))
                    except:
                        ValorNegociado = driver.find_element_by_xpath(
                            "//tr[@class='tex3'][{}]/td[6]".format(colocado)).text
                        BaseDado.iloc[actual_item-1, 5] = float(
                            ValorNegociado.replace(".", "").replace(",", "."))
                        BaseDado.iloc[actual_item-1,
                                      (5 + colocado)] = f"Valor Negociado: {ValorNegociado}\n"+BaseDado.iloc[actual_item-1, (5 + colocado)]

            except:
                break

        driver.switch_to.window(driver.window_handles[0])
        driver.switch_to.frame("main2")
        if actual_item == n:
            collecting = False
        elif item == 100:
            item -= 99
            page += 1
            driver.execute_script("javascript:PaginarItens('Proxima');")
        else:
            item += 1
        BaseDado.to_excel(
            f"C:/Users/Rodrigo/Desktop/Apuração-{Pregao}.xlsx", index=False)


def highlight_price(s):
    color = 'red'
    return 'background-color: %s' % color


def highlight_desert(s):
    color = 'orange'
    return 'background-color: %s' % color


def highlight_canceled(s):
    color = '#d65f5f'
    return 'background-color: %s' % color


def highlight_refused(s):
    color = '#8147d1'
    return 'background-color: %s' % color


def highlight_acepted(s):
    color = '#5cde35'
    return 'background-color: %s' % color


BaseDado = BaseDado.style.applymap(
    highlight_price, subset=pd.IndexSlice[list(BaseDado.query(
        'ValorReferência < ValorFornecedor').index), 'ValorFornecedor']
).applymap(
    highlight_desert, subset=pd.IndexSlice[list(
        BaseDado[BaseDado['Status'] == 'Item deserto'].index), 'Status']
).applymap(
    highlight_canceled, subset=pd.IndexSlice[list(
        BaseDado[BaseDado['Status'] == 'Cancelado no julgamento'].index), 'Status']
).applymap(
    highlight_refused, subset=pd.IndexSlice[list(
        BaseDado[BaseDado['Pos 1'].str.contains("Recusado", na=False)].index), 'Pos 1']
).applymap(
    highlight_refused, subset=pd.IndexSlice[list(
        BaseDado[BaseDado['Pos 2'].str.contains("Recusado", na=False)].index), 'Pos 2']
).applymap(
    highlight_refused, subset=pd.IndexSlice[list(
        BaseDado[BaseDado['Pos 3'].str.contains("Recusado", na=False)].index), 'Pos 3']
).applymap(
    highlight_refused, subset=pd.IndexSlice[list(
        BaseDado[BaseDado['Pos 4'].str.contains("Recusado", na=False)].index), 'Pos 4']
).applymap(
    highlight_refused, subset=pd.IndexSlice[list(
        BaseDado[BaseDado['Pos 5'].str.contains("Recusado", na=False)].index), 'Pos 5']
).applymap(
    highlight_acepted, subset=pd.IndexSlice[list(
        BaseDado[BaseDado['Pos 1'].str.contains("Aceito", na=False)].index), 'Pos 1']
).applymap(
    highlight_acepted, subset=pd.IndexSlice[list(
        BaseDado[BaseDado['Pos 2'].str.contains("Aceito", na=False)].index), 'Pos 2']
).applymap(
    highlight_acepted, subset=pd.IndexSlice[list(
        BaseDado[BaseDado['Pos 3'].str.contains("Aceito", na=False)].index), 'Pos 3']
).applymap(
    highlight_acepted, subset=pd.IndexSlice[list(
        BaseDado[BaseDado['Pos 4'].str.contains("Aceito", na=False)].index), 'Pos 4']
).applymap(
    highlight_acepted, subset=pd.IndexSlice[list(
        BaseDado[BaseDado['Pos 5'].str.contains("Aceito", na=False)].index), 'Pos 5']
)

BaseDado.to_excel(
    f"C:/Users/Rodrigo/Desktop/Apuração-{Pregao}.xlsx", index=False)

print("Extração Concluída \U0001F7E2")
