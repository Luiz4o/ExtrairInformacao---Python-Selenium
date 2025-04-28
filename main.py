from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import time
import os

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def main():
    logs = []
    options = webdriver.ChromeOptions()
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=options)


    driver.get("https://www.magazineluiza.com.br/")

    wait = WebDriverWait(driver, 20)

    for i in range(3):
        try:
            input_search_bar = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[data-testid="input-search"]'))
            )
            print("Barra de busca encontrada com sucesso.")
            break
        except TimeoutException:
            if i==2 :
                logs.append("Site fora do ar ou falha encontrar a barra de busca.")
                driver.quit()
                save(pd.DataFrame(), pd.DataFrame(), logs) 
                return 
            
            continue

    input_search_bar.send_keys("notebooks")
    input_search_bar.submit()

    search_itens(driver, logs)


def search_itens(driver, logs):
    melhores = []
    piores = []

    next_button_selector = (By.CSS_SELECTOR, 'button[aria-label="Go to next page"]')

    page_num = 1 
    while True:
        time.sleep(8)

        print(f"Processando página {page_num}...")

        try:
            product_items = driver.find_elements(By.XPATH, '/html/body/div[2]/div/main/section[4]/div[5]/div/ul/li')

            print(f"Encontrados {len(product_items)} produtos na página {page_num}.")
        except TimeoutException:
            logs.append(f"Produtos não carregaram na página {page_num}. Possível fim ou erro.")
            break 

        for i, product_card in enumerate(product_items):
            try:
                product_title = product_card.find_element(By.CSS_SELECTOR, '[data-testid="product-title"]').text

                reviews = product_card.find_elements(By.CSS_SELECTOR, '[data-testid="review"]')
                review_text = reviews[0].text if reviews else "0"

                qtd_avaliacoes = 0

                if "(" in review_text:
                    try:
                        qtd_avaliacoes = int(review_text.split("(")[-1].replace(")", "").strip())
                    except (ValueError, IndexError):
                        logs.append(f"Erro ao extrair quantidade de avaliações para '{product_title}' na página {page_num}.")
                        qtd_avaliacoes = 0 

                try:
                     url = product_card.find_element(By.TAG_NAME, 'a').get_attribute('href')
                except NoSuchElementException:
                     logs.append(f"Link do produto não encontrado para '{product_title}' na página {page_num}.")
                     url = "N/A"


                product_data = {
                    "PRODUTO": product_title,
                    "QTD_AVAL": qtd_avaliacoes,
                    "URL": url
                }

                if qtd_avaliacoes > 100:
                    melhores.append(product_data)
                else:
                    piores.append(product_data)

            except Exception as e:
                logs.append(f"Erro ao processar produto {i+1} na página {page_num}: {e}")
                continue 

        try:
            next_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(next_button_selector)
            )


            next_button.click()
            print(f"Clicado em 'Próxima página' na página {page_num}.")
            page_num += 1

        except (TimeoutException, NoSuchElementException):
            print(f"Botão 'Próxima página' não encontrado após a página {page_num}. Fim da paginação.")
            break 

    df_melhores = pd.DataFrame(melhores)
    df_piores = pd.DataFrame(piores)

    driver.quit()
    save(df_melhores, df_piores, logs)


def save(df_melhores, df_piores, logs=[]):
    df_logs = pd.DataFrame(logs, columns=["LOGS"])

    try:
        with pd.ExcelWriter("Notebook.xlsx", engine="openpyxl") as writer:
            if not df_melhores.empty:
                df_melhores.to_excel(writer, sheet_name="melhores", index=False)
            else:
                 pd.DataFrame(columns=["PRODUTO", "QTD_AVAL", "URL"]).to_excel(writer, sheet_name="melhores", index=False)
                 logs.append("Nenhum produto com mais de 100 avaliações encontrado.")

            if not df_piores.empty:
                df_piores.to_excel(writer, sheet_name="piores", index=False)
            else:
                 pd.DataFrame(columns=["PRODUTO", "QTD_AVAL", "URL"]).to_excel(writer, sheet_name="piores", index=False)
                 logs.append("Nenhum produto com 100 ou menos avaliações encontrado (ou erro no processamento).")


            df_logs.to_excel(writer, sheet_name="logs", index=False)

        print(f"Arquivo 'Notebooks.xlsx' salvo com sucesso!")

    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")
        with open("notebooks_logs_error.txt", "w") as f:
            for log in logs:
                f.write(log + "\n")

    send_email(os.getenv('EMAIL_FROM'), os.getenv('EMAIL_PASS') , os.getenv('EMAIL_TO'), "Relatório Notebooks", "Olá, aqui está o seu relatório dos notebooks extraídos da Magazine Luiza." , "Notebook.xlsx")

def send_email(From, password, To, subject, body, file):
    mensagem = MIMEMultipart()
    mensagem['From'] = From
    mensagem['To'] = To
    mensagem['Subject'] = subject

    mensagem.attach(MIMEText(body, 'plain'))

    try:
        with open(file, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())

            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file)}')
            mensagem.attach(part)
    except Exception as e:
        print(f"Erro ao anexar o arquivo: {e}")

    try:
        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()
        servidor.login(From, password)
        servidor.sendmail(From, To, mensagem.as_string())
        servidor.quit()
        print("Email enviado com sucesso!")

    except Exception as e:
        print(f"Erro ao enviar o email: {e}")


if __name__ == "__main__":
    main()