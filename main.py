import requests
from bs4 import BeautifulSoup
import lxml
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from time import sleep
import random
import os
import json

def get_site_pages(url, headers):
    if "pages" not in os.listdir():
            os.mkdir("pages")
    try:
        for page in range(1, 13):
            responce = requests.get(url=url + f"?PAGEN_2={page}", headers=headers)
            src = responce.text
            
            with open(f"pages/page_#{page}.html", "w") as file:
                file.write(src)
                print(f"[*] Page #{page} has been successfully saved. {12 - page} left...")

            sleep(random.random())

    except requests.ConnectionError as er:
            print("[!] Connection is failed. Please check your connection!")

def tables_loading():
    dirs = os.listdir("pages")

    if dirs == []:
        print(f"[!] Please, load pages previously!")
    else:
        if "tables" not in os.listdir():
                os.mkdir("tables")
        
        counter = 1
        for file_name in dirs:
            with open(f"pages/{file_name}", "r") as file:
                src = file.read()

            soup = BeautifulSoup(src, "lxml")
            dissovets = soup.find(class_="dissovet-list").find_all(class_="dissovet-name")
            
            for dissovet in dissovets:
                responce = requests.get("https://disser.niirosatom.ru" + dissovet.find("a").get("href"))

                soup = BeautifulSoup(responce.text, "lxml")
                info = soup.find(class_="dissovet-detail").find(class_="dissovet-info")
                contacts = info.find(class_="dissovet-contacts")
                description = info.find(class_="dissovet-descr")
                dissovet_name = info.find_previous_sibling("h1").text
                spec = description.find(class_="dissovet-spec").find(class_="prop-value").text
                science_direction = description.find(class_="dissovet-direction").find(class_="prop-value").text

                compos = soup.find(class_="dissovet-compos").find(class_="column")
                members = [member.text for member in compos.find(class_="dissovet-members").find_all(class_="prop-value")]

                df = pd.DataFrame({"Название диссовета": [dissovet_name],
                        "Институт": [description.find(class_="dissovet-org").find(class_="prop-value").text],
                        "Отрасль науки": [science_direction],
                        "Специальность": [spec],
                        "Шифр специальности": [description.find(class_="dissovet-code").find(class_="prop-value").text],
                        "Адрес": [contacts.find(class_="dissovet-address").text],
                        "Телефон": [contacts.find(class_="dissovet-phone").text],
                        "Сайт": [contacts.find(class_="dissovet-site").text],
                        "E-mail": [contacts.find(class_="dissovet-email").text],
                        "Председатель": [compos.find(class_="dissovet-president").find(class_="prop-value").text],
                        "Секретарь": [compos.find(class_="dissovet-secretary").find(class_="prop-value").text],
                        "Члены совета": [", ".join(members)]
                    })
                dissovet_name = dissovet_name + f"_({spec}_{science_direction})"
                dissovet_name = dissovet_name.replace(" ", "_").replace(".", "_").replace("-", "_").replace(",", "")
                df.to_excel(f"./tables/{dissovet_name}.xlsx")

                sleep(2)
                print(f"[*] Table #{counter}. {dissovet_name} table has been successfully saved!")
                counter += 1

def create_log():
    log = dict()
    for table_name in os.listdir("tables"):
        link = pd.read_excel(f"./tables/{table_name}", usecols=[8])["Сайт"].get(0)
        log.update({table_name: [False, link]})
        
    with open("log.json", "w") as file:
        json.dump(log, file, indent=4, ensure_ascii=False)

def start_working():
    dirs = os.listdir("tables")
    if dirs == []:
        print(f"[!] Please, load tables previously!")
    else:
        with open("log.json", "r") as log:
             dissoviets = json.load(log)
        
        counter = 0
        for soviet, item in dissoviets.items():
            counter += 1
            if item[0] == True:
                continue
            
            driver = webdriver.Firefox()
            driver.get(item[1])
            osCommandString = f"open 'tables/{soviet}'"
            os.system(osCommandString)
            try:
                element = WebDriverWait(driver, 3600).until(EC.number_of_windows_to_be(0))
                dissoviets[soviet] = [True, item[1]]
                print(f"[*] {soviet} has done ({counter}/{len(dissoviets)})…")
            finally:
                driver.quit()

                with open("log.json", "w") as log:
                    json.dump(dissoviets, log, ensure_ascii=False, indent=4)

def main():
    url = "https://disser.niirosatom.ru/dissovety/"
    headers = {
        "User-Agent" : "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)\
            AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.167 YaBrowser/22.7.3.799 Yowser/2.5 Safari/537.36", 
        "Accept": "*/*"
    }

    # get_site_pages(url, headers)
    # tables_loading()
    # create_log()
    start_working()


if __name__ == "__main__":
    main()