####################################################
# Sacrapy con Selenium 
# Rastrea en Amazon precio de Procesadores Ryzen 7
# sin ser detectado
# Autor: Enrique Estevez
####################################################

import requests
from requests_html import HTMLSession
from bs4 import BeautifulSoup as bs
import random
import time
import pandas as pd 
import numpy as np 
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import undetected_chromedriver as uc 
import time

s=HTMLSession()

browser=uc.Chrome()
url='https://www.amazon.com.mx/s?k=procesador+ryzen+7&sprefix=procesador%2Caps%2C334&ref=nb_sb_ss_ts-doa-p_3_10'
browser.get(url)


def get_data(url):
    r=s.get(url)
    r.html.render(timeout=20)
    soup=bs(r.html.hmtl, 'hmtl.parser')

    return soup

def get_object(soup):
    articulo=soup.find_all('div', {'data-component-type':'s-search-result'})
    for articulo in articulo:

        titulo=articulo.find('span', {'class':'a-size-base-plus'}).text
        envio=articulo.find('span', {'class':'a-color-base'}).text
        precio=articulo.find('span', {'class':'a-price-whole'}).text

        almacenador={

        'Titulo':titulo,
        'Envio':envio,
        'Precio':precio
        }

        lista_diccionarios.append(almacenador)
    return lista_diccionarios

print(get_object(get_data(url)))