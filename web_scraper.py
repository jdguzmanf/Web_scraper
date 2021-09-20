#####################
# Required packages #
#####################

from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import datetime as dt
import time
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import numpy as np
from openpyxl import load_workbook


#####################
#       code        #
#####################


#####################################################
#These are just some of the shares listed on the BVC.
stock_names = ["ecopetrol", "cemargos", "pfcemargos"]



##################################################################################################################################################################
#  The escraper will be programmed in selenium and will enter the BVC page to download the desired information and store it in the list  "chains". 
#
#  For them it is necessary to have the driver of the browser you want to use, I decided to use Chrome (https://chromedriver.chromium.org/downloads).
#
#  For Selenium to work it is necessary to define a driver and the URL to which it should access, in this case 
#  (https://www.bvc.com.co/pps/tibco/portalbvc/Home/Mercados/enlinea/acciones?action=dummy)

#  Once inside the page we tell the selenium program to search for an html element, the stock search engine (we name the python object search), and enter the
#  names of the stocks we are interested in, the ones defined in the chains list.

#  Now we obtain the data of interest for each of these stocks by instructing the bot to extract the information as a text string in another html element which
#  is a summary table. All the stock information is stored in the "chains" list.
##################################################################################################################################################################




while(True):
    for name in stock_names:
        chains = []

        ## Where you have installed the driver of your favorite navigator
        PATH = "C:/Program Files (x86)/chromedriver.exe"
        driver  = webdriver.Chrome(PATH)
        driver.get("https://www.bvc.com.co/pps/tibco/portalbvc/Home/Mercados/enlinea/acciones?action=dummy")

        ## The stock search engine, a text box
        search = driver.find_element_by_id("nemo")
        ## Click in text box and write stock name in stock_names
        search.send_keys(f"{name}")
        ## Click and search 
        search.send_keys(Keys.RETURN)

        ## The summary table, an html object
        tabla = driver.find_element_by_class_name("tabla_basica")
        ## Append all the information in chains
        chains.append(tabla.text.split(" "))

        print(chains)
        ## Sleep time to make sure there is no internet conexions problems
        time.sleep(5)
        ## The selenium program has done its job!!!
        driver.quit()




        ## Now that we have save all the important information we want it's time to organize it in a database for your use.
        rel_path = "C:/Users/User/Desktop/Esta"
        ## Using openpyxl, wich is a great library, we create a xlsx file for each stock and save the information
        day_stock = load_workbook(filename=rel_path+"/"+f"dia_{name}.xlsx")
        sheet = day_stock.active

        ## We can save as much information as we want, in this case i will just save this 5.

        #date
        date = chains[0][2]
        #hour
        hour = chains[0][3][:5]
        #volume
        volume = chains[0][25][:15].replace(".", "").replace(",","")
        volume = int(volume)
        #PRECIO CIERRE
        close_price = chains[0][26][:5].replace(".", "").replace(",","")
        close_price = int(close_price)
        #PRECIO CIERRE ANTERIOR
        close_price_ant = chains[0][27][:5].replace(".", "").replace(",","")
        close_price_ant = int(close_price_ant)


        sheet.append([date, hour, volume, close_price, close_price_ant])
        day_stock.save(rel_path+"/"+f"dia_{name}.xlsx")

    ## The program runs automatically every 15 minutes 
    time.sleep(900)



#################################################################################################################################################################
#  You could automate this code by simply creating a .bat file.
#  You just have to create a txt file like this and save it as .bat

#  .bat file
#  "Python directory" "directory of the program, in this case, web_scraper.py file"
#  pause

#  Then you just have to run the bat program and that it.
#################################################################################################################################################################

