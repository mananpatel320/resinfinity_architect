#!/usr/bin/env python
# coding: utf-8

# In[11]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import xlrd
import pandas as pd
from selenium.common.exceptions import NoSuchElementException        


# In[12]:


pathInv = 'architect_contacts.xlsx'
openFileInv = pd.read_excel(pathInv, engine='openpyxl', sheet_name='Sheet1')
ph_num = openFileInv['Phone']
name = openFileInv['Name']
total = 0
for i in ph_num:
    print(str(i) + " : " + name[total])
    total += 1
print("Total = " + str(total))


# In[13]:


def check_exists_by_xpath(xpath):
    try:
        driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True


# In[14]:


driver = webdriver.Chrome('D:\\resinfinity_architect\\chromedriver\\chromedriver.exe')
url = "https://web.whatsapp.com/"
driver.maximize_window()
driver.get(url)
val = input("Have you logged in to whatsapp? (y/n) : ")
if(val == 'n'):
    inp = input("Press enter to exit.")
    exit()
else:
    if(check_exists_by_xpath('//*[@id="side"]/header/div[2]/div/span/div[3]/div/span') == False):
        inp = input("[-] Error: It appears that you have not logged in. Press enter to exit.")
        exit()


# In[15]:


ctr = 0
sent = 0
notsent = 0
for i in ph_num:
    url = "https://web.whatsapp.com/send?phone=91" + str(i) + "&text&app_absent=0"
    driver.get(url)
    file_path = "D:\\resinfinity_architect\\main.jpeg"
    brochure_path = "D:\\resinfinity_architect\\Brochure-Resinfinity.pdf"
    catalog1_path = "D:\\resinfinity_architect\\Resin_Table_Catalogue-Resinfinity.pdf"
    catalog2_path = "D:\\resinfinity_architect\\Resin_Artifacts_Catalogue-Resinfinity.pdf"
    next = True
    print("[-] " + str(i) + " " + name[ctr] + " : Initiating Connection.")
    while(True):
        if(check_exists_by_xpath('//*[@id="main"]/footer/div[1]/div/div/div[2]/div[1]/div/div[2]')):
            break
        if(check_exists_by_xpath('//*[ text() = "Phone number shared via url is invalid." ]')):
            print("[-] " + str(i) + " " + name[ctr] + " : Error: number does not exist.")
            next = False
            notsent += 1
            break
    if(next):
        print("[+] " + str(i) + " " + name[ctr] + " : Connection Successful.")
        while(True):
            if(check_exists_by_xpath('//div[@title = "Attach"]')):
                break
        attachment_section = driver.find_element_by_xpath('//div[@title = "Attach"]')
        attachment_section.click()
        while(True):
            if(check_exists_by_xpath('//*[@id="main"]/footer/div[1]/div/div/div[1]/div[2]/div/span/div[1]/div/ul/li[1]/button/input')):
                break
        image_box = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/div/div[1]/div[2]/div/span/div[1]/div/ul/li[1]/button/input')
        image_box.send_keys(file_path)
        while(True):
            if(check_exists_by_xpath('//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[2]')):
                break
        msg_box = driver.find_element_by_xpath('//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[2]')
        msg = "Hi " + name[ctr] + ","
        msg_box.send_keys(msg)

        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = "Greetings for the day!"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = "We at *RESINFINITY* are luxury interior décor designers that fuse resin with conventional materials to produce beautiful, elegant and unique masterpieces."
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = "We manufacture designer products that are"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":infinity"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " One of a kind"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":infinity"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Completely customizable to your choice"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":infinity"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Sturdy and robust"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":infinity"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " 100% chemical safe"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":infinity"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " 100% waterproof"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":infinity"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Food grade safe"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":infinity"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Crystal clear"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":infinity"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Great visual expansion of space"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = "We have experience of manufacturing following types of designer art pieces working with resin."
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Dining tables"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Conference tables"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Centre tables"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Office tables (Boss Table)"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Coffee tables"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Corner tables"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Designer wall clocks"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Designer wash basins"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Restaurant tables"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Wall art"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Kitchen countertops"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Flush door surface art"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Closet doors surface art"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Wooden/RCC benches "
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":tick"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.ENTER)
        msg = " Designer base frame"
        msg_box.send_keys(msg)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = "Please find the attached Brochure and Catalogues below of our products."
        msg_box.send_keys(msg)

        msg_box.send_keys(Keys.ENTER)     

        while(True):
            if(check_exists_by_xpath('//div[@title = "Attach"]')):
                break
        attachment_section = driver.find_element_by_xpath('//div[@title = "Attach"]')
        attachment_section.click()
        while(True):
            if(check_exists_by_xpath('//*[@id="main"]/footer/div[1]/div/div/div[1]/div[2]/div/span/div[1]/div/ul/li[3]/button/input')):
                break
        file_box = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/div/div[1]/div[2]/div/span/div[1]/div/ul/li[3]/button/input')
        file_box.send_keys(brochure_path)

        while(True):
            if(check_exists_by_xpath('//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/input')):
                break
        catalog1_box = driver.find_element_by_xpath('//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/input')
        catalog1_box.send_keys(catalog1_path)

        while(True):
            if(check_exists_by_xpath('//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/input')):
                break
        catalog2_box = driver.find_element_by_xpath('//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/input')
        catalog2_box.send_keys(catalog2_path)

        while(True):
            if(check_exists_by_xpath('//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/div[2]/div[2]/div/div[1]')):
                break
        send_btn = driver.find_element_by_xpath('//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/div[2]/div[2]/div/div[1]')
        send_btn.click()

        while(True):
            if(check_exists_by_xpath('//*[@id="main"]/footer/div[1]/div/div/div[2]/div[1]/div/div[2]')):
                break
        msg_box2 = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/div/div[2]/div[1]/div/div[2]')
        msg = "Request you to please go through the same. Also, suggest suitable time for further discussion so that we can *show you some of our designs in person*. Please reach out to us if any additional information is required."
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = "Thanks and regards."
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = "Jugal Patel"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = "RESINFINITY"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = "304, Abhishek Shopping,"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = "Sector-11,"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = "Gandhinagar-382011"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = "(M) +91 8320 210 911"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box2.send_keys(Keys.ENTER)

        while(True):
            if(check_exists_by_xpath('//*[@id="main"]/footer/div[1]/div/div/div[2]/div[1]/div/div[2]')):
                break
        msg = "Please find our creations on social media accounts for further references."
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":internet"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.ENTER)
        msg = " www.resinfinity.com"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":internet"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.ENTER)
        msg = " https://wa.me/c/918320210911"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":internet"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.ENTER)
        msg = " www.facebook.com/resinfinity"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = ":internet"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.ENTER)
        msg = " www.instagram.com/resinfinity"
        msg_box2.send_keys(msg)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg_box2.send_keys(Keys.COMMAND, Keys.SHIFT, Keys.ENTER)
        msg = 'Since this is probably our first conversation, our social media links may not be visible as hyperlink. Just reply *"Hi"* to make all website hyperlinks clickable.'
        msg_box2.send_keys(msg)
        while(True):
            if(check_exists_by_xpath('//*[@id="main"]/footer/div[2]/div/div[5]/div[1]/div[1]/div[1]/div')):
                break
        msg_box2.send_keys(Keys.ENTER)
        time.sleep(5)
        sent += 1
        print("[+] " + str(i) + " " + name[ctr] + " : Successful")
        
    print("[~] Sent     : " + str(sent))  
    print("[~] Not Sent : " + str(notsent))
    print("[~] Pending  : " + str(total - ctr - 1))
    print("[~] Total    : " + str(total))
    ctr += 1
    
print("[+] You're all caught up!")


# In[ ]:




