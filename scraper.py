from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from openpyxl import Workbook
import time
import pandas as pd


######################### import path
url = "https://www.walmart.com/ip/Clorox-Disinfecting-Wipes-225-Count-Value-Pack-Crisp-Lemon-and-Fresh-Scent-3-Pack-75-Count-Each/14898365"
driver = webdriver.Chrome("chromedriver.exe")
driver.maximize_window()

driver.get(url)
driver.implicitly_wait(5)

driver.maximize_window()
time.sleep(5)

################ Scroll the window
scroll_review = driver.find_element_by_xpath("//*[@id='checkout-comments-container']")
driver.execute_script("arguments[0].scrollIntoView();", scroll_review)
time.sleep(3)

############click the All review button
driver.find_element_by_xpath("//*[@id='customer-reviews-header']/div[2]/div/div[3]/a[2]").click()
time.sleep(2)

###########scroll the down to sort reviews
scroll_sortby = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div[1]/div/div[5]/div/div[1]/div[2]/div/label[1]/input")
driver.execute_script("arguments[0].scrollIntoView();", scroll_sortby)
time.sleep(3)


########## sort review
sort_by = Select(driver.find_element(By.XPATH,"//select[@aria-label='Sort by']"))
sort_by.select_by_visible_text("newest to oldest")
time.sleep(3)

########## Fetch the All the Review informations
reviews_date1 = driver.find_elements(By.XPATH,"//*[@itemprop='datePublished']")
review_product_name1 = driver.find_elements(By.XPATH,"//*[@class='review-footer-userNickname']") 
review_title_name1= driver.find_elements(By.XPATH,"//*[@class='review-title font-bold']") 
review_description1= driver.find_elements(By.XPATH," //*[@class='review-text'] ") 
review_star_ratings1= driver.find_elements(By.XPATH,"//*[@class='arranger stars stars-small arranger--items-center']") 



########### All data stores in Seprate variables
dates1 = []
for date in reviews_date1:
    dates1.append(date.text)
    
names1 = []
for name in review_product_name1:
    names1.append(name.text)
    
title_names1 = []
for titles in review_title_name1:
    title_names1.append(titles.text)
    
descriptions1 = []
for describe in review_description1:
    descriptions1.append(describe.text)
    
review_stars1 = []
for star in review_star_ratings1:
    review_stars1.append(star.text)

 
###########All Variables seprate into one zip file 
finallist_reviews_firstpage = zip(dates1, names1, title_names1, descriptions1, review_stars1)


wb1 = Workbook()
wb1['Sheet'].title= 'Wallmart Reviews'
sh1 = wb1.active

########### Add the columns name
sh1.append(["Dates", "Names", "Title", "Description", "Rating"])
for x1 in list(finallist_reviews_firstpage):
    sh1.append(x1)

#######################Finally make the seprate one page xlsx file.    
wb1.save("Wallmart_reviews1.xlsx")

######################################################################################################### page2



############################### click the 2nd page
sort_by_page2 = Select(driver.find_element(By.XPATH,"//select[@aria-label='Sort by']"))
sort_by_page2.select_by_visible_text("most helpful")
sort_by = Select(driver.find_element(By.XPATH,"//select[@aria-label='Sort by']"))
sort_by.select_by_visible_text("newest to oldest")


###############################sort the final sort review
sort_by_page2 = Select(driver.find_element(By.XPATH,"//select[@aria-label='Sort by']"))
sort_by_page2.select_by_visible_text("most helpful")

################################Scrolldown to the 2ndpage line
scroll_sortby_downword = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div[1]/div/div[5]/div/div[1]")
driver.execute_script("arguments[0].scrollIntoView();", scroll_sortby_downword)
time.sleep(3)

button = driver.find_element_by_class_name("paginator-btn-next")
button.click()
time.sleep(2)

scroll_sortby_downword = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div[1]/div/div[6]/div[1]/div[20]")
driver.execute_script("arguments[0].scrollIntoView();", scroll_sortby_downword)
time.sleep(5)

scroll_sortby_downword = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div[1]/div/div[5]/div/div[1]")
driver.execute_script("arguments[0].scrollIntoView();", scroll_sortby_downword)
time.sleep(5)

#################################### Fetching All the 2nd Page reviews informations
reviews_date2 = driver.find_elements(By.XPATH,"//*[@itemprop='datePublished']")
review_product_name2 = driver.find_elements(By.XPATH,"//*[@class='review-footer-userNickname']") 
review_title_name2= driver.find_elements(By.XPATH,"//*[@class='review-title font-bold']") 
review_description2= driver.find_elements(By.XPATH," //*[@class='review-text'] ") 
review_star_ratings2= driver.find_elements(By.XPATH,"//*[@class='arranger stars stars-small arranger--items-center']") 


######################### Again store seprate variables
dates2 = []
for date in reviews_date2:
    dates2.append(date.text)
    
names2 = []
for name in review_product_name2:
    names2.append(name.text)
    
title_names2 = []
for titles in review_title_name2:
    title_names2.append(titles.text)
    
descriptions2 = []
for describe in review_description2:
    descriptions2.append(describe.text)
    
review_stars2 = []
for star in review_star_ratings2:
    review_stars2.append(star.text)
 
 
finallist_reviews_firstpage2 = zip(dates2, names2, title_names2, descriptions2, review_stars2)
wb2 = Workbook()
wb2['Sheet'].title= 'Wallmart Reviews'
sh2 = wb2.active

sh2.append(["Dates", "Names", "Title", "Description", "Rating"])
for x2 in list(finallist_reviews_firstpage2):
    sh2.append(x2)

############ store data in 2nd page seprate excel file    
wb2.save("Wallmart_reviews2.xlsx")


#############################   marge both Excels files into one excel file

df1 = pd.read_excel("wallmart_reviews1.xlsx")
df2 = pd.read_excel("wallmart_reviews2.xlsx")

values1 = df1         #[[" Dates_page", " Names_page", " Title_page", " Description_page", " Rating_page"]]
values2 = df2         #[[" Dates_page", " Names_page", " Title_page", " Description_page", " Rating_page"]]

dataframes = [values1, values2]
join = pd.concat(dataframes)
join.to_excel("output.xlsx")

##################showinf outpu
print(f'First page reviews is: {df1}')

print('*'*100)

print(f'Second page reviews is: {df2}')

driver.quit() 
############################################################################





