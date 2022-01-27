import xlwt 
from xlwt import Workbook 
from selenium import webdriver
import time
import selenium.webdriver.support.ui as ui
from selenium.webdriver.common.keys import Keys
driver = webdriver.Chrome('./chromedriver')
driver.implicitly_wait(10)
driver.get(
    'http://www.tsetmc.com/Loader.aspx?ParTree=15131F#')
time.sleep(70)
d=0
e=1
wb = Workbook() 
sheet1 = wb.add_sheet('Sheet 1') 
for element in driver.find_elements_by_class_name('inst') :
    if  d%2 ==0 :
        element.send_keys(Keys.CONTROL)
        element.send_keys(Keys.RETURN)
        driver.switch_to.window(driver.window_handles[0])
    d=d+1
time.sleep(5)
while d!=0:
    driver.switch_to.window(driver.window_handles[e])
    p=driver.find_elements_by_xpath("//*[@id='MainBox']//div[1]")
    p=p[0].text
    p=p[p.find("(")+1:p.rfind(")")]
    print(p)
    sheet1.write(e, 0, p) 
    for element in driver.find_elements_by_xpath("//*[@id='d11']/div"):
        z = (element.text)
        x=z[-1]
        z=z[:-1]
        z=z.replace(",","")
        z=float(z)
        print(z)
        print(x)
        sheet1.write(e, 1, z) 
        sheet1.write(e, 2, x)
        driver.find_element_by_class_name('lightslategray').click()
    c=0
    for element in driver.find_elements_by_xpath("//*[@id='PureData']/div[2]/div[2]/table/tbody/tr/td[3]"):
        c=c+float(element.text)
    print(100-c)
    sheet1.write(e, 3, 100-c) 
    wb.save('xlwt example.xls') 
    e=e+1
    d=d-2
