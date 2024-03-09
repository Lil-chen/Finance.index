from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
def yt_comment(title):
    options = Options()
    options.add_argument("--disable-notifications")
    options.add_experimental_option('detach', True)
    driver = webdriver.Chrome(options=options)


    driver.get("https://www.youtube.com/")
    WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.XPATH, "/html/body/ytd-app/div[1]/div/ytd-masthead/div[4]/div[2]/ytd-searchbox/form/div[1]/div[1]/input")))

    search_box=driver.find_element(By.XPATH, "/html/body/ytd-app/div[1]/div/ytd-masthead/div[4]/div[2]/ytd-searchbox/form/div[1]/div[1]/input")
    search_box.send_keys(title)
    WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.XPATH, "/html/body/ytd-app/div[1]/div/ytd-masthead/div[4]/div[2]/ytd-searchbox/button/yt-icon/yt-icon-shape/icon-shape/div")))
    search_buttom=driver.find_element(By.XPATH,"/html/body/ytd-app/div[1]/div/ytd-masthead/div[4]/div[2]/ytd-searchbox/button/yt-icon/yt-icon-shape/icon-shape/div")
    search_buttom.click()
    WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.XPATH,"/html/body/ytd-app/div[1]/ytd-page-manager/ytd-search/div[1]/ytd-two-column-search-results-renderer/div/ytd-section-list-renderer/div[2]/ytd-item-section-renderer/div[3]/ytd-video-renderer[1]/div[1]/div/div[1]/div/h3/a")))
    yt = driver.find_element(By.XPATH,'/html/body/ytd-app/div[1]/ytd-page-manager/ytd-search/div[1]/ytd-two-column-search-results-renderer/div/ytd-section-list-renderer/div[2]/ytd-item-section-renderer/div[3]/ytd-video-renderer[1]/div[1]/div/div[1]/div/h3/a')
    yt.click()
    time.sleep(3)
    for _ in range(20):
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.END)
        time.sleep(1)

    comments=driver.find_elements(By.XPATH,'//*[@id="content-text"]')
    gg=[]
    for comment in comments:
        gg.append(comment.text)
    
    

    driver.close()

    data=pd.DataFrame({"comment":gg})
    print(data.head(20))
    data.to_excel(title + ".xlsx")



yt_comment("bad and boujee")






