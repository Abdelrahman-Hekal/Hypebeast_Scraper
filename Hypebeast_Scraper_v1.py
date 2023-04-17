from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import undetected_chromedriver as uc
import time
import os
import re
from datetime import datetime
import pandas as pd
import warnings
import sys
import xlsxwriter
from multiprocessing import freeze_support
import calendar 
import shutil
warnings.filterwarnings('ignore')

def initialize_bot(translate):

    # Setting up chrome driver for the bot
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--headless')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    # configuring the driver
    driver = webdriver.Chrome(options=chrome_options, service=chrome_service)
    ver = int(driver.capabilities['chrome']['chromedriverVersion'].split('.')[0])
    driver.quit()
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--lang=en")
    chrome_options.add_argument('--headless=new')
    
    # disable location prompts & disable images loading
    if not translate:
        prefs = {"profile.default_content_setting_values.geolocation": 2, "profile.managed_default_content_settings.images": 2}  
        chrome_options.page_load_strategy = 'eager'
    else:
        prefs = {"profile.default_content_setting_values.geolocation": 2, "profile.managed_default_content_settings.images": 2, "profile.managed_default_content_settings.notifications": 1,   "translate_whitelists": {"zh-TW":"en"},"translate":{"enabled":"true"}}
        chrome_options.page_load_strategy = 'normal'

    chrome_options.add_experimental_option("prefs", prefs)
    driver = uc.Chrome(version_main = ver, options=chrome_options) 
    driver.set_window_size(1920, 1080)
    driver.maximize_window()
    driver.set_page_load_timeout(20000)

    return driver

def scrape_articles(driver, driver_en, output1, page, month, year):

    print('-'*75)
    print(f'Scraping The Articles Links from: {page}')
    # getting the full posts list
    links = []
    months = {month: index for index, month in enumerate(calendar.month_abbr) if month}
    full_months = {month: index for index, month in enumerate(calendar.month_name) if month}
    prev_month = month - 1
    if prev_month == 0:
        prev_month = 12
    driver.get(page)

    art_time = ''
    # handling lazy loading
    print('-'*75)
    print("Getting the previous month's articles..." )
    for _ in range(50):  
        try:
            height1 = driver.execute_script("return document.body.scrollHeight")
            driver.execute_script(f"window.scrollTo(0, {height1})")
            time.sleep(4)
            try:
                button = wait(driver, 1).until(EC.presence_of_element_located((By.XPATH, "//a[@id='load-more-btn']")))
                driver.execute_script("arguments[0].click();", button)
                time.sleep(1)
            except:
                pass
            try:
                art_time = wait(driver, 3).until(EC.presence_of_all_elements_located((By.TAG_NAME, "time")))[-1].get_attribute('textContent').strip()
            except:
                break
            try:
                art_month = months[art_time.split()[0]]
                art_year = int(art_time.split()[-1])
                # for articles from previous year
                if art_year < year and prev_month != 12:
                    break
                # for all months except Jan
                if art_month < prev_month and prev_month != 12 and art_year == year:
                    break
                # for Jan
                elif art_month < prev_month and prev_month == 12 and art_year < year:
                    break
            except:
                pass

            height2 = driver.execute_script("return document.body.scrollHeight")
            if height1 == height2: 
                break
        except Exception as err:
            break

    # scraping posts urls 
    try:
        posts = wait(driver, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class='post-box-content-container']")))    
    except:
        print('No posts are available')
        return

    for post in posts:
        try:
            date = wait(post, 2).until(EC.presence_of_element_located((By.TAG_NAME, "time"))).get_attribute('textContent') 
            art_month = int(months[date.split()[0]])
            art_year = int(date.split()[-1])
            if art_month != prev_month: continue
            if art_year < year and prev_month != 12: continue
            link = wait(post, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[class='title']"))).get_attribute('href')
            if link not in links:
                links.append(link)
        except:
            pass

    # scraping posts details
    print('-'*75)
    print('Scraping Articles Details...')
    print('-'*75)

    # reading previously scraped data for duplication checking
    scraped = []
    try:
        df = pd.read_excel(output1)
        scraped = df['unique_id'].values.tolist()
    except:
        pass

    n = len(links)
    data = pd.DataFrame()
    for i, link in enumerate(links):
        for _ in range(2):
            try:
                try:
                    driver.get(link)   
                except:
                    print(f'Warning: Failed to load the url: {link}')
                    continue

                art_id = ''
                try:
                    text = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class*='post-container post-status-publish ']"))).get_attribute('id').split('-')[-1].strip()
                    art_id = int(re.findall("\d+", text)[0])
                except:
                    pass
                if art_id in scraped or art_id == '': 
                    print(f'Article {i+1}\{n} is already scraped, skipping.')
                    break

                driver_en.get(link)
                time.sleep(1)
                # scrolling across the page for auto translation to be applied
                try:
                    htmlelement= wait(driver_en, 5).until(EC.presence_of_element_located((By.TAG_NAME, "html")))
                    total_height = driver_en.execute_script("return document.body.scrollHeight")
                    height = total_height/30
                    new_height = 0
                    for _ in range(20):
                        prev_hight = new_height
                        new_height += height             
                        driver_en.execute_script(f"window.scrollTo({prev_hight}, {new_height})")
                        driver.execute_script(f"window.scrollTo({prev_hight}, {new_height})")
                        time.sleep(0.2)
                except:
                    pass

                details = {}

                # English article author and date
                en_author, date = '', ''             
                try:
                    en_author = wait(driver_en, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[class='author-name']"))).get_attribute('textContent').split(':')[-1].strip()
                    date = wait(driver_en, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "time")))[1].get_attribute('textContent').strip()
                except Exception as err:
                    pass
            
                # checking if the article date is correct
                try:
                    if 'ago' in date:
                        art_month = datetime.now().month
                        art_year = datetime.now().year
                        art_day = datetime.now().day
                    else:
                        try:
                            art_month = int(full_months[date.split()[0]])
                        except:
                            art_month = int(months[date.split()[0]])
                        art_year = int(date.split()[-1])  
                        art_day = int(date.split()[1].replace(',', ''))       
                    if art_month != prev_month: 
                        print(f'skipping article with date {art_month}/{art_day}/{art_year}')
                        break
                    date = f'{art_month}/{art_day}/{art_year}'
                except:
                    #print(f'Warning: Failed to extract the date for link: {link}')
                    continue    

                details['sku'] = art_id
                details['unique_id'] = art_id
                details['articleurl'] = link

                # Chinese article title
                title = ''             
                try:
                    title = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1[class='post-body-title']"))).get_attribute('textContent').strip()
                except:
                    continue               
                
                details['articletitle'] = title            
            
                # Chinese article description
                des = ''             
                try:
                    des = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='post-body-content']"))).get_attribute('textContent').replace('  ', '').strip()
                except:
                    continue               
                
                details['articledescription'] = des
                                    
                # English article title
                en_title = ''             
                try:
                    en_title = wait(driver_en, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1[class='post-body-title']"))).get_attribute('textContent').strip()
                except:
                    continue  
                    
                #asian = re.findall(r'[\u3131-\ucb4c]+',en_title)
                #if asian: continue                   
                details['articletitle in English'] = en_title          
            
                # English article description
                en_des = ''             
                try:
                    en_des = wait(driver_en, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='post-body-content']"))).get_attribute('textContent').replace('  ', '').replace('\n\n', '').replace('read more', '').replace('閱讀全文', '').strip()
                except:
                    continue 
                
                #asian = re.findall(r'[\u3131-\ucb4c]+', en_des)
                #if asian: continue               
                details['articledescription in English'] = en_des
                details['articleauthor'] = en_author
                details['articledatetime'] = date            
            
                # English article category
                en_cat = ''             
                try:
                    en_cat = wait(driver_en, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='post-body-sidebar-category d-none d-lg-block']"))).get_attribute('textContent').replace('\n', '').strip()
                except:
                    pass 
            
                #asian = re.findall(r'[\u3131-\ucb4c]+',en_cat)
                #if asian: continue                 
                details['articlecategory'] = en_cat

                # article tags
                tags = ''
                try:
                    div = wait(driver_en, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='post-body-content-tags']")))
                    elems = wait(div, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))
                    for elem in elems:
                        try:
                            text = elem.get_attribute('textContent').strip()
                            if text[-1] == ',':
                                text = text[:-1].strip()
                            tags += text + ', '
                        except:
                            pass
                    tags = tags[:-2]
                except:
                    pass

                #asian = re.findall(r'[\u3131-\ucb4c]+',tags)
                #if asian: continue  

                # other columns
                details['domain'] = 'Hypebeast'
                hype = ''             
                try:
                    hype = wait(driver_en, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[class='hype-count']"))).get_attribute('textContent').replace('\n', '').split()[1].strip()
                except:
                    pass 
                details['hype'] = hype   
                details['articletags'] = tags
                details['articleheader'] = ''

                imgs = ''
                try:
                    pic = wait(driver_en, 2).until(EC.presence_of_element_located((By.TAG_NAME, "picture")))
                    try:
                        img = wait(pic, 2).until(EC.presence_of_element_located((By.TAG_NAME, "img"))).get_attribute('src')
                        imgs += img + ', '
                    except:
                        pass
                    sources = wait(pic, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "source")))
                    for source in sources:
                        try:
                            imgs += source.get_attribute('srcset') + ', '
                        except:
                            pass
                    imgs = imgs[:-2]
                except:
                    pass

                if imgs == '': continue
                details['articleimages'] = imgs
                details['articlecomment'] = ''

                # appending the output to the datafame       
                data = data.append([details.copy()])
                print(f'Scraping the details of article {i+1}\{n}')
                break
            except Exception as err:
                print(f'Warning: the below error occurred while scraping the article: {link}')
                print(str(err))
           
    # output to excel
    if data.shape[0] > 0:
        data['articledatetime'] = pd.to_datetime(data['articledatetime'])
        df1 = pd.read_excel(output1)
        df1 = df1.append(data)   
        df1 = df1.drop_duplicates()
        df1.to_excel(output1, index=False)
    else:
        print('-'*75)
        print('No New Articles Found')
        
def scrape_articles_English(driver, output1, page, month, year):

    print('-'*75)
    print(f'Scraping The Articles Links from: {page}')
    # getting the full posts list
    links = []
    months = {month: index for index, month in enumerate(calendar.month_abbr) if month}
    full_months = {month: index for index, month in enumerate(calendar.month_name) if month}
    prev_month = month - 1
    if prev_month == 0:
        prev_month = 12
    driver.get(page)

    art_time = ''
    # handling lazy loading
    print('-'*75)
    print("Getting the previous month's articles..." )
    for _ in range(50):
        try:
            height1 = driver.execute_script("return document.body.scrollHeight")
            driver.execute_script(f"window.scrollTo(0, {height1})")
            time.sleep(4)
            try:
                button = wait(driver, 1).until(EC.presence_of_element_located((By.XPATH, "//a[@id='load-more-btn']")))
                driver.execute_script("arguments[0].click();", button)
                time.sleep(1)
            except:
                pass
            try:
                art_time = wait(driver, 3).until(EC.presence_of_all_elements_located((By.TAG_NAME, "time")))[-1].get_attribute('textContent').strip()
            except:
                break
            try:
                art_month = months[art_time.split()[0]]
                art_year = int(art_time.split()[-1])
                # for articles from previous year
                if art_year < year and prev_month != 12:
                    break
                # for all months except Jan
                if art_month < prev_month and prev_month != 12 and art_year == year:
                    break
                # for Jan
                elif art_month < prev_month and prev_month == 12 and art_year < year:
                    break
            except:
                pass

            height2 = driver.execute_script("return document.body.scrollHeight")
            if height1 == height2: 
                break
        except Exception as err:
            break

    # scraping posts urls 
    try:
        posts = wait(driver, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class='post-box-content-container']")))    
    except:
        print('No posts are available')
        return

    for post in posts:
        try:
            date = wait(post, 2).until(EC.presence_of_element_located((By.TAG_NAME, "time"))).get_attribute('textContent') 
            art_month = int(months[date.split()[0]])
            art_year = int(date.split()[-1])
            if art_month != prev_month: continue
            if art_year < year and prev_month != 12: continue
            link = wait(post, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[class='title']"))).get_attribute('href')
            if link not in links:
                links.append(link)
        except:
            pass

    # scraping posts details
    print('-'*75)
    print('Scraping Articles Details...')
    print('-'*75)

    # reading previously scraped data for duplication checking
    scraped = []
    try:
        df = pd.read_excel(output1)
        scraped = df['unique_id'].values.tolist()
    except:
        pass

    n = len(links)
    data = pd.DataFrame()
    for i, link in enumerate(links):
        for _ in range(2):
            try:
                try:
                    driver.get(link)   
                except:
                    print(f'Warning: Failed to load the url: {link}')
                    continue

                art_id = ''
                try:
                    text = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class*='post-container post-status-publish ']"))).get_attribute('id').split('-')[-1].strip()
                    art_id = int(re.findall("\d+", text)[0])
                except:
                    pass
                if art_id in scraped or art_id == '': 
                    print(f'Article {i+1}\{n} is already scraped, skipping.')
                    break

                details = {}

                # English article author and date
                en_author, date = '', ''             
                try:
                    en_author = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[class='author-name']"))).get_attribute('textContent').split(':')[-1].strip()
                    date = wait(driver, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "time")))[1].get_attribute('textContent').strip()
                except Exception as err:
                    pass
            
                # checking if the article date is correct
                try:
                    if 'ago' in date:
                        art_month = datetime.now().month
                        art_year = datetime.now().year
                        art_day = datetime.now().day
                    else:
                        try:
                            art_month = int(full_months[date.split()[0]])
                        except:
                            art_month = int(months[date.split()[0]])
                        art_year = int(date.split()[-1])  
                        art_day = int(date.split()[1].replace(',', ''))       
                    if art_month != prev_month: 
                        print(f'skipping article with date {art_month}/{art_day}/{art_year}')
                        break
                    date = f'{art_month}/{art_day}/{art_year}'
                except:
                    #print(f'Warning: Failed to extract the date for link: {link}')
                    continue    

                details['sku'] = art_id
                details['unique_id'] = art_id
                details['articleurl'] = link

                # Chinese article title
                title = ''             
                try:
                    title = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1[class='post-body-title']"))).get_attribute('textContent').strip()
                except:
                    continue               
                
                details['articletitle'] = title            
            
                # Chinese article description
                des = ''             
                try:
                    des = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='post-body-content']"))).get_attribute('textContent').replace('  ', '').strip()
                except:
                    continue               
                
                details['articledescription'] = des
                details['articletitle in English'] = title          
                details['articledescription in English'] = des
                details['articleauthor'] = en_author
                details['articledatetime'] = date            
            
                # English article category
                en_cat = ''             
                try:
                    en_cat = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='post-body-sidebar-category d-none d-lg-block']"))).get_attribute('textContent').replace('\n', '').strip()
                except:
                    pass 
                 
                details['articlecategory'] = en_cat

                # article tags
                tags = ''
                try:
                    div = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='post-body-content-tags']")))
                    elems = wait(div, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))
                    for elem in elems:
                        try:
                            text = elem.get_attribute('textContent').strip()
                            if text[-1] == ',':
                                text = text[:-1].strip()
                            tags += text + ', '
                        except:
                            pass
                    tags = tags[:-2]
                except:
                    pass

                # other columns
                details['domain'] = 'Hypebeast'
                hype = ''             
                try:
                    hype = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[class='hype-count']"))).get_attribute('textContent').replace('\n', '').split()[1].strip()
                except:
                    pass 
                details['hype'] = hype   
                details['articletags'] = tags
                details['articleheader'] = ''

                imgs = ''
                try:
                    pic = wait(driver, 2).until(EC.presence_of_element_located((By.TAG_NAME, "picture")))
                    try:
                        img = wait(pic, 2).until(EC.presence_of_element_located((By.TAG_NAME, "img"))).get_attribute('src')
                        imgs += img + ', '
                    except:
                        pass
                    sources = wait(pic, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "source")))
                    for source in sources:
                        try:
                            imgs += source.get_attribute('srcset') + ', '
                        except:
                            pass
                    imgs = imgs[:-2]
                except:
                    pass

                if imgs == '': continue
                details['articleimages'] = imgs
                details['articlecomment'] = ''

                # appending the output to the datafame       
                data = data.append([details.copy()])
                print(f'Scraping the details of article {i+1}\{n}')
                break
            except Exception as err:
                print(f'Warning: the below error occurred while scraping the article: {link}')
                print(str(err))
           
    # output to excel
    if data.shape[0] > 0:
        data['articledatetime'] = pd.to_datetime(data['articledatetime'])
        df1 = pd.read_excel(output1)
        df1 = df1.append(data)   
        df1 = df1.drop_duplicates()
        df1.to_excel(output1, index=False)
    else:
        print('-'*75)
        print('No New Articles Found')
 
def get_inputs():

    # assuming the inputs to be in the same script directory
    path = os.getcwd()
    if '\\' in path:
        path += '\\Hypebeast_settings.xlsx'
    else:
        path += '/Hypebeast_settings.xlsx'

    if not os.path.isfile(path):
        print('Error: Missing the settings file "Hypebeast_settings.xlsx"')
        input('Press any key to exit')
        sys.exit(1)
    try:
        settings = {}
        df = pd.read_excel(path)
        cols = df.columns
        settings[cols[0]] = int(cols[1])
    except:
        print('Error: Failed to process the settings sheet')
        input('Press any key to exit')
        sys.exit(1)

    # checking the settings dictionary
    keys = ["Number of Posts"]
    for key in keys:
        if key not in settings.keys():
            print(f"Warning: the setting '{key}' is not present in the settings file")
            settings[key] = 3000

    if settings["Number of Posts"] < 1:
        settings[key] = 3000

    return settings

def initialize_output():

    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")
    path = os.getcwd() + '\\Scraped_Data\\' + stamp
    if os.path.exists(path):
        #os.remove(path) 
        shutil.rmtree(path)
    os.makedirs(path)

    file1 = f'Hypebeast_{stamp}.xlsx'

    # Windws and Linux slashes
    if os.getcwd().find('/') != -1:
        output1 = path.replace('\\', '/') + "/" + file1
    else:
        output1 = path + "\\" + file1  

    # Create an new Excel file and add a worksheet.
    workbook1 = xlsxwriter.Workbook(output1)
    workbook1.add_worksheet()
    workbook1.close()    

    return output1

def main():

    print('Initializing The Bot ...')
    freeze_support()
    start = time.time()
    output1 = initialize_output()
    urls = ["https://hypebeast.com/zh/travel", "https://hypebeast.com/zh/latest","https://hypebeast.com/zh", "https://hypebeast.com/zh/footwear", "https://hypebeast.com/zh/design", "https://hypebeast.com/zh/tech", "https://hypebeast.com/zh/automotive", "https://hypebeast.com/zh/food-beverage", "https://hypebeast.com/zh/fashion", "https://hypebeast.com/zh/footwear", "https://hypebeast.com/zh/entertainment"]

    English_urls = ["https://hypebeast.com/fashion", "https://hypebeast.com/footwear", "https://hypebeast.com/", "https://hypebeast.com/watches", "https://hypebeast.com/videos", "https://hypebeast.com/footwear", "https://hypebeast.com/art", "https://hypebeast.com/design", "https://hypebeast.com/tech", "https://hypebeast.com/automotive", "https://hypebeast.com/food-beverage"]
    month = datetime.now().month
    year = datetime.now().year
    try:
        driver = initialize_bot(False)
        driver_en = initialize_bot(True)
    except Exception as err:
        print('Failed to initialize the Chrome driver due to the following error:\n')
        print(str(err))
        print('-'*75)
        input('Press any key to exit.')
        sys.exit()

    print('-'*75)
    print('Getting the brands links from: https://hypebeast.com/brands')   
    # getting brands urls
    homepages = []
    try:
        driver.get("https://hypebeast.com/brands")
        brands = wait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[class='brand']")))
        for brand in brands:
            try:
                link = brand.get_attribute('href')
                if link not in homepages:
                    homepages.append(link)
            except:
                pass
    except:
        print('Warning: Failed to get brands links, skipping...')

    # English pages
    homepages += English_urls
    for page in homepages:
        try:
            scrape_articles_English(driver, output1, page, month, year)
        except Exception as err: 
            print(f'Warning: the below error occurred:\n {err}')
            driver.quit()
            time.sleep(5)
            driver = initialize_bot(False)

    driver.quit()
    driver = initialize_bot(False)
    # Non English pages
    for page in urls:
        try:
            scrape_articles(driver, driver_en, output1, page, month, year)
        except Exception as err: 
            print(f'Warning: the below error occurred:\n {err}')
            driver.quit()
            driver_en.quit()
            time.sleep(5)
            driver = initialize_bot(False)
            driver_en = initialize_bot(True) 

    driver.quit()
    driver_en.quit()
    print('-'*75)
    elapsed_time = round(((time.time() - start)/60), 2)
    input(f'Process is completed in {elapsed_time} mins, Press any key to exit.')

if __name__ == '__main__':

    main()