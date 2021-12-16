from selenium import webdriver
from selenium.webdriver.chrome import options
import undetected_chromedriver.v2 as uc,webbrowser
import undetected_chromedriver as uc
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
import logging, traceback
from configparser import ConfigParser
from time import sleep
import os.path
import xlsxwriter
import warnings

warnings.filterwarnings("ignore", category=DeprecationWarning) 


config = ConfigParser()
config.read('config.ini')
class Request:
    def __init__(self):
        self.database_option = config['Default']['ignore_database']
        self.database_path = config['Path']['database_file_path']
        self.row_counter = 1
        self.letter_list = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U']
        logging.basicConfig(
            filename=f"{self.database_path}/logfile.txt",
            format="%(asctime)s %(levelname)s -  %(message)s",
            filemode="w",
            level=logging.INFO
        )
        self.logger = logging.getLogger()
    def save_product(self,browser,tr_counter):
        controller = browser.find_element_by_xpath(f'/html/body/section/div/div[2]/section/div/div[2]/div[1]/div[4]/table/tbody/tr[{tr_counter}]')
           
        with open(self.database_path+"/part1database.txt", "a") as f:
            f.write(controller.get_attribute("data-pb-id") +"\n")
            f.close

    def scrapper(self,browser,tr_counter,row,sheet,table_cell_format,label_format,row_counter):
        self.row_counter = row_counter
        self.row_counter += 1
        self.logger.info('Product Number: '+str(row))
        sheet.write(f'A{self.row_counter}', 'Image',label_format)
        sheet.write(f'B{self.row_counter}', 'Name',label_format)
        sheet.write(f'C{self.row_counter}', 'Core Count',label_format)
        sheet.write(f'D{self.row_counter}', 'Performance Core Clock',label_format)
        sheet.write(f'E{self.row_counter}', 'Performance Boost Clock',label_format)
        sheet.write(f'F{self.row_counter}', 'TDP',label_format)
        sheet.write(f'G{self.row_counter}', 'Integrared Graphics',label_format)
        sheet.write(f'H{self.row_counter}', 'SMT',label_format)
        sheet.write(f'I{self.row_counter}', 'Rating',label_format)
        sheet.write(f'J{self.row_counter}', 'Price',label_format) 
        self.row_counter += 1                       
        image = browser.find_element_by_xpath(f"/html/body/section/div/div[2]/section/div/div[2]/div[1]/div[4]/table/tbody/tr[{tr_counter}]/td[2]/a/div[1]/div/img").get_attribute("src")
        sheet.write(f'A{self.row_counter}', image)
        name = browser.find_element_by_xpath(f"//*[@id='category_content']/tr[{tr_counter}]/td[2]/a/div[2]")
        sheet.write(f'B{self.row_counter}', name.text)
        coreCounts = browser.find_element_by_xpath(f"//*[@id='category_content']/tr[{tr_counter}]/td[3]")
        coreCount = browser.execute_script('return arguments[0].lastChild.textContent;', coreCounts)
        sheet.write(f'C{self.row_counter}', coreCount)
        performanceCores = browser.find_element_by_xpath(f"//*[@id='category_content']/tr[{tr_counter}]/td[4]")
        performanceCore = browser.execute_script('return arguments[0].lastChild.textContent;', performanceCores)
        sheet.write(f'D{self.row_counter}', performanceCore)
        performanceBoosts = browser.find_element_by_xpath(f"//*[@id='category_content']/tr[{tr_counter}]/td[5]")
        performanceBoost = browser.execute_script('return arguments[0].lastChild.textContent;', performanceBoosts)
        if (performanceBoost == "Performance Boost Clock"):
            sheet.write(f'E{self.row_counter}', "-")
        else:
            sheet.write(f'E{self.row_counter}', performanceBoost)
        Tdps = browser.find_element_by_xpath(f"//*[@id='category_content']/tr[{tr_counter}]/td[6]")
        Tdp = browser.execute_script('return arguments[0].lastChild.textContent;', Tdps)
        sheet.write(f'F{self.row_counter}', Tdp)
        integratedGraphics = browser.find_element_by_xpath(f"//*[@id='category_content']/tr[{tr_counter}]/td[7]")
        integratedGraphic = browser.execute_script('return arguments[0].lastChild.textContent;', integratedGraphics)
        sheet.write(f'G{self.row_counter}', integratedGraphic)
        smts = browser.find_element_by_xpath(f"//*[@id='category_content']/tr[{tr_counter}]/td[8]")
        smt = browser.execute_script('return arguments[0].lastChild.textContent;', smts)
        sheet.write(f'H{self.row_counter}', smt)     
        ratings = browser.find_element_by_xpath(f"//*[@id='category_content']/tr[{tr_counter}]/td[9]")
        rating = str(ratings.get_attribute("innerHTML")).count('shape-star-full')-3
        rating_count = browser.execute_script('return arguments[0].lastChild.textContent;', ratings)
        sheet.write(f'I{self.row_counter}', f'Rating: {rating}, Rating Count: {rating_count}')
        prices = browser.find_element_by_xpath(f"//*[@id='category_content']/tr[{tr_counter}]/td[10]")
        price = browser.execute_script('return arguments[0].firstChild.textContent;', prices)
        if (price == "Add"):
            sheet.write(f'J{self.row_counter}', "-")
        else:
            sheet.write(f'J{self.row_counter}', price)                        
        name.click()
        sleep(15)
        #images
        image_checker = False
        try:
            browser.find_element_by_class_name('gallery__imageWrapper')
            image_checker = True
        except:
            image_checker = False
        if image_checker == True:
            extra_images = browser.find_elements_by_xpath('//*[@class="gallery__imageWrapper"]/img')
            for extra_image in extra_images:
                self.row_counter += 1
                sheet.write(f'A{self.row_counter}', extra_image.get_attribute("src"))
        self.row_counter += 1
        #prices
        price_checker = False
        global_price_checker = False
        self.row_counter +=1
        sheet.merge_range(f'A{self.row_counter}:F{self.row_counter}','Prices',label_format)
        self.row_counter +=1
        sheet.write(f'A{self.row_counter}','Merhant',table_cell_format)
        sheet.write(f'B{self.row_counter}','Base',table_cell_format)
        sheet.write(f'C{self.row_counter}','Promo',table_cell_format)
        sheet.write(f'D{self.row_counter}','Shipping',table_cell_format)
        sheet.write(f'E{self.row_counter}','Availability',table_cell_format)
        sheet.write(f'F{self.row_counter}','Total',table_cell_format)
        for price_counter in range(1,15):
            try:
                browser.find_element_by_xpath(f'//*[@id="prices"]/table/tbody/tr[{price_counter}]/td[1]/a/img')
                price_checker = True
                global_price_checker = True
            except (Exception):
                price_checker = False
            if (price_checker == True):
                self.row_counter += 1
                merchant = browser.find_element_by_xpath(f'//*[@id="prices"]/table/tbody/tr[{price_counter}]/td[1]/a/img').get_attribute('alt')
                base_price = browser.find_element_by_xpath(f'//*[@id="prices"]/table/tbody/tr[{price_counter}]/td[2]').text
                promo = browser.find_element_by_xpath(f'//*[@id="prices"]/table/tbody/tr[{price_counter}]/td[3]').text
                shipping = browser.find_element_by_xpath(f'//*[@id="prices"]/table/tbody/tr[{price_counter}]/td[4]').text
                availability = browser.find_element_by_xpath(f'//*[@id="prices"]/table/tbody/tr[{price_counter}]/td[6]').text
                total_price = browser.find_element_by_xpath(f'//*[@id="prices"]/table/tbody/tr[{price_counter}]/td[7]/a').text
                sheet.write(self.row_counter-1,0,merchant)
                sheet.write(self.row_counter-1,1,base_price)
                sheet.write(self.row_counter-1,2,promo)
                sheet.write(self.row_counter-1,3,shipping)
                sheet.write(self.row_counter-1,4,availability)
                sheet.write(self.row_counter-1,5,total_price)
        #spesification
        self.row_counter += 1
        sheet.merge_range(f'A{self.row_counter}:F{self.row_counter}','Spesifications',label_format)
        self.row_counter += 1
        spesification_titles = browser.find_elements_by_class_name("group__title")
        spesification_counter = 0
        for spesification_title in spesification_titles:
            sheet.write(f'{self.letter_list[spesification_counter]}{self.row_counter}',spesification_title.text,table_cell_format)
            spesification_counter += 1
            if (spesification_counter == 20):
                break
        spesification_counter = 0
        spesification_datas = browser.find_elements_by_class_name("group__content")
        for spesification_data in spesification_datas:
            sheet.write(f'{self.letter_list[spesification_counter]}{self.row_counter+1}',spesification_data.text)
            spesification_counter += 1
            if (spesification_counter == 20):
                break
        self.row_counter += 2
        #reviews
        sheet.merge_range(f'A{self.row_counter}:F{self.row_counter}','Reviews',label_format)
        self.row_counter += 1
        reviews_checker = False
        reviews_section_controller = 2
        try:
            browser.find_element_by_class_name('partReviews__review')
            reviews_checker = True
        except:
            reviews_checker = False
        if (reviews_checker == True):
            reviews = browser.find_elements_by_class_name('partReviews__writeup')
            review_counter = 1
            try:
                browser.find_element_by_xpath(f'//*[@id="product-page"]/section/div[2]/section[2]/div/div[2]/div[6]/div[{review_counter+1}]/div[2]/div/ul')
                reviews_section_controller = 2
            except:
                reviews_section_controller = 1
            for review in reviews:
                if(global_price_checker == True):
                    raw_rating = browser.find_element_by_xpath(f'//*[@id="product-page"]/section/div[2]/section[{reviews_section_controller}]/div/div[2]/div[6]/div[{review_counter+1}]/div[2]/div/ul')
                    rating = str(raw_rating.get_attribute("innerHTML")).count('shape-star-full')-5
                    sheet.write(self.row_counter-1,0,f'Review Rate: {rating}')
                else:        
                    raw_rating = browser.find_element_by_xpath(f'//*[@id="product-page"]/section/div[2]/section[{reviews_section_controller}]/div/div[2]/div[5]/div[{review_counter+1}]/div[2]/div/ul')
                    rating = str(raw_rating.get_attribute("innerHTML")).count('shape-star-full')-5
                    sheet.write(self.row_counter-1,0,f'Review Rate: {rating}')
                self.row_counter += 1
                sheet.merge_range(f'A{self.row_counter}:F{self.row_counter}',review.text)
                self.row_counter += 2
                review_counter += 1
                sheet.merge_range(f'A{self.row_counter}:F{self.row_counter}',f'Done - {row}',label_format)                           
        browser.back()
        sleep(15)
        self.save_product(browser,tr_counter) 
        with open(self.database_path+"/lastrow.txt", "w") as f:
            f.write(f"{self.row_counter}")
            f.close


    def run_script(self):          
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--incognito')
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("user-agent=naber-selam-heyy")

        excel = xlsxwriter.Workbook(self.database_path+"/informations.xlsx")
        sheet = excel.add_worksheet("part1")
        #column width
        sheet.set_column(1,30,20)
        #labels formats
        label_format = excel.add_format()
        label_format.set_font_size(14)
        label_format.set_bold()
        label_format.set_bg_color('#58aa9e')
        #table formats
        table_cell_format = excel.add_format()
        table_cell_format.set_bold()
        table_cell_format.set_bg_color('#C2EDDA')
        browser = uc.Chrome(executable_path="http://3.227.145.50:4444",options=chrome_options)
        browser.get("https://pcpartpicker.com/products/cpu")
        browser.set_window_size(1920, 1080)
        sleep(10)
        limit = browser.find_element_by_xpath('/html/body/section/div/div[2]/section/div/div[2]/div[2]/section/ul/li[7]/a').text
        self.logger.info('limit:'+str(limit))
        #create excel if not exits
        xlsxwriter.Workbook(self.database_path+"/informations.xlsx").close()
        if (self.database_option == "False"):
            if (os.path.exists(self.database_path+"/part1database.txt")):
                for page in range(1,int(limit)+1):
                    try:
                        pagination = browser.find_element_by_xpath(f"//a[contains(@href,'#page={page}')]")
                        pagination.click() 
                        sleep(15)
                        self.logger.info("Pagination:"+str(page))  
                        self.logger.info('Page Ready!')                                                  
                        rows = browser.find_elements_by_class_name("tr__product")
                        tr_counter = 0
                        for row in range(1,len(rows)):
                            tr_counter += 1
                            controller = browser.find_element_by_xpath(f'/html/body/section/div/div[2]/section/div/div[2]/div[1]/div[4]/table/tbody/tr[{tr_counter}]')
                            with open(self.database_path+"/part1database.txt") as f:
                                with open(self.database_path+"/lastrow.txt") as a:
                                    row_count = int(a.read())
                                    a.close()
                                if str(controller.get_attribute("data-pb-id")) in f.read():
                                    self.logger.info(str(controller.get_attribute("data-pb-id"))+" is exists.") 
                                else:
                                    self.logger.info(str(controller.get_attribute("data-pb-id"))+" is not exists.") 
                                    self.scrapper(browser,tr_counter,row,sheet,table_cell_format,label_format,row_count)                                                                    
                    except (TimeoutException, WebDriverException):
                        self.logger.error(traceback.format_exc())
                        sleep(6)
                        return
            else:
                with open(self.database_path+"/lastrow.txt", "w") as f:
                    f.write("1")
                    f.close
                for page in range(1,int(limit)+1): 
                    try:
                        pagination = browser.find_element_by_xpath(f"//a[contains(@href,'#page={page}')]")
                        pagination.click() 
                        sleep(15)
                        self.logger.info("Pagination:"+str(page)) 
                        self.logger.info('Page Ready!')                                                  
                        rows = browser.find_elements_by_class_name("tr__product")
                        tr_counter = 0
                        for row in range(1,len(rows)):
                            tr_counter += 1
                            controller = browser.find_element_by_xpath(f'/html/body/section/div/div[2]/section/div/div[2]/div[1]/div[4]/table/tbody/tr[{tr_counter}]')                          
                            with open(self.database_path+"/lastrow.txt") as a:
                                row_count = int(a.read())
                                a.close()
                            self.scrapper(browser,tr_counter,row,sheet,table_cell_format,label_format,row_count)                                                                    
                    except (TimeoutException, WebDriverException):
                        self.logger.error(traceback.format_exc())
                        sleep(6)
                        return
        else:
            if (os.path.exists(self.database_path+"/part1database.txt")):
                os.remove(self.database_path+"/part1database.txt")
                os.remove(self.database_path+"/informations.xlsx")
                xlsxwriter.Workbook(self.database_path+"/informations.xlsx").close()
            with open(self.database_path+"/lastrow.txt", "w") as f:
                f.write("1")
                f.close
            for page in range(1,int(limit)+1):
                try:
                    pagination = browser.find_element_by_xpath(f"//a[contains(@href,'#page={page}')]")
                    pagination.click() 
                    sleep(7)
                    self.logger.info("Pagination:"+str(page))      
                    self.logger.info('Page Ready!')                                                  
                    rows = browser.find_elements_by_class_name("tr__product")
                    tr_counter = 0
                    for row in range(1,len(rows)):
                        tr_counter += 1
                        with open(self.database_path+"/lastrow.txt") as a:
                            row_count = int(a.read())
                            a.close()
                        self.scrapper(browser,tr_counter,row,sheet,table_cell_format,label_format,row_count)                                                                    
                except (TimeoutException, WebDriverException):
                    self.logger.error(traceback.format_exc())
                    sleep(6)
                    return
        excel.close()
        browser.close()


Request().run_script()
