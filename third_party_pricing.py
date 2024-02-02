from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from fuzzywuzzy import process
import time

chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=chrome_options)


def grainger_search(input):
    driver.get("https://www.grainger.com/")
    search_input = driver.find_element(By.NAME, value="searchQuery")
    search_input.clear()
    item_number = input
    search_input.send_keys(item_number)
    search_input.send_keys(Keys.RETURN)
    time.sleep(10)
    vendor_part = driver.find_element(By.XPATH, value='(//dd[contains(@class, "rOM8HV") and contains(@class, "hRRBwT")])[2]')
    item_desc = driver.find_element(By.CLASS_NAME, value='lypQpT')
    price = driver.find_element(By.XPATH, '//*[@class="rbqU0E lVwVq5"]')
    return vendor_part.text, item_desc.text, price.text

def zoro_search(man_part, grainger_desc):
    driver.get(f"https://www.zoro.com/search?q={man_part}")
    time.sleep(3)
    grid_list =[]
    if len(driver.current_url) < 50:
        button = driver.find_element(By.XPATH, '//*[@aria-label="Display results as Grid"]')
        button.click()
        time.sleep(3)
        element = driver.find_element(By.XPATH, '//*[@aria-label="Search Results"]')
        item_list =[]
        for i in range(1,4):
            try:
                i = str(i)
                time.sleep(3)
                desc = element.find_element(By.XPATH,value=f'//*[@id="search-container"]/div[1]/div[2]/section[1]/div/section/div[{i}]/div[3]')
                desc = desc.text
                item_list.append(desc)
            except NoSuchElementException:
                print("No More Items")
                break
        desc_list = []
        for items in item_list:
            items = items.split('\n')
            desc_list.append(items[1])
        print(desc_list)
        best_match, score = process.extractOne(grainger_desc, desc_list)
        match_xpath = f"//*[text()='{best_match}']"
        if score > 80:
            search_link = driver.find_element(By.XPATH, value=match_xpath)
            search_link.click()
            time.sleep(5)
            price = driver.find_element(By.XPATH, value='//*[@id="v-app"]/div/div/div/div[1]/div[2]/div[4]/div/div/div/div')
            price = price.text
            print(price)
        else:
            print("No Similar Items")
    print("Search Ended")
def nasco_login():
    username_element = driver.find_element(By.NAME, 'usr_name')
    username_element.send_keys("701792.Gabriel")
    password_element = driver.find_element(By.NAME, 'usr_password')
    password_element.send_keys("Happy!2023")
    password_element.send_keys(Keys.RETURN)
    login_buttom = driver.find_element(By.ID, 'searchGoButton')
    login_buttom.click()

def nasco_search(item_description, brand):
    url = 'https://www.orsnasco.com/storefrontCommerce/home.do'
    driver.get(url)
    time.sleep(3)
    remove_popup = driver.find_element(By.XPATH, '//*[@id="modal-1"]/div/div/div[1]/button')
    remove_popup.click()
    time.sleep(3)
    nasco_login()
    search_input = driver.find_element(By.ID, 'searchInput')
    search_input.send_keys(item_description)
    search_input.send_keys(Keys.RETURN)
    time.sleep(5)
    detailed_view = driver.find_element(By.ID, 'toggle_link_detailed')
    detailed_view.click()
    # Locate the table with id 'card-table'
    card_table = driver.find_element(By.ID, 'card-table')
    brand_name = brand
    try:
        try:
            if card_table:
                table_text = card_table.text
                rows = table_text.split('\n')
                items_to_exclude = ['ADD TO CART', '  Special pricing only while supplies last']
                filtered_list = [item for item in rows if item not in items_to_exclude]
                keys = ['description', 'brand', 'list_price', 'price', 'unit_price', 'pack', 'part_number', 'available']
                item_dicts = [dict(zip(keys, filtered_list[i:i + len(keys)])) for i in range(0, len(filtered_list), len(keys))]
                filtered_item_descriptions = [item_dict['description'] for item_dict in item_dicts if item_dict['brand'] == brand_name]
                input_string = item_description
                best_match, score = process.extractOne(input_string, filtered_item_descriptions)
                if score > 80:
                    print("Best Match:", best_match)
                    print("Matching Score:", score)
                    match_xpath = f"//*[text()='{best_match}']"
                    item_link = driver.find_element(By.XPATH, match_xpath)
                    item_link.click()
                    time.sleep(3)
                    unit_price = driver.find_element(By.XPATH, '//*[@id="itemDetailContainer"]/div[2]/table/tbody/tr[9]/td[2]')
                    unit_price = unit_price.text
                    print("Unit Price: ", unit_price)
        except:
                print("No significant matches")
    except:
        print("Table with id 'card-table' not found.")

input = '792YD2'
print(grainger_search(input))
manu_part, item_desc, price = grainger_search(input)
zoro_search(manu_part, item_desc)
nasco_search(item_desc, brand="DeWalt")
driver.close()
