from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# Loading selenium
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 5)

# Function to open Google Maps, input location, and submit
def search_location(location):
    search_box = driver.find_element(By.XPATH, '//*[@id="searchboxinput"]')
    search_box.clear()
    search_box.send_keys(location)
    search_box.send_keys(Keys.RETURN)
    sleep(2)

# Open Google Maps
url = "https://www.google.com/maps"
driver.get(url)
sleep(2)

# Manual input for location
location_input = input("Masukkan lokasi yang ingin dicari di Google Maps: ")
search_location(location_input)

def ButtonReview():
    ButtonReview= driver.find_element(By.CLASS_NAME, "wNNZR fontTitleSmall")
    ButtonReview.click()
    sleep(2)

try:
    # Click on all reviews
    element = driver.find_element(By.XPATH, '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[3]/div/div/button[2]/div[2]')
    print("Element ditemukan")
    try:
        element.click()
        sleep(5)
    except:
        print("Not Loaded")
        raise Exception("Failed to load reviews")
except:
    print("Element is not present")
    raise Exception("Failed to find the reviews button")
finally:
    reviews = driver.find_elements(By.CLASS_NAME, 'jftiEf')
    sleep(2)

    min_r = 0
    max_reviews = float('inf')

    # List untuk menyimpan data ulasan
    data_reviews = []

    try:
        while min_r < max_reviews:
            # Memperbarui daftar ulasan setelah melakukan scroll
            reviews = driver.find_elements(By.CLASS_NAME, 'jftiEf')

            # Iterasi melalui ulasan yang baru ditambahkan
            for i in range(min_r, len(reviews)):
                review = reviews[i]
                try:
                    moreButton = review.find_element(By.CLASS_NAME, 'w8nwRe')
                    moreButton.click()
                    sleep(2)
                except:
                    print("")

                # Menemukan elemen yang menyimpan rating bintang
                star_rating_element = review.find_element(By.CLASS_NAME, 'kvMYJc')
                star_rating = star_rating_element.get_attribute('aria-label')

                # Menemukan elemen yang menyimpan teks ulasan (jika ada)
                try:
                    text_class = review.find_element(By.CLASS_NAME, 'wiI7pd')
                    text = text_class.text
                except:
                    text = ""

                # Menambahkan data ulasan ke dalam list
                data_reviews.append({'Star Rating': star_rating, 'Review': text})

                print("\n\n------------------", i, "--------------\n\n", "Star Rating:", star_rating, "\n Ulasan:", text)
                driver.execute_script('arguments[0].scrollIntoView(true);', reviews[i])
                sleep(2)

                min_r += 1

    finally:
        # Simpan data ke dalam dataframe pandas
        df_reviews = pd.DataFrame(data_reviews)

        # Simpan dataframe ke dalam file Excel
        excel_file = "Reviews.xlsx"
        df_reviews.to_excel(excel_file, index=False)

        # Tutup browser setelah selesai
        driver.quit()
