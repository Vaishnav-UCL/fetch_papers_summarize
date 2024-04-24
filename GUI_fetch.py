import tkinter as tk
from tkinter import simpledialog, messagebox
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re

def fetch_papers(scientist_name, years, keywords):
    results = []
    options = Options()
    #options.add_argument('--headless')  # Uncomment for headless mode
    driver = webdriver.Chrome(options=options)

    try:
        driver.get(f"https://scholar.google.com/scholar?q={scientist_name}")
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="gs_res_ccl_mid"]/div[1]/table/tbody/tr/td[2]/h4/a'))).click()
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, 'YEAR'))).click()
        time.sleep(5)

        articles = WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.XPATH, '//*[@id="gsc_a_b"]/tr/td[1]/a')))

        for article in articles:
            driver.execute_script("window.open(arguments[0]);", article.get_attribute('href'))
            driver.switch_to.window(driver.window_handles[1])

            try:
                WebDriverWait(driver, 20).until(lambda d: d.execute_script('return document.readyState') == 'complete')
                title = driver.find_element(By.XPATH, '//*[@id="gsc_oci_title"]/a').text
                pub_info = driver.find_element(By.XPATH, '//*[@id="gsc_oci_table"]/div[2]/div[2]').text

                # Extract year from the publication date string using regex
                publication_year_match = re.search(r'\b(20\d{2})\b', pub_info)
                if publication_year_match:
                    publication_year = int(publication_year_match.group())
                else:
                    print(f"No valid year found for {title}. Skipping...")
                    continue

                abstract = driver.find_element(By.XPATH, '//*[@id="gsc_oci_descr"]').text
                results.append({
                    "Scientist Name": scientist_name,
                    "Title": title,
                    "Publication Date": publication_year,
                    "Abstract": abstract
                })
            finally:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

    finally:
        driver.quit()

    return results

def main():
    root = tk.Tk()
    root.withdraw()

    scientist_name = simpledialog.askstring("Input", "Enter the scientist's name:", parent=root)
    years = simpledialog.askinteger("Input", "Enter the number of years (e.g., 5 for the last 5 years):", parent=root)
    keywords = simpledialog.askstring("Input", "Enter keywords separated by commas (leave blank for all):", parent=root)
    keywords = [keyword.strip().lower() for keyword in keywords.split(',')] if keywords else []

    if scientist_name and years is not None:
        results = fetch_papers(scientist_name, years, keywords)
        if results:
            df = pd.DataFrame(results)
            df.to_excel(f"{scientist_name}_papers.xlsx", index=False)
            messagebox.showinfo("Success", "Data written to Excel.")
        else:
            messagebox.showinfo("No Results", "No matching papers found.")
    else:
        messagebox.showerror("Error", "Scientist name and years are required!")

if __name__ == "__main__":
    main()
