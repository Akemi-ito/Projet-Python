from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import os
import json

class SpaceXScraper:
    def __init__(self, headless=True):
        self.chrome_options = Options()
        if headless:
            self.chrome_options.add_argument("--headless") #ouverture cacher de la page
            self.chrome_options.add_argument("--disable-gpu") # desactiver l'accelerationn gpu

        # Chemin ChromeDriver
        self.chrome_driver_path = "C:/Users/mathi/Downloads/chromedriver-win64/chromedriver-win64/chromedriver.exe"

        # Vérification que le fichier existe
        if not os.path.exists(self.chrome_driver_path):
            raise FileNotFoundError(f"ChromeDriver non trouvé à l'emplacement: {self.chrome_driver_path}")

    def scrape_launches(self):
        try:
            driver = webdriver.Chrome(
                service=Service(self.chrome_driver_path),
                options=self.chrome_options
            )
            url = "https://www.spacex.com/launches/"
            driver.get(url)
            time.sleep(5)

            launch_items = driver.find_elements(By.CSS_SELECTOR, "div.item")
            launch_data = []

            for item in launch_items:
                try:
                    mission_name = item.find_element(By.CSS_SELECTOR, "div.label").text
                    mission_date = item.find_element(By.CSS_SELECTOR, "div.date").text
                    mission_link = item.find_element(By.CSS_SELECTOR, "a.article-header").get_attribute("href")

                    launch_data.append({
                        "mission_name": mission_name,
                        "mission_date": mission_date,
                        "mission_url": mission_link
                    })
                except Exception as e:
                    print(f"Erreur lors du scraping d'un élément: {e}")

            self.save_to_json(launch_data, "spacex_launches.json")
            return launch_data

        except Exception as e:
            raise Exception(f"Erreur lors du scraping: {str(e)}")
        finally:
            if 'driver' in locals():
                driver.quit()

    def save_to_json(self, data, filename):
        """Sauvegarde les données scrapées dans un fichier JSON."""
        try:
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            print(f"Données sauvegardées dans {filename}")
        except Exception as e:
            print(f"Erreur lors de la sauvegarde en JSON: {e}")
