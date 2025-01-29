import time, json
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

path = r'D:\WechatAI\WebContent\PDF'  # Path of PDF saved

# chrome option，call the printer，and save as pdf
chrome_options = webdriver.ChromeOptions()
# Printer Setting
settings = {"recentDestinations": [{"id": "Save as PDF",
                                    "origin": "local",
                                    "account": ""
                                    }],
            "selectedDestinationId": "Save as PDF",
            "version": 2,
            "isHeaderFooterEnabled": True,
            "isLandscapeEnabled": False,  # landscape=horizontal，portrait=vertical，default=portrait
            "isCssBackgroundEnabled": True,
            "mediaSize": {"height_microns": 297000,
                          "name": "ISO_A4",
                          "width_microns": 210000,
                          "custom_display_name": "A4 210 x 297 mm"
                          },
            }

#chrome_options.add_argument('--headless') # headless = window invisible
chrome_options.add_argument('--enable-print-browser')
chrome_options.add_argument('--kiosk-printing')  # silent printing, user does not need to press the printing bottom

prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
         'savefile.default_directory': path  # saving path
         }
chrome_options.add_experimental_option('prefs', prefs)


#  Print function
def web_print_save_pdf(url, filename):
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    driver.get(url)
    driver.maximize_window()
    time.sleep(3)
    driver.execute_script('document.title="{}";window.print();'.format(
        filename))  # file name
    driver.close()


# Call the function
df = pd.read_excel('Article Summary.xlsx')
titles = df.iloc[:, 0] # read the title column
urls = df.iloc[:, 2] # read the link column
count = 1
for url, title in zip(urls, titles):
    print('Processing {}'.format(count))
    web_print_save_pdf(url, title)
    count += 1
print('Finished')
