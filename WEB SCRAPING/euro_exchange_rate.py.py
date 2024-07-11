import requests
from bs4 import BeautifulSoup
import pandas as pd

url = 'https://www.ecb.europa.eu/stats/policy_and_exchange_rates/euro_reference_exchange_rates/html/index.en.html' 

headers = {'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.135 Safari/537.36 Edge/12.246"}
# Here the user agent is for Edge browser on windows 10. You can find your browser user agent from the above given link.

r = requests.get(url=url, headers=headers)

response = requests.get(url)
if response.status_code == 200:

    soup = BeautifulSoup(response.content, 'html.parser')

    table = soup.find('table', {'class': 'forextable'})
    if table:
        headers = ["currency","currency_fullname","exchange_rate"]
        print("Headers:", headers)  
        rows = table.find_all('tr')
        data = []
        for row in rows[1:]:
            cols = row.find_all('td')
            if len(cols) >= 3: 
                currency = cols[0].text.strip()
                currency_fullname = cols[1].text.strip() 
                exchange_rate = cols[2].text.strip() 
                data.append([currency, currency_fullname, exchange_rate])
                print("Row data:", [currency, currency_fullname ,exchange_rate]) 

        df = pd.DataFrame(data, columns=headers)

        cls=soup.find('div',attrs={'class':'content-box'})
        df['date'] = cls.h3.text
        breakpoint()
        df.to_csv(r'C:\Users\komalkumari.b\Downloads\scraped_data_euro_exchange.csv', index=False)
        print("Data saved to scraped_data.csv")
    else:
        print("No table found on the page")
else:
    print(f"Failed to retrieve the webpage. Status code: {response.status_code}")



