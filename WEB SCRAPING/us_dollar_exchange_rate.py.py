import pandas as pd 
import requests
from bs4 import BeautifulSoup

url = "https://www.x-rates.com/historical/?from=USD&amount=1&date=2024-06-26"
r = requests.get(url)
print(r)

headers = {'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.135 Safari/537.36 Edge/12.246"}
# Here the user agent is for Edge browser on windows 10. You can find your browser user agent from the above given link.

r = requests.get(url=url, headers=headers)

response = requests.get(url)
if response.status_code == 200:
    soup = BeautifulSoup(response.content, 'html.parser')
    cls=soup.find('span',attrs={'class':'ratesTimestamp'}).text

    table = soup.find('table', {'class': 'tablesorter ratesTable'})
    if table:
        headers = ["currency_fullname","exchange_rate","inv_exchange_rate"]
        print("Headers:", headers) 
        rows = table.find_all('tr')
        data = []
        for row in rows[1:]:
            cols = row.find_all('td')
            if len(cols) >= 3:
                currency_fullname = cols[0].text.strip()
                exchange_rate = cols[1].text.strip() 
                inv_exchange_rate = cols[2].text.strip()
                data.append([currency_fullname, exchange_rate, inv_exchange_rate])
                print("Row data:", [currency_fullname, exchange_rate ,inv_exchange_rate])

        df = pd.DataFrame(data, columns=headers)

        cls=soup.find('span',attrs={'class':'ratesTimestamp'})
        df['date'] = cls.text
        breakpoint()
        df.to_csv(r'C:\Users\komalkumari.b\Downloads\scraped_data_US_dollar.csv', index=False)
        print("Data saved to scraped_data.csv")
    else:
        print("No table found on the page")
else:
    print(f"Failed to retrieve the webpage. Status code: {response.status_code}")

