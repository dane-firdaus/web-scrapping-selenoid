from bs4 import BeautifulSoup
import pandas as pd

def parse_html_file(file_path):
    # Initialize the list to hold data
    haji = []

    # Open and read the HTML file
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # Parse the HTML content with BeautifulSoup
    soup = BeautifulSoup(content, 'html.parser')
    
    # Find the table with ID 'example1'
    table = soup.find('table', id='example1')
    
    if not table:
        raise LookupError("The table with ID 'example1' is not found")
    
    # Iterate over each row in the table (skipping the header row)
    for row in table.find_all('tr')[1:]:
        cols = row.find_all('td')
        if len(cols) < 6:
            continue

        number = cols[0].text.strip()
        company_info = cols[1]
        company_name = company_info.find('a').text.strip()
        company_link = company_info.find('a')['href']
        status_daftar_hitam = company_info.find('span').text.strip()
        
        sk_info = cols[2].find('small').text.strip()
        alamat = cols[3].find('small').text.strip()
        status = cols[4].text.strip()
        detail_link = cols[5].find('a')['href']
    
        haji.append(
            {
                'No': number,
                'Company Name': company_name,
                'Company Link': company_link,
                'Status Daftar Hitam': status_daftar_hitam,
                'SK Info': sk_info,
                'Alamat': alamat,
                'Status': status,
                'Detail Link': detail_link
            }
        )
    
    return haji

if __name__ == '__main__':
    # Define the path to the HTML file
    file_path = 'dataHtml.html'
    
    data = parse_html_file(file_path)
    
    # Convert the data to a pandas DataFrame
    df = pd.DataFrame(data)
    
    # Save the DataFrame to an Excel file
    df.to_excel('scrapp_data_travel_haji.xlsx', index=False, engine='openpyxl')
    
    print("Data has been successfully parsed and saved to scrapp_data_travel_haji.xlsx")
