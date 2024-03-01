# You may need to install the pandas library if you haven't already. Just go to your terminal and run the following command:
# pip3 install pandas

import requests
import pandas as pd
import random
import os

### API Functions
def get_card_info(cert_number, access_token):
    card_info = f'https://api.psacard.com/publicapi/cert/GetByCertNumber/{cert_number}'

    headers = {
        "Authorization": f"bearer {access_token}"
    }

    response = requests.get(card_info, headers=headers)

    if response.status_code == 200:
        data = response.json()
        print(f'\nScraping data for Cert Number #{cert_number}')
        return data
    else:
        print(f"Card does not exist. Error: {response.status_code}")
        return None
    
def get_card_images(cert_number, access_token):
    images = f'https://api.psacard.com/publicapi/cert/GetImagesByCertNumber/{cert_number}'

    headers = {
        "Authorization": f"bearer {access_token}"
    }

    response = requests.get(images, headers=headers)

    if response.status_code == 200:
        data = response.json()
        return data
    else:
        # Handle API error
        print(f"Card does not have available images. Error: {response.status_code}")
        return None
    
def download_image(url, folder_path, filename):
    os.makedirs(folder_path, exist_ok=True)

    response = requests.get(url)
    if response.status_code == 200:
        with open(os.path.join(folder_path, filename), "wb") as f:
            f.write(response.content)
        print(f"Image downloaded: {filename}")
    else:
        print(f"Failed to download image: {filename}")
    
### Cert Number Generator
def generate_certification_number(consulted_numbers):
    while True:
        if not consulted_numbers:
            return random.randint(60000000, 87000000)
        
        cert_number = random.randint(60000000, 87000000)
        if cert_number not in consulted_numbers:
            return cert_number
        
### Read and update cache file
def load_cache(cache_file):
    consulted_numbers = set()
    if os.path.exists(cache_file):
        with open(cache_file, "r") as f:
            for line in f:
                consulted_numbers.add(int(line.strip()))
    return consulted_numbers

def update_cache(cache_file, consulted_numbers):
    with open(cache_file, "w") as f:
        for cert_number in consulted_numbers:
            f.write(str(cert_number) + "\n")

# Defining variables
api_token = 'nOWT2DI_WAXp1HOcHN57gLvp3NuzT0QOxuZjRtRBnnGJOSD_3JVPGFjx0HBYxK_lwHi8OpF_XZyesTcOctZXV2E9wcGIe4U8B-r79qHMCBbBQ5dt3UvadEyNGmkxMHGfFMy2TgpBxbY2ReSR0EAePSQfrjeo9pjkpkW1GrcnuStdZLqeojbNy2sqA6XeEEqHF2w82wGOdf6fpXNCQi00VNi09lsIIzhfkl5xsUw7nMKJyP1PM5iOkEf_pqOaLmYbbMZsVOgi5QG6eGSLzVcS2tcW4RnQccKWtGqewRjTxT4NTn2H' # Kevin, enter your token here between the single quotes
card_info_list = []
cache_file = "consulted_cert_numbers.txt"

# Check if the existing file exists
if not os.path.exists(cache_file):
    # Create a blank text file named "example.txt"
    with open("consulted_cert_numbers.txt", "w") as f:
        pass


# Loading consulted certification numbers from the cache file
consulted_numbers = load_cache(cache_file)
new_run = len(consulted_numbers) + 2

# Reading database
database_path = "Card_database.xlsx"

# Check if the existing file exists
if os.path.exists(database_path):
    existing_df = pd.read_excel(database_path)
else:
    existing_df = pd.DataFrame()


# Pulling information for 400 cards
while len(consulted_numbers) < new_run:
    # Generates the cert number
    card_year, card_brand, card_sport, card_number, card_player, card_variety, card_grade, back_hyperlink, front_hyperlink = ('',) * 9
    card_grade_desc, card_spec_num, card_qual_code, card_label_type, card_psadna, card_dualcert, card_total_pop, card_total_pop_higher, card_item_era = ('',) * 9
    cert_number = generate_certification_number(consulted_numbers)
    try:
        # Calls the API to get card info
        card_info = get_card_info(cert_number, api_token)

        if not card_info:
            consulted_numbers.add(cert_number)
            continue  # Skip this card if data doesn't exist
        
        cert_number = card_info['PSACert']['CertNumber']
        card_year = card_info['PSACert']['Year']
        card_brand = card_info['PSACert']['Brand']
        card_sport = card_info['PSACert']['Category']
        card_number = card_info['PSACert']['CardNumber']
        card_player = card_info['PSACert']['Subject']
        card_variety = card_info['PSACert']['Variety']
        card_grade = card_info['PSACert']['CardGrade']
        card_grade_desc = card_info['PSACert']['GradeDescription']
        card_spec_num = card_info['PSACert']['SpecNumber']
        card_label_type = card_info['PSACert']['LabelType']
        card_psadna = card_info['PSACert']['IsPSADNA']
        card_dualcert = card_info['PSACert']['IsDualCert']
        card_total_pop = card_info['PSACert']['TotalPopulation']
        card_total_pop_higher = card_info['PSACert']['PopulationHigher']
        try: card_item_era = card_info['PSACert']['ItemEra']
        except: card_item_era = ''
        try: card_qual_code = card_info['PSACert']['QualifierCode']
        except: card_qual_code = ''

        card_images = get_card_images(cert_number, api_token)

        try:
            # URL of the image to download
            back_image = card_images[0]['ImageURL']
            front_image = card_images[1]['ImageURL']

            # Folder path to save the image
            folder_path = 'images'

            # Filename for the image
            filename_back = f'{cert_number}_back.jpg'
            filename_front = f'{cert_number}_front.jpg'

            download_image(back_image, folder_path, filename_back)
            download_image(front_image, folder_path, filename_front)

            back_hyperlink = f'=HYPERLINK("images/{filename_back}", "Image")'
            front_hyperlink = f'=HYPERLINK("images/{filename_front}", "Image")'

        except:
            image_hyperlink = 'No Image Available'

        # Add card information to the list
        card_info_list.append({
            'Cert number': cert_number,
            'Year': card_year,
            'Brand': card_brand,
            'Card number': card_number,
            'Player': card_player,
            'Variety': card_variety,
            'Grade': card_grade,
            'Grade description': card_grade_desc,
            'Spec number': card_spec_num,
            'Qualifier code':card_qual_code,
            'Label type': card_label_type,
            'Is PSADNA': card_psadna,
            'Is Dual Cert': card_dualcert,
            'Total population': card_total_pop,
            'Population higher': card_total_pop_higher,
            'Item era': card_item_era,
            'Front image': front_hyperlink,
            'Back image': back_hyperlink
        })

        consulted_numbers.add(cert_number)

    except:
        consulted_numbers.add(cert_number)

# Updating cache and database
update_cache(cache_file, consulted_numbers)
new_df = pd.DataFrame(card_info_list)

# Concatenate existing and new DataFrames
final_df = pd.concat([existing_df, new_df], ignore_index=True)

# Save the combined DataFrame to the Excel file
final_df.to_excel(database_path, index=False)
    