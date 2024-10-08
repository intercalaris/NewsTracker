import datetime
from newsapi import NewsApiClient
import pandas as pd
import openpyxl
import requests
import os
import dotenv

dotenv.load_dotenv()

# now using get_everything() function instead of newsapi library
def get_everything(api_key, params):
    endpoint = "https://newsapi.org/v2/everything"
    headers = {'Authorization': f'Bearer {api_key}'}
    response = requests.get(endpoint, headers=headers, params=params)
    return response.json()


# API Key has limit of 100 calls a day. If new API needed, go to https://newsapi.org/ and click "get API key"
api_key = os.getenv('NEWSAPI_KEY')
if not api_key:
    raise Exception("API key not found. Please set it in your environment variables or .env file.")

previous_day_object = datetime.datetime.now() - datetime.timedelta(days=1)
previous_day = previous_day_object.strftime('%Y-%m-%d')
past_day_backfill = '2024-09-10'

from_date = past_day_backfill
to_date = previous_day


params = {
    'q': 'Linguistics',
    'language': 'en',
    'sortBy': 'publishedAt',
    'from': from_date,
    'to': to_date,
    'pageSize': 100
}

all_articles = get_everything(api_key, params)
print(f"API Response: {all_articles['totalResults']} articles fetched from {len(all_articles['articles'])} sources.")


# extract data into dictionary lists
articles = []
for article in all_articles['articles']:
    if "[Removed]" not in article['source']['name']:
        articles.append({
            'Source': article['source']['name'],
            'Author': article['author'],
            'Title': article['title'],
            'Description': article['description'],
            'URL': article['url'],
            'Date Published': article['publishedAt'][:10],
            'Content': article['content']
        })

# convert dict list to dataframe
newsfeed_df = pd.DataFrame(articles)
newsfeed_df['Date Fetched'] = pd.Timestamp.now().date()

file_name = "NewsFeed_Linguistics.xlsx"

# try to open excel file and append new data, if doesn't exist create new file 
try:
    existing_df = pd.read_excel(file_name)
    combined_df = pd.concat([newsfeed_df, existing_df], ignore_index=True)
    combined_df.to_excel(file_name, index=False)
    print(f"Updated {file_name} with new articles.")
except FileNotFoundError:
    newsfeed_df.to_excel(file_name, index=False)
    print(f"Created new file and saved news to {file_name}")


workbook = openpyxl.load_workbook(file_name)
worksheet = workbook.active
for column in worksheet.columns:
    worksheet.column_dimensions[column[0].column_letter].width = 16
workbook.save(file_name)