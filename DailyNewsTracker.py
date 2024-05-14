import datetime
from newsapi import NewsApiClient
import pandas as pd
import openpyxl

newsapi = NewsApiClient(api_key='ad131bb2cd094aeeb1748bf2f2b5a24a')

previous_day_object = datetime.datetime.now() - datetime.timedelta(days=1)
previous_day = previous_day_object.strftime('%Y-%m-%d')
past_day_backfill = '2024-04-09'

from_date = previous_day
to_date = previous_day

all_articles = newsapi.get_everything(q='"Linguistics"',
                                      language='en',
                                      sort_by='publishedAt',
                                      from_param=from_date,
                                      to=to_date,
                                      page_size=100)

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