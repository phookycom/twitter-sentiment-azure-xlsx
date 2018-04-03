# Twitter Sentiment Analysis with Azure Text Analytics API
# (c) 2018 Daniel Schötz

from pymongo import MongoClient
import re
import sys
import time
import json
import xlsxwriter
from progressbar import Bar, Percentage, ProgressBar
import requests

# Azure connection data (Create your Cognitive Services APIs account in the Azure portal)
subscription_key = ''
sentiment_api_url = ''

# Connect to MongoDB
connection = MongoClient("mongodb://localhost")
db = connection.bundestagswahl.tweets

# Set maximum tweets
maxtweets = 5000

# Temporary dictionaries
original_tweets = {'documents' : []}  
tidied_tweets = {'documents' : []}     


# Use a progress bar in your shell (for larger data sets.)
pbar1 = ProgressBar(widgets=['Analyze Tweets: ', Percentage(), Bar()], maxval=maxtweets).start()
pbar2 = ProgressBar(widgets=['Write Excel File: ', Percentage(), Bar()], maxval=maxtweets).start()

def tidy_tweet(tweet):
      '''
      Helper function to tidy the tweet text. 
      The regex removes special characters and links, etc.
      '''
      return ' '.join(re.sub("(@[A-Za-z0-9äöüÄÖÜß]+)|([^0-9A-Za-z9äöüÄÖÜß \t])|(\w+:\/\/\S+)", " ", tweet).split())


def write_excel_result(sentiments):
      '''
      Helper function to write the Excel file and 
      visualize a simple sentiment result.
      '''
      workbook = xlsxwriter.Workbook('Azure_Sentiment_Analysis.xlsx')
      worksheet = workbook.add_worksheet('Sentiment Analysis')
      ws_row = 0
      ws_col = 0 
      cell_format = workbook.add_format({'bg_color': '#ffffff', 'font_color': '#000000', 'border': True, 'border_color': 'silver', 'bold': True})
      worksheet.write(ws_row, ws_col, "Tweet Text", cell_format)
      worksheet.write(ws_row, ws_col + 1, "Azure Sentiment Value", cell_format)
      worksheet.write(ws_row, ws_col + 2, "Sentiment Analysis", cell_format)
      ws_row = 1
      i = 0
      for document in original_tweets['documents']:
            pbar2.update(i+1)
            for result in sentiments['documents']:
                  if document['id'] == result['id']:
                        worksheet.write(ws_row, ws_col, document['text'].replace('\n', ' ').replace('\r', ''))
                        worksheet.write(ws_row, ws_col + 1, str(result['score']))
                        if 0.5 > result['score'] >= 0.25:
                              cell_format = workbook.add_format({'bg_color': '#fec8d0', 'font_color': '#ba001a', 'border': True, 'border_color': 'silver'})
                              worksheet.write(ws_row, ws_col + 2, "Somewhat negative", cell_format)
                        elif 0.25 > result['score'] >= 0:
                              cell_format = workbook.add_format({'bg_color': '#fa9fac', 'font_color': '#ba001a', 'border': True, 'border_color': 'silver'})
                              worksheet.write(ws_row, ws_col + 2, "Negative", cell_format)
                        elif 1 >= result['score'] >= 0.75:
                              cell_format = workbook.add_format({'bg_color': '#93f882', 'font_color': '#066e15', 'border': True, 'border_color': 'silver'})
                              worksheet.write(ws_row, ws_col + 2, "Positive", cell_format)
                        elif 0.75 > result['score'] > 0.5:
                              cell_format = workbook.add_format({'bg_color': '#d9fcd2', 'font_color': '#066e15', 'border': True, 'border_color': 'silver'})
                              worksheet.write(ws_row, ws_col + 2, "Somewhat positive", cell_format)      
                        elif result['score'] == 0.5:
                              cell_format = workbook.add_format({'bg_color': 'white', 'font_color': 'black', 'border': True, 'border_color': 'silver'})
                              worksheet.write(ws_row, ws_col + 2, "Neutral", cell_format)
                        ws_row += 1
      workbook.close()            
      pbar2.finish()
def analyze_tweets():
      '''
      Fetch tweets from MongoDB, write dictionary for Azure Text Analytics, 
      send data to Azure and receive sentiment results.
      '''
      tweets = []
      row = 1  
      try:
            fetched_tweets = db.find().limit(maxtweets)
            i = 0
            for tweet in fetched_tweets:
                  parsed_tweet = {}
                  pbar1.update(i+1)
                  parsed_tweet['text'] = tweet['text']

                  try:
                        if tweet['originalTweet']['retweeted_status']['retweet_count'] > 0:
                              if parsed_tweet['text'] not in tweets:
                                    new_original_item = {"id": str(row), "language": "de", "text": tweet['text']}
                                    original_tweets['documents'].append(new_original_item)
                                    new_tidied_item = {"id": str(row), "language": "de", "text": tidy_tweet(tweet['text'])}
                                    tidied_tweets['documents'].append(new_tidied_item)
                                    row += 1

                        else:
                              new_original_item = {"id": str(row), "language": "de", "text": tweet['text']}
                              original_tweets['documents'].append(new_original_item)
                              new_tidied_item = {"id": str(row), "language": "de", "text": tidy_tweet(tweet['text'])}
                              tidied_tweets['documents'].append(new_tidied_item)
                              row += 1
                  except KeyError:
                        pass
                  tweets.append(parsed_tweet['text'])
                  i = i + 1
            connection.close()
            pbar1.finish()

            # Connect to Azure Text Analytics API
            headers   = {"Ocp-Apim-Subscription-Key": subscription_key}
            response  = requests.post(sentiment_api_url, headers=headers, json=tidied_tweets)
            sentiments = response.json()

            write_excel_result(sentiments)
            return tweets

      except ConnectionError:
            print("Can't connect to MongoDB :-(")
      

def main():
      # Start to analyze...
      analyze_tweets()

if __name__ == "__main__":
	# Call main 
	main()
 
