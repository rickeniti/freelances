# Trading Card Scraper (API to Excel)


## Project goal
The client needed a Python script that, whenever run, would pull information from multiple trading cards from PSA's website. The information should be recorded in an Excel spreadsheet after each run, including front and back images.
There was no specific list of cards that the client wanted to scrape. The program should be able to generate card ID's (certification numbers) and use them as a parameter to call the API as many times as the quota allowed.

## Solution
My solution can be broken down into 2 parts:


### 1) Create a function that generates certification numbers
I went to different websites and forums to understand the pattern used to generate certification numbers. It's an 8-digit identifier composed only by numbers and, in general, the higher the number the more recently the card was graded. I found a range of numbers that generated cards with images and used random.choice to pick cards at random.


### 2) Call different API methods and join the results
I used two API methods, one to get card information and another one to get image URLs.
The script calls the API using the certification number, populates an existing Excel file with card information and downloads the images in a separate folder linking them to their respective card through Excel's formula HYPERLINK.
The script also keeps a record of all the certification numbers already consulted through a cache file to avoid duplicates
