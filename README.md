# What is "ScrapeZoomURLfromOutlook"?

This program enables you to scrape zoom urls from mails in Outlook Mail application. 

- Mail app we use in outlook:
    - @yahoo.ne.jp
    - @gmail.com
    - @st.#####-u.ac.jp


# Scraping Step
- Set a data collection period

- Per mail address
    - Create a mail list that contains URLs in the collection period
    - Create a mail list from it that contains zoom URLS
    - Input contents of the mail in the zoom mail list to TABLE
        - Received time (TEXT)
        - Mail subject (TEXT)
        - Sender name (TEXT)
        - Sender email address (TEXT)
        - URLs (https://...zoom.us/....|https://...zoom.us/....|.......) (TEXT)
        - Numbers of URLs (INTEGER)

- Delete rows whose URL is equal to the following URLs:
    - https://zoom.us/
    - https://zoom.us/support/download
    - https://zoom.us/test


- Show the numbers of zoom URLs and a lump of URLs per mail like
    - {number} zoom url(S): 
    - https://...zoom.us/....|https://...zoom.us/....|.......
    - 
    - similarly.....


# Finally
We hope you could use this program to effectivelly help carry out your creative activity.