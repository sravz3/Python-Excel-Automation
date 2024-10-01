# Excel Dashboard/Report Automation with Python
This repository showcases how to use xlwings module to manipulate excel files to automate the update process of reports/dashboards created in excel. In this particular case study,
I scraped google playstore reviews of LinkedIn Learning App with the help of google_play_scraper module and I built a small dashboard in excel to analyse the reviews of last week, especially the one with low scores.


## Usage

1. Download all the files.

2. Update the 'APP_ID' in main file with the ID of app you wish to scrape reviews of. Refer to this article to get the app id of any app - https://support.google.com/admanager/answer/11382876?hl=en

3. Update the 'MAX_REVIEWS' with a high enough number to ensure you are scraping all reviews for the timeframe you are looking for. Higher the number, higher is the execution time.

4. Run the file and a new excel report will be generated updated with the reviews from last week.
