import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import xlwings as xw
import fetchreviews as fr


APP_ID = 'com.linkedin.android.learning'
MAX_REVIEWS = 5000

result = fr.fetch_reviews(APP_ID, MAX_REVIEWS)


# Set the start date from which to scrape reviews

# Get today's date
today = datetime.today()
# Calculate the no. of days since the start of the current week (Monday)
days_since_monday = today.weekday()  
# Get the previous week's Monday date by subtracting 7 more days
start_date = today - timedelta(days=days_since_monday + 7)
start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)


# Filter reviews based on start date
filtered_reviews = [review for review in result if datetime.strptime(review['at'].strftime('%Y-%m-%d'), '%Y-%m-%d') >= start_date]

# Convert to a DataFrame
df = pd.DataFrame(filtered_reviews)
df = df[["reviewId","score","content","reviewCreatedVersion","at","replyContent","repliedAt","appVersion"]]

# Calculations for identifying bad reviews
df['bad_review'] = df['score'] < 3
df['bad_review_replied'] = np.where(pd.notna(df['replyContent']), df['score'] < 3, False)

# # Save the DataFrame to an Excel file for the first time and create a template file
# df.to_excel('Template.xlsx', index=False)

## Run below commands after the excel template file is already created

# Open existing excel template file
FILENAME = "Template.xlsx"

wb = xw.Book(FILENAME)
wb.sheets["Raw Data"].activate()
ws = wb.sheets["Raw Data"]

# First clear existing data and then paste the new values of the dataframe from B2 cell
ws.range("A2:J1048576").clear_contents()
ws.range("A2").value = df.values

# Select the worksheet that contains the Pivot Tables to refresh it
wb.sheets['Calculations'].select()
wb.api.active_sheet.refresh_all(wb.api)

# Select the dashboard as active sheet before closing the report
wb.sheets['Dashboard'].activate()

# Save as a new excel workbook
report_file_name = today.strftime('%Y-%m-%d')+ "-" + "Weekly Review Analysis.xlsx"
wb.save(report_file_name)

# Close the Excel workbook
wb.close()