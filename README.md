# 📊 Excel Dashboard/Report Automation with Python

This repository demonstrates how to automate Excel report updates using Python, with a focus on scraping app review data and updating a dashboard seamlessly.

In this case study, we use the `google_play_scraper` module to extract weekly reviews of the **LinkedIn Learning App** from the Google Play Store, and then automate the population of an Excel-based dashboard using the powerful `xlwings` library.

👉 [Read the full case study on Medium](https://medium.com/python-in-plain-english/python-automation-for-excel-dashboards-a-case-study-of-linkedin-learning-app-reviews-cb9d45c80b7a)

---

## 🚀 Features

- Scrapes recent user reviews from the Google Play Store using `google_play_scraper`
- Analyzes low-rated reviews for weekly trends and feedback
- Updates an Excel dashboard automatically using `xlwings`
- Replaces manual data copy-pasting with a clean, reproducible script
- Customizable for any Android app on the Play Store

---

## 📦 Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/<your-username>/excel-dashboard-automation.git
   cd excel-dashboard-automation
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

---

## ⚙️ Configuration

Before running the script, update the following values in the main script:

- **`APP_ID`**  
  Replace with the ID of the app you wish to analyze (e.g. `'com.linkedin.android.learning'`)  
  📘 Refer to this [Google guide](https://support.google.com/admanager/answer/11382876?hl=en) to find any app’s ID.

- **`MAX_REVIEWS`**  
  Set a high enough value to ensure you're capturing all reviews from the last week.  
  Keep in mind: More reviews = Longer execution time.

---

## ▶️ Usage

Run the script from the terminal or your preferred IDE:

```bash
python main.py
```

A new Excel report will be generated automatically, populated with the latest weekly reviews, analysis summaries, and charts.

---

## 📈 Output

- `LinkedInLearningReport.xlsx` (or your chosen file) will be updated or created with:
  - All scraped reviews from the last 7 days
  - Filtered list of low-rated reviews (e.g. 1 or 2 stars)
  - Summary insights for manual or automated review

---

## 🛠️ Customization Ideas

- Connect to other app stores or APIs (Apple App Store, Trustpilot, etc.)
- Add sentiment analysis for more nuanced feedback interpretation
- Schedule the script with cron jobs or Task Scheduler for weekly automation
- Email the updated report using `smtplib` or integrations like Zapier

