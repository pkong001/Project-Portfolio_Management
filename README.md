### Project Information
This project's objective is to create a dashboard for managing investment funds, particularly in Thailand, where some of the underlying assets are foreign. To accomplish this, I will pull information from various sources using public API, combine, transform, and build the dashboard.

### Project Overview
It is difficult to track multiple funds from multiple ports of funds that are provided by various brokers and investment companies; therefore, in order to effectively track each fund, we must manually access each broker's website and simply copy and paste the information we need. Despite the time-consuming task of accessing each fund's information individually, some funds serve as feeder funds, so you must also examine the fed fund, which is on a different market, in a different time zone, and uses a different currency. Some individuals may also find it difficult to track the performance of their funds due to their unfamiliarity with the type of price graph, as not all funds use candlestick charts, which are preferred by many investors.

### Solution
This project was developed to alleviate the time-consuming task of monitoring each fund individually. First, I created a timeseries plot of price between the feeder fund and the fed fund in order to observe the price movement between the two.

Then I also create a weekly candlestick for the fed fund, as many websites do not have one.
### Prepare Data
We use API to retrieve data from  [Finnomena](https://www.finnomena.com/) and [Yahoo Finance](https://finance.yahoo.com/). Some value from Finnomena we use web scraping to retrieve since API doesn't provided. The scraping of data from Finnomena websites is permitted because the data are publicly accessible and the websites contain no indication that web scraping algorithms are prohibited. The data we receive are in various forms and formats; therefore, we will transform and clean the data before combining it for use in chart plotting and updating the current price (NAV)

### Process
1. Retrieve Data from websites utilizing API and BeautifulSoup (web scraping library)
2. Transform Data into the required format, then combine two data sources. During the data transformation, I also performed data cleaning; however, we will clean the data again to ensure that it is accurate (the data set is not particularly large, so this should not be a problem).
3. After retrieving data, we will transform and then clean it to ensure its integrity.
The steps to verify data integrity are as follows:
* Ensure that there are no missing, duplicate, or corrupted data.
* Verify changes as data is collected over time. 
* Ensure data is in the correct form and format. 
* Ensure data is sorted and ready for use.

### Result
1. Excel automation: The resulting data set will contain three columns titled Funds Code, NAV(correct price), and NAV. Date
2. [Tableau Chart](https://public.tableau.com/app/profile/pongpisut.kongdan/viz/FundsDashboard_16752377846780/Dashboard1): The data set will contain historical prices for both feeder funds and fed funds, which will be stored in 9 columns, 'Open', 'High', 'Low', 'Close Price for Feeder Funds', Close price for Fed Funds, 'Funds Code', 'Funds Filters', 'Date', 'Year', and 'Categories'. These data will be used to visualize time series and candlestick charts on the tableau.


### Additional Improvements and Limitations 
* Use paid API to retrieve data for global funds or fed funds, since the API we use is free, provided by [Yahoo Finance](https://finance.yahoo.com/); however, some data from prior years are missing. 
* For plotting charts, we can use other methods that are more powerful, such as Plotly, DASH, or API provided by Trandingview; which is able to add indicators and perform other trading-related tasks.

