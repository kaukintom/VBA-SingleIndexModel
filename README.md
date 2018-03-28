# VBA-SingleIndexModel
Download stock data to a particular Index (ex. ^GSPTSE - S&amp;P/TSX Composite index) using Alpha Vantage API keys and then calculates the market portfolio and the integer values of how much to invest in each stock.
# API key
Can receive an API key from alphavantage.com or https://www.alphavantage.co/support/#api-key 
# Data Used
- Daily Adjusted close return data for all calculations to determine % allocation
- Daily close price (not adjusted) used for integer allocation
- Up to 4 months of return data is used (can choose to use 1 to 4 months in portfolio analysis)
- Risk free rate (Rf) determined from Bank of Canada T-Bill rates at https://www.bankofcanada.ca/rates/interest-rates/t-bill-yields/
  - T-Bill rate is converted to a daily return based average number of data points from daily stock returns in a year
# Calculations
- OLS regression between each stock and the index is used to determine alpha (𝜶<sub>i</sub>)  and beta (Β<sub>i</sub>)
- Variance of a stock determined by  Β<sup>2</sup><sub>ie</sub>𝝈<sup>2</sup><sub>m</sub>+𝝈<sup>2</sup><sub>ie</sub>
- covariance between stocks determined by  Β<sub>i</sub>Β<sub>j</sub>𝝈<sup>2</sup><sub>m</sub>
- Phi (𝚽) = (𝝈<sup>2</sup><sub>m</sub>∑<sup>n</sup><sub>j=1</sub>Β<sub>j</sub>(µ<sub>j</sub> − Rf)/𝝈<sup>2</sup><sub>je</sub>)/(1+𝝈<sup>2</sup><sub>m</sub>∑<sup>n</sup><sub>j=1</sub>B<sup>2</sup><sub>j</sub>/𝝈<sup>2</sup><sub>ie</sub>)
- For short selling to not be permitted, securities are ranked by excess returns over beta ((µ<sub>1</sub>-Rf)/(Β<sub>1</sub>))
  - Securities are added until 𝚽 is maximized
- Security weights in tangent portfolio Z<sub>i</sub>=B<sub>i</sub>/(𝝈<sup>2</sup><sub>ie</sub>)*((µ<sub>i</sub>-Rf)/B<sub>i</sub>-𝚽), where x<sub>i</sub> = z<sub>i</sub>/(∑<sup>n</sup><sub>i=1</sub>z<sub>i</sub>)
# VBA Macros
Module: Collect_Stock_Data<br />
Sub Procedure: TickerDataCollect
 - Collects stock and index data from Alpha Vantage from API key
 - Drawback: Currently uses a fixed 5 second delay per a security to allow the csv to open and display the data
  - Time consuming
  
Module: Portfolio<br />
Sub Procedure: OLSRegression
 - Calculates the portfolio composition of stocks and their weight as a %
 - Determines the integer value of each stock an investor should invest in to minimize variance or maximize return
 - Quick runtime (< 1 second)

Module: TransferPriceData<br />
Sub Procedure: Transfer_Table
 - The allocation table created from OLSRegression can be moved to avoid being deleted if OLSRegression is used again

Module: UpdatePriceDate<br />
Sub Procedure: GetPriceData
 - Update the table saved from Transfer_Table with the most current price data
 - Calculates the return to date and dollar value return
# Acknowledgements
McMaster University: Commerce 4FF3 - Portfolio Theory and Management
