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
- OLS regression between the stock and index is used to determine alpha (𝜶<sub>i</sub>)  and beta (Β<sub>i</sub>)
- Variance of a stock determined by  Β<sup>2</sup><sub>ie</sub>𝝈<sup>2</sup><sub>m</sub>+𝝈<sup>2</sup><sub>ie</sub>
- covariance between stocks determined by  Β<sub>i</sub>Β<sub>j</sub>𝝈<sup>2</sup><sub>m</sub>
- Phi (𝚽) = (𝝈<sup>2</sup><sub>m</sub>∑<sup>n</sup><sub>j=1</sub>Β<sub>j</sub>(µ<sub>j</sub> − Rf)/𝝈<sup>2</sup><sub>je</sub>)/(1+𝝈<sup>2</sup><sub>m</sub>∑<sup>n</sup><sub>j=1</sub>B<sup>2</sup><sub>j</sub>/𝝈<sup>2</sup><sub>ie</sub>)
- For short selling to not be permitted, securities are ranked by excess returns over beta ((µ<sub>1</sub>-Rf)/(Β<sub>1</sub>))
  - Securities are added until 𝚽 is maximized
- Security weights in tangent portfolio Z<sub>i</sub>=B<sub>i</sub>/(𝝈<sup>2</sup><sub>ie</sub>)*((µ<sub>i</sub>-Rf)/B<sub>i</sub>-𝚽), where x<sub>i</sub> = z<sub>i</sub>/(∑<sup>n</sup><sub>i=1</sub>z<sub>i</sub>)
