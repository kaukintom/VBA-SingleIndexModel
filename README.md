# VBA-SingleIndexModel
Download stock data to a particular Index (ex. ^GSPTSE - S&amp;P/TSX Composite index) using Alpha Vantage API keys and then calculates the market portfolio and the integer values of how much to invest in each stock.
# Calculations
- OLS regression between the stock and index is used to determine alpha (𝜶<sub>i</sub>)  and beta (Β<sub>i</sub>)
- Variance of a stock determined by  Β<sup>2</sup><sub>ie</sub>𝝈<sup>2</sup><sub>m</sub>+𝝈<sup>2</sup><sub>ie</sub>
- covariance betweeen stocks determined by  Β<sub>i</sub>Β<sub>j</sub>𝝈<sup>2</sup><sub>m</sub>
- Phi (𝚽) = 𝝈<sup>2</sup><sub>m</sub>∑<sub>j=1</sub><sup>n</sup>Β<sub>j</sub>(µ<sub>j</sub> − Rf)/𝝈<sup>2</sup><sub>je</sub>
