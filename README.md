# Portfolio-Risk-Analysis-Excel

This project analyzes a **hypothetical portfolio** constructed from three major stock indexes: **S&P 500, FTSE 100, and DAX**, using **two years of historical data**. Relevant **exchange rates** were also included to account for currency effects.  

The raw data was **extracted from Google Sheets using GOOGLEFINANCE functions**, then **cleaned, converted, and prepared** for analysis. The project visualizes the **loss distribution** through a histogram, calculates statistical measures of risk, and constructs a **risk matrix** including **Value at Risk (VaR)** and **Expected Shortfall (ES)**.

Additionally, the **loss distribution is compared against Normal and t-distributions** to identify which distribution best models the portfolio. The impact of assuming each distribution on **VaR and ES estimates** is also analyzed. 
This project is inspired by the **Risk Management Specialization course on Coursera**.


## Method
1. Extract data from Google Sheets using the GOOGLEFINANCE function.
2. Clean, convert, and prepare data using Excel functions (e.g., VLOOKUP, date alignment, and filtering).
To ensure comparability, only trading days where all three markets (S&P 500, FTSE 100, DAX) were open were retained. Days with missing data due to holidays or closures in any index were removed.
3. Construct a hypothetical portfolio using the three indexes.  
4. Set up bins covering the full range of losses.  
5. Calculate relative frequencies for each bin.  
6. Create a histogram in Excel to visualize the loss distribution.  
7. Fit **Normal and t-distributions** to the loss data and compare which better models the portfolio.  
8. Compute statistical measures including:
   - Mean, standard deviation, kurtosis  
   - **Value at Risk (VaR)**  
   - **Expected Shortfall (ES)**  
   - Assess changes in VaR and ES under Normal vs t-distribution assumptions.  
9. Construct a **risk matrix** summarizing portfolio risk.


## Tools
- Microsoft Excel



The project is still in progress. A partially completed Excel file has been uploaded, and detailed analysis and explanations will be added upon project completion.
