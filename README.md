# Portfolio-Risk-Analysis-Excel

This project analyzes a **hypothetical portfolio** constructed from three major stock indexes: **S&P 500, FTSE 100, and DAX**, using **two years of historical data**. Relevant **exchange rates** were also included to account for currency effects.  

The raw data was **extracted from Google Sheets using GOOGLEFINANCE functions**, then **cleaned, converted, and prepared** for analysis. The project visualizes the **loss distribution** through a histogram, calculates statistical measures of risk, and constructs a **risk matrix** including **Value at Risk (VaR)** and **Expected Shortfall (ES)**.

Additionally, the **loss distribution is compared against Normal and Student's t-distributions** to identify which distribution best models the portfolio. The impact of assuming each distribution on **VaR and ES estimates** is also analyzed. 
This project is inspired by the **Risk Management Specialization course on Coursera**.


## Method
1. Extract data from Google Sheets using Excel functions.  
2. Clean, convert, and prepare data for analysis.  
3. Construct a hypothetical portfolio using the three indexes and exchange rates.  
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
