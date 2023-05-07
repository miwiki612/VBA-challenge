* code source: retrieval_and_summary.bas Module export from .xlsm. 
  - SUB retrieval
    - sort data for the selected sheet, record each Ticker
    - Yearly Change = close_price - open_price
    - Percent Change = (close_price - open_price) / open_price
    - sum vol
  - SUB summary
    - read result in SUB retrieval, then select Ticker with max per, min per, max vol
    
* location: https://github.com/miwiki612/VBA-challenge/blob/main/retrieval_and_summary.bas
