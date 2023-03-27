# VBA-Challenge
#Module 2 VBA 
#Week 2 Challenge to create a VBA script that loops through all the stocks for one year and outputs the following information:
#The ticker symbol
#Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
#The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
#The total stock volume of the stock.
#Ran the VBA Script using the 'alphabetical_listing.xlxs' file for practice before adding to the Multiple Year Stock List, which is too large to load here.
#Took screen shots of each worksheet after successfully running script on workbook.  
#** Used chat GPT to correct loop issue for Yearly Change in ticker price.  Initial script was not looping through all open/close prices for each ticker and only using the #first open and close price for each ticker, making the yearly change incorrect.**
#The missing loop logic was: Dim ticker_row As Long
                             #ticker_row = i
                             #Do While ws.Cells(ticker_row - 1, 1).Value = ticker
                             #ticker_row = ticker_row -1
                             #Loop
                             #opening_price = ws.Cells(ticker_row, 3).Value
                             #closing_price = ws.Cells(i, 6).Value
                             
                             #and then calculate yearly change with
                             #yearly_change = closing_price - opening_price
                            #https://chat.openai.com/chat/8d6a76f4-bf84-4948-809a-a8beb6d10617

