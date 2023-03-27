# VBA-Challenge
#Module 2 VBA 
# Used chat GPT to correct loop issue for Yearly Change in ticker price.  Initial script was initially using the first open and close price for each ticker, not #beginning of year and end of year so the yearly change was way off.
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

