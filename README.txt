PRICE PARITY TRACKER  -  Nat Habit prices on 10 channels vs the ALIGNED PRICE
  Marketplaces (one all-India price): Amazon, Flipkart, Myntra, Nykaa, FirstCry
  Quick commerce (pincode by pincode): Blinkit, Zepto, Instamart, BigBasket, FK Minutes
==============================================================================

FIRST TIME ONLY:  double-click  FIRST TIME SETUP.bat

YOU ONLY EDIT 4 EXCEL FILES  (folder 1_Input)  - same idea as the SLA tracker
------------------------------------------------------------------------------
 1. Links.xlsx           Channel Name | Channel ID
                         = which products to check. Channel ID can be the id OR the full product link.
 2. Product Master.xlsx  Channel Name | Channel ID | NH SKU | Product Name
                         = which NH SKU each channel id is. Every id in Links should be here.
 3. Pincodes.xlsx        Channel Name | Pincode
                         = where to check the quick-commerce apps.
                           Channel Name = All QC (all 5 apps) or one app.
 4. Pincode Master.xlsx  Pincode | Tier | City | State | Region   (city names)

 Channel Name is always one of:
   Amazon, Flipkart, Myntra, Nykaa, FirstCry, Blinkit, Zepto, Instamart, BigBasket, FK Minutes
 Each file has a "How to fill" tab.

 Aligned price: taken from the newest "Aligned Price*.xlsx" in the Price Check folder next to
 this one. Update it there when the aligned prices change.
 (NH Products.xlsx = product names only. QC Mapping.xlsx is not used any more.)

BUTTONS
------------------------------------------------------------------------------
 TEST RUN (2 pincodes, windows visible).bat   quick test, you can watch the browser
 RUN PRICE TRACKER.bat                        one full run now
 SCHEDULE DAILY RUN.bat                       Windows runs it by itself every day at the time you type
 STOP DAILY RUN.bat                           stop the daily run
 OPEN DASHBOARD.bat                           the dashboard
 FK MINUTES - LOGIN ONCE.bat                  only if FK Minutes says "could not set location"

RESULTS (folder 2_Output) - don't edit
------------------------------------------------------------------------------
 Price_Parity_DD-MM-YY-HH-MM.xlsx
   Parity Report    one row per NH SKU: lowest / highest price on each channel + status
   Needs Attention  what to fix in your Excel files - READ AFTER EACH RUN:
                      1. ids in Links but missing in Product Master
                      2. Nat Habit products an app shows that are NOT in your Links (add them if needed)
                      3. rows that could not be read
   Detail           every channel x pincode x product
   Summary          counts per channel
 history\qc_price_history.csv   all runs (the dashboard reads this). Kept 180 days.
 _Logs   run logs (send the newest one to Claude if something looks wrong)

STATUS MEANING
------------------------------------------------------------------------------
 At par               within Rs 1 of aligned price
 Below aligned        cheaper than aligned (margin loss / pulls sales from other channels)
 Above aligned        costlier than aligned (sales loss on that channel)
 Out of stock         listed but not buyable
 Not available here   in your Links, but the app does not show it at this pincode
 Not live / not read  marketplace product whose price could not be read (delisted, or site blocked)
 Not in Links         the app shows this Nat Habit product but it is not in your Links
 Not mapped           no NH SKU for this id in Product Master
 Check failed         the app could not be read at that pincode (see _Logs\debug)

NEEDS THESE FOLDERS NEXT TO THIS ONE (same Desktop)
------------------------------------------------------------------------------
 Price Check, Amazon, Flipkart, Myntra, Nykaa, FirstCry
 (Don't run anything inside them on this laptop - Price Check mail runs from the main laptop.)

TIPS
------------------------------------------------------------------------------
 - Close the Excel files in 1_Input before a run.
 - One channel only:      RUN PRICE TRACKER.bat --only zepto
 - Only quick commerce:   --only qc          Only marketplaces:  --only marketplaces
