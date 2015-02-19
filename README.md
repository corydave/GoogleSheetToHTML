# GoogleSheetToHTML
Parses content in a Google Sheet and publishes as an HTML file (with anchor tags)

STEP 1:
Copy and paste this code directly into a blank Google Script (bound to a Spreadsheet). 


STEP 2:
Populate the three global variables (LOGO, ACTUALURL, and HTMLFILE) in the beginning of the code.


STEP 3:
Make sure there are three tabs, named appropriately:

"Chronological" - This tab is where the content goes (it can be formatted with blank lines, colors, borders, etc.).

"forHTML" - This tab is where the content for the HTML page will be drawn from. Content from "Chronological" will be parsed
            and and placed here. 
 
 "Standards" - This tab is where a nice, readable, formatted version of the content from "Chronological" resides.
 
 
 STEP 4:
 Make sure the file referenced in forHTML is in a folder that is open to the public to view.
 
 
 STEP 5: 
 Have fun.
