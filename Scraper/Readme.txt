   ____      _ _  __                  _         _   _ _       _                          ____       _             _ 
  / ___|__ _| (_)/ _| ___  _ __ _ __ (_) __ _  | | | (_) __ _| |____      ____ _ _   _  |  _ \ __ _| |_ _ __ ___ | |
 | |   / _` | | | |_ / _ \| '__| '_ \| |/ _` | | |_| | |/ _` | '_ \ \ /\ / / _` | | | | | |_) / _` | __| '__/ _ \| |
 | |__| (_| | | |  _| (_) | |  | | | | | (_| | |  _  | | (_| | | | \ V  V | (_| | |_| | |  __| (_| | |_| | | (_) | |
  \____\__,_|_|_|_|  \___/|_|  |_| |_|_|\__,_| |_| |_|_|\__, |_| |_|\_/\_/ \__,_|\__, | |_|   \__,_|\__|_|  \___/|_|
                                                        |___/                    |___/                                  
  __  __  ___  _   _      _                                            _     _   _           _       _            
 |  \/  |/ _ \| | | |    / \   __ _ _ __ ___  ___ _ __ ___   ___ _ __ | |_  | | | |_ __   __| | __ _| |_ ___ _ __ 
 | |\/| | | | | | | |   / _ \ / _` | '__/ _ \/ _ | '_ ` _ \ / _ | '_ \| __| | | | | '_ \ / _` |/ _` | __/ _ | '__|
 | |  | | |_| | |_| |  / ___ | (_| | | |  __|  __| | | | | |  __| | | | |_  | |_| | |_) | (_| | (_| | ||  __| |   
 |_|  |_|\___/ \___/  /_/   \_\__, |_|  \___|\___|_| |_| |_|\___|_| |_|\__|  \___/| .__/ \__,_|\__,_|\__\___|_|   
                              |___/                                               |_|                             
                                                                         

How to run:
This is a Python file, so you will need to install Python to run the script. You can download at Python.org.
The scraper should run anytime there’s a excel file named 'MOU agreements' in the same folder as the program. In the event
the file is missing or named incorrectly the program will throw an exception and state 'MOU agreements.xlsx' is not found.
The excel file is expected to have the first row as column headers. I would suggest inputting a copy of the blank template
file I've provided and saving the results after the script has run. Rename the file and place it into a separate folder.

How it functions:
This program is designed to load a premade .xslx file, it opens hyperlinks associated with the callsigns of our allied
agencies. While visiting the hyperlinks, it looks for certain queues to know what information to grab and what information 
to skip. The queues I've chosen are all the boxes between the Licensee section's 'Type' and the Ownership and 
Qualifications section’s 'Radio Service Type'. The only other information that is grabbed from the hyperlink is the 
expiration dates, the expiration dates are hard coded based on the order they are placed on the website. If results 
become inaccurate or include more than the desired number of results it is most likely due to a change in the website and
will require the code to be updated.

WARNING:
This script is a web scraper, so it loads all the pages just like a normal browser would, but it does it very fast. Due to 
the speed it's opening the pages at, it can put a strain on the servers. For this reason I would suggest refraining from 
running the script more than once per day, and ideally no more than once per month. I've ran the script a few times in the 
same day and it will make the pages inaccessible at a certain point, either because our access was cut off or because the
page stops working, not sure which. 


