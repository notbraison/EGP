# install necessary

# egp.py
python -m venv venv
venv\Scripts\activate
pip install openpyxl
## scrape_egp.py
pip install playwright pandas requests
playwright install chromium
## sync_gsuite.py
pip install openpyxl


# start venv
venv\Scripts\activate
# execute program
python main.py
# deactivate
deactivate  #when finished

