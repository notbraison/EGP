# install necessary

# egp.py
python -m venv venv
venv\Scripts\activate
pip install openpyxl
## scrape_egp.p
pip install playwright pandas requests
playwright install chromium


# start venv
venv\Scripts\activate
# execute program
python egp.py
# deactivate
deactivate  #when finished

