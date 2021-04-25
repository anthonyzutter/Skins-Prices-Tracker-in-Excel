# Skins Prices Tracker in Excel
[![made-with-python](https://img.shields.io/badge/Made%20with-Python-1f425f.svg)](https://www.python.org/)

### Why should you use it?
If you have a lot of steam items it's hard to know the price you paid for each skin and also you have to manually go to your inventory to see the current prices of each skin, with this Python script you just write the name of each skin you want to track and the price you paid for it, the script will then get the current price for your.

### Sample
![Sample Img](imgs/sample.png)

### How to use it
First install the requirements
```sh
python -m pip install -r requirements.txt
```

To write the skins names go to your inventory choose a skin then click on View in community market
![Inventory Img](imgs/inv.png)

Now you just have to copy the name in the URL like shown below and paste it to the excel
![Url Img](imgs/url.png)

Just run the script every time before you open the excel file

If you want to change the currency, game or file name you can change it at the bottom of the code
![Url Img](imgs/change.png)

[Link](https://steam-community-market.readthedocs.io/en/latest/pages/enums.html#esteamcurrency) of Suported Currencys and games
