# Solartis-Web-Automation
This is a Web Automation script written in python with the help of Selenium.
It is meant to extract data of 51 people from an excel sheet and fill the insurance form [here](https://enrollmentdemo.solartis.net/Quote.xhtml).

#### PRE-REQUISITES TO RUN THE SCRIPT:
1. Download and unzip **Chrome Driver** . You can download it [here](https://chromedriver.chromium.org/downloads) according to your version.
2. Download **Selenium**. [here](https://www.selenium.dev/)
3. Download **openpyxl**. You can do _pip install openpyxl_ if you have pip.
4. In line number 210 of the code, you must give the path of your folder where you want to download the final recipts of each insurance form. 
   **By default the recipts will not get downloaded** but you can choose to download them by uncommenting lines 268 and 275 of the code. (The recipts are only 300 KBs. Dont worry :-)  
   An example would be: "download.default_directory": "D:\Py Project\Output"

#### NOTE:
1. The website might not get loaded properly when you execute the code first time and you may recieve a timeout error . Consider executing it once again to get the correct results
2. For some of the test cases, the purchase won't be successful. You can then wait for a minute for the code to automatically go into next text case or you can close the browser window, then the next test case would automatically be implemented.
3. All the data(names,addresses,numbers) used in the script is fake and the insurance site made by Solartis is only for project purposes and are not real.
