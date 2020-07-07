# uOttawa Timetable-Scraper 

This is a python script that extracts information about University of Ottawa courses such as their name, subject, semester and times. This script uses Selenium, a web automation tool and xlwt, an excel module for python.

# Installation

To use this script, you can simply download or clone the repository. After that, you'll have a folder containing the files including courses.py, which is the script used to scrape. 

You'll also need to install the Selenium WebDriver, which can be downloaded here: [Link](https://www.seleniumhq.org/)

The script also uses the FireFox browser, which can be downloaded here: [Link](https://www.mozilla.org/en-CA/firefox/)

The script also uses geckoDriver, which allows Selenium to perform its operations in a FireFox browser. This can be downloaded here: [Link](https://github.com/mozilla/geckodriver/releases). You'll have to put the geckodriver into your PATH for it to work. 

Of course, you can use Google Chrome and ChromeDriver instead, but you'll have to modify the script accordingly. 

# Usage

If your installation is successful, you will be able to use the script by going into the repository folder and typing into your terminal:

```
python courses.py
```

When your browser launches and begins changing pages, then you know it works. The script will take between 15-17 minutes depending on your connection. Messages will be printed in the terminal to inform you of its progress. 