# Overview

This is a simple script that pulls down Michigan construction contract data from https://mdotjboss.state.mi.us/CCI. The intended purpose is to help automate the process of gathering this data for analysts that are interested in performing analytics on construction data. 

# Implementation

Because the MDOT site uses javascript to dynamically generate the data that is seen on the site, the use of Selenium is necessary in order to mimic the actions of a user. As of right now, Chrome is the only browser that is supported, but support for Firefox, Edge, IE, Opera, Brave, or any other major browser could be implemented.

The script also attempts to do a lot of the heavy lifting for users, as the expectation is that users are not programmers by trade, and thus shouldn't be expected to do things like install Python libraries.

# Usage

In order to run the script, you must have Python installed on your computer. If you're using Windows, you can find a good tutorial on how to install Python [here](https://www.digitalocean.com/community/tutorials/install-python-windows-10). If you're on Mac, you can use [this article](https://docs.python-guide.org/starting/install3/osx/).

Once you have Python installed, the script can be run with the following command.

```bash
python mdot_scraper.py
```

You'll be asked two questions:

1. Where is the file with the project numbers located, and
2. Where should the data be saved once it's downloaded.

Both questions expect you to type file paths into the terminal. The filepath for the project numbers should point to a CSV or a TXT file, where each project number is on its own line. The filepath for the data output should be a folder where you want the CSV file(s) to be saved.

# Handling Errors

If you encounter any errors while using this script, it's almost certainly because the CSS/HTML for the site has changed. Feel free to [open an issue](https://github.com/bryanbritten/scrape_mdot/issues) with a screenshot of your Powershell window with as much output as you can show. 
