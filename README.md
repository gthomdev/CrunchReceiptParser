# CrunchReceiptParser
Parse Crunch receipt page to correlate all receipt fields

## Installation

* Make sure you have Python installed (https://www.python.org/downloads/)
* Open a command terminal and navigate to the folder where you have downloaded this project
* Run the following command to install the required packages:
* `pip install -r requirements.txt`

## Usage

* Download the receipt page from Crunch 
  * To do this navigate to the page in a browser 
  * Press CTRL + S to save the HTML file

In the space in the main method of the "CrunchParser.py" file, add the path to the receipt page you have downloaded

I have provided an example of what that path should look like.

Also, add the path where you would like the output file to be saved, an example has also been provided.

To run the script, open up a command shell (e.g. CMD, PowerShell) and run:
`python <path_here>CrunchParser.py`

* The output file will be created in the location you specified.

