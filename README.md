# Data-Collection-and-Comparison
#### Collecting Data content of three databases and check for differences or mismatches

The goal of this project was to collect data content of three different databases abount one certain product and store them into one new file. Excel was the primary software, so all code was written within VBA.

The content was completely in German. Comparing to the project for content management, only comments will be translated into English in order not to corrupt any code.

Information about a product was given via Excel or XML files. In most cases, the XML files contained about 1000 lines/rows so the integrated xml parser of Excel was well suited to handle the information.

All code is written within one Excel file and can be accessed via buttons on the data sheet. So, the project can be seen as a toolset data collection but not necessary analytics.

The goals of the toolset were the following:
- Download necessary XML files from within the excel file by pushing a button
- Download list of attributes as Excel file and remove unnecessary attributes by pushing a button
- Read all Information of XML file, compare with attribute list and check for matches, mismatches and missing information in database with attributes. Create IDs for database if necessary
- Do that in mass by not selecting a single XML file but a directory
- Do another comparison with a third database where the information will also be given as an Excel file.
- Prepare Import of missing data by positioning attributes, data description and default values within a new Excel file for a product
- Repeat the step above for a mass of products stored in different files each by selecting a directory

## Installation
No specific installation required. However, all userforms and modules are to be imported into the same Excel data file.

## Download XML Files
As mentioned, all buttons are inserted the first data sheet. On the second data sheet, one can find names of products in column a and their associated addresses in column B. The button for XML file download opens a userform that lets the user look for a product and download its XML file. Before the userform appears, all product names from column A will be stored in an array and this array will hand over its content into the combobox of the userform. When a product is selected, an object called ```MSXML2.DOMDocument``` will be created and the XML file will be downloaded. However, there will be no error, if this XML file does not exist but a file with zero kb will appear.

## Filter Attributes
The database that is used does not allow any kind of grouped filtering to remove unnecessary attributes from a list. However, attributes that shall be used for comparison and matching follow a specific pattern where their ID ends with -Product, -Article or other key words with capital letters. So the entire list of attributes will be downloaded as an excel file and filtered subsequently.

## Create Data file and do Comparison
If the xml file is not too big, it can be parsed with the built in XML import function. Thus, this XML file can be loaded into a userform. Furthermore, the attribute list and another one called "DIM" will be loaded. All three are necessary for a successful data collection. The macro goes through the XML file and stores Data into a new file. This list contains attributes. Thes will be compared with the attribute list and certain components of attributes will be checked for a match. If there is no match then this attribute does not exist and a new ID will be created. The third list checks if dimensioning is needed. The color set of excel is used to make mismatches visible.

## Apply Mass of Data and Mass Comparison
This procedure follows creating data file and doing comparison. However, instead of choosing a single XML file, one can choose an entire directory and loop through all XML files.

## Prepare Import
All the data collected and analyzed with the procedures above have to be imported into database. All data can be imported via Excel files but attributes, default values and characteristics have to be aranged into a certain structure. Furthermore, the shape of all cells that contains data have to be changed so that the data will be recognized by the database software.
Multiple protocols are allowed to be loaded into one input file. So if an input file has been created and is still open and the user wants to load another protocol, all information will be stored additionally in that input file.

## Prepare Mass Import
Instead of loading a single file, load an entire directory, i.e. the directory where the protocols are stored and create an import file by looping through all files.
