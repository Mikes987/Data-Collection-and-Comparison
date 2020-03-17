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

## Download XML Files
As mentioned, all buttons are inserted the first data sheet. On the second data sheet, one can find names of products in column a and their associated addresses in column B. The button for XML file download opens a userform that lets the user look for a product and download its XML file. Before the userform appears, all product names from column A will be stored in an array and this array will hand over its content into the combobox of the userform.
