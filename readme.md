~ Lien Autofiller ~
Check the included template file. Note that the company names and addresses (see header)
have been redacted and should be replaced before using. 

Input is in the form of a list object containing another list with all of the lien details:
Usage:   [Company name, date, suite #, $ amount, $ amount with commas and decimal]
Example: ['Enterprise', '08/22/2024', '100', 1234.56, '1,234.56']
All fields but the $ amount are strings

The script uses the docx library to scan the document for specific strings to be replaced.
Comtypes is used to open the doc in word and export as pdf. This library is WINDOWS ONLY
Which means the script must be ran in POWERSHELL

Output is placed in the doc_files folder as both a pdf and .docx with the same name.
