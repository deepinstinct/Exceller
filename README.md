# Exceller
Replaces VBA cells function calls with their content, to make the VBA code less obfuscated and easier to detect by AV vendors.

## Using Exceller
Use Python to run 'exceller.py' and provide it with the following parameters:
* vba_file: a text file containing VBA code. You can use oledump to extract VBA code from Office files.
* excel_file: a path to the Excel file from which the VBA code was extracted. **Note: currently, Exceller only supports OOXML Excel files. OLE support will be added at a later stage.**
* edited_vba_file: the path in which you want to save Exceller's output, which will be a text file containig a less obfuscated version of 'vba_file'


## Extracting VBA Code from Office Files
As mentioned above, you can use oledump to extract VBA from Office files.
Example: `python oledump.py -v -sa my_office_file> my_vba_output`

Link to Didider Stevens' oledump: https://github.com/DidierStevens/DidierStevensSuite/blob/master/oledump.py

## An Example for a Deobfuscation Performed by Exceller

![Alt text](Deobfuscation_Flow.png?raw=True)
