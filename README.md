# Language Checker

Language Checker uses a terminology sheet to collate the user inputs such as the search term and its suggested replacement. Language Checker uses the terminology sheet to parse the source XML files and generate a detailed report regarding its findings. The report captures the current context and the suggested context of the search terms. You can use the information contained in this report to make an informed decision on whether or not to replace the offensive occurrence of the term. You can update the report to indicate whether you wish to make the suggested replacement in the source file. You can also edit the replacement phrase in the report so that when the offensive term is replaced in the source file, it can be done so by considering the context of its occurrence as well. 
Tip: When the report is generated, Language Checker only replaces the search term with the suggested replacement term in the report. If you notice any inaccuracies in the replacement phrase (for example use of articles around the replaced term), you can update the report to correct it.

Once you are done reviewing the report, correcting it wherever required, and indicating whether or not you wish to replace the offensive occurrence (by putting a simple ‘yes’ in the specified column of the report), you can use Language Checker again to implement the replacements in the source files. Language Checker generates a detailed log of all the replacements done in the source file. This log comes in handy if you wish to undo the replacements done in future. 
Tip: To undo a replacement, you only need to switch the contents of the Original Phrase and Suggested Phrase columns in the report and run the Language Checker again using the updated report.

Using Language Checker:
The input terminology sheet must have two sheets, Sheet1 and Sheet2. If you rename the sheets, you need rename them in the script as well. Sheet1 is where you need to give your inputs.

Required Libraries:
Install the following libraries using a package manager like pip.
1. openpyxl
2. PySimpleGUI

 To find and replace the deprecated terms in the source files:
1.	Double click the easygui.py file.
2.	Enter the documentation folder path and the input terminology sheet path. The terminology sheet path should include full file name of the sheet, including the file extension:
3.	Click Find. Language Checker saves its findings in Sheet2 of the input terminology sheet.
4.	After looking at the context of the findings reported in Sheet2, type yes in the column F (Change?) if you wish to implement the suggested change. Alternatively, you can also update the Suggested Phrase column and enter yes in the Change? Column. 
5.	Save the input terminology sheet.
6.	To replace the instances against which you entered “yes” (in Change? Column), click the Replace button.
Language Checker replaces all the chosen occurrences and generates a log indicating the changes it made.
