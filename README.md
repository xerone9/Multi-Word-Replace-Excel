# Multi-Word-Replace-Excel

When you have to replace MULTIPLE VALUES into MULTIPLE VALUES

Microsoft Excel comes with so many features yet it has so many limits in its options. If you want to replace a value into another value it can easily be done by pressing CTRL + H and replace but what if you have 10 values that is needed to be replace in another 10 values. Let say you have 100 names and you have another file that contains 100,000 names and their NIC numbers. Now you need to replace those 100 names into their NIC numbers. How’d you do that?
There are many tricks for that but every single trick comes with a downside as well. Let say you add a macro. The macro will search for values and replace them accordingly but the downside is let say you have 3 names “Ana”, “Smith” and “Ana Smith”. Most of the macros I found on internet replaces the last name Ana Smith with 2 NIC numbers of Ana and Smith.
Other trick is to copy both sheets into one sheet and then select duplicate values in both sheets. That will change the color of duplicated values and you can resort it. But again what if the initial sheet (that contained 100 names) has also duplicate values inside them. Then the result of duplicated values in both sheet will also give inaccurate results

Working:

Prepare an excel sheet that contains the values needed to be translated and make sure in excel file values are in first sheet (Best practice is to create a new file). Then add the excel file (Dictionary) that contains translations/replacements make sure again that in excel file values are in first sheet. 1st column contain the actual values (the values those are common in both sheets) where as the 2nd column contains the translations). Press Go.
At the end of the process it will open the initial file (The file that contains values those are needed to be translated) with the translation on their very right next cell

https://www.youtube.com/watch?v=AzhZdptpao4

How it Works:

The code is written in python. With the help of py2pdf library it’ll start a loop that reads the value from the initial file and read it on the second file once the value is found it will copy the value present at the very next cell of the 2nd sheet and add it to the initial sheet’s very next column

Compatibility:

Requires Windows 7 SP1 or above

Requires Microsoft Office 2010 or any Excel File Reader as the output file will be in .xlsx format

Password: softwares.rubick.org
