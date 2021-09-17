# REDCap Versions

## Brief

This is a VBA script for use with Microsoft Excel in order to produce a historical progression of the current fields in your REDCap project.

## Overview

REDCap is a data collection package that operates via a web interface. The software is available free of charge, but is licensed; you have to be a qualifying institution. For more information on REDCap visit [Project REDCap](https://projectredcap.org/)

Projects in REDCap are updated regularly with changes to the project form structure being saved in data dictionaries which are preserved historically in

Project Home > Project Revision History

Combine each of these historical data dictionaries into a single spreadsheet, use the script to explore differences and produce a detailed HTML file that lists each of the current fields and how it has been altered over time.

## Compatibility

This code has been developed using the latest version of Microsoft Excel 64 bit, but should work on older 32bit versions.

## Setup

### Download all of the historical data dictionaries

Project Home > Project Revision History

### Add each downloaded data dictionary into the one Excel file

You should use the example .xlsm file which already has the VBA file added, unless you're happy to add the VBA module yourself.

### Rename each tab to r1 .. rNN

By default the name of the sheet might be something like "20180421023845_DataDictionary_71a90f" but should be the letter r with the revision number at the end, starting with "r1".

### Format "Revisions" tab

I like to put a lot of information here, but the code will use column 1 (revision display name) and column 3 (revision date)

Example:
Production revision #1    r1    12-03-2020
Production revision #2    r2    13-03-2020
Production revision #3    r3    14-03-2020

The first row of this data (Production revision #1) should start on row 3 (or edit VBA "Revisions_Start_Row")

### Method to activate the VBA code

In the example, there is a button on the "Notes" tab that activates the "BuildCodebook" sub by a macro call.

## Operation

Check if there are any new revisions to download and add to the spreadsheet.

Check the embedded VBA code for the correct values in the constants
Developer > Visual Basic 

```vba
Const Base_Rev = 10 ' The most recent revision number
Const Revisions_Start_Row = 3 ' row on Revisions tab that contains r1
Const OutFileFolder = "c:\tmp\"  ' HTML output file is saved here, ensure ends with \
Const OutFileNamePrefix = "CodeBookVersions" ' the start of the file name
```

If you are using the example xlsm then Click the button on the "Notes" tab.

## NOTE

To add/edit VBA, you will need to reveal the developer tab: File tab > Options > Customize Ribbon. Under Customise the Ribbon and under Main Tabs, select the Developer check box.

Can take a long time to process, the time take will be listed at the end of the html file.

Fields lists are only those that appear in the current version (the base revision), fields that have been deleted will not be listed.

## Contributing

I use this for my own projects and as a utility for projects that contact me when they run into the "gateway timeout" issue.

Please, if you have any ideas, just [open an issue][issues] and tell me what you think.

If you'd like to contribute, please fork the repository and make changes as you need.

## Licensing

The script is provided as is. This project is licensed under Unlicense license. This license does not require you to take the license with you to your project.
