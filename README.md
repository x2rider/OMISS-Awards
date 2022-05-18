# OMISS-Awards
Excel and LibreOffice calc templates to calculate OMISS awards based on imported ADI contact logs.
For more information on these awards and how to become a member, please visit https://www.omiss.net/

A video walkthrough on how to use the spreadsheets is located here: https://www.youtube.com/watch?v=EBXyA3ZQAcg

Originally written by Mike Maddox and recently updated by Paul Reedy (KA5PMV), this template will read your exported ADI files from Netlogger and generate results for each OMISS award.

* Download the files.  
* Add your call sign and OMISS number to the OC_Config tab, 
* Import your ADI file from Netlogger (or other logging software such as N3FJP logger using the "Other" field using the format: 
`"OMISS OM#"` where OM# is the Omiss number of the station you worked)'
* Click "Run All Awards" button, to see how close you are to completing all of the awards.
* From then on, just add your new contacts from each day as you make them, select your new contacts ADI file, and read those in.  Do Not import the same contacts again, as that will create duplicates.

## NEW!! LibreOffice OMISS menu bar
A new LibreOffice menu bar has been added, which has similar features as the Excel Add-In without having to install anything.

An Excel Add-In is available in the Excel folder to make it easier to run all awards and selected scripts.
View how to install and use the Add-In here https://youtu.be/2YXONHvJadY

To Install:
* Exit Excel
* Download OMISS.xlam and InstallOmissAdd-in.cmd to the your computer.
* Right click on OMISS.xlam and select Properties.
* Check the "Un Block" checkbox to remove the "Mark of the Web".  If you don't do this, the tool will not load.
* Double-Click the install command file, it should put the tool in the %APPDATA%\Roaming\Microsoft\AddIns directory

## Compatibility

This has been tested on Windows and Mac with the latest version of Excel and LibreOffice Calc.
It does not work on Excel 2010 or older.

NOTE:  While the Alternative to Excel was originally OpenOffice, There is a bug with OpenOffice, which prevents the full membership roster from loading, which will affect awards.  Please use LibreOffice instead of OpenOffice.

Paul - KA5PMV
