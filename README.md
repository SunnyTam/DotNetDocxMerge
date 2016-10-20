# DotNetDocxMerge
This is a gui project for merge data to  Microsoft Word document(.docx) from csv. I start this program because there is no solid program to solve mail merge in gui

## Features
* Merge from csv
* Auto new page
* Will not change text format
* Keep template size/margins
* Progress bar display


## How to use
All you need is the  Microsoft Word document (.docx) which act as a template and an Excel (.csv) file which contains rows of data to bind to template.

The word document should contain "\<\<yourData\>\>" as a template marker. e.g. \<\<myName\>\> \<\<id-no\>\>

The first row of csv will count as header. Please make sure each header exist only once. 

## How to build
Build in Visual Studio.

## Screens
![alt tag](https://raw.githubusercontent.com/SunnyTam/DotNetDocxMerge/master/DotNetDocxMerge/mail-merge-generator.png)

## How to config
You can setup your default path at DotNetDocxMerge/DotNetDocxMerge/App.config
```
        <DotNetDocxMerge.Properties.Settings>
            <setting name="template" serializeAs="String">
                <value>C:\yourpath\template.docx</value>
            </setting>
            <setting name="csv" serializeAs="String">
                <value>C:\yourpath\data.csv</value>
            </setting>
            <setting name="dist" serializeAs="String">
                <value>C:\yourpath\result.docx</value>
            </setting>
        </DotNetDocxMerge.Properties.Settings>
```
