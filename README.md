# <img src="https://github.com/ruddj/SportsAdmin/raw/sa-pages/images/sports.png?raw=true" height=132 alt="Sports Administrator" />

[![Download Latest Release][download-badge]][DOWNLOAD] [![Current Version][version-badge]][CHANGELOG] [![license][license-badge]][LICENSE] 

Sports Carnival Management Software 

This is an easy to use access database for managing results for a school Athletics or Swimming carnival.
 
It has many [features](https://github.com/ruddj/SportsAdmin/wiki/Feature-List) with some of the main ones listed below:
* Easy entry of event results 
* Tracking of event records 
* Cumulative house points 
* Age Champions
* Multiple Reports to print results
* HTML export of results
* Export of top athletes for Meet Manager Swim and Track & Field

This database requires Microsoft Access 2010 or newer. If you do not have this installed you can download the [Microsoft Access 2016 Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=50040) for free.

The database was [originally developed by Andrew Rogers](https://github.com/ruddj/SportsAdmin/wiki/Background), in conjunction with Christian Outreach College, Brisbane. He has generously allowed it to be made open source, to allow further development and allow it to run on more modern systems.

## Download
You can download the most [recent release](https://github.com/ruddj/SportsAdmin/releases/latest) from [Releases tab](https://github.com/ruddj/SportsAdmin/releases/latest).
A list of changes between versions can be seen here [CHANGELOG].

## Upgrading
### From Previous Commercial Release
Before version 5, Sports Administrator was written using Access 97 databases. 
Support for this database format was dropped from Access 2013 and later versions.
To open a carnival file created in the previous version you will need to convert it to a new Access format. 
Instructions for different Access versions provided below.

#### Access 2010 ####
1. Open Database in Access
2. File -> Save & Publish
3. Choose either *Access Database (\*.accdb)* or *Access 2002-2003 Database (\*.mdb)*
4. Click *Save As* and save your file.

#### Access 2003 ####
1. Open Database in Access
2. Click OK to warning
3. Tools -> Database Utilities -> Convert Database -> To Access 2002-2003 File Format..
4. Rename and *Save* your file.

### From Previous Open Source Release
You can download the most recent Sports.accdb from code above and change file extension to .accdr to load in runtime only mode. 

You can reload the list of past carnivals by clicking *Import Carnival List* in Maintain Carnivals.

## Source Code
All coding occurs in the Sports.accdb file. 

The Source folder contains a text dump of contents to allow diff comparison of changes. 
This is generated through use of [msaccess-vcs-integration](https://github.com/timabell/msaccess-vcs-integration) scripts.
To export source run "ExportAllSource" in Immediate window.

To contribute please generate a Pull request against the Source Export and I will import it into main accdb.

For end users you can change the Sports.accdb file extension to .accdr to load in runtime only mode.

## Screens

### Main Screen
![Sports Admin Main Screen](https://github.com/ruddj/SportsAdmin/blob/sa-pages/images/welcome-screen.png?raw=true)


### Setup Carnival Guide
![Setup Carnival Guide](https://github.com/ruddj/SportsAdmin/blob/sa-pages/images/setup-carnival.png?raw=true)


### Competitor Results Entry
![Competitor Results Entry](https://github.com/ruddj/SportsAdmin/blob/sa-pages/images/event-results.png?raw=true)


### Report Options
![Report Options](https://github.com/ruddj/SportsAdmin/blob/sa-pages/images/Reports.png?raw=true)

[CHANGELOG]: ./CHANGELOG.md
[LICENSE]: ./License.txt
[DOWNLOAD]: https://github.com/ruddj/SportsAdmin/releases/latest
[license-badge]: https://img.shields.io/badge/license-MIT-blue.svg

[download-badge]: https://img.shields.io/badge/download-latest-blue.svg
[version-badge]: https://img.shields.io/badge/version-5.2.0-blue.svg