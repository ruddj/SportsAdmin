# Change Log
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [Unreleased]

## [5.3.0] - 2023-12-05
### Added
- Events can now have a range of ages, use - to separate, e.g. 14-16

## [5.2.2] - 2019-02-27
### Fixed
- Catch errors in MeetManager export where no time was entered for event
- Correctly close Meet Manager export file handle if function crashes

### Changed
- Escape now closes Maintain Competitors Form

## [5.2.1] - 2018-05-24
### Changed
- Reordered Ribbon to open on Entry Tab rather than Setup tab
- Recreated Main Menu screen Graphs in newer MS Graph version

## [5.2.0] - 2018-02-09
### Added
- Access x64 Support. Should now work on Access 2010+ x64. 

### Changed
- In Results entry pressing up and down arrow in name selection will bring up dropdown list of names, making it easier to select name without using mouse.
- When Importing students will verify Age is number and DoB is Date. Useful to remove header line.

### Fixed
- When choosing Sort by Place it would run the script to update Final Status generating an error if setting it to completed.
- Maintain Event Order can now start from 1 when no events have been numbered

## [5.1.4] - 2017-07-25
### Changed
- Default location that Import Carnivals List views to Database directory. 
- Changed Event Status from constant to enum
- Changed web ordering of Competitor Results based on result within heat level

### Fixed
- Add Results list was not updating after closing an event (Caused by code making forms resizable)
- Improved strong typing by declaring variables functions as types
- Changed Q and Response from globals to locally defined
- Started enabling Option Explicit for some forms to check variable declarations

### Added
- Warning if Web export does not fill in data due to faulty template
- Reset extra allocated Team points button with other reset options

## [5.1.3] - 2017-06-05
### Fixed
- Meet Manager Export was incorrectly formatting times > 60 secs
- Installer Desktop Shortcuts

## [5.1.2] - 2017-05-08
### Added
- Meet Manager Export. New Age -> Division mapping form to allow exporting event results to Meet Manager Track and Field.

### Changed
- Database Schema of Carnival files modified to add Meet Manager Mappings. 
- New field added in CompetitorEventAge: Mdiv. 

### Fixed
- Adding text fields in schema for MeetManager added a default vale of False, instead of blank.

## [5.1.1] - 2017-05-05
### Added
- Competitor export for Meet Manager Swimming. Can now generate a RE1 Registration file for import into Meet Manager Swimming.

### Fixed
- Meet Manager T&F competitor only export was including duplicates.

## [5.1.0] - 2017-05-04
### Added
- Meet Manager Export. Can now generate a semicolon delimited file for import into Meet Manager Track and Field.

### Changed
- Database Schema of Carnival files modified to add Meet Manager Mappings. 
 - New fields added in Miscellaneous: Mteam,  Mcode, Mtop. 
 - New fields added in EventType: Mevent.
 - Updated Sample Databases

## [5.0.2] - 2017-04-28
### Added
- Maintain Competitors listbox now supports type-ahead search

### Changed
- Update MsgBox parameters to use names instead of magic numbers

### Security
- In Runtime prompt user to add Sports DB folder as trusted location

## [5.0.1] - 2017-04-03
### Changed
- Updated UI font to Tahoma

### Fixed
- Fixed bug importing Carinval list from older data file

## [5.0.0] - 2017-04-01
### Added
- Released under MIT license
- Migrated software from Access 97 .MDB/.MDE to Access 2013/2016 .ACCDB
- Support for .ACCDB carnival files
- Ribbon Interface
- Temporary trusted location for new datafile to fix security warnings
- When started in Runtime mode adds current location as trusted to prevent future prompts
- CSS Support for HTML ouput
- Multiple Forms made resizable

### Changed
- Adjusted Age Champions - All Div to use student age rather than event age
- HTML export VBA script instead of Report hooks. Start using CSS output.

### Fixed
- Multiple minor bugs relating to migration changes in Access 97 to 2013 code

## Following changes were made by Andrew Rogers
## 4.1.0 - 2010-05-19
### Changed
- Changed the order of the unlimited lane reports back to name order (at request of HCC-V)

## 4.0.0 - 2010-03-01
### Fixed
- Fixed an import problem for carnival disks with OPEN age competitors

## 3.9.0 - 2006-09-01
### Added
- Added 3 column program summary.

### Changed
- Changed look of forms to flat.  

## 3.8.0 - 2006-08-28
### Changed
- Updated minor things.

## 3.7.2 - 2005-08-23
### Fixed
 - Fixed the reports that were showing Boys when it should have shown mixed.

## 3.7.1 - 2003-02-16
### Changed
- Changed the way EventAges are determined.  Means the age champion report should work better.

## 3.7.0 - 2003-02-12
### Fixed
- Fixed bug in determing places across all heats

## 3.6.1 - 2002-07-22
### Added
- Added field to stop records from being generated for certain heats.

## 3.6.0 - 2002-07-16
### Added
- Added removal of competitors from events

## 3.5.2 - 2001-10-29
### Added
- Added checkbox to allow places to be calculated across entire final level, not just heat

### Changed
- Completely redid how the places were calculated.  Added checkbox to stop recalculation of places.
- Speeded up the promote competitors option

## 3.5.1 - 2001-09-29
### Added
- Added quick add of competitor
- Add ID field
- Added permanent opening of linked database
- Added preview report from main menu

### Changed
- Changed Age data type to Numeric

### Fixed
- Tidied up various routines
- Tidied up the statistical report generation

## 3.5.0 - 2001-03-14
### Added
- Added jpeg graph export to web page generation

## 3.4.0 - 2001-02-10
### Changed
- Improved the import of carnival disks.  Better error handling.

## 3.3.1 - 2001-01-31
### Changed
- Modified the Team / Event report slightly

## 3.3.0 - 2000-11-13
### Fixed
- Fixed the Null error in Order competitor events

## 3.2.0 - 2000-11-10
### Changed
- Updated the version number

## 3.1.2 - 2000-11-02
### Fixed 
- Fixed the relationship between Heats and CompEvents so that updates would cascade.
- Fixed the bulk maintain of competitiors missing table bug

## 3.1.1 - 2000-10-17
### Changed
- Improved the PopupWindow option for reports by allowing it to remember its last position.
- Re-Created the installation routine seeing I lost the previous one.

## 3.1.0 - 2000-09-01
### Fixed 
- Fixed Utilities|Remove Empty Heats.  Based on fields that no longer existed in original table.

### Changed
- Modified Generate Reports:  Maximises reports and added Popup Window menu option.

## 3.0.0 - 2000-08-01
### Added
- Added New Setup Carnival form
- Added AddEvent Wizard
- Add various informational messages

### Fixed 
- Checked all Msgbox dialogs
- Fixed control colours to refelct default system colours.

### Changed
- Modified relationship creation: No longer deletes relationships if they all exists
- Ordered Competitors handled differently: Table held locally.  Does not delete old but overwrites them.


[Unreleased]: https://github.com/ruddj/SportsAdmin/compare/v5.3.0...HEAD
[5.3.0]: https://github.com/ruddj/SportsAdmin/compare/v5.2.2...v5.3.0
[5.2.2]: https://github.com/ruddj/SportsAdmin/compare/v5.2.1...v5.2.2
[5.2.1]: https://github.com/ruddj/SportsAdmin/compare/v5.2.0...v5.2.1
[5.2.0]: https://github.com/ruddj/SportsAdmin/compare/v5.1.4...v5.2.0
[5.1.4]: https://github.com/ruddj/SportsAdmin/compare/v5.1.3...v5.1.4
[5.1.3]: https://github.com/ruddj/SportsAdmin/compare/v5.1.2...v5.1.3
[5.1.2]: https://github.com/ruddj/SportsAdmin/compare/v5.1.1...v5.1.2
[5.1.1]: https://github.com/ruddj/SportsAdmin/compare/v5.1.0...v5.1.1
[5.1.0]: https://github.com/ruddj/SportsAdmin/compare/v5.0.2...v5.1.0
[5.0.2]: https://github.com/ruddj/SportsAdmin/compare/v5.0.1...v5.0.2
[5.0.1]: https://github.com/ruddj/SportsAdmin/compare/v5.0.0...v5.0.1
[5.0.0]: https://github.com/ruddj/SportsAdmin/tree/v5.0.0
