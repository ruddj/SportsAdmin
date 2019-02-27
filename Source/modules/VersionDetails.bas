Option Compare Database
Option Explicit


Global Const VersionNumber = "5.2.2"
Global Const VersionDate = "(27/Feb/2019)"

' Version 5.2.2 - 2019-02-27
' Catch errors in MeetManager export where no time was entered for event
' Correctly close Meet Manager export file handle if function crashes
' Escape now closes Maintain Competitors Form

' Version 5.2.1 - 2018-05-24
' Reordered Ribbon to open on Entry Tab rather than Setup tab
' Recreated Main Menu screen Graphs in newer MS Graph version

' Version 5.2.0 - 2018-02-09

' Access x64 Support. Should now work on Access 2010+ x64.
' In Results entry pressing up and down arrow in name selection will bring up dropdown list of names, making it easier to select name without using mouse.
' When Importing students will verify Age is number and DoB is Date. Useful to remove header line.
' When choosing Sort by Place it would run the script to update Final Status generating an error if setting it to completed.
' Maintain Event Order can now start from 1 when no events have been numbered

' Version 5.1.4 - 2017-07-25
'
' Default location that Import Carnivals List views to Database directory.
' Changed Event Status from constant to enum
' Changed web ordering of Competitor Results based on result within heat level
' Changed Q and Response from globals to locally defined
' Add Results list was not updating after closing an event (Caused by code making forms resizable)
' Improved strong typing by declaring variables functions as types
' Warning if Web export does not fill in data due to faulty template
' Reset extra allocated Team points button with other reset options

' Version 5.1.3
'   (5/June/2017)
'   Meet Manager Export was incorrectly formatting times > 60 secs

' Version 5.1.2
'   (8/May/2017)
'   Meet Manager T&F Event Export - Age Division Mapping

' Version 5.1.1
'   (5/May/2017)
'   Added Swim Meet Manager Export

' Version 5.1.0
'   (4/May/2017)
'   Added Meet Manager Export

' Version 5.0.2
'   (28/Apr/2017)
'   Maintain Competitors listbox now supports type-ahead search
'   Update MsgBox parameters to use names instead of magic numbers

' Version 5.0.1
'   (03/Apr/2017)
'   Updated UI font to Tahoma
'   Fixed bug importing Carinval list from older data file

' Version 5.0.0
'   (01/Apr/2017)
'   Released under MIT license
'   Migrated software from Access 97 .MDB/.MDE to Access 2013/2016 .ACCDB
'       Adjusted code functions to use more modern equivalents (on-going)
'   Added support for .ACCDB carnival files
'   Added Ribbon Interface
'   Adjusted Age Champions - All Div to use student age rather than event age
'   Added temporary trusted location for new datafile to fix security warnings
'   When started in Runtime mode adds current location as trusted to prevent future prompts
'   HTML export VBA script instead of Report hooks. Start using CSS output.
'   Multiple Forms made resizable



' Version 4.1
'   (19/5/2010)
'   Changed the order of the unlimited lane reports back to name order (at request of HCC-V)

' Version 4.0
'   (1/3/2010)
'   Fixed an import problem for carnival disks with OPEN age competitors

' Version 3.9
'   (1/9/2006)
'   Changed look of forms to flat.  Added 3 column program summary.

' Version 3.8
'   (28/8/2006)
'   Updated minor things.

' Version 3.7
'   (23/8/2005)
'   Fixed the reports that were showing Boys when it should have shown mixed.

' Version 3.7
'   (16/2/2003)
'   Changed the way EventAges are determined.  Means the age champion report should work better.
'   (12/2/2003)
'   Fixed bug in determing places across all heats

' Version 3.6
'   (22/7/2002)
'   Added field to stop records from being generated for certain heats.
'   (16/7/2002)
'   Added removal of competitors from events

' Version 3.5
'   (29/10/2001)
'   Completely redid how the places were calculated.  Added checkbox to stop recalculation of places.
'   Added checkbox to allow places to be calculated across entire final level, not just heat
'   Speeded up the promote competitors option
'
'   (29/9/2001)
'   Added quick add of competitor
'   Changed Age data type to Numeric
'   Tidied up various routines
'   Add ID field
'   Added permanent opening of linked database
'   Added preview report from main menu
'   Tidied up the statistical report generation
'
' Version 3.05
'   (14/3/2001)
'   Added jpeg graph export to web page generation
'
' Version 3.04
'   (10/2/2001)
'   Improved the import of carnival disks.  Better error handling.
'
' Version 3.03
'   (31/1/2001)
'   Modified the Team / Event report slightly
'   (13/11/2000)
'   Fixed the Null error in Order competitor events
'
' Version 3.02
'   (10/11/2000)
'   Updated the version number
' Version 3.01
'   (2/11/2000)
'   Fixed the relationship between Heats and CompEvents so that updates would cascade.
'   Fixed the bulk maintain of competitiors missing table bug
'   (17/10/2000)
'   Improved the PopupWindow option for reports by allowing it to remember its last position.
'   Re-Created the installation routine seeing I lost the previous one.
'   (1/9/2000)
'   Fixed Utilities|Remove Empty Heats
'     Based on fields that no longer existed in original table.
'   Modified Generate Reports
'     Maximises reports and added Popup Window menu option.

' Version 3.00
'   Added New Setup Carnival form
'   Added AddEvent Wizard
'   Add various informational messages
'   Checked all Msgbox dialogs
'   Modified relationship creation
'     No longer deletes relationships if they all exists
'   Ordered Competitors handled differently
'     Table held locally.  Does not delete old but overwrites them.
'   Fixed control colours to refelct default system colours.