CREATE TABLE [ReportTypes] (
  [R_Code] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Desc] VARCHAR (50),
  [Report] VARCHAR (50),
  [EventReport] BIT ,
  [LimitedLanes] BIT ,
  [Flag] BIT ,
  [Relay] BIT ,
  [SummaryReport] VARCHAR (50)
)
