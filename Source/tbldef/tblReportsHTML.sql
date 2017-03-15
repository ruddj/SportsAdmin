CREATE TABLE [tblReportsHTML] (
  [repID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [repShortCode] VARCHAR (10) CONSTRAINT [repShortCode] UNIQUE ,
  [repTitle] VARCHAR (255),
  [repCaption] VARCHAR (255),
  [repQuery] VARCHAR (255),
  [repFields] VARCHAR (255),
  [repHeaders] VARCHAR (255),
  [repGroup] VARCHAR (255),
  [repGroupHeader] VARCHAR (255),
  [repDisplayLimit] LONG ,
  [repAgeChamp] BIT ,
  [repFinalLev] VARCHAR (255),
  [repPlace] VARCHAR (255)
)
