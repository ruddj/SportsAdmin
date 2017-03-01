CREATE TABLE [ReportList] (
  [ReportName] VARCHAR (255) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ReportCaption] VARCHAR (255),
  [Open] BIT 
)
