CREATE TABLE [zz~House] (
  [H_Code] VARCHAR (7) CONSTRAINT [H_Code] UNIQUE  CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [H_NAme] VARCHAR (50),
  [HT_Code] LONG ,
  [Include] BIT ,
  [Details] LONGTEXT ,
  [Lane] SHORT ,
  [CompPool] LONG ,
  [Flag] BIT ,
  [H_ID] AUTOINCREMENT
)
