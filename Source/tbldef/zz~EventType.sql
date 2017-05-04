CREATE TABLE [zz~EventType] (
  [ET_Code] AUTOINCREMENT CONSTRAINT [ET_Code] UNIQUE  CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ET_Des] VARCHAR (30),
  [Units] VARCHAR (10),
  [Lane_Cnt] SHORT ,
  [R_Code] LONG ,
  [Include] BIT ,
  [EntrantNum] SHORT ,
  [Flag] BIT ,
  [PlacesAcrossAllHeats] BIT ,
  [Mevent] VARCHAR (10)
)
