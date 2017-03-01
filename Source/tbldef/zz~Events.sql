CREATE TABLE [zz~Events] (
  [E_Code] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ET_Code] LONG ,
  [Sex] VARCHAR (1),
  [Age] VARCHAR (10),
  [nRecord] DOUBLE ,
  [Include] BIT ,
  [Record] VARCHAR (15),
  [RecName] VARCHAR (50),
  [RecHouse] LONG 
)
