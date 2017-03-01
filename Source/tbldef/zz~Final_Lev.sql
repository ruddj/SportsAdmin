CREATE TABLE [zz~Final_Lev] (
  [ET_Code] LONG ,
  [F_Lev] BYTE ,
  [NoHeats] SHORT ,
  [PtScale] VARCHAR (10),
  [ProType] VARCHAR (15),
  [UseTimes] BIT ,
  [ProNum] SHORT ,
   CONSTRAINT ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([ET_Code], [F_Lev])
)
