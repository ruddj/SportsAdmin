CREATE TABLE [zz~Heats] (
  [HE_Code] AUTOINCREMENT,
  [E_Code] LONG ,
  [Heat] SHORT ,
  [PtScale] VARCHAR (10),
  [E_Number] LONG ,
  [E_Time] DATETIME ,
  [F_Lev] BYTE ,
  [Pro_Type] VARCHAR (15),
  [UseTimes] BIT ,
  [Completed] BIT ,
  [Status] BYTE ,
  [AllNames] BIT ,
  [DontOverridePlaces] BIT ,
   CONSTRAINT [Ind1] PRIMARY KEY ([E_Code], [F_Lev], [Heat])
)
