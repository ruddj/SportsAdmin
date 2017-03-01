CREATE TABLE [zzz~Relationships Main] (
  [R ID] LONG  CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Relationship Name] VARCHAR (50) CONSTRAINT [Relationship Name] UNIQUE ,
  [First Table] VARCHAR (50),
  [Second Table] VARCHAR (50),
  [Type] LONG ,
  [Description] VARCHAR (255)
)
