CREATE TABLE [usys_AppFiles] (
  [id] VARCHAR (255) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [version] VARCHAR (10),
  [file] LONGBINARY,
  [SccRev] VARCHAR (50),
  [url] VARCHAR (255)
)
