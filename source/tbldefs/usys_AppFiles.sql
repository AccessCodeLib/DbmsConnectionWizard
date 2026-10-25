CREATE TABLE [usys_AppFiles] (
  [id] VARCHAR (20) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [version] VARCHAR (10),
  [file] LONGBINARY,
  [SccRev] VARCHAR (20),
  [url] VARCHAR (255)
)
