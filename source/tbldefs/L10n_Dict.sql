CREATE TABLE [L10n_Dict] (
  [LangCode] VARCHAR (2),
  [KeyText] VARCHAR (255),
  [LngText] VARCHAR (255),
   CONSTRAINT [PK_L10n_Dict] PRIMARY KEY ([LangCode], [KeyText])
)
