CREATE TABLE [tabSqlLangFormat] (
  [SqlLang] VARCHAR (20) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [SqlDateFormat] VARCHAR (255),
  [SqlBooleanTrueString] VARCHAR (255),
  [SqlWildCardString] VARCHAR (255)
)
