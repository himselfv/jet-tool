/*
Auto-increment fields, all supported types listed explicitly.
*/
CREATE TABLE [AutoIncTypeTest] (
[ID_I4] COUNTER NOT NULL,
[ID_GUID] UNIQUEIDENTIFIER DEFAULT GenGUID()
)
