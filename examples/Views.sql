/*
Views are implemented in Access as simple procedures.
*/

CREATE TABLE [ViewTestTable] (
[ID] COUNTER NOT NULL,
[IntValue] INTEGER DEFAULT 0,
[StringValue] TEXT(128)
);
INSERT INTO [ViewTestTable] ([IntValue], [StringValue]) VALUES (10, 'Item A');
INSERT INTO [ViewTestTable] ([IntValue], [StringValue]) VALUES (20, 'Item B');
INSERT INTO [ViewTestTable] ([IntValue], [StringValue]) VALUES (30, 'Item C');
INSERT INTO [ViewTestTable] ([IntValue], [StringValue]) VALUES (40, 'Item D');
INSERT INTO [ViewTestTable] ([IntValue], [StringValue]) VALUES (50, 'Item E');

CREATE PROCEDURE [ViewTestView] AS
SELECT [ID], [IntValue], [StringValue] AS TextValue
FROM [ViewTestTable]
WHERE [IntValue] > 25;