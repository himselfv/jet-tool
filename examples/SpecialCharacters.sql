/*
Test special character handling.

Jet escapes quotes and double-quotes by doubling them (' -> '', " -> ""),
and only the kind of quotes used to wrap the string should be escaped.

Strings with any other suspicious characters should just be passed as binary data (Jet supports this).
*/

CREATE TABLE [SpecialCharacterTest1] (
[ID] COUNTER,
[Content] TEXT(64)
);

INSERT INTO [SpecialCharacterTest1] ([Content]) VALUES ('Normal text');
INSERT INTO [SpecialCharacterTest1] ([Content]) VALUES ('Text with '' a single quote');
INSERT INTO [SpecialCharacterTest1] ([Content]) VALUES ('Text with two '''' single quotes');
/* Double quote is not escaped because we use single quotes for wrapping */
INSERT INTO [SpecialCharacterTest1] ([Content]) VALUES ('Text with " a double quote');
INSERT INTO [SpecialCharacterTest1] ([Content]) VALUES ('Text with two "" double quotes');
/* Text as binary data.
This text contains no special characters and will be dumped as normal text (digits and spaces): */
INSERT INTO [SpecialCharacterTest1] ([Content]) VALUES (0x37003500320020003100300032003900);
/* This one has weird characters incl. zero character, and should be dumped verbatim */
INSERT INTO [SpecialCharacterTest1] ([Content]) VALUES (0x37003500320000003100300012001500);
