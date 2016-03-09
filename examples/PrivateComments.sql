/*
Tests private extension comment parsing with jet-tool.
*/

CREATE TABLE [PrivateCommentsTest] (
[ID] COUNTER NOT NULL /**COMMENT* Counter (auto-increment) field */,
[IntValue] INTEGER DEFAULT 0 /**COMMENT* Integer field */,
[StringValue] TEXT(128) /**COMMENT* String field */
) /**COMMENT* Comment text for the whole table */;
