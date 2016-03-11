/*
Tests for a bug in comment handling code where --style comments could fail to be stripped when they were the last thing in the file, or the last thing before the end of the command (EOFS weren't treated as --comment terminators).
Currently fixed.
*/

CREATE TABLE [AccessAudit] (
[ID] COUNTER NOT NULL
); -- commented;