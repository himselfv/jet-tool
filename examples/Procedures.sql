/*
Tests procedure creation and calls.
This requires a fair amount of other functionality working, so run later in the tests.
*/

CREATE TABLE [Users] (
[ID] COUNTER NOT NULL,
[Username] TEXT(64),
[Password] TEXT(64),
[FullName] TEXT(128),
[Age] INTEGER,
[Revision] INTEGER
);

CREATE PROCEDURE NewUser(
 aUsername VARCHAR,
 aPassword VARCHAR,
 aFullName VARCHAR,
 aAge INTEGER
) AS
INSERT INTO [Users](
 [Username], [Password], FullName, Age, Revision
) VALUES (
 aUsername, aPassword, aFullName, aAge, 1
);

EXECUTE NewUser 'Login1', 'PassA', 'User name 1', 25;
EXECUTE NewUser 'Login2', 'PassB', 'User name 2', 26;
EXECUTE NewUser 'Login3', 'PassC', 'User name 3', 27;
EXECUTE NewUser 'Login4', 'PassD', 'User name 4', 29;
EXECUTE NewUser 'Login5', 'PassE', 'User name 5', 29;