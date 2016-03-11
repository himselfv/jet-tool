/*
Tests datatype support in Jet.
1. This must run without errors.
2. The resulting DB must dump more or less to the same SQL.
*/

CREATE TABLE [DataTypeTest_NoDefaults] (
[ID] COUNTER,
[I1] TINYINT,
[I2] SMALLINT,
[I4] INTEGER,
/*[TI8] BIGINT, -- not supported in Jet */
[UI1] BYTE,
/*[UI2] SMALLINT UNSIGNED,
[UI4] INTEGER UNSIGNED, -- unsigned types not supported in Jet */
/*[UI8] BIGINT UNSIGNED, -- not supported in Jet */
[CY] MONEY,
[R4] REAL,
[R8] FLOAT,
[T_GUID] UNIQUEIDENTIFIER,
[T_DATE] DATE,
[T_DECIMAL] DECIMAL(18,4),
[T_BOOL] BIT,
[T_BYTES] BINARY,
[T_BYTESL] LONGBINARY,
[T_TEXT_64] TEXT(64),
[T_TEXT_DYN] TEXT,
[T_TEXT_L] LONGTEXT
);

CREATE TABLE [DataTypeTest_Defaults] (
[ID] COUNTER,
[I1] TINYINT DEFAULT 11,
[I2] SMALLINT DEFAULT 12,
[I4] INTEGER DEFAULT 13,
/*[I8] BIGINT DEFAULT 14, -- not supported in Jet */
[UI1] BYTE DEFAULT 15,
/*[UI2] SMALLINT UNSIGNED DEFAULT 16,
[UI4] INTEGER UNSIGNED DEFAULT 17, -- unsigned types not supported in Jet */
/*[UI8] BIGINT UNSIGNED DEFAULT 18, -- not supported in Jet */
[CY] MONEY DEFAULT 19,
[R4] REAL DEFAULT 20,
[R8] FLOAT DEFAULT 21,
[T_GUID] UNIQUEIDENTIFIER DEFAULT GenGUID(), /* This is how GUID-counters are dumped too */
[T_DATE] DATE DEFAULT Now(),
[T_DECIMAL] DECIMAL(18,4) DEFAULT 24,
[T_BOOL] BIT DEFAULT False,
[T_BYTES] BINARY DEFAULT "26 A", /* Stored but invisible in Access */
[T_BYTESL] LONGBINARY DEFAULT "27 B", /* Stored but invisible in Access */
[T_TEXT_64] TEXT(64) DEFAULT "28 C",
[T_TEXT_DYN] TEXT DEFAULT "29 D",
[T_TEXT_L] LONGTEXT DEFAULT "30 E"
);


CREATE TABLE [DataTypeTest_NotNull] (
[ID] COUNTER NOT NULL,
[I1] TINYINT NOT NULL,
[I2] SMALLINT NOT NULL,
[I4] INTEGER NOT NULL,
/*[I8] BIGINT NOT NULL, -- not supported in Jet */
[UI1] BYTE NOT NULL,
/*[UI2] SMALLINT UNSIGNED NOT NULL,
[UI4] INTEGER UNSIGNED NOT NULL, -- unsigned types not supported in Jet */
/*[UI8] BIGINT UNSIGNED NOT NULL, -- not supported in Jet */
[CY] MONEY NOT NULL,
[R4] REAL NOT NULL,
[R8] FLOAT NOT NULL,
[T_GUID] UNIQUEIDENTIFIER NOT NULL,
[T_DATE] DATE NOT NULL,
[T_DECIMAL] DECIMAL(18,4) NOT NULL,
[T_BOOL] BIT NOT NULL,
[T_BYTES] BINARY NOT NULL,
[T_BYTESL] LONGBINARY NOT NULL,
[T_TEXT_64] TEXT(64) NOT NULL,
[T_TEXT_DYN] TEXT NOT NULL,
[T_TEXT_L] LONGTEXT NOT NULL
);
