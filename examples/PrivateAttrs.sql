/*
Tests attributes supported by jet-tool's private extensions.
*/

/* WEAK: ignore failure. This should not crash. */
DROP TABLE [NameOfNonExistentTable] /**WEAK**/;
DROP TABLE [NameOfNonExistentTable] /**WEAK**/;