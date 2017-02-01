# jet-tool

This tool lets you dump MS Access/MS Jet database schema as an SQL file containing the commands needed to recreate the database from the scratch.

It also lets you execute those files, or any other SQL files, or SQL commands from keyboard.

Usage information is available by doing "`jet help`".

## Syntax ##
Syntax:
`jet <command> [params]`

Available comands:

  * `jet help` :: displays help
  * `jet touch` :: connect to database and quit (useful for creating an empty DB or verifying connection)
  * `jet dump` :: dump sql schema
  * `jet exec` :: executes keyboard input. To execute an sql file, pass it as input with "jet exec < filename.sql". See also: `-stdi` option.
  * `jet schema` :: dumps "openSchema()" tables.

In case you somehow cannot redirect the tool input and output through the usual command-line piping, you can specify redirects with these options:

  * `-stdi [filename]`
  * `-stdo [filename]`
  * `-stde [filename]`

There are three ways to specify database connection. Some of them are less preferable than others because they're limited in what you can do. For certain operations this tool uses DAO, which cannot connect to Jet databases by anything other than file name (and cannot connect to ANY databases by ADO ConnectionString).

If the tool cannot use DAO, it'll silently disable the optional features relying on it, but if you explicitly enable those through the command-line, it'll fail.

Database connection options:

  * `-f [file.ext]` :: recommended way as it's supported by ADO, ADOX, DAO.
  * `-dsn [data-source-name]` :: supported by ADO and ADOX.
  * `-c [ado-connection-string]` :: supported by ADO and ADOX.
  * `-u [username]`
  * `-p [password]`
  * `-dp [database password]`
  * `-new` :: creates a new database. This works only with filename type connections.
  * `-force` :: overwrite exitsting database (requires `-new`)

What to include in dumps:

  * `--tables` (or `--no-tables` to disable, same with the rest of dump options)
  * `--views`
  * `--procedures`
  * `--relations` (cross-table links, with or without constraints)
  * `--data`
  * `--query "SQL Query" TableName` (same as `--data`, but records come from the query)

Limit dump to only selected tables: `--data table1,table2` (`--no-tables table1` not allowed).

Shortcuts: `--all`, `--none`, `--default`; may be overriden: `--all --no-procedures`

Additional options:

  * `--comments` :: dump comments or ignore them

By default, Jet gives no way to set table/field comments ("descriptions" you write for them in MS Access) via SQL. They can be read though. If you enable this, they will be dumped in SQL file, but how they'll be dumped, that depends.

By default they're dumped just as comments, which means, when you execute the resulting SQL they will be ignored. With `--private-extensions` on, they're dumped in a private comment format:
```
  /**COMMENT* [content] */
```

If, and only if, you execute the SQL file with Jet-tool, these comments will be parsed and the tool will manually set the field/table/view "descriptions" via DAO. Consequently, you'll also need Filename-type connection, when executing.

Back to dumping options:

  * `--private-extensions` (or `--no-private-extensions`) :: enables/disabled the private comment generation (with `dump`) or parsing (with `exec`)
  * `--drop` :: adds "`DROP TABLE [tablename]`" and "`DROP VIEW [tablename]`" commands before "`CREATE TABLE/VIEW`".
With `--private-extensions` on, these are marked as `/**WEAK**/`, which means errors on these commands will be ignored even in `--stop-on-errors` mode.
  * `--no-create` (`--create` is by default) :: disables the creation commands (if you want only DROP commands)
  * `--enable-if-exists` :: adds `IF EXISTS` to `DROP` commands. Jet does not support this, so only when exporting to other engines.

If you specify none of the above options, default dump takes place:

  `--tables --views --comments --relations --procedures --drop --private-extensions`

Logging options:

  * `--silent` :: do not print anything at all
  * `--verbose` :: echo commands which are being executed and otherwise act like a spammer

Error handling:

  * `--ignore-errors` :: continue even when some of the commands fail
  * `--stop-on-errors` :: exit with error code immediately when a command fails

Input handling in EXEC:

  * `--crlf-break` (`--no-crlf-break`) :: treat linefeeds as command terminators.
By default linefeeds are ignored and only ';' terminates the command. That's fine as long as you parse SQL file, but when you type commands by hand you often want them executed immediately as you hit Enter.

In CRLF-Break mode ';' is still considered a command terminator. '`asd;bsd[return]`' will result in two commands being executed, '`asd`' and '`bsd`'.

By default when the tool is sure input is not redirected (comes from console), these are set:

  `--verbose --ignore-errors --crlf-break`

When the tool is unsure and thinks the input might be coming from file, these are set:

  `--stop-on-errors --no-crlf-break`

You can explicitly set these options via command-line in which case no guessing is made.

## Limitations ##
See [Limitations](Unsupported) of what the tool can support in the database.

## Development ##
This tool is written in Object Pascal and should probably compile on Delphi 2005+ or equiualent FreePascal. If you want to build it from the sources, read [[Building]] notes.

This tool is not being actively developed, but you can contact me if you need an option which is not available.

If you're interested in contributing, feel free to join.