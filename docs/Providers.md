## Short version

Use `-f filename.mdb` / `-f filename.accdb` and everything will just work.

MDB: This tool will likely work on XP+ out of the box.

ACCDB: You need Office 2010 (32 bit) or later installed, or install various providers separately (see below).


## Long version

This tool supports several database formats (ACCDB and several versions of MDB). To access these formats it uses a number of _access methods_:

* ADO/OLEDB: required.

* ADOX: optional, needed for database creation.

* DAO: optional, needed for some extended functions.

Several providers can be present for each of these methods:

* ADO/ADOX both need a Jet/ACE OLEDB provider (12.0 or 4.0)
* DAO needs DAO library (12.0 or 3.6)

Older providers (Jet OLEDB 4.0 and DAO 3.6) are universally available and support older databases (MDB) better but do not support newer databases (ACCDB).

Newer providers may need to be installed with Office or separately.

This tool auto-selects the best available provider for the job, but you can override this (see below).


### Connection options

There are three ways to connect to a database with this tool:

* By file name

* By data source name

* By connection string

File name is preferred and is the most functional and automatic of all. All access methods can be employed by file name.

Data Source name is acceptable but somewhat less compatible.

Finally, with a Connection String you can specify a lot of things by hand but the corresponding settings (OLEDB/DAO providers, user name, password) are ignored. Only ADO/ADOX is available in this mode.


### OLEDB Providers

The following providers are supported by default:

  * `Microsoft.Jet.OLEDB.4.0` (MDB only)
  * `Microsoft.ACE.OLEDB.12.0` (MDB/ACCDB)

Jet.OLEDB.4.0 provider is almost universally available on Windows XP and later.

To access accdb databases you need ACE.OLEDB provider 12.0. It is installed with Office 2010 or later or can be downloaded separately, e.g. here:
  http://www.microsoft.com/en-us/download/details.aspx?id=13255

The tool auto-selects the best available provider. It prefers Jet.OLEDB.4.0 for MDB files (for compatibility) and ACE.OLEDB.12.0 for ACCDB, but will fall back to other versions if needed.

To override this, pass `--oledb-eng [provider id]` (IDs are listed above). You may pass unrelated OLEDB providers if you're adventurous, but some features will not work.


#### ACE OLEDB bitness confusion

Due to the way Microsoft set this up, you can only have ONE bitness version of ACE.OLEDB.12 provider installed - either 32 bit (x86) or 64 bit. Not both (except see below).

This tool needs the 32 bit provider when compiled as 32 bit.

Go here, lament the things that Microsoft has done and figure out what to do in your case:
  https://social.msdn.microsoft.com/Forums/en-US/abf34eea-1029-429a-b88e-4671bffcee76/why-cant-32-and-64-bit-access-database-engine-aceoledb-dataproviders-coexist?forum=adodotnetdataproviders

The only way to install both bitness versions is to install, say, Office 2013 64 bit and a separate 32 bit 2010 provider.


### DAO Engine

The following DAO engines are supported:

  * `DAO.Engine.36` (32 bit only, MDB only)
  * `DAO.Engine.120` (32/64 bit, MDB/ACCDB)

DAO.Engine.36 is almost universally present on Windows XP and later. DAO.Engine.120 may need to be installed with Office 2010 and later.

The tool auto-selects the best available engine. It prefers DAO.Engine.36 for MDB files (for compatibility) and DAO.Engine.120 for ACCDB, but will fall back to any engine it finds.

To override this, pass `--dao-eng [engine progid]` (the above engine names are ProgIDs). DAO engine version needs not to match OLEDB engine version.
