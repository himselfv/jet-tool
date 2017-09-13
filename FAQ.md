# FAQ

#### **Q**: What database formats are supported?

**A**: mdb (all versions: 4.x, 3.x, 2.0, 1.1, 1.0), accdb. See `jet help` for the latest info.


#### **Q**: Can I create Jet databases of older versions (3.x, 2.0, ...)

**A**: Yes, `--db-format [jet3x / jet20 / ...]`.


#### **Q**: Is .accdb format supported?

**A**: Yes, but not its extended features. You can do everything you can do with MDB.


#### **Q**: I'm getting "provider not found" (0x800A0E7A) or "database not supported" errors with .accdb files.

**A**: To access accdb databases you need ACE.OLEDB provider 12.0 and DAO.Engine.120. It's installed with Office 2010 and later. [See here](Providers) for details.


#### **Q**: I'm using a x64 version of this tool and I'm getting "provider not found" errors with any files.

**A**: x64 version [needs DAO.Engine.120](Providers) even for MDB databases.