Some notes about building the project from the sources

To build the Jet tool, you'll need to import DAO_TLB. To do that, open Component> Import Component, choose "Type Library" and select "Microsoft DAO 3.6 Object Library".

You might also need ADOX_TLB, for this import "Microsoft ADO Ext. 2.8 for DDL and Security".

It's recommended to build this in 32 bit, but 64 bit versions will work too. ([DAO 3.6 provider](Providers.md) is unavailable for 64 bit so you'll need DAO 12.0).
