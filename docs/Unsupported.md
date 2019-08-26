# Introduction #
This page lists various limitations and unsupported configurations for the Jet tool.

## Comments ##
Comments are exported only in private-extensions mode.

## Foreign keys ##
Foreign keys are exported unless they're "dummy" foreign keys (pure table relationships, without any integrity checks attached). It is impossible to create dummy foreign keys with SQL, and a method of doing so via DAO/private extensions is not implemented yet.

For the time being, for dummy foreign keys, creation commands are generated which create "integrity-verified foreign keys" instead of "dummy foreign keys". These commands are written as comments in the SQL.

There should be no impact on the database from losing such a foreign key, except for aesthetic impact from not seeing the relation anymore.

## Check constraints ##
Check constraints are not exported at all yet.

## Database schema ##
Visual database schema is not exported at all. Still, all the relations between the tables are exported (unless they're subject to other limitations). The only thing that's lost is visual arrangement of tables on the schema.

There should be no impact on the database from losing visual schema informaton.

## GUID autoincrement fields ##
These are exported just fine, but as Jet DDL has no way of properly defining a COUNTER GUID field, they show up in the resulting database as "GUID fields with DEFAULT=GenGuid()" instead of "COUNTERS".

This does not affect how they function.