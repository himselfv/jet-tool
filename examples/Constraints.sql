/*
Constraints are a way to define relations between tables. Some constraints may enforce rules, which is why they are so called.

Constraints are the principal part of the database schema, but not all of it. Schema also defines graphical positions of the tables on a map. There's no way to preserve this information, but it does not affect the database functions.

Constraints require that "one" sides have unique indices.

This file defines a sample database with SecurityPrincipals (Users and UserGroups) that can be granted Permissions to access Resources (only one type is Document). The access is to be Audited.
*/

CREATE TABLE [SecurityPrincipals] (
[ID] COUNTER NOT NULL,
[DisplayName] TEXT(128)
);
CREATE UNIQUE INDEX [ID] ON [SecurityPrincipals] (ID ASC) WITH PRIMARY DISALLOW NULL;

-- User = class(SecurityPrincipal)
CREATE TABLE [Users] (
[ID] INTEGER NOT NULL,
[Username] TEXT(64),
[Password] TEXT(64),
[FullName] TEXT(128),
[Age] INTEGER
);
CREATE UNIQUE INDEX [ID] ON [Users] (ID ASC) WITH PRIMARY DISALLOW NULL;
ALTER TABLE [Users] ADD CONSTRAINT [UserIsSecurityPrincipal] FOREIGN KEY (ID) REFERENCES [SecurityPrincipals] (ID) ON UPDATE CASCADE ON DELETE CASCADE;

-- UserGroup = class(SecurityPrincipal)
CREATE TABLE [UserGroups] (
[ID] INTEGER NOT NULL
);
CREATE UNIQUE INDEX [ID] ON [UserGroups] (ID ASC) WITH PRIMARY DISALLOW NULL;
ALTER TABLE [UserGroups] ADD CONSTRAINT [UserGroupIsSecurityPrincipal] FOREIGN KEY (ID) REFERENCES [SecurityPrincipals] (ID) ON UPDATE CASCADE ON DELETE CASCADE;

CREATE TABLE [UserGroupMembership] (
[ID] COUNTER NOT NULL,
[GroupID] INTEGER NOT NULL,
[UserID] INTEGER NOT NULL
);
ALTER TABLE [UserGroupMembership] ADD CONSTRAINT [UserGroupMembershipGroupID_References_UserGroup] FOREIGN KEY (GroupID) REFERENCES [UserGroups] (ID) ON UPDATE CASCADE ON DELETE CASCADE;
ALTER TABLE [UserGroupMembership] ADD CONSTRAINT [UserGroupMembershipUserID_References_User] FOREIGN KEY (UserID) REFERENCES [Users] (ID) ON UPDATE CASCADE ON DELETE CASCADE;


CREATE TABLE [Resources] (
[ID] COUNTER NOT NULL
);
CREATE UNIQUE INDEX [ID] ON [Resources] (ID ASC) WITH PRIMARY DISALLOW NULL;

-- Document = class(Resource)
CREATE TABLE [Documents] (
[ID] INTEGER NOT NULL,
[Name] TEXT(64),
[Contents] TEXT,
[CreatorID] INTEGER  -- can be null if Creator is deleted or unknown
);
CREATE UNIQUE INDEX [ID] ON [Documents] (ID ASC) WITH PRIMARY DISALLOW NULL;
CREATE INDEX [CreatorID] ON [Documents] (ID ASC);
ALTER TABLE [Documents] ADD CONSTRAINT [DocumentIsResource] FOREIGN KEY (ID) REFERENCES [Resources] (ID) ON UPDATE CASCADE ON DELETE CASCADE;
ALTER TABLE [Documents] ADD CONSTRAINT [DocumentsCreatorID_References_User] FOREIGN KEY (ID) REFERENCES [Users] (ID) ON UPDATE CASCADE ON DELETE SET NULL; -- just reset the link if the user is deleted

CREATE TABLE [AccessPermissions] (
[ID] COUNTER NOT NULL,
[PrincipalID] INTEGER NOT NULL,
[ResourceID] INTEGER NOT NULL,
[Allow] INTEGER,
[Deny] INTEGER
);
ALTER TABLE [AccessPermissions] ADD CONSTRAINT [AccessPermissionsPrincipalID_References_SecurityPrincipal] FOREIGN KEY (PrincipalID) REFERENCES [SecurityPrincipals] (ID) ON UPDATE CASCADE ON DELETE CASCADE;
ALTER TABLE [AccessPermissions] ADD CONSTRAINT [AccessPermissionsResourceID_References_Resource] FOREIGN KEY (ResourceID) REFERENCES [Resources] (ID) ON UPDATE CASCADE ON DELETE CASCADE;

CREATE TABLE [AccessAudit] (
[ID] COUNTER NOT NULL,
[AccessTime] DATETIME,
[ResourceID] INTEGER NOT NULL,
[PrincipalID] INTEGER NOT NULL,
[Action] INTEGER
);
ALTER TABLE [AccessAudit] ADD CONSTRAINT [AccessAuditResourceID_References_Resource] FOREIGN KEY (ResourceID) REFERENCES [Resources] (ID) ON UPDATE CASCADE;
ALTER TABLE [AccessAudit] ADD CONSTRAINT [AccessAuditPrincipalID_References_SecurityPrincipal] FOREIGN KEY (PrincipalID) REFERENCES [SecurityPrincipals] (ID) ON UPDATE CASCADE ON DELETE SET NULL;