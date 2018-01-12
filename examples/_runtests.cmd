@echo off
md Databases
md Dumps
call :process BasicTables.sql
call :process Comments.sql
call :process DataTypes.sql
call :process AutoIncFields.sql
call :process Indices.sql
call :process PrivateComments.sql
call :process PrivateAttrs.sql
call :process Procedures.sql
call :process Views.sql
call :process Constraints.sql
call :process SpecialCharacters.sql
goto :eof

:process
echo %1 -- %~dp1Databases\%~n1.mdb
..\jet.exe exec -f %~dp1Databases\%~n1.mdb -new -force -stdi %1
if errorlevel 1 (
  echo %1: execution failed.
  goto :eof
)
..\jet.exe dump --all -f %~dp1Databases\%~n1.mdb -stdo %~dp1Dumps\%~nx1
if errorlevel 1 (
  echo %1: dump failed.
  goto :eof
)
goto :eof