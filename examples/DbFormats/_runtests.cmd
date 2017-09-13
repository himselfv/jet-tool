@echo off
rem Creates empty databases in various supported formats, via various supported OLEDB/DAO providers.

setlocal
set PATH=%CD%\..\..;%PATH%
set JETNEW=jet touch -new -force -f

echo Creating the databases using default OLEDB/DAO providers:
call :new_default jet4x
call :new_default jet3x
call :new_default jet20
call :new_default jet11
call :new_default jet10
call :new_default ace12

echo Testing format shortcuts...
rem These should be identical to jet4x
%JETNEW% mdb1.mdb --mdb
%JETNEW% mdb2.mdb
rem These should be identical to ace12
%JETNEW% accdb1.accdb --accdb
%JETNEW% accdb2.accdb

echo Creating the databases using explicitly set ACE provider.
echo Older formats may not be supported, depending on the ACE version installed.
rem This one should be identical to ace12_def
call :new_ace ace12
call :new_ace jet4x
call :new_ace jet3x
call :new_ace jet20
call :new_ace jet11
call :new_ace jet10

echo Creating the databases using explicitly set Jet 4.0 provider.
echo Accdb format should not be supported.
call :new_jet4 ace12
call :new_jet4 jet4x
call :new_jet4 jet3x
call :new_jet4 jet20
call :new_jet4 jet11
call :new_jet4 jet10

echo All tests completed.
echo Verify that the appropriate databases are identical.
goto :eof


:new_default
echo %1 via default engines:
%JETNEW% %1_def.mdb --db-format %1
goto :eof

:new_ace
echo %1 via modern engines (ACE/DAO120):
%JETNEW% %1_eng12.mdb --db-format %1 --oledb-eng Microsoft.ACE.OLEDB.12.0 --dao-eng DAO.Engine.120
goto :eof

:new_jet4
echo %1 via older engines (Jet4/DAO36):
%JETNEW% %1_eng4.mdb --db-format %1 --oledb-eng Microsoft.Jet.OLEDB.4.0 --dao-eng DAO.Engine.36
goto :eof
