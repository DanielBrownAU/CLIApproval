Installation Guide:

Install SPMCASTSADM.dll into the GAC (C:\Windows\assembly")

Exmaple: gacutil /i SPMCASTSADM.dll


Copy stsadmcommands.DBSPMCA.xml into the following location:

%PROGRAMFILES%\Common Files\Microsoft Shared\web server extensions\12\CONFIG


Commands:

stsadm -o spmcahelp

Check out my blog at http://www.danielbrown.id.au/