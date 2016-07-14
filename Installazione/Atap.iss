[ISTool]
EnableISX=false


[Files]

;RunTime Library
;NOT INCLUDE BECAUSE IT CAN CAUSE SYSTEM CRASH
;Source: ..\..\..\DeployFiles\Runtime\MFC42.DLL; DestDir: {sys}; Flags: regserver sharedfile uninsneveruninstall promptifolder onlyifdoesntexist
;Source: ..\..\..\DeployFiles\Runtime\MSVCIRT.DLL; DestDir: {sys}; Flags: regserver sharedfile uninsneveruninstall promptifolder onlyifdoesntexist
;Source: ..\..\..\DeployFiles\Runtime\MSVCP60.DLL; DestDir: {sys}; Flags: regserver sharedfile uninsneveruninstall promptifolder onlyifdoesntexist
;Source: ..\..\..\DeployFiles\Runtime\MFCANS32.DLL; DestDir: {sys}; Flags: regserver sharedfile uninsneveruninstall promptifolder onlyifdoesntexist
;

;START DAO
Source: ..\..\..\DeployFiles\DAO\dao360.dll; DestDir: {dao}; Flags: regserver sharedfile uninsneveruninstall promptifolder
;End DAO

; START VISUAL BASIC 6.0
Source: ..\..\..\DeployFiles\VB6_Runtime\STDOLE2.TLB; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: ..\..\..\DeployFiles\VB6_Runtime\MSVBVM60.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver

Source: ..\..\..\DeployFiles\VB6_Runtime\OleAut32.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: ..\..\..\DeployFiles\VB6_Runtime\OlePro32.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: ..\..\..\DeployFiles\VB6_Runtime\AsycFilt.dll; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile

Source: ..\..\..\DeployFiles\VB6_Runtime\VB6IT.DLL; DestDir: {sys}; Flags: promptifolder sharedfile


Source: ..\..\..\DeployFiles\Runtime\MSDERUN.DLL; DestDir: {sys}; Flags: promptifolder regserver sharedfile

Source: ..\..\..\DeployFiles\Runtime\MSBIND.DLL; DestDir: {sys}; Flags: promptifolder regserver sharedfile
Source: ..\..\..\DeployFiles\Runtime\MSSTDFMT.DLL; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; END VISUAL BASIC


; CRYSTAL REPORT
Source: ..\..\..\DeployFiles\Crystal\Crystl32.OCX; DestDir: {sys}\Crystal; Flags: regserver sharedfile
Source: ..\..\..\DeployFiles\Crystal\crpe32.dll; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\implode.dll; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\crpaig80.dll; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\cpeaut32.dll; DestDir: {sys}\Crystal; Flags: 

Source: ..\..\..\DeployFiles\Crystal\crxlat32.dll; DestDir: {sys}\Crystal; Flags: regserver sharedfile noregerror
Source: ..\..\..\DeployFiles\Crystal\p2sodbc.dll; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\P2BDAO.DLL; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\P2CTDAO.DLL; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\P2IRDAO.DLL; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\P2SMON.DLL; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\P2LODBC.DLL; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\P2SOLEDB.DLL; DestDir: {sys}\Crystal; Flags: 

Source: ..\..\..\DeployFiles\Crystal\u2ddisk.dll; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\u2fodbc.dll; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\u2fcr.dll; DestDir: {sys}\Crystal; Flags: 

Source: ..\..\..\DeployFiles\Crystal\CRXF_PDF.DLL; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\CRXF_RTF.DLL; DestDir: {sys}\Crystal; Flags: 

Source: ..\..\..\DeployFiles\Crystal\u2fhtml.dll; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\u2ftext.dll; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\u2fwordw.dll; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\u2fxls.dll; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\u2fxml.dll; DestDir: {sys}\Crystal; Flags: 

Source: ..\..\..\DeployFiles\Crystal\U2L2000.DLL; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\U252000.DLL; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Crystal\U2LCOM.DLL; DestDir: {sys}\Crystal; Flags: 
Source: ..\..\..\DeployFiles\Seagate Software\Viewers\ActiveXViewer\crviewer.dll; DestDir: {sys}; Flags: regserver
Source: ..\..\..\DeployFiles\Seagate Software\Report Designer Component\craxdrt.dll; DestDir: {sys}; Flags: regserver

Source: ..\..\..\DeployFiles\Databases\DBMSSOCN.DLL; DestDir: {sys}; Flags: regserver sharedfile noregerror






;ComponentOne OCX

Source: ..\..\..\DeployFiles\ComponentOne\vsflex8.ocx; DestDir: {sys}; Flags: regserver sharedfile
Source: ..\..\..\DeployFiles\ComponentOne\tibase6.dll; DestDir: {sys}
Source: ..\..\..\DeployFiles\ComponentOne\tibase8.dll; DestDir: {sys}
Source: ..\..\..\DeployFiles\ComponentOne\tidate8.ocx; DestDir: {sys}; Flags: regserver sharedfile
Source: ..\..\..\DeployFiles\ComponentOne\titext8.ocx; DestDir: {sys}; Flags: regserver sharedfile
Source: ..\..\..\DeployFiles\ComponentOne\tishare8.dll; DestDir: {sys}; Flags: regserver sharedfile
Source: ..\..\..\DeployFiles\ComponentOne\tinumb8.ocx; DestDir: {sys}; Flags: regserver sharedfile
Source: ..\..\..\DeployFiles\ComponentOne\todl8.ocx; DestDir: {sys}; Flags: regserver sharedfile
Source: ..\..\..\DeployFiles\ComponentOne\timask8.ocx; DestDir: {sys}; Flags: regserver sharedfile
Source: ..\..\..\DeployFiles\ComponentOne\todgub8.dll; DestDir: {sys}; Flags: regserver sharedfile noregerror
Source: ..\..\..\DeployFiles\ComponentOne\tdbg8.ocx; DestDir: {sys}; Flags: regserver sharedfile noregerror
Source: ..\..\..\DeployFiles\ComponentOne\tdbgpp8.dll; DestDir: {sys}; Flags: regserver sharedfile noregerror
Source: ..\..\..\DeployFiles\ComponentOne\tdcl8.ocx; DestDir: {sys}; Flags: regserver sharedfile noregerror
Source: ..\..\..\DeployFiles\ComponentOne\todg8.ocx; DestDir: {sys}; Flags: regserver sharedfile noregerror



;Menu Creator

;Source: C:\WINDOWS\system32\MenuCreator.dll; DestDir: {sys}; Flags: regserver sharedfile
;Source: C:\WINDOWS\system32\MenuExtended.dll; DestDir: {sys}; Flags: regserver sharedfile
;Source: C:\WINDOWS\system32\MenuEx.DLL; DestDir: {sys}; Flags: regserver sharedfile


;MS Common Dialog
Source: ..\..\..\DeployFiles\comdlg32.ocx; DestDir: {sys}; Flags: regserver sharedfile


;MS Common Controls
Source: ..\..\..\DeployFiles\MSCOMCTL.OCX; DestDir: {sys}; Flags: regserver sharedfile

;MS Common Control 2
Source: ..\..\..\DeployFiles\mscomct2.ocx; DestDir: {sys}; Flags: regserver sharedfile

;MS Common Control 3
Source: ..\..\..\DeployFiles\COMCT332.OCX; DestDir: {sys}; Flags: regserver sharedfile

;MS Tabbed Dialog Control
Source: ..\..\..\DeployFiles\tabctl32.ocx; DestDir: {sys}; Flags: regserver sharedfile



;VB Accellerator ImmageList
Source: ..\..\..\DeployFiles\vbAccellerator\vbalIml6.ocx; DestDir: {sys}; Flags: regserver sharedfile


;MS ADO
;Source: ..\..\..\DeployFiles\ado\msado21.tlb; DestDir: {sys}; Flags: regtypelib
;Source: C:\Programmi\File comuni\System\ado\msado27.tlb; DestDir: {sys}; Flags: regtypelib

;Source: C:\WINDOWS\system32\hhctrl.ocx; DestDir: {sys}; Flags: regserver


; START MDAC
Source: MDAC_TYP.EXE; DestDir: {tmp}; MinVersion: 4.0,4.0; OnlyBelowVersion: 0,5.0; Components: MDAC; Flags: ignoreversion

Source: ..\Atap.exe; DestDir: {app}
Source: ..\Logo.jpg; DestDir: {app}
Source: ..\sto.0; DestDir: {app}
Source: ..\zip.exe; DestDir: {app}
Source: ..\unzip.exe; DestDir: {app}
Source: ..\Atap.chm; DestDir: {app}\Help



;Source: ..\atap.mdb; DestDir: {app}

;REPORTS
Source: ..\Report\AnagraficaC_Nome.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\AnagraficaCondensata.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\AnagraficaDettagliata.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\AnticipiEuro.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\AssegniCircolari.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\EstrattoConto.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\EstrattoContoAdempim.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\EstrattoContoLiquidazione.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\Fattura.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\FatturaProv.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\GiornalieraAdempimenti.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\GiornalieraDecreti.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\GiornalieraNotifiche.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\GiornalieraSfrattiPig.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\Pignoramenti.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\Saldi.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SaldiProvvisori.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\Sospesi.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SospesiAttivit‡.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SospesiAttivit‡New.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SospesiAttTrib.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SospesiAttTribNew.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SospesiLiquidazione.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SospesiLiquidazioneNew.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SospesiNew.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\Tribunali.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\Tribunali.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\GiornalieraNotificheUNEP.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\GiornalieraSfrattiPigUNEP.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\EstrattoContoUNEP.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\AssegniCircolariUNEP.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SaldiUNEP.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SaldiProvvisoriUNEP.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\FatturaUNEP.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SospesiLiquidazione.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SospesiLiquidazioneUNEP.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SospesiAttTribNewUNEP.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SospesiNewUNEP.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion
Source: ..\Report\SospesiAttivit‡NewUNEP.rpt; DestDir: {app}\Report; Flags: ignoreversion overwritereadonly replacesameversion

Source: ..\Dati\listacomuni.txt; DestDir: {app}\Dati; Flags: ignoreversion overwritereadonly replacesameversion





[Icons]
Name: {group}\Atap; Filename: {app}\Atap.exe; WorkingDir: {app}; IconFilename: {app}\Atap.exe; IconIndex: 0
Name: {commondesktop}\Atap; Filename: {app}\Atap.exe; WorkingDir: {app}; IconFilename: {app}\Atap.exe; IconIndex: 0


[Components]
Name: MDAC; Description: Microsoft Data Access Components; Flags: disablenouninstallwarning restart fixed; MinVersion: 4.0,4.0; OnlyBelowVersion: 0,5.0; Types: custom compact full

[Run]

; START MDAC
Filename: {tmp}\mdac_typ.exe; Parameters: "/Q /C:""setup /QNT"""; WorkingDir: {tmp}; Flags: skipifdoesntexist; Components: MDAC; MinVersion: 4.0,4.0; OnlyBelowVersion: 0,5.0
; END MDAC



[Setup]

AppName=Atap
AppVerName=Atap
DisableDirPage=false
DefaultGroupName=Atap
Compression=zip/9
DisableStartupPrompt=true
DefaultDirName={pf}\Atap


AppPublisherURL=
AppSupportURL=
AppUpdatesURL=
WizardImageFile=C:\Programmi\Inno Setup 5\WizModernSmallImage-IS.bmp
WizardImageBackColor=clSilver
WizardSmallImageFile=C:\Programmi\Inno Setup 5\WizModernImage-IS.bmp
OutputBaseFilename=Setup
PrivilegesRequired=none
AppID={{CBFFB0F9-9168-43F6-A372-51DE7D464BD0}



[Languages]
Name: default; MessagesFile: compiler:Languages\Italian.isl
