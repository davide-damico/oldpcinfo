#include <FileConstants.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#RequireAdmin

$vOldPC = InputBox("OldPcInfo - v. 1.0", "Utility per leggere le configurazioni di una installazione Windows OffLine. " & @CRLF & @CRLF &  "by Davide D'Amico. " & @CRLF & @CRLF & "Inseirire lettera drive di Windows:", "C:")
If @Error Then Exit
$vOldPC = StringLeft($vOldPC, 1) & ":"

; Creo una mascherina per far vedere lo stato dell'elaborazione
$gui=GUICreate("OldPcInfo", 260, 100, -1, -1)
$progressbar = GUICtrlCreateProgress(10, 50, 240, 20)
GUICtrlSetState ($progressbar,$GUI_HIDE)
$label = GUICtrlCreateLabel("Attendere. Sto elaborando...", 10, 10, 210)
GUICtrlSetColor(-1, 32250)
GUISetState()


;Se voglio vedere il mio C: devo seguire una strada diversa
If $vOldPC = "C:" Then

   ;Metto tutto su un file Report
   _FileCreate(@DesktopDir & "\OldPCInfo.txt")
   $hFileOpen = FileOpen(@DesktopDir & "\OldPCInfo.txt", $FO_OVERWRITE)

   ;Nome del PC
   $sOldPCName = RegRead("HKLM\System\ControlSet001\Control\ComputerName\ComputerName", "ComputerName")
   FileWriteLine($hFileOpen, "Nome PC: " & $sOldPCName & @CRLF & @CRLF)


   FileWriteLine($hFileOpen, "===============================Connessioni di Rete===============================" & @CRLF & @CRLF)

   GUICtrlSetState ($progressbar,$GUI_SHOW)
   GUICtrlSetData($progressbar, (1/5)*100)

   ;Interfacce
   For $i = 1 To 100
	  $var = RegEnumKey("HKLM\System\ControlSet001\Services\Tcpip\Parameters\Interfaces", $i)
	  If @error <> 0 Then ExitLoop
	  $sOldInterfaceName = RegRead("HKLM\System\ControlSet001\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & $var & "\Connection", "Name")
	  If $sOldInterfaceName == "" Then ContinueLoop
	  FileWriteLine($hFileOpen, $sOldInterfaceName & ":")
	  $sEnableDHCP = RegRead("HKLM\System\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "EnableDHCP")
	  If $sEnableDHCP Then
		 $sEnableDHCP = "SI"
	  Else
		 $sEnableDHCP = "NO"
	  EndIf
	  If $sEnableDHCP = "SI" Then
		 FileWriteLine($hFileOpen, "DHCP Abilitato: " & $sEnableDHCP)
		 $sDhcpIPAddress = RegRead("HKLM\System\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "DhcpIPAddress")
		 FileWriteLine($hFileOpen, "DHCP IP Address: " & $sDhcpIPAddress)
		 $sDhcpSubnetMask = RegRead("HKLM\System\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "DhcpSubnetMask")
		 FileWriteLine($hFileOpen, "DHCP Subnet Mask: " & $sDhcpSubnetMask)
		 $sDhcpDefaultGateway = RegRead("HKLM\System\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "DhcpDefaultGateway")
		 FileWriteLine($hFileOpen, "DHCP Default Gateway: " & $sDhcpDefaultGateway)
		 $sDhcpNameServer = RegRead("HKLM\System\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "DhcpNameServer")
		 FileWriteLine($hFileOpen, "DHCP DNS: " & $sDhcpNameServer)
	  Else
		 FileWriteLine($hFileOpen, "DHCP Abilitato: " & $sEnableDHCP)
		 $sIPAddress = RegRead("HKLM\System\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "IPAddress")
		 FileWriteLine($hFileOpen, "IP Address: " & $sIPAddress)
		 $sSubnetMask = RegRead("HKLM\System\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "SubnetMask")
		 FileWriteLine($hFileOpen, "Subnet Mask: " & $sSubnetMask)
		 $sDefaultGateway = RegRead("HKLM\System\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "DefaultGateway")
		 FileWriteLine($hFileOpen, "Default Gateway: " & $sDefaultGateway)
		 $sNameServer = RegRead("HKLM\System\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "NameServer")
		 FileWriteLine($hFileOpen, "DNS: " & $sNameServer)
	  EndIf
	  FileWriteLine($hFileOpen, @CRLF)
   Next

   FileWriteLine($hFileOpen, "===============================Stampanti===============================" & @CRLF & @CRLF)

   GUICtrlSetData($progressbar, (2/5)*100)

   ;Stampanti
   For $i = 1 To 100
	  $sOldPrinterName = RegEnumKey("HKLM\Software\Microsoft\Windows NT\CurrentVersion\Print\Printers", $i)
	  If @error <> 0 Then ExitLoop
	  $sOldPrinterDriverName = RegRead("HKLM\Software\Microsoft\Windows NT\CurrentVersion\Print\Printers\" & $sOldPrinterName, "Printer Driver")
	  $sOldPrinterPort = RegRead("HKLM\Software\Microsoft\Windows NT\CurrentVersion\Print\Printers\" & $sOldPrinterName, "Port")
	  FileWriteLine($hFileOpen, $sOldPrinterName & ":")
	  FileWriteLine($hFileOpen, "Nome Driver: " & $sOldPrinterDriverName)
	  FileWriteLine($hFileOpen, "Porta: " & $sOldPrinterPort)
	  FileWriteLine($hFileOpen, @CRLF)
   Next


   FileWriteLine($hFileOpen, "===============================Condivisioni===============================" & @CRLF & @CRLF)

   GUICtrlSetData($progressbar, (3/5)*100)

   ;Condivisioni
   For $i = 1 To 100
	  $sOldSharesEnum = RegEnumVal("HKLM\System\ControlSet001\services\LanmanServer\Shares", $i)
	  $sOldShare = StringSplit(RegRead("HKLM\System\ControlSet001\services\LanmanServer\Shares", $sOldSharesEnum), @LF)
	  If @error <> 0 Then ExitLoop
	  For $element = 1 To Ubound($sOldShare) - 1
		 If StringInStr($sOldShare[$element], "Path") > 0 Then
			FileWriteLine($hFileOpen, $sOldShare[$element] & " - Nome=" & $sOldSharesEnum)
		 endif
	  Next
   Next

   FileWriteLine($hFileOpen, @CRLF)

   FileWriteLine($hFileOpen, "===============================Elenco Programmi Installati===============================" & @CRLF & @CRLF)

   GUICtrlSetData($progressbar, (4/5)*100)

   For $i = 1 To 500
	  $sOldProgramEnum = RegEnumKey("HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall", $i)
	  If @error <> 0 Then ExitLoop
	  $sOldProgramName = RegRead("HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & $sOldProgramEnum, "DisplayName")
	  If $sOldProgramName Then
		 FileWriteLine($hFileOpen, $sOldProgramName)
		 FileWriteLine($hFileOpen, @CRLF)
	  Endif
   Next

   FileWriteLine($hFileOpen, "===============================Elenco Utenti===============================" & @CRLF & @CRLF)

   ;Elenco utenti
   $aOldUsers = _FileListToArray($vOldPC & "\Users", "*", $FLTA_FOLDERS)
   For $i = 1 to Ubound($aOldUsers) - 1

	  GUICtrlSetData($progressbar, ((4/5)*100) + ((1/5)*100/(Ubound($aOldUsers) - 1) * $i))

	  FileWriteLine($hFileOpen, $aOldUsers[$i])
	  ;Per ogni utente vado a vedere se aveva delle mappature
	  FileWriteLine($hFileOpen, "Mappature:")

	  For $x = 1 To 100
		 $sOldMapName = RegEnumKey("HKCU\Network", $x)
		 If @error <> 0 Then ExitLoop
		 $sOldMapPath = RegRead("HKCU\Network\" & $sOldMapName, "RemotePath")
		 FileWriteLine ($hFileOpen, $sOldMapName & "==>" & $sOldMapPath & @CRLF)
	  Next

	  FileWriteLine($hFileOpen, @CRLF)
   Next

   ;Chiudo il File Report
   FileClose($hFileOpen)

   ;Apro il File Report
   ShellExecute(@DesktopDir & "\OldPCInfo.txt")
   Exit
EndIf

;Carico gli Hive dall'installazione esterna
Run(@ComSpec & " /c reg.exe load HKLM\OldSystem " & $vOldPC & "\Windows\System32\Config\system","",@SW_HIDE)
Sleep(1000)
Run(@ComSpec & " /c reg.exe load HKLM\OldSoftware " & $vOldPC & "\Windows\System32\Config\software","",@SW_HIDE)
Sleep(1000)
Run(@ComSpec & " /c reg.exe load HKLM\OldSoftware " & $vOldPC & "\Windows\System32\Config\software","",@SW_HIDE)
Sleep(1000)

;Metto tutto su un file Report
_FileCreate(@DesktopDir & "\OldPCInfo.txt")
$hFileOpen = FileOpen(@DesktopDir & "\OldPCInfo.txt", $FO_OVERWRITE)

;Nome del PC
$sOldPCName = RegRead("HKLM\OldSystem\ControlSet001\Control\ComputerName\ComputerName", "ComputerName")
FileWriteLine($hFileOpen, "Nome PC: " & $sOldPCName & @CRLF & @CRLF)

FileWriteLine($hFileOpen, "===============================Connessioni di Rete===============================" & @CRLF & @CRLF)

GUICtrlSetState ($progressbar,$GUI_SHOW)
GUICtrlSetData($progressbar, (1/5)*100)

;Interfacce
For $i = 1 To 100
   $var = RegEnumKey("HKLM\OldSystem\ControlSet001\Services\Tcpip\Parameters\Interfaces", $i)
   If @error <> 0 Then ExitLoop
   $sOldInterfaceName = RegRead("HKLM\OldSystem\ControlSet001\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & $var & "\Connection", "Name")
   If $sOldInterfaceName == "" Then ContinueLoop
   FileWriteLine($hFileOpen, $sOldInterfaceName & ":")
   $sEnableDHCP = RegRead("HKLM\OldSystem\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "EnableDHCP")
   If $sEnableDHCP Then
	  $sEnableDHCP = "SI"
   Else
	  $sEnableDHCP = "NO"
   EndIf
   If $sEnableDHCP = "SI" Then
	  FileWriteLine($hFileOpen, "DHCP Abilitato: " & $sEnableDHCP)
	  $sDhcpIPAddress = RegRead("HKLM\OldSystem\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "DhcpIPAddress")
	  FileWriteLine($hFileOpen, "DHCP IP Address: " & $sDhcpIPAddress)
	  $sDhcpSubnetMask = RegRead("HKLM\OldSystem\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "DhcpSubnetMask")
	  FileWriteLine($hFileOpen, "DHCP Subnet Mask: " & $sDhcpSubnetMask)
	  $sDhcpDefaultGateway = RegRead("HKLM\OldSystem\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "DhcpDefaultGateway")
	  FileWriteLine($hFileOpen, "DHCP Default Gateway: " & $sDhcpDefaultGateway)
	  $sDhcpNameServer = RegRead("HKLM\OldSystem\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "DhcpNameServer")
	  FileWriteLine($hFileOpen, "DHCP DNS: " & $sDhcpNameServer)
   Else
	  FileWriteLine($hFileOpen, "DHCP Abilitato: " & $sEnableDHCP)
	  $sIPAddress = RegRead("HKLM\OldSystem\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "IPAddress")
	  FileWriteLine($hFileOpen, "IP Address: " & $sIPAddress)
	  $sSubnetMask = RegRead("HKLM\OldSystem\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "SubnetMask")
	  FileWriteLine($hFileOpen, "Subnet Mask: " & $sSubnetMask)
	  $sDefaultGateway = RegRead("HKLM\OldSystem\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "DefaultGateway")
	  FileWriteLine($hFileOpen, "Default Gateway: " & $sDefaultGateway)
	  $sNameServer = RegRead("HKLM\OldSystem\ControlSet001\Services\Tcpip\Parameters\Interfaces\" & $var, "NameServer")
	  FileWriteLine($hFileOpen, "DNS: " & $sNameServer)
   EndIf
   FileWriteLine($hFileOpen, @CRLF)
Next

FileWriteLine($hFileOpen, "===============================Stampanti===============================" & @CRLF & @CRLF)

GUICtrlSetData($progressbar, (2/5)*100)

;Stampanti
For $i = 1 To 100
   $sOldPrinterName = RegEnumKey("HKLM\OldSoftware\Microsoft\Windows NT\CurrentVersion\Print\Printers", $i)
   If @error <> 0 Then ExitLoop
   $sOldPrinterDriverName = RegRead("HKLM\OldSoftware\Microsoft\Windows NT\CurrentVersion\Print\Printers\" & $sOldPrinterName, "Printer Driver")
   $sOldPrinterPort = RegRead("HKLM\OldSoftware\Microsoft\Windows NT\CurrentVersion\Print\Printers\" & $sOldPrinterName, "Port")
   FileWriteLine($hFileOpen, $sOldPrinterName & ":")
   FileWriteLine($hFileOpen, "Nome Driver: " & $sOldPrinterDriverName)
   FileWriteLine($hFileOpen, "Porta: " & $sOldPrinterPort)
   FileWriteLine($hFileOpen, @CRLF)
Next

FileWriteLine($hFileOpen, "===============================Condivisioni===============================" & @CRLF & @CRLF)

GUICtrlSetData($progressbar, (3/5)*100)

;Condivisioni
For $i = 1 To 100
   $sOldSharesEnum = RegEnumVal("HKLM\OldSystem\ControlSet001\services\LanmanServer\Shares", $i)
   $sOldShare = StringSplit(RegRead("HKLM\OldSystem\ControlSet001\services\LanmanServer\Shares", $sOldSharesEnum), @LF)
   If @error <> 0 Then ExitLoop
   For $element = 1 To Ubound($sOldShare) - 1
	  If StringInStr($sOldShare[$element], "Path") > 0 Then
		 FileWriteLine($hFileOpen, $sOldShare[$element] & " - Nome=" & $sOldSharesEnum)
	  endif
   Next
Next

FileWriteLine($hFileOpen, @CRLF)

FileWriteLine($hFileOpen, "===============================Elenco Programmi Installati===============================" & @CRLF & @CRLF)

GUICtrlSetData($progressbar, (4/5)*100)

For $i = 1 To 500
   $sOldProgramEnum = RegEnumKey("HKLM\OldSoftware\Microsoft\Windows\CurrentVersion\Uninstall", $i)
   If @error <> 0 Then ExitLoop
   $sOldProgramName = RegRead("HKLM\OldSoftware\Microsoft\Windows\CurrentVersion\Uninstall\" & $sOldProgramEnum, "DisplayName")
   If $sOldProgramName Then
	  FileWriteLine($hFileOpen, $sOldProgramName)
	  FileWriteLine($hFileOpen, @CRLF)
   Endif
Next

;Scarico tutti gli Hive
RunWait(@ComSpec & " /c reg.exe unload HKLM\OldSystem","",@SW_HIDE)
RunWait(@ComSpec & " /c reg.exe unload HKLM\OldSoftware","",@SW_HIDE)

FileWriteLine($hFileOpen, "===============================Elenco Utenti===============================" & @CRLF & @CRLF)

;Elenco utenti
$aOldUsers = _FileListToArray($vOldPC & "\Users", "*", $FLTA_FOLDERS)
For $i = 1 to Ubound($aOldUsers) - 1

   GUICtrlSetData($progressbar, ((4/5)*100) + ((1/5)*100/(Ubound($aOldUsers) - 1) * $i))

   FileWriteLine($hFileOpen, $aOldUsers[$i])
   ;Per ogni utente vado a vedere se aveva delle mappature
   RunWait(@ComSpec & " /c reg.exe load HKLM\OldUser " & $vOldPC & "\Users\" & $aOldUsers[$i] & "\NTUSER.DAT","",@SW_HIDE)
   ;Sleep(1000)
   FileWriteLine($hFileOpen, "Mappature:")

   For $x = 1 To 100
	  $sOldMapName = RegEnumKey("HKLM\OldUser\Network", $x)
	  If @error <> 0 Then ExitLoop
	  $sOldMapPath = RegRead("HKLM\OldUser\Network\" & $sOldMapName, "RemotePath")
	  FileWriteLine ($hFileOpen, $sOldMapName & "==>" & $sOldMapPath & @CRLF)
   Next

   FileWriteLine($hFileOpen, @CRLF)
   RunWait(@ComSpec & " /c reg.exe unload HKLM\OldUser","",@SW_HIDE)
   ;Sleep(1000)
Next

;Chiudo il File Report
FileClose($hFileOpen)

;Apro il File Report
ShellExecute(@DesktopDir & "\OldPCInfo.txt")