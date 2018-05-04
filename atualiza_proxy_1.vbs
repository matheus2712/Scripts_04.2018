dim oShell
set oShell = Wscript.CreateObject("Wscript.Shell")


oShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD"
oShell.RegWrite "HKCU\Software\Microsoft\Windows\currentVersion\Internet Settings\ProxyServer", "192.168.0.2:8080", "REG_SZ"
oShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride","officeclient.microsoft.com; outlook.mandic.com.br; officeclient.microsoft.com; cdn.odc.officeapps.live.com; templateservice.office.com; autodiscover.microsoft.com; outlook.mandic.com.br; jwmtransportes.com.br; autodiscover.jwmtransportes.com.br; omextemplates.content.office.net; odc.officeapps.live.com; office15client.microsoft.com; store.office.com; clienttemplates.content.office.net; office.com; office.net; microsoft.com; live.com; contentstorage.osi.office.net; ocws.officeapps.live.com; jwmtransportes-my.sharepoint.com; jwmtransportes.onmicrosoft.com; autodiscover-s.outlook.com; messaging.office.com; watson.microsoft.com; client-office365-tas.msedge.net; nexusrules.officeapps.live.com; jwm.local; autodiscover.jwm.local; login.windows.net; login.microsoftonline.com; nl.osi.office.net; omextemplates.content.office.net; templateservice.oofice.com; contentstorage.osi.office.net; ocws.officeapps.live.com; contentstorage.osi.office.net; odc.officeapps.live.com; omextemplates.content.office.net; <local>"



Set oShell = Nothing