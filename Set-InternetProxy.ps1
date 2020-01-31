##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2U3hSWElUcqQhZKo04+w8OvoqBnJW5UMUBp+lSac
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiW5
##OsHQCZGeTiiZ4dI=
##OcrLFtDXTiW5
##LM/BD5WYTiiZ4tI=
##McvWDJ+OTiiZ4tI=
##OMvOC56PFnzN8u+Vs1Q=
##M9jHFoeYB2Hc8u+Vs1Q=
##PdrWFpmIG2HcofKIo2QX
##OMfRFJyLFzWE8uK1
##KsfMAp/KUzWJ0g==
##OsfOAYaPHGbQvbyVvnQX
##LNzNAIWJGmPcoKHc7Do3uAuO
##LNzNAIWJGnvYv7eVvnQX
##M9zLA5mED3nfu77Q7TV64AuzAgg=
##NcDWAYKED3nfu77Q7TV64AuzAgg=
##OMvRB4KDHmHQvbyVvnQX
##P8HPFJGEFzWE8tI=
##KNzDAJWHD2fS8u+Vgw==
##P8HSHYKDCX3N8u+Vgw==
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTlaDjofG5iZk2U3hSWElUcqQhZKi14qo8PrQmCncSIgkW1tvj2f5HE7d
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
# Function For Set Internet Explorer Proxy and ACS
# For Use You Can Set Set-
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function Show-Console
{
    # 4 SHOW
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 4)
}

function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}

Function Set-InternetProxy
{

            $regKey="HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
            Set-ItemProperty -path $regKey AutoConfigURL -Value ""        
            Set-ItemProperty -path $regKey ProxyEnable -value ""               
            Set-ItemProperty -path $regKey ProxyServer -value ""
}

Set-InternetProxy 
