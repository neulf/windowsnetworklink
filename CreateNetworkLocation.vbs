Option Explicit

Function CreateNetworkLocation( networkLocationName, networkLocationTarget )
    Const ssfNETHOOD  = &H13&
    Const fsATTRIBUTES_READONLY = 1
    Const fsATTRIBUTES_HIDDEN = 2
    Const fsATTRIBUTES_SYSTEM = 4

    CreateNetworkLocation = False 

    ' Instantiate needed components
    Dim fso, shell, shellApplication
        Set fso =               WScript.CreateObject("Scripting.FileSystemObject")
        Set shell =             WScript.CreateObject("WScript.Shell")
        Set shellApplication =  WScript.CreateObject("Shell.Application")

    ' Locate where NetworkLocations are stored
    Dim nethoodFolderPath, networkLocationFolder, networkLocationFolderPath
        nethoodFolderPath = shellApplication.Namespace( ssfNETHOOD ).Self.Path

    ' Create the folder for our NetworkLocation and set its attributes
        networkLocationFolderPath = fso.BuildPath( nethoodFolderPath, networkLocationName )
        If fso.FolderExists( networkLocationFolderPath ) Then 
            Exit Function 
        End If 
        Set networkLocationFolder = fso.CreateFolder( networkLocationFolderPath )
        networkLocationFolder.Attributes = fsATTRIBUTES_READONLY

    ' Write the desktop.ini inside our NetworkLocation folder and change its attributes    
    Dim desktopINIFilePath
        desktopINIFilePath = fso.BuildPath( networkLocationFolderPath, "desktop.ini" )
        With fso.CreateTextFile(desktopINIFilePath)
            .Write  "[.ShellClassInfo]" & vbCrlf & _ 
                    "CLSID2={0AFACED1-E828-11D1-9187-B532F1E9575D}" & vbCrlf & _ 
                    "Flags=2" & vbCrlf
            .Close
        End With 
        With fso.GetFile( desktopINIFilePath )
            .Attributes = fsATTRIBUTES_HIDDEN + fsATTRIBUTES_SYSTEM
        End With 

    ' Create the shortcut to the target of our NetworkLocation
    Dim targetLink
        targetLink = fso.BuildPath( networkLocationFolderPath, "target.lnk" )
        With shell.CreateShortcut( targetLink )
            .TargetPath = networkLocationTarget
            .Save
        End With        

    ' Done
        CreateNetworkLocation = True 

End Function

'Please call the function for your need.
'CreateNetworkLocation "阿李王图书馆", "\\192.168.68.12\ali_library"
'CreateNetworkLocation "阿李档案馆", "\\192.168.68.12\safe_share\阿李档案馆"
'CreateNetworkLocation "阿李软件城", "\\192.168.68.12\safe_share\阿李软件城"
'CreateNetworkLocation "阿李照相馆", "\\192.168.68.12\safe_share\阿李照相馆"
'CreateNetworkLocation "桥头博物馆", "\\192.168.68.12\safe_share\桥头博物馆"
'CreateNetworkLocation "阿力运动", "\\192.168.68.12\safe_share\阿力运动"
'CreateNetworkLocation "半岛文库", "\\192.168.68.12\safe_share\半岛文库"
'CreateNetworkLocation "萧太后影视城", "\\192.168.68.12\safe_share\萧太后影视城"
'CreateNetworkLocation "红灯区", "\\192.168.68.12\safe_share\红灯区"
'CreateNetworkLocation "半岛音乐厅", "\\192.168.68.12\safe_share\半岛音乐厅"
'CreateNetworkLocation "王子学院", "\\192.168.68.12\safe_share\王子学院"
'CreateNetworkLocation "其他资料", "\\192.168.68.12\safe_share\其他资料"