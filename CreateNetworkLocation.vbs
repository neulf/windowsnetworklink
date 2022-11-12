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
'CreateNetworkLocation "������ͼ���", "\\192.168.68.12\ali_library"
'CreateNetworkLocation "�������", "\\192.168.68.12\safe_share\�������"
'CreateNetworkLocation "���������", "\\192.168.68.12\safe_share\���������"
'CreateNetworkLocation "���������", "\\192.168.68.12\safe_share\���������"
'CreateNetworkLocation "��ͷ�����", "\\192.168.68.12\safe_share\��ͷ�����"
'CreateNetworkLocation "�����˶�", "\\192.168.68.12\safe_share\�����˶�"
'CreateNetworkLocation "�뵺�Ŀ�", "\\192.168.68.12\safe_share\�뵺�Ŀ�"
'CreateNetworkLocation "��̫��Ӱ�ӳ�", "\\192.168.68.12\safe_share\��̫��Ӱ�ӳ�"
'CreateNetworkLocation "�����", "\\192.168.68.12\safe_share\�����"
'CreateNetworkLocation "�뵺������", "\\192.168.68.12\safe_share\�뵺������"
'CreateNetworkLocation "����ѧԺ", "\\192.168.68.12\safe_share\����ѧԺ"
'CreateNetworkLocation "��������", "\\192.168.68.12\safe_share\��������"