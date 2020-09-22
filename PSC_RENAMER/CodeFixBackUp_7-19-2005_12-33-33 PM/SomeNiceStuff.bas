Attribute VB_Name = "SomeNiceStuff"
Option Explicit
' Shell variables
Public Enum ShowWindowType
  SW_HIDE = 0
  SW_NORMAL = 1
  SW_MINIMIZED = 2
  SW_MAXIMIZED = 3
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private SW_HIDE, SW_NORMAL, SW_MINIMIZED, SW_MAXIMIZED
#End If
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As ShowWindowType) As Long


Public Sub RunAssociated(sFileName As String, _
                          lHwnd As Long, _
                          Optional sParams As String = vbNullString, _
                          Optional sDefaultDir As String = vbNullString)

  Dim lrc As Long

  'Parameters passed to ShellExecute
  'hWnd - active form
  'lpOperation - "Open" or "Print" (vbNullString defaults to "Open"
  'lpFile - Program name or name of a for to print or open using the associated program
  'lpParameters - Command line if lpFile is a program to run
  'lpDirectory - Default directory to use
  'nShowCmd - Constant specifying how to show the launched program (maximized, minimized, normal)
  lrc = ShellExecute(lHwnd, "Open", sFileName, sParams, sDefaultDir, SW_NORMAL)

End Sub
Public Function SafeFileName(ByVal oldFileName As String, _
                             ByVal FileExt As String) As String

  'Func by DiGiTaIErRoR
  '@ http://www.vbforums.com/archive/index.php/t-255172.html
  'MODDED BY THE AUTHOR of PSC_RENAMER
  
  Dim ServDir  As String
  Dim BadChars As String
  Dim BadChar  As String
  Dim x        As Long

  BadChars = "\/:*?""<>|"
  ServDir = oldFileName
  For x = 1 To Len(BadChars)
    BadChar = Mid$(BadChars, x, 1)
    ServDir = Replace$(ServDir, BadChar, "_")
  Next
  If Len(ServDir) + 1 + Len(FileExt) > 255 Then 'FAT32 and up file path length limitation
    ServDir = Left$(ServDir, 255 - 1 - Len(FileExt) - 1)
  End If
  SafeFileName = ServDir & "." & FileExt

End Function
Public Function SaveResItemToDisk( _
            ByVal iResourceNum As Integer, _
            ByVal sResourceType As Integer, _
            ByVal sDestFileName As String _
            ) As Long
    '=============================================
    'Saves a resource item to disk
    'Returns 0 on success, error number on failure
    'PROVIDED BY MICROSOFT MSDN DOWNLOADS
    '=============================================
    
    'Example Call:
    ' iRetVal = SaveResItemToDisk(101, "CUSTOM", "C:\myImage.gif")
    
    Dim bytResourceData()   As Byte
    Dim iFileNumOut         As Integer
    
    On Error GoTo SaveResItemToDisk_err
    
    'Retrieve the resource contents (data) into a byte array
    bytResourceData = LoadResData(iResourceNum, sResourceType)
    
    'Get Free File Handle
    iFileNumOut = FreeFile
    
    'Open the output file
    Open sDestFileName For Binary Access Write As #iFileNumOut
        
        'Write the resource to the file
        Put #iFileNumOut, , bytResourceData
    
    'Close the file
    Close #iFileNumOut
    
    'Return 0 for success
    SaveResItemToDisk = 0
    
    Exit Function
SaveResItemToDisk_err:
    'Return error number
    SaveResItemToDisk = Err.Number
End Function
