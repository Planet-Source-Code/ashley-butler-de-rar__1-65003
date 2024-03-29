Attribute VB_Name = "Extraction"
Const ERAR_END_ARCHIVE = 10
Const ERAR_NO_MEMORY = 11
Const ERAR_BAD_DATA = 12
Const ERAR_BAD_ARCHIVE = 13
Const ERAR_UNKNOWN_FORMAT = 14
Const ERAR_EOPEN = 15
Const ERAR_ECREATE = 16
Const ERAR_ECLOSE = 17
Const ERAR_EREAD = 18
Const ERAR_EWRITE = 19
Const ERAR_SMALL_BUF = 20

Const RAR_OM_LIST = 0
Const RAR_OM_EXTRACT = 1
 
Const RAR_SKIP = 0
Const RAR_TEST = 1
Const RAR_EXTRACT = 2
 
Const RAR_VOL_ASK = 0
Const RAR_VOL_NOTIFY = 1

Enum RarOperations
    OP_EXTRACT = 0
    OP_TEST = 1
    op_list = 2
End Enum
 
Private Type RARHeaderData
    ArcName As String * 260
    FileName As String * 260
    flags As Long
    PackSize As Long
    UnpSize As Long
    HostOS As Long
    FileCRC As Long
    FileTime As Long
    UnpVer As Long
    Method As Long
    FileAttr As Long
    CmtBuf As String
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
End Type
 
Private Type RAROpenArchiveData
    ArcName As String
    OpenMode As Long
    OpenResult As Long
    CmtBuf As String
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
End Type
 
Private Declare Function RAROpenArchive Lib "unrar.dll" (ByRef ArchiveData As RAROpenArchiveData) As Long
Private Declare Function RARCloseArchive Lib "unrar.dll" (ByVal hArcData As Long) As Long
Private Declare Function RARReadHeader Lib "unrar.dll" (ByVal hArcData As Long, ByRef HeaderData As RARHeaderData) As Long
Private Declare Function RARProcessFile Lib "unrar.dll" (ByVal hArcData As Long, ByVal Operation As Long, ByVal DestPath As String, ByVal DestName As String) As Long
Private Declare Sub RARSetChangeVolProc Lib "unrar.dll" (ByVal hArcData As Long, ByVal mode As Long)
Private Declare Sub RARSetPassword Lib "unrar.dll" (ByVal hArcData As Long, ByVal Password As String)
    'above is the Unrar.Dll stuff, below is mine
    'Dim FullPathToRar As String




' here lieth my code to find filenames and whatnot
'aims: _
(tick) filename list _
(tick)packed and unpacked size _
(tick)progress bar _
(tick) Use of file flags more _
unrar "one" file (not sure its possible) _
(half tick)understand the Rar code more!!! _
(tick) read and display comments



Function RARExtract(ByVal ReqdFunction As String, ByVal sRARArchive As String, Optional ByVal sDestPath As String, Optional ByVal sPassword As String, Optional ByVal ReqdFolder) As Integer

' Description:-
' Exrtact file(s) from RAR archive.

' Parameters:-
' sRARArchive   = RAR Archive filename
' sDestPath     = Destination path for extracted file(s)
' sPassword     = Password [OPTIONAL]

' Returns:-
' Integer       = 0  Failed (no files, incorrect PW etc)
'                 -1 Failed to open RAR archive
'                 >0 Number of files extracted
    
Dim lHandle As Long
Dim lStatus As Long
Dim uRAR As RAROpenArchiveData
Dim uHeader As RARHeaderData
Dim Ret As Long 'if not used, it only shows two items in the list
Dim sStat As String ' the filename

Dim TheIcon As Integer

Dim PropFilename As String
Dim PropFlags As String
Dim PropPassword As Integer
Dim PropFolder As String
Dim PropComment As String
Dim TotalUnpacked As Double
Dim PropCarriesOn As Boolean

Dim Path As String


'PropComment = uHeader.flags And &H8
'MsgBox PropComment

'File count
ArchiveObjects = 0
ExtractedObjects = 0

If Mid(sRARArchive, Len(sRARArchive) - 3, 4) <> ".rar" Then 'if the ext on fullpathtorar is .rar
    MsgBox "You have bypassed the 'select only Rar Archive' function, you cheeky monkey!" & vbCrLf & vbCrLf & "But I am smart and knew you would try!", vbExclamation + vbOKOnly, "Not a Rar Archive"
    Exit Function
End If

'displaying comments by frozenpea of vbcity.com
uRAR.CmtBuf = Space(16384)
uRAR.CmtBufSize = 16384

'Dim PropFileComp As Double
    RARExtract = -1
    
    ' Open the RAR

    uRAR.ArcName = sRARArchive
    uRAR.OpenMode = RAR_OM_EXTRACT
    lHandle = RAROpenArchive(uRAR) 'RAROpen(uRAR)

CorruptFile = False
    ' Failed to open RAR ?
    

        If uRAR.OpenResult <> 0 Then
            GoTo ErrDefs
            'MsgBox "Corrupt File", vbCritical + vbOKOnly
            'CorruptFile = True
            'Exit Function
        End If
    
    ' Password ?
    
    If sPassword <> "" Then
        RARSetPassword lHandle, sPassword
    End If
    
    ' Extract file(s)...
    
    ' Is there at lease one archived file to extract ?
    lStatus = RARReadHeader(lHandle, uHeader) 'RARReadHdr(lHandle, uHeader)




'/////////////////////////////
'TotalUnpacked = 0
    Do Until lStatus <> 0
    DoEvents ' keep it responsive
        If StopProcessing = True Then
            frmProgress.lblProcessingFile.Caption = "Stopping. Please be patient"
            Exit Do 'exit loop early
        End If

        If ReqdFunction = "Extract" Then
            'Process (extract) the current file within the archive

           
            If RARProcessFile(lHandle, RAR_EXTRACT, "", sDestPath + uHeader.FileName) = 0 Then
                Debug.Print uHeader.flags And &H1
                
                'Checks to see if the archive spans volumes, if it does, the size is reported wrongly (size*no of volumes wrongly)
                'if it does span volumes, don't count the size
                PropCarriesOn = uHeader.flags And &H1
                If PropCarriesOn = True Then
                    'do nothing
                Else
                    TotalUnpacked = TotalUnpacked + uHeader.UnpSize 'how much data has been currently extracted
                    ExtractedObjects = ExtractedObjects + 1
                End If
                '/////////// progress bar code
                With frmProgress
                        .lblProcessingFile = (Left(uHeader.FileName, InStr(1, uHeader.FileName, vbNullChar) - 1))
                    If TotalArchiveSize <> 0 Then 'if its 0, will cause an error + if its 0, the archive unpacked is 0 which is daft
                        .lblProgress.Caption = CInt((TotalUnpacked / TotalArchiveSize) * 100) & " %" 'change %age text
                        .lblFileNumber.Caption = ExtractedObjects & " / " & ArchiveObjectsStatic
                        'frmOutput.Refresh 'needed to show the value thoughout process
                        .ProgressBar1.Max = TotalArchiveSize 'set progressbar to 100 max
                        .ProgressBar1.Value = TotalUnpacked ' advance the progressbar to how much has been unpacked
                        .ProgressBar2.Max = ArchiveObjectsStatic
                        .ProgressBar2.Value = ExtractedObjects
                    Else
                        .ProgressBar1 = 100
                        .lblProgress = "100 %"
                        
                    End If
                End With
                
                ' Is there another archived file in this RAR ?
                lStatus = RARReadHeader(lHandle, uHeader) 'generates a code, the Defs are at the top.Allows exiting of loop 'RARReadHdr(lHandle, uHeader)
            
            'below code checks if there is an error and exits the function
            'IMPORTANT needs to be last otherwise it exits prematurely
            ElseIf RARProcessFile(lHandle, RAR_EXTRACT, "", sDestPath + uHeader.FileName) <> 0 Then 'extracts the rar archive  'RARProcFile(lHandle, RAR_EXTRACT, "", sDestPath + uHeader.FileName) = 0 Then
                Debug.Print RARProcessFile(lHandle, RAR_EXTRACT, "", sDestPath + uHeader.FileName)
                MsgBox "Unexpected end of archive", vbExclamation, "Error"
                RARCloseArchive lHandle 'close archive
                Exit Function
            End If
'/////////////////////////////////
        ElseIf ReqdFunction = "ObtainList" Then
        'list file code
        

            sStat = Left(uHeader.FileName, InStr(1, uHeader.FileName, vbNullChar) - 1) 'allows listbox to have more than set of data. sStat is the data for listbox
            PropFlags = uHeader.flags
            
            
            
            PropPassword = uHeader.flags And &H400 ' the password flag
            If PropPassword = 0 Then
                PropPassword = 0
            ElseIf PropPassword > 0 Then 'check for password flag
                PropPassword = 3 ' if the flag is set to true tell the user
                PasswordStatus = True
            
            End If
            
            PropFolder = uHeader.flags And &HE0
            
            'places the items in the archive into the listview and gives them their proper icon
            If PropFolder = &HE0 Then 'folder flag
                TheIcon = PropPassword + 2
            Else
                TheIcon = PropPassword + 1
            End If
            
            'checks to see if it is a toplevel folder. Prevents errors with long ass path code
            If InStr(1, sStat, "\") = 0 Then
            
                Set Lisx = frmSource.ListView.ListItems.Add(, , sStat, TheIcon, TheIcon)
            Else
                Set Lisx = frmSource.ListView.ListItems.Add(, , Right(sStat, InStr(1, StrReverse(sStat), "\") - 1), TheIcon, TheIcon)
                
            End If
                'fill the listboxes properties up
                Lisx.SubItems(1) = FilesSize(uHeader.PackSize)
                Lisx.SubItems(2) = FilesSize(uHeader.UnpSize)
                Lisx.SubItems(3) = ProcessDate(uHeader.FileTime)
                Lisx.SubItems(4) = Hex(uHeader.FileCRC)
                'checks to see if the object is a toplevel folder or not
                Path = Left(sStat, InStr(1, sStat, Right(sStat, InStr(1, StrReverse(sStat), "\"))))
                If Len(Path) = 1 Then Path = ""
                Lisx.SubItems(5) = Path

            
            Debug.Print "Carries on from before:"; PropFlags And &H1
            
            PropCarriesOn = PropFlags And &H1
            If PropCarriesOn = True Then
                'do not add the file size to the total size
            Else
                TotalArchiveSize = TotalArchiveSize + uHeader.UnpSize 'calculate uncompressed size
                ArchiveObjects = ArchiveObjects + 1
                ArchiveObjectsStatic = ArchiveObjects
            End If
            
            Ret = RARProcessFile(lHandle, RAR_SKIP, "", "")
            'FilesInArchive.List1.List(FilesInArchive.List1.ListCount - 1) = FilesInArchive.List1.List(FilesInArchive.List1.ListCount - 1)

            lStatus = RARReadHeader(lHandle, uHeader) 'SCROLLS THROUTH THE LIST & gereates a code, defs are at top. Allows exiting of loop
            
            frmSource.ListView.View = lvwReport
            
        End If
'/////////////////////

 '///////////////////////////////
    Loop

If ReqdFunction = "ObtainList" Then
'shows the comment if needed
    Debug.Print "Comment:"; uRAR.CmtState
    If uRAR.CmtState = 1 Then 'there is an archive comment so display it
        ArchiveComment = uRAR.CmtBuf
        'MsgBox ArchiveComment, vbOKOnly, "Archive Comment"
    Else 'there isn't a comment so clear the "buffer"
        ArchiveComment = ""
    End If
End If

'filenames are encrypted?
        If lStatus = 21 Then
            MsgBox "Filenames encrypted. To see the filenames, enter a password in the 'Password' box" _
            & vbCrLf & vbCrLf & "then click on the 'Reload Archive Data' button at the bottom of window ", vbInformation + vbOKOnly
            PasswordStatus = True
        End If
    ' Close the RAR
    RARCloseArchive lHandle 'RARClose lHandle

    ' Return

    RARExtract = iFileCount

If PasswordStatus = True Then
    'bottom lines enables the password field
    frmSource.txtPassword.Visible = True
    frmSource.lblpassword.Visible = True
    frmSource.btnPasswordHelp.Visible = True
    frmSource.btnReload.Enabled = True
    'frmSource.txtPassword.BackColor = vbWhite
End If

    
Exit Function

ErrDefs:
Select Case uRAR.OpenResult
    Case 10
        MsgBox "Unexpected End of Archive", vbExclamation + vbOKOnly, "Error code 10"
    Case 11
        MsgBox "Not enough memory to open the archive", vbOKOnly + vbExclamation, "Error code 11"
    Case 12
        MsgBox "The archive header corrupt or damaged", vbOKOnly + vbCritical, "Error code 12"
        CorruptFile = True
    
    Case 13
        MsgBox "The archive is corrupt or damaged", vbOKOnly + vbCritical, "Error code 13"
        CorruptFile = True
    Case 14
        MsgBox "The Comment is in an unknown format", vbExclamation + vbOKOnly, "Error code 14"
    Case 15
        MsgBox "There was an error that occured when the archive was opened", vbOKOnly + vbCritical, "Error code 15"
    Case 16
        MsgBox "There was an error when the file was created", vbCritical + vbOKOnly, "Error code 16"
    Case 17
        MsgBox "There was an error closing the archive meaning it is still in the memory. Please terminate the De-rar.exe process to claim this memory space back", vbCritical + vbOKOnly, "Error code 17"
End Select

RARCloseArchive lHandle 'close archive


End Function

