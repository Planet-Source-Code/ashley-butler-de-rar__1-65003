Attribute VB_Name = "modLinkBetweenForms"
Public Password As String ' needed so that encrypted files can be decrypted in one go
Public TotalArchiveSize As Double 'used for progressbar, Total unpacked archive size
Public PasswordStatus As Boolean ' test if theres a password so that user can't change page until one is supplied
Public CorruptFile As Boolean 'is tested when the user tries to select output form
Public StopProcessing As Boolean
Public ArchiveObjects As Long
Public ArchiveObjectsStatic As Long
Public ExtractedObjects As Long
Public Lisx As ListItem
Public Colx As ColumnHeader

Public ArchiveComment As String
