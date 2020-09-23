VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileCopyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "fast file copying"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarCacheSize As Long 'local copy
Private mvarPercentREady As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarSucessfull As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarSecondsTaken As Long 'local copy
Public Property Let SecondsTaken(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SecondsTaken = 5
    mvarSecondsTaken = vData
End Property


Public Property Get SecondsTaken() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SecondsTaken
    SecondsTaken = mvarSecondsTaken
End Property



Public Property Let Sucessfull(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sucessfull = 5
    mvarSucessfull = vData
End Property


Public Property Get Sucessfull() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sucessfull
    Sucessfull = mvarSucessfull
End Property



Public Sub Copy(SourceFile As String, Destination As String)
Attribute Copy.VB_Description = "guess :)"

If mvarCacheSize = 0 Then
ReadChunkSize& = 32000
Else
ReadChunkSize& = mvarCacheSize
End If

BufferSize& = ReadChunkSize& * 2

If Dir(SourceFile) = "" Then              '» If source file does not exist,

    mvarSucessfull = False                   '» Set return value;
    Exit Sub                              '»  well, bye bye;
End If

If Dir(Destination) <> "" Then            '» If the destination file already exists,
Kill Destination                          '» lets kill it.
End If

FCStTime = Timer                          '» lets time how long it takes...

infil = FreeFile                          '» gets a FreeFile handle;
Open SourceFile For Binary As #infil      '» Opens the input file;
outfil = FreeFile                         '» get another freefile handle;
Open Destination For Binary As #outfil    '» Open the output file;


If ReadChunkSize& > LOF(infil) Then   '» Basically, if the file size is
ReadChunkSize& = LOF(infil)           '» smaler than the read chuck size;
r$ = String$(ReadChunkSize&, Chr$(0)) '»
Get #infil, , r$                      '» Just read the file all at once,
Put #outfil, , r$                     '» & flush it now;
GoTo CLoseFiles                       '» Most people say using GOTO is a bad Qbasic habit
End If                                '» but hell it works fine here.


Dim Buffer$                           '» QB habit, don't complain;

Do While Not Loc(infil) = LOF(infil)
r$ = String$(ReadChunkSize&, Chr$(0))
Get #infil, , r$
Buffer$ = Buffer$ + r$

        If Len(Buffer$) >= BufferSize& Then
                Put #outfil, , Buffer$
                Buffer$ = ""
        End If


                CurPosOn = Loc(outfil) / 100
                Totalto = LOF(infil) / 100
                mvarPercentREady = (CurPosOn * 100) / Totalto
                
                DoEvents

If LOF(infil) - Loc(infil) < ReadChunkSize& Then   '» This changes the read chunk size as we
ReadChunkSize& = LOF(infil) - Loc(infil)           '» get to the end of the file;
End If

Loop
If Len(Buffer$) <> 0 Then          '» Check if there is anything else on the buffer;
Put #outfil, , Buffer$             '» If thats the case, lets flush it.
End If

CLoseFiles:
Close #infil, #outfil
mvarSucessfull = True
EndOfCTime = Timer

mvarSecondsTaken = EndOfCTime - FCStTime  '» Boom, end timer..

End Sub

Public Property Let PercentREady(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PercentREady = 5
    mvarPercentREady = vData
End Property


Public Property Get PercentREady() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PercentREady
    PercentREady = mvarPercentREady
End Property



Public Property Let CacheSize(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CacheSize = 5
    mvarCacheSize = vData
End Property


Public Property Get CacheSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CacheSize
    CacheSize = mvarCacheSize
End Property



