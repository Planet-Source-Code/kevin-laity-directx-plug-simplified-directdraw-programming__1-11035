Attribute VB_Name = "PRF"

'This BAS file is all you need for using a file
'packaging system I (quickly) devised called
'Packaged Resource Format (PRF)
'
'Basically the idea behind this is that most commercial
'products store all their graphic files and data files
'in big packages.  Its just more neat and professional
'looking than shoving all your graphics and sounds into
'a directory for your user to see and fiddle with.
'
'Anyways, you can use the packager program to make your
'resource files, just select the files you want and
'click "add", then click "pack".  This will add all the
'files you have selected (make sure the files are in
'the same directory as the packager) to a file called
'outfile.prf
'
'To use the package, add the prf.bas file to your
'project, and use the unpack function.  The function
'returns a string containing the exact path and file
'name of the file after it is extracted.
'
'If you extract another file, the first file will be
'deleted.  When you are done extracting the files you
'need, call Clear_Previous to delete the last file you
'extracted,  and viola!

Dim msg As String
Dim lastfileunpacked As String

Public Sub Clear_Previous()

If lastfileunpacked = "" Then Exit Sub

On Error Resume Next
Kill lastfileunpacked


End Sub


Function Close_File()

Close 1

End Function


Function Open_File_For_Packing(outfile As String)

Open outfile For Binary As 1


End Function


Function Pack_File(ByVal filename As String)


Dim msg As String
Dim inty As Single

inty = Len(filename)
Put #1, , inty
Put #1, , filename
inty = FileLen(App.Path & "\" & filename)
Put #1, , inty

Open App.Path & "\" & filename For Binary As 2

msg = String$(inty, " ")
Get #2, , msg
Put #1, , msg

Close #2
msg = ""


End Function


Function Unpack(infile, outfile) As String

Clear_Previous

Dim i As Single
Dim inty As Single
Dim msg As String

Open infile For Binary As 1

i = 1

Do While Not EOF(1)
Get #1, i, inty
msg = String$(inty, " ")
i = i + 4
Get #1, i, msg
i = i + inty
Get #1, i, inty
i = i + 4

    If UCase(msg) = UCase(outfile) Then
        Open App.Path & "\data\" & outfile For Binary As 2
           msg = String$(inty, " ")
           Get #1, i, msg
           Put #2, , msg
        Close 2
        Close 1
        lastfileunpacked = App.Path & "\data\" & outfile
        Unpack = lastfileunpacked
        Exit Function
    End If
i = i + inty
Loop
Close 1



End Function


