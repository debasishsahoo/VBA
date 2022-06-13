Selection.EndOf Unit:=wdStory, Extend:=wdMove 
Set myRev = Selection.PreviousRevision 
If Not (myRev Is Nothing) Then MsgBox myRev.Date



Set myRev = Selection.PreviousRevision(Wrap:=True) 
If Not (myRev Is Nothing) Then 
 Select Case myRev.Type 
 Case wdRevisionDelete 
 myRev.Reject 
 Case wdRevisionInsert 
 myRev.Reject 
 Case wdRevisionStyle 
 myRev.Accept 
 End Select 
End If