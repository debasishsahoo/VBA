Select Case Selection.Range.Kana 
 Case wdKanaHiragana 
 MsgBox "This text is hiragana." 
 Case wdKanaKatakana 
 MsgBox "This text is katakana." 
 Case wdUndefined 
 MsgBox "This text is a mix of " _ 
 & "hiragana and katakana." 
End Select