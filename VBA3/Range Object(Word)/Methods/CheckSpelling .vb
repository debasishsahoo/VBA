Set range2 = Documents("MyDocument.doc").Sections(2).Range 
range2.CheckSpelling IgnoreUpperCase:=False, _ 
 CustomDictionary:="MyWork.Dic", _ 
 CustomDictionary2:="MyTechnical.Dic"