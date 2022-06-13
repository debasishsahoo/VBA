Dim strXML As String 
 
strXML = "<"xml version=""1.0""><abc:books xmlns:abc=""urn:books"" " & _ 
 "xmlns:xsi=""https://www.w3.org/2001/XMLSchema-instance"" " & _ 
 "xsi:schemaLocation=""urn:books books.xsd""><book>" & _ 
 "<author>Matt Hink</author><title>Migration Paths of the Red " & _ 
 "Breasted Robin</title><genre>non-fiction</genre>" & _ 
 "<price>29.95</price><pub_date>2006-05-01</pub_date>" & _ 
 "<abstract>You see them in the spring outside your windows. " & _ 
 "You hear their lovely songs wafting in the warm spring air. " & _ 
 "Now follow their path as they migrate to warmer climes in the fall, " & _ 
 "and then back to your back yard in the spring.</abstract></book></abc:books>" 
 
Selection.InsertXML strXML