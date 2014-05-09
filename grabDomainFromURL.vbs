Function Rgx(p,i,g,l):Dim o:Set o=New RegExp:o.Pattern=p:o.IgnoreCase=CBool(i):o.Global=CBool(g):o.MultiLine=CBool(l):Set Rgx=o:End Function

Function grabDomainFromURL(t)
 Dim o,s
 s=""
 Set o=Rgx("^.*://(?:[wW]{3}\.)?([^:/]*).*$",1,1,0)
 Set o=o.Execute(t)
 if o.Count > 0 Then
	Set o=o.Item(0).SubMatches
	if o.Count > 0 Then
		s=o.Item(0)
	End if
 End if
 grabDomainFromURL=s
End Function

Response.Write  grabDomainFromURL("http://192.168.25.40:81/landing/zengclub.it/extras/core/step1.asp")

