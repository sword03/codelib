'函数：设置ini值（ini路径，目标节点，目标键，目标值）
'注：若ini文件不存在则创建；节点或键不存在则添加
Function SetIniValue(path, sectionName, keyName, value)

Dim fso,file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile(path, 1,true)

Dim line, cache, inSection, sectionExist, keyExist

Do Until file.AtEndOfStream
	line = file.Readline
	if StrComp(Trim(line),"["+sectionName+"]",1)=0 Then
	  inSection=True
	  sectionExist=True
	End If
	if inSection And Left(LTrim(line),1)="[" And StrComp(Trim(line),"["+sectionName+"]",1)<>0 Then
	  inSection=False
	  If Not keyExist Then
		cache = cache + keyName+"="+value+vbCrLf
		keyExist=True
	  End If
	End If

	if inSection And InStr(line,"=")<>0 Then
	  ss = Split(line,"=")
	  If StrComp(Trim(ss(0)),keyName,1)=0 Then
		line = ss(0)+"="+value
		keyExist = True
	  End If
	End If

	cache=cache+line+vbcrlf
Loop

file.Close

If not sectionExist Then
  cache = cache + "["+sectionName+"]"+vbCrLf
  cache = cache + keyName+"="+value+vbCrLf
ElseIf Not keyExist Then
  cache = cache + keyName+"="+value+vbCrLf
End If

Set file = fso.OpenTextFile(path, 2, True)
	file.Write(cache)
	file.Close
End Function

Function DelSection(path, sectionName )

Dim fso,file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile(path, 1,true)

Dim line, cache, inSection, sectionExist, keyExist

Do Until file.AtEndOfStream
	line = file.Readline
	if StrComp(Trim(line),"["+sectionName+"]",1)=0 Then
	  inSection=True
	  sectionExist=True
	End If
	if inSection And Left(LTrim(line),1)="[" And StrComp(Trim(line),"["+sectionName+"]",1)<>0 Then
	  inSection=False
	End If

	If Not inSection Then 
		cache=cache+line+vbcrlf
	End If
Loop

file.Close

Set file = fso.OpenTextFile(path, 2, True)
	file.Write(cache)
	file.Close
End Function

SetIniValue "E:\gpt.ini","Other","ccc","aaa"