﻿Imports System
Imports System.Text
Imports System.Xml


Module SCRename
	Sub Main(ByVal CmdArgs() As String)
		'SCRename Ver. 6.0
		Console.WriteLine("======================================")
		Console.WriteLine("SCRename Ver. 6.0")
		Console.WriteLine("======================================")

		'変数定義
		Const sep As String = "_"
		Const nind As String = "#"
		Const char1 As String = ChrW(&H20) & ChrW(&HFF1A) & ChrW(&HFF1B) & ChrW(&HFF0F) & ChrW(&H27) & ChrW(&H22) & ChrW(&H201D) & ChrW(&H2010)
		Const char2 As String = ChrW(&H3000) & ChrW(&HFF1A) & ChrW(&HFF1B) & ChrW(&HFF0F) & ChrW(&H2019) & ChrW(&H201D) & ChrW(&H2010)
		Const char3 As String = "!？！～…『』"
		Const char4 As String = "([（〔［｛〈《「【＜"
		Const char5 As String = ")]）〕］｝〉》」】＞"
		Const char6 As String = "/:*?!""<>|"
		Const char7 As String = ChrW(&H2215) & ChrW(&HFF1A) & ChrW(&HFF0A) & ChrW(&HFF1F) & ChrW(&HFF01) & ChrW(&H201D) & ChrW(&HFF1C) & ChrW(&HFF1E) & ChrW(&HFF5C)
		Dim i As Object, j As Object, k As Object, l As Object, argc As Integer, opt As Object, elen As Object, days As Object, pos As Object, serv As Object
		Dim str1 As Object, str2 As Object, path As Object, rpath As Object, ext As Object, title As Object, title2 As Object, ftitle As Object, number As Object, number1 As Object, number2 As Object, number3 As Object, number4 As Object, part As Object, subtitle As Object
		Dim yr As Object, mon As Object, dy As Object, hr As Object, min As Object, sec As Object
		Dim dt1 As Date?, dt2 As Date?, tgtdt As Date?, stdt As Date?, eddt As Date?
		Dim dtflag As Object
		Dim service(,) As String
		Dim tid(,) As String
		Dim objFSO As Object, objFile As Object, objHTTP As Object
		Dim str8 As String = "I II III IV V VI VII VIII IX X"
		Dim str9 As String = "quot amp #039 lt gt"
		Dim str10 As String = ChrW(&H201C) & " & " & Chr(39) & " ＜ ＞"
		Dim str11 As String = "Sun Mon Tue Wed Thu Fri Sat"
		Dim char8() As String = str8.Split(" ")
		Dim char9() As String = str9.Split(" ")
		Dim char10() As String = str10.Split(" ")
		Dim char11() As String = str11.Split(" ")
		Dim xmlFilePath As String = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SCRename.xml")


		'引数処理
		For Each str1 In CmdArgs
			str1 = str1.Replace("""", "")
			If (str1).ToLower() = "-h" Or str1 = "-?" Then
				Console.WriteLine(Environment.NewLine & "SCRename.vbs [オプション] ファイル リネーム書式")
				Console.WriteLine(" [タイトル開始位置] [検索文字数]" & Environment.NewLine)
				Environment.Exit(1)
			ElseIf (str1).ToLower() = "-t" Then
				opt = opt Or 1
			ElseIf (str1).ToLower() = "-n" Then
				opt = opt Or 2
			ElseIf (str1).ToLower() = "-f" Then
				opt = opt Or 4
			ElseIf (str1).ToLower() = "-s" Then
				opt = opt Or 8
			ElseIf (str1).ToLower() = "-a" Then
				opt = opt Or 16
			ElseIf (str1).ToLower() = "-a1" Then
				opt = opt Or 32
			Else
				ReDim Preserve CmdArgs(argc)
				CmdArgs(argc) = str1
				argc = argc + 1
			End If
		Next

		'起動時処理
		Threading.Thread.Sleep(1000)
		Console.Error.WriteLine(Environment.NewLine + "SCRename 動作中..." + Environment.NewLine)
		If argc < 2 Then
			Console.WriteLine(CmdArgs(0))
			Console.Error.WriteLine("パラメータが足りません。")
			Threading.Thread.Sleep(1000)
			Environment.Exit(1)
		ElseIf CmdArgs(0) = "" Then
			Console.Error.WriteLine("処理対象のファイルが指定されていません。")
			Threading.Thread.Sleep(1000)
			Environment.Exit(1)
		ElseIf CmdArgs(1) = "" Then
			Console.WriteLine(CmdArgs(0))
			Console.Error.WriteLine("リネーム書式が指定されていません。")
			Threading.Thread.Sleep(1000)
			Environment.Exit(1)
		End If
		Console.WriteLine("起動時処理終了")

		'しょぼかるユーザ名取得
		Dim xmlDoc As New XmlDocument()
		xmlDoc.Load(xmlFilePath)
		Dim usrNode As XmlNode = xmlDoc.SelectSingleNode("/config/settings/setting[@name='usr']")
		If usrNode IsNot Nothing Then
			Console.WriteLine("user: " & usrNode.InnerText)
		End If
		Console.WriteLine("しょぼかるユーザ名取得終了")

		'実体名最大文字数取得
		elen = 0
		For i = 0 To UBound(char9)
			j = Len(char9(i))
			If j > elen Then
				elen = j
			End If
		Next
		elen = elen + 2
		Console.WriteLine("実体名最大文字数取得終了")

		'SCRename.exc 読み込み
		path = Left(System.Reflection.Assembly.GetExecutingAssembly().Location, InStrRev(System.Reflection.Assembly.GetExecutingAssembly().Location, "\"))
		str1 = path + "SCRename.exc"
		objFSO = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.FileSystemObject"))
		If objFSO.FileExists(str1) Then
			objFile = objFSO.OpenTextFile(str1, 1, False, -2)
			Do While Not objFile.AtEndOfStream
				str1 = objFile.ReadLine
				If Left(str1, 1) <> ":" Then
					If InStr((CmdArgs(0).ToUpper()), (str1).ToUpper()) > 0 Then
						i = -1
					End If
				End If
			Loop
			objFile.Close
			objFile = Nothing
			If i < 0 Then
				Console.WriteLine(CmdArgs(0))
				Console.Error.WriteLine("対象外のファイルのため処理しませんでした。")
				Environment.Exit(1)
			End If
		End If
		Console.WriteLine("SCRename.exc 読み込み終了")

		'リネーム元ファイル存在確認
		If (opt And 1) = 0 Then
			If Not objFSO.FileExists(CmdArgs(0)) Then
				Console.Error.WriteLine(CmdArgs(0) + " がありません。")
				Threading.Thread.Sleep(1000)
				Environment.Exit(1)
			End If
		End If
		Console.WriteLine("リネーム元ファイル存在確認終了")

		'SCRename.srv 読み込み
		str1 = path + "SCRename.srv"
		objFile = objFSO.OpenTextFile(str1, 1, False, -2)
		If Err.Number <> 0 Then
			Console.WriteLine(CmdArgs(0))
			Console.Error.WriteLine(str1 + " がありません。")
			Threading.Thread.Sleep(1000)
			Environment.Exit(1)
		End If
		i = 0
		Dim ilinecount As Integer = 0
		Do While Not objFile.AtEndOfStream
			str1 = objFile.ReadLine
			If Left(str1, 1) <> ":" Then
				ReDim Preserve service(3, ilinecount)
				ilinecount = ilinecount + 1
				j = 0
				k = 0
				Do While j < 4
					l = InStr(k + 1, str1, ",")
					If l > 0 Then
						service(j, i) = Mid(str1, k + 1, l - k - 1)
					Else
						service(j, i) = Mid(str1, k + 1)
						Exit Do
					End If
					k = l
					j = j + 1
				Loop
				If service(0, i) <> "" Then
					i = i + 1
				End If
			End If
		Loop
		objFile.Close
		objFile = Nothing
		Console.WriteLine("SCRename.srv 読み込み終了")

		'ファイルパス、ファイル名、拡張子、タイトル開始位置取得
		i = InStrRev(CmdArgs(0), "\")
		If i > 0 Then
			rpath = Left(CmdArgs(0), i)
		ElseIf Mid(CmdArgs(0), 2, 1) = ":" Then
			rpath = Left(CmdArgs(0), 2)
			i = 2
		End If
		title = Mid(CmdArgs(0), i + 1)
		i = InStrRev(title, sep)
		If i > 0 Then
			i = InStr(i + 1, title, ".")
		Else
			i = InStrRev(title, ".")
		End If
		If i > 0 Then
			k = 0
			For j = Len(title) To i Step -1
				str1 = Mid(title, j, 1)
				If str1 = "." Then
					k = j
				End If
				If str1 <> "." And (str1 < "0" Or str1 > "9") And (str1 < "A" Or str1 > "Z") And (str1 < "a" Or str1 > "z") Then
					Exit For
				End If
			Next
			If k > 1 Then
				ext = Mid(title, k)
				title = Left(title, k - 1)
			End If
		End If
		If argc > 2 Then
			If IsNumeric(CmdArgs(2)) Then
				pos = CInt(CmdArgs(2))
			End If
		End If
		Console.WriteLine("ファイルパス、ファイル名、拡張子、タイトル開始位置取得終了")

		'日付取得
		days = 1
		l = Len(title)
		If l > 8 Then
			yr = Year(DateTime.Now)
			For i = 1 To l - 5
				If IsNumeric(Mid(title, i, 1)) Then
					For j = i + 1 To l
						If Not IsNumeric(Mid(title, j, 1)) Then
							Exit For
						End If
					Next
					If j - i > 5 Then
						k = yr - CInt(Mid(title, i, 4))
						If j - i < 8 Or k < -2 Or k > 99 Then
							str1 = Left(CStr(yr), 2) + Mid(title, i, 6)
							k = i + 6
						Else
							str1 = Mid(title, i, 8)
							k = i + 8
						End If
						If CInt(Left(str1, 4)) < yr + 3 Then
							str1 = Left(str1, 4) + "/" + Mid(str1, 5, 2) + "/" + Mid(str1, 7, 2)
							If IsDate(str1) Then
								tgtdt = CDate(str1)
								If i = 1 Then
									pos = j
								End If
								Exit For
							End If
						End If
					End If
				End If
			Next
		End If
		If pos = 0 Then
			pos = 1
		End If
		If tgtdt Is Nothing And k < l - 2 Then
			str1 = Mid(title, k, 1)
			If Not IsNumeric(str1) And str1 <> sep Then
				k = k + 1
			End If
			If IsNumeric(Mid(title, k, 4)) Then
				i = CInt(Mid(title, k, 2))
				j = 0
				If i > 23 Then
					i = i - 24
					j = 1
				End If
				str1 = CStr(i) + ":" + Mid(title, k + 2, 2)
				If IsDate(str1) Then
					tgtdt = DateAdd("d", j, tgtdt) + CDate(str1)
					dtflag = 1
				End If
			End If
		End If
		If dtflag = 0 And objFSO.FileExists(CmdArgs(0)) Then
			dt1 = objFSO.GetFile(CmdArgs(0)).DateCreated
			dt2 = objFSO.GetFile(CmdArgs(0)).DateLastModified
			If dt1 < dt2 Then
				dt2 = dt1
				dtflag = 1
			End If
			If tgtdt Is Nothing Then
				tgtdt = dt2
				days = 7
			Else
				tgtdt = tgtdt + New TimeSpan(Hour(dt2), Minute(dt2), Second(dt2))
			End If
		End If
		If tgtdt Is Nothing Then
			tgtdt = DateTime.Now
		End If
		Console.WriteLine("日付取得終了")

		'ファイル名先頭部分削除
		If pos > 1 Then
			title = Mid(title, pos)
		End If
		Console.WriteLine("ファイル名先頭部分削除終了")

		'SCRename.rp1 読み込み＆ファイル名置換
		str1 = path + "SCRename.rp1"
		If objFSO.FileExists(str1) Then
			objFile = objFSO.OpenTextFile(str1, 1, False, -2)
			Do While Not objFile.AtEndOfStream
				str1 = objFile.ReadLine
				If Left(str1, 1) <> ":" Then
					i = InStr(str1, ",")
					If i > 0 Then
						title = (title).Replace(Left(str1, i - 1), Mid(str1, i + 1))
					End If
				End If
			Loop
			objFile.Close
			objFile = Nothing
		End If
		Console.WriteLine("SCRename.rp1 読み込み＆ファイル名置換終了")

		'先頭部分記号削除
		i = 1
		Do While True
			str1 = Mid(title, i, 1)
			If InStr(sep + char1 + char2 + char3 + "・", str1) < 1 Then
				j = InStr(char4, str1)
				If j < 1 Then
					Exit Do
				Else
					k = InStr(i + 1, title, Mid(char5, j, 1))
					If k > 0 Then
						i = k
					Else
						Exit Do
					End If
				End If
			ElseIf str1 = "『" Then
				k = InStr(i + 1, title, "』")
				If k > 1 Then
					title = Left(title, k - 1) + " " + Mid(title, k + 1)
				End If
			End If
			i = i + 1
			If i > Len(title) Then
				Console.WriteLine(CmdArgs(0))
				Console.Error.WriteLine("タイトルを取得出来ませんでした。")
				Threading.Thread.Sleep(1000)
				Environment.Exit(1)
			End If
		Loop
		title = Mid(title, i)
		Console.WriteLine("先頭部分記号削除終了")

		'全角半角ローマ数字記号変換
		str1 = ""
		For i = 1 To Len(title)
			str2 = Mid(title, i, 1)
			j = InStr(char2, str2)
			If j > 0 Then
				str1 = str1 + Mid(char1, j, 1)
			Else
				k = Asc(str2)
				If k >= Asc("０") And k <= Asc("９") Then
					str1 = str1 + Chr(k - Asc("０") + Asc("0"))
				ElseIf k >= Asc("Ａ") And k <= Asc("Ｚ") Then
					str1 = str1 + Chr(k - Asc("Ａ") + Asc("A"))
				ElseIf k >= Asc("ａ") And k <= Asc("ｚ") Then
					str1 = str1 + Chr(k - Asc("ａ") + Asc("a"))
				ElseIf k >= Asc("Ⅰ") And k <= Asc("Ⅹ") Then
					str1 = str1 + char8(k - Asc("Ⅰ"))
				Else
					str1 = str1 + str2
				End If
			End If
		Next
		Console.WriteLine("全角半角ローマ数字記号変換終了")

		'タイトル取得
		ftitle = str1
		l = Len(str1)
		i = 4
		If argc > 3 Then
			If IsNumeric(CmdArgs(3)) Then
				i = CInt(CmdArgs(3))
			End If
		End If
		If i < 1 Then
			i = 4
		End If
		If i < l Then
			l = i
		End If
		title = Left(str1, 1)
		For i = 2 To l
			str2 = Mid(str1, i, 1)
			If InStr(" " + sep + char3 + char4 + char5, str2) > 0 Then
				Exit For
			Else
				title = title + str2
			End If
		Next
		Console.WriteLine("タイトル取得終了")

		'放送局名取得
		i = InStrRev(str1, sep)
		If pos < 7 And i > 3 Then
			j = InStrRev(str1, sep, i - 2)
			If j > 1 Then
				i = j
			End If
		End If
		str2 = Mid(str1, i + 1)
		For i = 1 To 4
			If i < 3 Then
				j = 0
			Else
				j = 2
			End If
			For serv = 0 To UBound(service, 2)
				If i = 1 Or i = 3 Then
					k = InStrRev((str2).ToUpper(), (service(j, serv).ToUpper()))
				Else
					k = InStrRev((str1).ToUpper(), (service(j, serv).ToUpper()))
				End If
				If k > 0 Then
					i = 0
					Exit For
				End If
			Next
			If i = 0 Then
				Exit For
			End If
		Next
		If i > 0 Then
			serv = -1
			Console.Error.WriteLine("放送局が不明のためすべての放送局を対象にします。")
		End If
		Console.WriteLine("放送局名取得終了")

		'検索開始
		If (opt And 32) = 0 Then
			If days = 7 Then
				str1 = "ファイル名から日付が取得できないためファイルの"
				If dtflag = 1 Then
					str1 = str1 + "作成"
				Else
					str1 = str1 + "更新"
				End If
				Console.Error.WriteLine(str1 + "日から一週間遡って、")
			End If
			If dtflag = 1 Then
				str1 = "開始"
			Else
				str1 = "終了"
			End If
			Console.Error.WriteLine(str1 + "日時が " + tgtdt + " に最も近い")
			If serv < 0 Then
				str1 = ""
			Else
				str1 = "（" + service(1, serv) + "）"
			End If
			Console.Error.WriteLine("「" + title + "」" + str1 + "を検索します。" + Environment.NewLine)
			If Err.Number <> 0 Then
				Console.WriteLine(CmdArgs(0))
				Console.Error.WriteLine("検索前処理でエラーが発生しました。")
				Threading.Thread.Sleep(1000)
				Environment.Exit(1)
			End If
		End If
		Console.WriteLine("検索開始終了")

		'XMLHTTP オブジェクト作成
		objHTTP = Activator.CreateInstance(Type.GetTypeFromProgID("Msxml2.XMLHTTP"))
		If Err.Number <> 0 Then
			objHTTP = Activator.CreateInstance(Type.GetTypeFromProgID("Microsoft.XMLHTTP"))
		End If
		If objHTTP Is Nothing Then
			Console.WriteLine(CmdArgs(0))
			Console.Error.WriteLine("XMLHTTP オブジェクトを作成できませんでした。")
			Threading.Thread.Sleep(1000)
			Environment.Exit(1)
		End If
		Console.WriteLine("XMLHTTP オブジェクト作成終了")

		'しょぼいカレンダーより情報取得
		If (opt And 32) = 0 Then
			dt1 = DateAdd("d", -days, tgtdt)
			Dim YMDLong As Long = (Year(dt1) * 10000) + (Month(dt1) * 100) + (Day(dt1))
			str1 = YMDLong.ToString() + "0000"
			For i = 0 To 2
				If i > 0 Then
					Threading.Thread.Sleep(1000)
				End If
				Dim URL As String = "http://cal.syoboi.jp/rss2.php?start=" + str1.ToString() + "&days=" + (days + 1).ToString() + "&usr=" + usrNode.InnerText + "&titlefmt=$(Title)|$(ChName)|$(EdTime)|$(SubTitleB)"
				objHTTP.Open("Get", URL, False)
				objHTTP.Send
				If objHTTP.Status >= 200 And objHTTP.Status < 300 Then
					Exit For
				End If
			Next
			If i > 2 Then
				Console.WriteLine(CmdArgs(0))
				Console.Error.WriteLine("しょぼいカレンダーにアクセスできませんでした。")
				Threading.Thread.Sleep(1000)
				Environment.Exit(1)
			End If
			str1 = objHTTP.responseText
			str1 = Mid(str1, InStr(str1, "<item>") + 6)
			For i = 1 To Len(str1)
				str2 = Mid(str1, i, 1)
				If str2 = "\" Then
					str1 = Left(str1, i - 1) + "＼" + Mid(str1, i + 1)
				ElseIf str2 = "&" Then
					Do
						For j = 1 To elen - 1
							str2 = Mid(str1, i + j, 1)
							If str2 = "&" Then
								i = i + j
								Exit For
							ElseIf str2 = ";" Then
								str2 = Mid(str1, i + 1, j - 1)
								For k = 0 To UBound(char9)
									If str2 = char9(k) Then
										Dim strLeft As String = Left(str1, i - 1)
										Dim strMid As String = Mid(str1, i + j + 1)
										str1 = Left(str1, i - 1) + char10(k) + Mid(str1, i + j + 1)
										j = elen
										Exit For
									End If
								Next
							End If
						Next
					Loop While j < elen
				End If
			Next
		End If
		Console.WriteLine("しょぼいカレンダーより情報取得終了")

		'番組情報検索
		pos = 0
		Dim ii As Integer = 0
		i = InStr(str1, "<item>") + 6
		Do While (opt And 32) = 0
			ii = InStr(i, str1, "<title>")
			If ii < 1 Then
				Exit Do
			End If
			i = ii + 7
			ii = ii + 7
			j = InStr(ii + 1, str1, "|")
			If j < 1 Then
				Exit Do
			End If
			If InStr((Mid(str1, ii, j - ii).ToUpper()), (title).ToUpper()) > 0 Then
				k = InStr(j + 1, str1, "|")
				If serv < 0 Then
					j = -1
				ElseIf InStr((Mid(str1, j + 1, k - j - 1).ToUpper()), (service(1, serv).ToUpper())) > 0 Then
					j = -1
				End If
				If j = -1 Then
					j = InStr(k + 1, str1, "<pubDate>") + 9
					str2 = Mid(str1, j, InStr(j + 10, str1, "+") - j)
					j = InStr(str2, "T")
					str2 = Left(str2, j - 1) + " " + Mid(str2, j + 1)
					If IsDate(str2) Then
						dt1 = CDate(str2)
						str2 = Left(str2, j)
					Else
						dt1 = CDate("0")
						str2 = ""
					End If
					str2 = str2 + Mid(str1, k + 1, 5)
					If IsDate(str2) Then
						dt2 = CDate(str2)
						If dt1 >= dt2 Then
							dt2 = DateAdd("d", 1, dt2)
						End If
					Else
						dt2 = CDate("0")
					End If

					If dtflag = 1 Then
						If stdt Is Nothing Then
							stdt = dt1
							eddt = dt2
							pos = i
						End If
						If Math.Abs(CType(CType((tgtdt - dt1), TimeSpan).Days, SByte)) < Math.Abs(CType(CType((tgtdt - stdt), TimeSpan).Days, SByte)) Then
							stdt = dt1
							eddt = dt2
							pos = i
						End If

					Else
						If eddt Is Nothing Then
							stdt = dt1
							eddt = dt2
							pos = i
						End If
						If Math.Abs(CType(CType((tgtdt - dt2), TimeSpan).Days, SByte)) < Math.Abs(CType(CType((tgtdt - eddt), TimeSpan).Days, SByte)) Then
							stdt = dt1
							eddt = dt2
							pos = i
						End If
					End If
				End If
			End If
		Loop
		Console.WriteLine("番組情報検索終了")

		'番組情報取得
		If pos > 0 Then
			i = InStr(pos, str1, "|")
			title = Mid(str1, pos, i - pos)
			j = InStr(i + 1, str1, "|")
			If serv < 0 Then
				str2 = Mid(str1, i + 1, j - i - 1)
				For k = 0 To UBound(service, 2)
					If InStr(str2, service(1, k)) > 0 Then
						serv = k
						Exit For
					End If
				Next
				If serv < 0 Then
					serv = 0
					service(2, 0) = str2
				End If
			End If
			i = InStr(j + 1, str1, "|")
			j = InStr(i + 1, str1, "</title>")
			If i < j - 1 Then
				str2 = Mid(str1, i + 1, j - i - 1)
				i = InStr(str2, "「")
				If i > 0 Then
					If i > 2 And Left(str2, 1) = "#" Then
						For j = 2 To i - 1
							If Mid(str2, j, 1) = " " Then
								Exit For
							End If
						Next
						number = Left(str2, j - 1)
					ElseIf i > 1 Then
						part = Left(str2, i - 1)
						part = (part).Trim()
					End If
					subtitle = Mid(str2, i + 1)
					If Right(subtitle, 1) = "」" Then
						subtitle = Left(subtitle, Len(subtitle) - 1)
					End If
				ElseIf Left(str2, 1) = "#" Then
					i = 1
					Do While True
						j = InStr(i + 1, str2, " ")
						If j < 1 Then
							number = number + "," + Mid(str2, i)
							Exit Do
						Else
							number = number + "," + Mid(str2, i, j - i)
							i = InStr(j, str2, " / #")
							If i < 1 Then
								If j < Len(str2) Then
									subtitle = subtitle + " ／ " + (Mid(str2, j + 1).Trim())
								End If
								Exit Do
							ElseIf i > j + 1 Then
								subtitle = subtitle + " ／ " + (Mid(str2, j + 1, i - j - 1).Trim())
							End If
						End If
						i = i + 3
					Loop
					number = Mid(number, 2)
					If subtitle <> "" Then
						subtitle = Mid(subtitle, 4)
					End If
				Else
					part = str2
				End If
			End If
		Else
			'話数検索開始
			If (opt And 16) > 0 Or (opt And 32) > 0 Then
				If (opt And 32) = 0 Then
					Console.Error.WriteLine("番組情報が見つかりませんでした。")
				End If
				Console.Error.WriteLine("話数検索を行います。" + Environment.NewLine)
				k = -1
				For i = 2 To Len(ftitle) - 2
					str1 = Mid(ftitle, i, 1)
					If InStr("「『", str1) > 0 Then
						Exit For
					ElseIf InStr(" " + sep, str1) > 0 Then
						str1 = Mid(ftitle, i + 1, 1)
						If str1 = nind Or str1 = "第" Then
							For j = i + 2 To Len(ftitle)
								str2 = Mid(ftitle, j, 1)
								If Not IsNumeric(str2) Then
									Exit For
								End If
							Next
							If j > i + 2 Then
								If str1 = nind And InStr(" " + sep, str2) > 0 Or str1 = "第" And str2 = "話" Then
									k = CInt(Mid(ftitle, i + 2, j - i - 2))
									Exit For
								End If
							End If
						End If
					End If
				Next
				If k = -1 Then
					Console.Error.WriteLine("ファイル名から話数を取得できませんでした。" + Environment.NewLine)
				Else
					number = CStr(k)
					For j = 2 To i - 1
						If InStr(sep + char4 + char5 + "～", Mid(ftitle, j, 1)) > 0 Then
							Exit For
						End If
					Next
					title = (Left(ftitle, j - 1).TrimEnd())
					title2 = ((title).Replace(" ", "").ToUpper())
					k = -1
					str1 = path + "SCRename.tid"
					If objFSO.FileExists(str1) Then
						objFile = objFSO.OpenTextFile(str1, 1, False, -2)
						i = 0
						Do While Not objFile.AtEndOfStream
							ReDim Preserve tid(1, i)
							str1 = objFile.ReadLine
							j = InStr(str1, ",")
							tid(0, i) = Left(str1, j - 1)
							tid(1, i) = Mid(str1, j + 1)
							i = i + 1
						Loop
						objFile.Close
						objFile = Nothing
						For j = 0 To i - 1
							'If (Left((tid(0).Replace(j).ToUpper(), " ",""),Len(title2))) = title2 Then
							'If (Left((tid(0, j)).Replace(" ", "").ToUpper(), ), Len(title2))) = title2 Then
							If (UCase(tid(0, j).Replace(" ", ""))) = title2 Then

								Console.Error.WriteLine("SCRename.tid から")
								title = tid(0, j)
								k = CInt(tid(1, j))
								Exit For
							End If
						Next
					End If
					If k < 0 Then
						With Activator.CreateInstance(Type.GetTypeFromProgID("ADODB.Stream"))
							.Open
							.Charset = "UTF-8"
							.WriteText(title)
							.Position = 0
							.Type = 1
							str1 = .Read
							.Close
						End With
						str2 = ""
						For i = 4 To System.Text.Encoding.Default.GetByteCount(str1)
							str2 = str2 + "%" + CStr(Hex(System.Text.Encoding.Default.GetBytes(System.Text.Encoding.Default.GetString(str1, i, 1))))
						Next
						For i = 0 To 2
							If i > 0 Then
								Threading.Thread.Sleep(1000)
							End If
							objHTTP.Open("Get", "http://cal.syoboi.jp/find?kw=" + str2, False)
							objHTTP.Send
							If objHTTP.Status >= 200 And objHTTP.Status < 300 Then
								Exit For
							End If
						Next
						If i > 2 Then
							Console.Error.WriteLine()
							Console.WriteLine(CmdArgs(0))
							Console.Error.WriteLine("しょぼいカレンダーにアクセスできませんでした。")
							Threading.Thread.Sleep(1000)
							Environment.Exit(1)
						End If
						str1 = objHTTP.responseText
						For i = 1 To Len(str1)
							str2 = Mid(str1, i, 1)
							If str2 = "\" Then
								str1 = Left(str1, i - 1) + "＼" + Mid(str1, i + 1)
							ElseIf str2 = "+" Then
								Do
									For j = 1 To elen - 1
										str2 = Mid(str1, i + j, 1)
										If str2 = "+" Then
											i = i + j
											Exit For
										ElseIf str2 = ";" Then
											str2 = Mid(str1, i + 1, j - 1)
											For k = 0 To UBound(char9)
												If str2 = char9(k) Then
													str1 = Left(str1, i - 1) + char10(k) + Mid(str1, i + j + 1)
													j = elen
													Exit For
												End If
											Next
										End If
									Next
								Loop While j < elen
							End If
						Next
						i = Len(str1)
						Do
							i = InStrRev(str1, "/tid/", i)
							If i > 0 Then
								j = i + 5
								Do While IsNumeric(Mid(str1, j, 1))
									j = j + 1
								Loop
								k = CInt(Mid(str1, i + 5, j - i - 5))
								j = j + 2
								l = InStr(j, str1, "</a>")
								If l > j Then
									str2 = Mid(str1, j, l - j)
									str2 = (str2).Replace("?", "？")
									str2 = (str2).Replace("!", "！")
									If (Left((str2).Replace(" ", "").ToUpper(), Len(title2))) = title2 Then
										Console.Error.WriteLine("しょぼいカレンダーから")
										title = str2
										Exit Do
									End If
								End If
							End If
						Loop While i > 0
						If i > 0 Then
							i = 0
							j = 0
							l = 0
							If objFSO.FileExists(path + "SCRename.tid") Then
								j = UBound(tid, 2) + 1
								For i = 0 To j - 1
									If CInt(tid(1, i)) = k Then
										j = j - 1
										l = -1
										Exit For
									ElseIf CInt(tid(1, i)) > k Then
										Exit For
									End If
								Next
							End If
							If l > -1 Then
								ReDim Preserve tid(1, j)
								For l = j To i + 1 Step -1
									tid(0, l) = tid(0, l - 1)
									tid(1, l) = tid(1, l - 1)
								Next
							End If
							tid(0, i) = title
							tid(1, i) = CStr(k)
							objFile = objFSO.OpenTextFile(path + "SCRename.tid", 2, True, -2)
							For i = 0 To j
								objFile.WriteLine(tid(0, i) + "," + tid(1, i))
							Next
							objFile.Close
							objFile = Nothing
						End If
					End If
					If i < 1 Then
						Console.Error.WriteLine("「" + title + "」の TID を取得できませんでした。" + Environment.NewLine)
					Else
						str1 = ""
						str2 = ""
						If serv > -1 Then
							If service(3, serv) <> "" Then
								str1 = "（" + service(1, serv) + "）"
								str2 = "+ChID=" + service(3, serv)
							End If
						End If
						Console.Error.WriteLine("「" + title + "」の TID（" + k + "）を取得しました。")
						Console.Error.WriteLine("第" + number + "話" + str1 + "の情報を検索します。" + Environment.NewLine)
						For i = 0 To 2
							If i > 0 Then
								Threading.Thread.Sleep(1000)
							End If
							objHTTP.Open("Get", "http://cal.syoboi.jp/db.php?Command=ProgLookup+TID=" + k + str2 + "+Count=" + number + "+Fields=StTime,EdTime,ChID,STSubTitle+JOIN=SubTitles", False)
							objHTTP.Send
							If objHTTP.Status >= 200 And objHTTP.Status < 300 Then
								Exit For
							End If
						Next
						If i > 2 Then
							Console.WriteLine(CmdArgs(0))
							Console.Error.WriteLine("しょぼいカレンダーにアクセスできませんでした。")
							Threading.Thread.Sleep(1000)
							Environment.Exit(1)
						End If
						str1 = objHTTP.responseText
						i = InStr(str1, "<StTime>")
						If i > 0 Then
							i = i + 8
							j = InStr(i, str1, "</StTime>")
							'str2 = (Mid(str1).Replace(i, j - i), "-", "/")
							str2 = Mid(str1, i, j - i).Replace("-", "/")
							If IsDate(str2) Then
								stdt = CDate(str2)
							End If
							i = InStr(j + 9, str1, "<EdTime>") + 8
							j = InStr(i, str1, "</EdTime>")
							'str2 = (Mid(str1).Replace(i, j - i), "-", "/")
							str2 = Mid(str1, i, j - i).Replace("-", "/")
							If IsDate(str2) Then
								eddt = CDate(str2)
							End If
							i = InStr(j + 9, str1, "<ChID>") + 6
							j = InStr(i, str1, "</ChID>")
							str2 = Mid(str1, i, j - i)
							If IsNumeric(str2) Then
								For i = 0 To UBound(service, 2)
									If service(3, i) = str2 Then
										serv = i
									End If
								Next
							End If
							i = InStr(j + 7, str1, "<STSubTitle>") + 12
							j = InStr(i, str1, "</STSubTitle>")
							subtitle = Mid(str1, i, j - i)
							For i = 1 To Len(subtitle) - 3
								If Mid(subtitle, i, 1) = "+" Then
									Do
										For j = 1 To elen - 1
											str2 = Mid(subtitle, i + j, 1)
											If str2 = "+" Then
												i = i + j
												Exit For
											ElseIf str2 = ";" Then
												str2 = Mid(subtitle, i + 1, j - 1)
												For k = 0 To UBound(char9)
													If str2 = char9(k) Then
														subtitle = Left(subtitle, i - 1) + char10(k) + Mid(subtitle, i + j + 1)
														j = elen
														Exit For
													End If
												Next
											End If
										Next
									Loop While (j < elen)
								End If
							Next
							number = "#" + number
							pos = 1
						End If
					End If
				End If
			End If
		End If
		objHTTP = Nothing
		If pos = 0 Then
			If (opt And 4) = 0 Then
				Console.WriteLine(CmdArgs(0))
			End If
			Console.Error.WriteLine("番組情報が見つかりませんでした。")
			If (opt And 4) = 0 Then
				Threading.Thread.Sleep(1000)
				Environment.Exit(1)
			Else
				'強制リネーム
				Console.Error.WriteLine("強制リネームを行います。" + Environment.NewLine)
				number = ""
				stdt = tgtdt
				eddt = tgtdt
				i = InStr(2, ftitle, sep)
				If i > 1 Then
					title = Left(ftitle, i - 1)
				Else
					title = ftitle
				End If
				l = Len(title)
				For i = 2 To l - 1
					str1 = Mid(title, i, 1)
					j = InStr("「『", str1)
					If j > 0 Then
						k = InStr(i + 1, title, Mid("」』", j, 1))
						If k > 0 Then
							Exit For
						End If
					End If
				Next
				If i < l Then
					If i < k - 1 Then
						subtitle = Mid(title, i + 1, k - i - 1)
					End If
					title = (Left(title, i - 1).TrimEnd())
				End If
				title2 = title
				i = 2
				Do While i < Len(title)
					str1 = Mid(title, i, 1)
					j = InStr(char4, str1)
					If j > 0 Then
						k = InStr(i + 1, title, Mid(char5, j, 1))
						If k > 0 Then
							str1 = Left(title, i - 1)
							If k < Len(title) Then
								str1 = str1 + " " + Mid(title, k + 1)
							End If
							title2 = str1
						End If
					End If
					i = i + 1
				Loop
				pos = 1
			End If
		End If
		Console.WriteLine("番組情報取得終了")

		'リネーム書式設定
		str1 = CmdArgs(1)
		If number <> "" Then
			k = Len(number)
			For i = 1 To k
				If Mid(number, i, 1) = "#" Then
					For j = i + 1 To k
						If Not IsNumeric(Mid(number, j, 1)) Then
							Exit For
						End If
					Next
					If j > i + 1 Then
						str2 = CStr(CInt(Mid(number, i + 1, j - i - 1)))
						l = Len(str2)
						Dim strtmp = number1
						number1 = strtmp + str2
						If l < 2 Then
							str2 = "0" + str2
						End If
						strtmp = number2
						number2 = strtmp + str2
						If l < 3 Then
							str2 = "0" + str2
						End If
						strtmp = number3
						number3 = strtmp + str2
						If l < 4 Then
							str2 = "0" + str2
						End If
						strtmp = number4
						number4 = strtmp + str2
					End If
					i = j - 1
				Else
					number2 = number2 + Mid(number, i, 1)
					number3 = number3 + Mid(number, i, 1)
					number4 = number4 + Mid(number, i, 1)
				End If
			Next
		End If
		str1 = (str1).Replace("$SCnumber1$", number1)
		str1 = (str1).Replace("$SCnumber$", number2)
		str1 = (str1).Replace("$SCnumber2$", number2)
		str1 = (str1).Replace("$SCnumber3$", number3)
		str1 = (str1).Replace("$SCnumber4$", number4)
		i = Hour(stdt)
		yr = Year(stdt)
		mon = Right("0" + Month(stdt), 2)
		dy = Right("0" + Day(stdt), 2)
		hr = Right("0" + i, 2)
		min = Right("0" + Minute(stdt), 2)
		sec = Right("0" + Second(stdt), 2)
		str1 = (str1).Replace("$SCdate$", Right(CStr(yr), 2) + mon + dy)
		str1 = (str1).Replace("$SCdate2$", yr + mon + dy)
		str1 = (str1).Replace("$SCyear$", Right(CStr(yr), 2))
		str1 = (str1).Replace("$SCyear2$", yr)
		str1 = (str1).Replace("$SCmonth$", mon)
		str1 = (str1).Replace("$SCday$", dy)
		str1 = (str1).Replace("$SCquarter$", DatePart("q", stdt))
		j = Weekday(stdt)
		str1 = (str1).Replace("$SCweek$", WeekdayName(j, True))
		str1 = (str1).Replace("$SCweek2$", char11(j - 1))
		str1 = (str1).Replace("$SCweek3$", (char11(j - 1).ToUpper()))
		str1 = (str1).Replace("$SCtime$", hr + min)
		str1 = (str1).Replace("$SCtime2$", hr + min + sec)
		str1 = (str1).Replace("$SChour$", hr)
		str1 = (str1).Replace("$SCminute$", min)
		str1 = (str1).Replace("$SCsecond$", sec)
		If i < 5 Then
			stdt = DateAdd("d", -1, stdt)
			i = i + 24
		End If
		yr = Year(stdt)
		mon = Right("0" + Month(stdt), 2)
		dy = Right("0" + Day(stdt), 2)
		hr = Right("0" + i, 2)
		str1 = (str1).Replace("$SCdates$", Right(CStr(yr), 2) + mon + dy)
		str1 = (str1).Replace("$SCdate2s$", yr + mon + dy)
		str1 = (str1).Replace("$SCyears$", Right(CStr(yr), 2))
		str1 = (str1).Replace("$SCyear2s$", yr)
		str1 = (str1).Replace("$SCmonths$", mon)
		str1 = (str1).Replace("$SCdays$", dy)
		str1 = (str1).Replace("$SCquarters$", DatePart("q", stdt))
		j = Weekday(stdt)
		str1 = (str1).Replace("$SCweeks$", WeekdayName(j, True))
		str1 = (str1).Replace("$SCweek2s$", char11(j - 1))
		str1 = (str1).Replace("$SCweek3s$", (char11(j - 1).ToUpper()))
		str1 = (str1).Replace("$SCtimes$", hr + min)
		str1 = (str1).Replace("$SCtime2s$", hr + min + sec)
		str1 = (str1).Replace("$SChours$", hr)
		i = Hour(eddt)
		yr = Year(eddt)
		mon = Right("0" + Month(eddt), 2)
		dy = Right("0" + Day(eddt), 2)
		hr = Right("0" + i, 2)
		min = Right("0" + Minute(eddt), 2)
		sec = Right("0" + Second(eddt), 2)
		str1 = (str1).Replace("$SCeddate$", Right(CStr(yr), 2) + mon + dy)
		str1 = (str1).Replace("$SCeddate2$", yr + mon + dy)
		str1 = (str1).Replace("$SCedyear$", Right(CStr(yr), 2))
		str1 = (str1).Replace("$SCedyear2$", yr)
		str1 = (str1).Replace("$SCedmonth$", mon)
		str1 = (str1).Replace("$SCedday$", dy)
		str1 = (str1).Replace("$SCedquarter$", DatePart("q", eddt))
		j = Weekday(eddt)
		str1 = (str1).Replace("$SCedweek$", WeekdayName(j, True))
		str1 = (str1).Replace("$SCedweek2$", char11(j - 1))
		str1 = (str1).Replace("$SCedweek3$", (char11(j - 1).ToUpper()))
		str1 = (str1).Replace("$SCedtime$", hr + min)
		str1 = (str1).Replace("$SCedtime2$", hr + min + sec)
		str1 = (str1).Replace("$SCedhour$", hr)
		str1 = (str1).Replace("$SCedminute$", min)
		str1 = (str1).Replace("$SCedsecond$", sec)
		If i < 5 Then
			eddt = DateAdd("d", -1, eddt)
			i = i + 24
		End If
		yr = Year(eddt)
		mon = Right("0" + Month(eddt), 2)
		dy = Right("0" + Day(eddt), 2)
		hr = Right("0" + i, 2)
		min = Right("0" + Minute(eddt), 2)
		sec = Right("0" + Second(eddt), 2)
		str1 = (str1).Replace("$SCeddates$", Right(CStr(yr), 2) + mon + dy)
		str1 = (str1).Replace("$SCeddate2s$", yr + mon + dy)
		str1 = (str1).Replace("$SCedyears$", Right(CStr(yr), 2))
		str1 = (str1).Replace("$SCedyear2s$", yr)
		str1 = (str1).Replace("$SCedmonths$", mon)
		str1 = (str1).Replace("$SCeddays$", dy)
		str1 = (str1).Replace("$SCedquarters$", DatePart("q", eddt))
		j = Weekday(eddt)
		str1 = (str1).Replace("$SCedweeks$", WeekdayName(j, True))
		str1 = (str1).Replace("$SCedweek2s$", char11(j - 1))
		str1 = (str1).Replace("$SCedweek3s$", (char11(j - 1).ToUpper()))
		str1 = (str1).Replace("$SCedtimes$", hr + min)
		str1 = (str1).Replace("$SCedtime2s$", hr + min + sec)
		str1 = (str1).Replace("$SCedhours$", hr)
		If serv < 0 Then
			str2 = ""
		Else
			str2 = service(2, serv)
		End If
		str1 = (str1).Replace("$SCservice$", str2)
		str1 = (str1).Replace("$SCpart$", part)
		str1 = (str1).Replace("$SCtitle$", title)
		str1 = (str1).Replace("$SCtitle2$", title2)
		str1 = (str1).Replace("$SCsubtitle$", subtitle)
		Console.WriteLine("リネーム書式設定終了")

		'リネーム中止
		If (opt And 2) > 0 And subtitle = "" Then
			Console.WriteLine(CmdArgs(0))
			Console.Error.WriteLine("サブタイトルを取得できなかったため処理を中止しました。")
			Threading.Thread.Sleep(1000)
			Environment.Exit(1)
		End If
		Console.WriteLine("リネーム中止終了")

		'SCRename.rp2 読み込み＆リネーム名置換
		str2 = path + "SCRename.rp2"
		If objFSO.FileExists(str2) Then
			objFile = objFSO.OpenTextFile(str2, 1, False, -2)
			Do While Not objFile.AtEndOfStream
				str2 = objFile.ReadLine
				If Left(str1, 1) <> ":" Then
					i = InStr(str2, ",")
					If i > 0 Then
						str1 = (str1).Replace(Left(str2, i - 1), Mid(str2, i + 1))
					End If
				End If
			Loop
			objFile.Close
			objFile = Nothing
		End If
		Console.WriteLine("SCRename.rp2 読み込み＆リネーム名置換終了")

		'使用不可文字置換
		str2 = ""
		If Mid(CmdArgs(1), 2, 1) = ":" And Mid(str1, 2, 1) = ":" Then
			str2 = Left(str1, 2)
			str1 = Mid(str1, 3)
		End If
		For i = 1 To Len(char6)
			str1 = (str1).Replace(Mid(char6, i, 1), Mid(char7, i, 1))
		Next
		str1 = str2 + str1
		Console.WriteLine("使用不可文字置換終了")

		'不要空白削除
		i = 2
		Do While i <= Len(str1)
			i = InStr(i, str1, "\")
			If i < 1 Then
				Exit Do
			End If
			For j = i - 1 To 1 Step -1
				If Mid(str1, j, 1) <> " " Then
					Exit For
				End If
			Next
			If j < i - 1 Then
				str1 = Left(str1, j) + Mid(str1, i)
			End If
			i = i + 1
		Loop
		If (opt And 8) = 0 Then
			str1 = (str1).Trim()
			i = 1
			Do While i <= Len(str1)
				str2 = Mid(str1, i, 1)
				If str2 = " " Or str2 = "　" Then
					For j = i + 1 To Len(str1)
						str2 = Mid(str1, j, 1)
						If str2 <> " " And str2 <> "　" Then
							Exit For
						End If
					Next
					str2 = Mid(str1, i - 1, 1)
					If str2 = ":" Or str2 = "\" Then
						i = i - 1
					End If
					str1 = Left(str1, i) + Mid(str1, j)
				End If
				i = i + 1
			Loop
		End If
		Console.WriteLine("不要空白削除終了")

		'フルパス生成
		i = 0
		If Left(str1, 2) = "\\" Then
			i = InStr(4, str1, "\")
		Else
			If Left(str1, 1) = "\" And Mid(CmdArgs(0), 2, 1) = ":" Then
				str1 = Left(CmdArgs(0), 2) + str1
				i = 3
			ElseIf Mid(str1, 2, 1) <> ":" Then
				str1 = rpath + str1
				i = Len(rpath)
			End If
		End If

		'256文字以上ファイルパス削除
		j = 255 - Len(ext)
		If Len(str1) > j Then
			Console.Error.WriteLine("ファイルパスが256文字以上のため切り詰めます。")
			str1 = Left(str1, j)
		End If

		'フォルダ作成
		If (opt And 1) = 0 Then
			Do
				i = InStr(i + 2, str1, "\")
				If i > 1 Then
					str2 = Left(str1, i - 1)
					If objFSO.FolderExists(str2) = False Then
						objFSO.CreateFolder(str2)
					End If
				End If
			Loop While i > 0
			If Err.Number <> 0 Then
				Console.WriteLine(CmdArgs(0))
				Console.Error.WriteLine("フォルダ " + str2 + " を作成できませんでした。")
				Threading.Thread.Sleep(1000)
				Environment.Exit(1)
			End If
		End If

		'移動およびリネーム
		str1 = str1 + ext
		If (opt And 1) = 0 Then
			If objFSO.FileExists(str1) Then
				Console.WriteLine(CmdArgs(0))
				Console.Error.WriteLine(str1 + " はすでに存在しています。")
				Threading.Thread.Sleep(1000)
				Environment.Exit(1)
			End If
			objFSO.MoveFile(CmdArgs(0), str1)
			If Err.Number <> 0 Then
				Console.WriteLine(CmdArgs(0))
				Console.Error.WriteLine("ファイル処理のエラーのため完了できませんでした。")
				Threading.Thread.Sleep(1000)
				Environment.Exit(1)
			End If
		End If
		Console.WriteLine(str1)
		Console.Error.WriteLine("処理が完了しました。")
		Environment.Exit(0)
	End Sub
End Module
