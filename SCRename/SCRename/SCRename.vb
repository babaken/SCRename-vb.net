Imports System
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text
Imports System.Xml

Module SCRename

	''' <summary>
	''' エラーメッセージ表示
	''' </summary>
	''' <param name="prm"></param>
	''' <param name="message"></param>
	Sub ErrGuidance(prm As String, message As String)
		If prm <> "" Then
			Console.WriteLine(prm)
		End If
		Console.Error.WriteLine(message)
		Threading.Thread.Sleep(1000)
	End Sub

	''' <summary>
	''' フォーマットに値をセット
	''' </summary>
	''' <param name="prm"></param>
	''' <returns></returns>
	Function SetFmt(prm() As String) As String
		Dim fmt = prm(0)
		Dim StartDate = DateTime.Parse(prm(1))
		Dim EndDate = DateTime.Parse(prm(2))
		Dim number1 = prm(3)
		Dim number2 = prm(4)
		Dim number3 = prm(5)
		Dim number4 = prm(6)
		Dim part = prm(7)
		Dim title = prm(8)
		Dim title2 = prm(9)
		Dim subtitle = prm(10)
		Dim chan = prm(11)
		Dim StartDateDash = DateTime.Parse(prm(12))
		Dim StartHH = Integer.Parse(prm(13))
		Dim EndDateDash = DateTime.Parse(prm(14))
		Dim EndHH = Integer.Parse(prm(15))
		Dim char11() As String = prm(16).Split(" ")
		Dim tmpfmt As String

		'StartDate
		tmpfmt = (fmt).Replace("$SCnumber1$", number1)
		fmt = (tmpfmt).Replace("$SCnumber$", number2)
		tmpfmt = (fmt).Replace("$SCnumber2$", number2)
		fmt = (tmpfmt).Replace("$SCnumber3$", number3)
		tmpfmt = (fmt).Replace("$SCnumber4$", number4)
		fmt = (tmpfmt).Replace("$SCdate$", Right(CStr(Year(StartDate)), 2) + Right("0" + Month(StartDate), 2) + Right("0" + Day(StartDate), 2))
		tmpfmt = (fmt).Replace("$SCdate2$", Year(StartDate) + Right("0" + Month(StartDate), 2) + Right("0" + Day(StartDate), 2))
		fmt = (tmpfmt).Replace("$SCyear$", Right(CStr(Year(StartDate)), 2))
		tmpfmt = (fmt).Replace("$SCyear2$", Year(StartDate))
		fmt = (tmpfmt).Replace("$SCmonth$", Right("0" + Month(StartDate), 2))
		tmpfmt = (fmt).Replace("$SCday$", Right("0" + Day(StartDate), 2))
		fmt = (tmpfmt).Replace("$SCquarter$", DatePart("q", StartDate))
		tmpfmt = (fmt).Replace("$SCweek$", WeekdayName(Weekday(StartDate), True))
		fmt = (tmpfmt).Replace("$SCweek2$", char11(Weekday(StartDate) - 1))
		tmpfmt = (fmt).Replace("$SCweek3$", (char11(Weekday(StartDate) - 1).ToUpper()))
		fmt = (tmpfmt).Replace("$SCtime$", Right("0" + Hour(StartDate), 2) + Right("0" + Minute(StartDate), 2))
		tmpfmt = (fmt).Replace("$SCtime2$", Right("0" + Hour(StartDate), 2) + Right("0" + Minute(StartDate), 2) + Right("0" + Second(StartDate), 2))
		fmt = (tmpfmt).Replace("$SChour$", Right("0" + Hour(StartDate), 2))
		tmpfmt = (fmt).Replace("$SCminute$", Right("0" + Minute(StartDate), 2))
		fmt = (tmpfmt).Replace("$SCsecond$", Right("0" + Second(StartDate), 2))

		StartDate = StartDateDash
		tmpfmt = (fmt).Replace("$SCdates$", Right(CStr(Year(StartDate)), 2) + Right("0" + Month(StartDate), 2) + Right("0" + Day(StartDate), 2))
		fmt = (tmpfmt).Replace("$SCdate2s$", Year(StartDate) + Right("0" + Month(StartDate), 2) + Right("0" + Day(StartDate), 2))
		tmpfmt = (fmt).Replace("$SCyears$", Right(CStr(Year(StartDate)), 2))
		fmt = (tmpfmt).Replace("$SCyear2s$", Year(StartDate))
		tmpfmt = (fmt).Replace("$SCmonths$", Right("0" + Month(StartDate), 2))
		fmt = (tmpfmt).Replace("$SCdays$", Right("0" + Day(StartDate), 2))
		tmpfmt = (fmt).Replace("$SCquarters$", DatePart("q", StartDate))
		fmt = (tmpfmt).Replace("$SCweeks$", WeekdayName(Weekday(StartDate), True))
		tmpfmt = (fmt).Replace("$SCweek2s$", char11(Weekday(StartDate) - 1))
		fmt = (tmpfmt).Replace("$SCweek3s$", (char11(Weekday(StartDate) - 1).ToUpper()))
		tmpfmt = (fmt).Replace("$SCtimes$", Right("0" + StartHH, 2) + Right("0" + Minute(StartDate), 2))
		fmt = (tmpfmt).Replace("$SCtime2s$", Right("0" + StartHH, 2) + Right("0" + Minute(StartDate), 2) + Right("0" + Second(StartDate), 2))
		tmpfmt = (fmt).Replace("$SChours$", Right("0" + StartHH, 2))

		'EndDate
		fmt = (tmpfmt).Replace("$SCeddate$", Right(CStr(Year(EndDate)), 2) + Right("0" + Month(EndDate), 2) + Right("0" + Day(EndDate), 2))
		tmpfmt = (fmt).Replace("$SCeddate2$", Year(EndDate) + Right("0" + Month(EndDate), 2) + Right("0" + Day(EndDate), 2))
		fmt = (tmpfmt).Replace("$SCedyear$", Right(CStr(Year(EndDate)), 2))
		tmpfmt = (fmt).Replace("$SCedyear2$", Year(EndDate))
		fmt = (tmpfmt).Replace("$SCedmonth$", Right("0" + Month(EndDate), 2))
		tmpfmt = (fmt).Replace("$SCedday$", Right("0" + Day(EndDate), 2))
		fmt = (tmpfmt).Replace("$SCedquarter$", DatePart("q", EndDate))
		tmpfmt = (fmt).Replace("$SCedweek$", WeekdayName(Weekday(EndDate), True))
		fmt = (tmpfmt).Replace("$SCedweek2$", char11(Weekday(EndDate) - 1))
		tmpfmt = (fmt).Replace("$SCedweek3$", (char11(Weekday(EndDate) - 1).ToUpper()))
		fmt = (tmpfmt).Replace("$SCedtime$", Right("0" + Hour(EndDate), 2) + Right("0" + Minute(EndDate), 2))
		tmpfmt = (fmt).Replace("$SCedtime2$", Right("0" + Hour(EndDate), 2) + Right("0" + Minute(EndDate), 2) + Right("0" + Second(EndDate), 2))
		fmt = (tmpfmt).Replace("$SCedhour$", Right("0" + Hour(EndDate), 2))
		tmpfmt = (fmt).Replace("$SCedminute$", Right("0" + Minute(EndDate), 2))
		fmt = (tmpfmt).Replace("$SCedsecond$", Right("0" + Second(EndDate), 2))

		EndDate = EndDateDash
		tmpfmt = (fmt).Replace("$SCeddates$", Right(CStr(Year(EndDate)), 2) + Right("0" + Month(EndDate), 2) + Right("0" + Day(EndDate), 2))
		fmt = (tmpfmt).Replace("$SCeddate2s$", Year(EndDate) + Right("0" + Month(EndDate), 2) + Right("0" + Day(EndDate), 2))
		tmpfmt = (fmt).Replace("$SCedyears$", Right(CStr(Year(EndDate)), 2))
		fmt = (tmpfmt).Replace("$SCedyear2s$", Year(EndDate))
		tmpfmt = (fmt).Replace("$SCedmonths$", Right("0" + Month(EndDate), 2))
		fmt = (tmpfmt).Replace("$SCeddays$", Right("0" + Day(EndDate), 2))
		tmpfmt = (fmt).Replace("$SCedquarters$", DatePart("q", EndDate))
		fmt = (tmpfmt).Replace("$SCedweeks$", WeekdayName(Weekday(EndDate), True))
		tmpfmt = (fmt).Replace("$SCedweek2s$", char11(Weekday(EndDate) - 1))
		fmt = (tmpfmt).Replace("$SCedweek3s$", (char11(Weekday(EndDate) - 1).ToUpper()))
		tmpfmt = (fmt).Replace("$SCedtimes$", Right("0" + EndHH, 2) + Right("0" + Minute(EndDate), 2))
		fmt = (tmpfmt).Replace("$SCedtime2s$", Right("0" + EndHH, 2) + Right("0" + Minute(EndDate), 2) + Right("0" + Second(EndDate), 2))
		tmpfmt = (fmt).Replace("$SCedhours$", Right("0" + EndHH, 2))
		fmt = (tmpfmt).Replace("$SCservice$", chan)
		tmpfmt = (fmt).Replace("$SCpart$", part)
		fmt = (tmpfmt).Replace("$SCtitle$", title)
		tmpfmt = (fmt).Replace("$SCtitle2$", title2)
		fmt = (tmpfmt).Replace("$SCsubtitle$", subtitle)

		Return fmt
	End Function

	Sub Main(ByVal CmdArgs() As String)
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
		Dim i As Object
		Dim j As Object
		Dim k As Integer = 0
		Dim l As Object
		Dim argc As Integer
		Dim opt As Integer = 0
		Dim elen As Object
		Dim days As Object
		Dim pos As Integer = 0
		Dim serv As Integer = 0
		Dim str1 As Object
		Dim str2 As Object
		Dim rpath As String = ""
		Dim ext As String = ""
		Dim title As Object
		Dim title2 As String = ""
		Dim ftitle As Object
		Dim Number As String = ""
		Dim number1 As String = ""
		Dim number2 As String = ""
		Dim number3 As String = ""
		Dim number4 As String = ""
		Dim part As String = ""
		Dim subtitle As String = ""
		Dim yr As Object
		Dim dt1 As Date?
		Dim dt2 As Date?
		Dim tgtdt As Date?
		Dim StartDate As Date?
		Dim EndDate As Date?
		Dim dtflag As Integer = 0
		Dim service(,) As String
		Dim tid(,) As String
		Dim objFile As Object, objHTTP As Object
		Dim str8 As String = "I II III IV V VI VII VIII IX X"
		Dim str9 As String = "quot amp #039 lt gt"
		Dim str10 As String = ChrW(&H201C) & " & " & Chr(39) & " ＜ ＞"
		Dim str11 As String = "Sun Mon Tue Wed Thu Fri Sat"
		Dim char8() As String = str8.Split(" ")
		Dim char9() As String = str9.Split(" ")
		Dim char10() As String = str10.Split(" ")
		Dim char11() As String = str11.Split(" ")
		Dim xmlFilePath As String = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SCRename.xml")
		Dim objFSO As Object = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.FileSystemObject"))
		Dim prm As String
		Dim msg As String
		Dim counter As Integer
		Dim ilinecount As Integer = 0
		Dim path As Object = Left(System.Reflection.Assembly.GetExecutingAssembly().Location, InStrRev(System.Reflection.Assembly.GetExecutingAssembly().Location, "\"))
		Dim excpath As String = path + "SCRename.exc"
		Dim srvpath As String = path + "SCRename.srv"
		Dim rp1path As String = path + "SCRename.rp1"
		Dim tidpath As String = path + "SCRename.tid"
		Dim rp2path As String = path + "SCRename.rp2"
		Dim CreateDate As Object
		Dim ModifieDate As Object
		Dim outputfile As String = ""
		Dim fmt As String = ""

		'引数処理
		For Each prm In CmdArgs
			prm = prm.Replace("""", "")
			If (prm).ToLower() = "-h" Or prm = "-?" Then
				Console.WriteLine(Environment.NewLine & "SCRename.vbs [オプション] ファイル リネーム書式")
				Console.WriteLine(" [タイトル開始位置] [検索文字数]" & Environment.NewLine)
				Environment.Exit(1)
			ElseIf (prm).ToLower() = "-t" Then
				opt = opt Or 1
			ElseIf (prm).ToLower() = "-n" Then
				opt = opt Or 2
			ElseIf (prm).ToLower() = "-f" Then
				opt = opt Or 4
			ElseIf (prm).ToLower() = "-s" Then
				opt = opt Or 8
			ElseIf (prm).ToLower() = "-a" Then
				opt = opt Or 16
			ElseIf (prm).ToLower() = "-a1" Then
				opt = opt Or 32
			Else
				ReDim Preserve CmdArgs(argc)
				CmdArgs(argc) = prm
				argc = argc + 1
			End If
		Next

		'起動時処理
		Threading.Thread.Sleep(1000)
		Console.Error.WriteLine(Environment.NewLine + "SCRename 動作中..." + Environment.NewLine)
		If argc < 2 Then
			msg = "パラメータが足りません。"
			ErrGuidance(CmdArgs(0), msg)
			Environment.Exit(1)
		ElseIf CmdArgs(0) = "" Then
			msg = "処理対象のファイルが指定されていません。"
			ErrGuidance("", msg)
			Environment.Exit(1)
		ElseIf CmdArgs(1) = "" Then
			msg = "リネーム書式が指定されていません。"
			ErrGuidance(CmdArgs(0), msg)
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
		For counter = 0 To UBound(char9)
			j = Len(char9(counter))
			If j > elen Then
				elen = j
			End If
		Next
		elen = elen + 2
		Console.WriteLine("実体名最大文字数取得終了")

		'SCRename.exc 読み込み
		If objFSO.FileExists(excpath) Then
			objFile = objFSO.OpenTextFile(excpath, 1, False, -2)
			Do While Not objFile.AtEndOfStream
				excpath = objFile.ReadLine
				If Left(excpath, 1) <> ":" Then
					If InStr((CmdArgs(0).ToUpper()), (excpath).ToUpper()) > 0 Then
						counter = -1
					End If
				End If
			Loop
			objFile.Close
			objFile = Nothing
			If counter < 0 Then
				msg = "対象外のファイルのため処理しませんでした。"
				ErrGuidance(CmdArgs(0), msg)
				Environment.Exit(1)
			End If
		End If
		Console.WriteLine("SCRename.exc 読み込み終了")

		'リネーム元ファイル存在確認
		If (opt And 1) = 0 Then
			If Not objFSO.FileExists(CmdArgs(0)) Then
				msg = CmdArgs(0) + " がありません。"
				ErrGuidance("", msg)
				Environment.Exit(1)
			End If
		End If
		Console.WriteLine("リネーム元ファイル存在確認終了")

		'SCRename.srv 読み込み
		objFile = objFSO.OpenTextFile(srvpath, 1, False, -2)
		If Err.Number <> 0 Then
			msg = srvpath + " がありません。"
			ErrGuidance(CmdArgs(0), msg)
			Environment.Exit(1)
		End If
		i = 0
		Do While Not objFile.AtEndOfStream
			srvpath = objFile.ReadLine
			If Left(srvpath, 1) <> ":" Then
				ReDim Preserve service(3, ilinecount)
				ilinecount = ilinecount + 1
				j = 0
				k = 0
				Do While j < 4
					l = InStr(k + 1, srvpath, ",")
					If l > 0 Then
						service(j, i) = Mid(srvpath, k + 1, l - k - 1)
					Else
						service(j, i) = Mid(srvpath, k + 1)
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
			CreateDate = objFSO.GetFile(CmdArgs(0)).DateCreated
			ModifieDate = objFSO.GetFile(CmdArgs(0)).DateLastModified
			If CreateDate < ModifieDate Then
				ModifieDate = CreateDate
				dtflag = 1
			End If
			If tgtdt Is Nothing Then
				tgtdt = ModifieDate
				days = 7
			Else
				tgtdt = tgtdt + New TimeSpan(Hour(ModifieDate), Minute(ModifieDate), Second(ModifieDate))
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
		If objFSO.FileExists(rp1path) Then
			objFile = objFSO.OpenTextFile(rp1path, 1, False, -2)
			Do While Not objFile.AtEndOfStream
				rp1path = objFile.ReadLine
				If Left(rp1path, 1) <> ":" Then
					i = InStr(rp1path, ",")
					If i > 0 Then
						title = (title).Replace(Left(rp1path, i - 1), Mid(rp1path, i + 1))
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
				msg = "タイトルを取得出来ませんでした。"
				ErrGuidance(CmdArgs(0), msg)
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
			msg = "放送局が不明のためすべての放送局を対象にします。"
			ErrGuidance("", msg)
		End If
		Console.WriteLine("放送局名取得終了")

		'検索開始
		If (opt And 32) = 0 Then
			msg = ""
			If days = 7 Then
				msg = "ファイル名から日付が取得できないためファイルの"
				If dtflag = 1 Then
					msg = msg + "作成"
				Else
					msg = msg + "更新"
				End If
				msg = msg + "日から一週間遡って、" + Environment.NewLine
			End If
			If dtflag = 1 Then
				msg = msg + "開始"
			Else
				msg = msg + "終了"
			End If
			msg = msg + "日時が " + tgtdt + " に最も近い"
			ErrGuidance("", msg)
			If serv < 0 Then
				str1 = ""
			Else
				str1 = "（" + service(1, serv) + "）"
			End If
			msg = "「" + title + "」" + str1 + "を検索します。" + Environment.NewLine
			ErrGuidance("", msg)
			If Err.Number <> 0 Then
				msg = "検索前処理でエラーが発生しました。"
				ErrGuidance(CmdArgs(0), msg)
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
			msg = "XMLHTTP オブジェクトを作成できませんでした。"
			ErrGuidance(CmdArgs(0), msg)
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
				msg = "しょぼいカレンダーにアクセスできませんでした。"
				ErrGuidance(CmdArgs(0), msg)
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
						If StartDate Is Nothing Then
							StartDate = dt1
							EndDate = dt2
							pos = i
						End If
						If Math.Abs(CType(CType((tgtdt - dt1), TimeSpan).Days, SByte)) < Math.Abs(CType(CType((tgtdt - StartDate), TimeSpan).Days, SByte)) Then
							StartDate = dt1
							EndDate = dt2
							pos = i
						End If

					Else
						If EndDate Is Nothing Then
							StartDate = dt1
							EndDate = dt2
							pos = i
						End If
						If Math.Abs(CType(CType((tgtdt - dt2), TimeSpan).Days, SByte)) < Math.Abs(CType(CType((tgtdt - EndDate), TimeSpan).Days, SByte)) Then
							StartDate = dt1
							EndDate = dt2
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
						Number = Left(str2, j - 1)
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
							Number = Number + "," + Mid(str2, i)
							Exit Do
						Else
							Number = Number + "," + Mid(str2, i, j - i)
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
					Number = Mid(Number, 2)
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
				msg = ""
				If (opt And 32) = 0 Then
					msg = "番組情報が見つかりませんでした。"
				End If
				msg = msg + "話数検索を行います。" + Environment.NewLine
				ErrGuidance(CmdArgs(0), msg)
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
					Number = CStr(k)
					For j = 2 To i - 1
						If InStr(sep + char4 + char5 + "～", Mid(ftitle, j, 1)) > 0 Then
							Exit For
						End If
					Next
					title = (Left(ftitle, j - 1).TrimEnd())
					title2 = ((title).Replace(" ", "").ToUpper())
					k = -1
					If objFSO.FileExists(tidpath) Then
						objFile = objFSO.OpenTextFile(tidpath, 1, False, -2)
						i = 0
						Do While Not objFile.AtEndOfStream
							ReDim Preserve tid(1, i)
							tidpath = objFile.ReadLine
							j = InStr(tidpath, ",")
							tid(0, i) = Left(tidpath, j - 1)
							tid(1, i) = Mid(tidpath, j + 1)
							i = i + 1
						Loop
						objFile.Close
						objFile = Nothing
						For j = 0 To i - 1
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
							msg = "しょぼいカレンダーにアクセスできませんでした。"
							ErrGuidance(CmdArgs(0), msg)
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
							If objFSO.FileExists(tidpath) Then
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
							objFile = objFSO.OpenTextFile(tidpath, 2, True, -2)
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
						msg = "「" + title + "」の TID（" + k + "）を取得しました。"
						msg = msg + "第" + Number + "話" + str1 + "の情報を検索します。" + Environment.NewLine
						ErrGuidance("", msg)
						For i = 0 To 2
							If i > 0 Then
								Threading.Thread.Sleep(1000)
							End If
							objHTTP.Open("Get", "http://cal.syoboi.jp/db.php?Command=ProgLookup+TID=" + k + str2 + "+Count=" + Number + "+Fields=StTime,EdTime,ChID,STSubTitle+JOIN=SubTitles", False)
							objHTTP.Send
							If objHTTP.Status >= 200 And objHTTP.Status < 300 Then
								Exit For
							End If
						Next
						If i > 2 Then
							msg = "しょぼいカレンダーにアクセスできませんでした。"
							ErrGuidance(CmdArgs(0), msg)
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
								StartDate = CDate(str2)
							End If
							i = InStr(j + 9, str1, "<EdTime>") + 8
							j = InStr(i, str1, "</EdTime>")
							'str2 = (Mid(str1).Replace(i, j - i), "-", "/")
							str2 = Mid(str1, i, j - i).Replace("-", "/")
							If IsDate(str2) Then
								EndDate = CDate(str2)
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
							Number = "#" + Number
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
			msg = "番組情報が見つかりませんでした。"
			ErrGuidance("", msg)
			If (opt And 4) = 0 Then
				Threading.Thread.Sleep(1000)
				Environment.Exit(1)
			Else
				'強制リネーム
				msg = "強制リネームを行います。" + Environment.NewLine
				ErrGuidance("", msg)
				Number = ""
				StartDate = tgtdt
				EndDate = tgtdt
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
		fmt = CmdArgs(1)
		If Number <> "" Then
			k = Len(Number)
			For i = 1 To k
				If Mid(Number, i, 1) = "#" Then
					For j = i + 1 To k
						If Not IsNumeric(Mid(Number, j, 1)) Then
							Exit For
						End If
					Next
					If j > i + 1 Then
						str2 = CStr(CInt(Mid(Number, i + 1, j - i - 1)))
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
					number2 = number2 + Mid(Number, i, 1)
					number3 = number3 + Mid(Number, i, 1)
					number4 = number4 + Mid(Number, i, 1)
				End If
			Next
		End If


		Dim chan As String
		If serv < 0 Then
			chan = ""
		Else
			chan = service(2, serv)
		End If

		Dim param() As String = {}
		Array.Resize(param, 17)
		param(0) = fmt
		param(1) = StartDate.ToString
		param(2) = EndDate.ToString
		param(3) = number1
		param(4) = number2
		param(5) = number3
		param(6) = number4
		param(7) = part
		param(8) = title.ToString
		param(9) = title2
		param(10) = subtitle
		param(11) = chan
		Dim StartHH = Hour(StartDate)
		If StartHH < 5 Then
			param(12) = DateAdd("d", -1, StartDate).ToString
			StartHH = StartHH + 24
		Else
			param(12) = StartDate.ToString
		End If
		param(13) = StartHH.ToString
		Dim EndHH = Hour(EndDate)
		If EndHH < 5 Then
			param(14) = DateAdd("d", -1, EndDate).ToString
			EndHH = EndHH + 24
		Else
			param(14) = EndDate.ToString
		End If
		param(15) = EndHH.ToString
		param(16) = str11
		fmt = SetFmt(param)

		Console.WriteLine("リネーム書式設定終了")

		'リネーム中止
		If (opt And 2) > 0 And subtitle = "" Then
			msg = "サブタイトルを取得できなかったため処理を中止しました。"
			ErrGuidance(CmdArgs(0), msg)
			Environment.Exit(1)
		End If
		Console.WriteLine("リネーム中止終了")

		'SCRename.rp2 読み込み＆リネーム名置換
		If objFSO.FileExists(rp2path) Then
			objFile = objFSO.OpenTextFile(rp2path, 1, False, -2)
			Do While Not objFile.AtEndOfStream
				rp2path = objFile.ReadLine
				If Left(fmt, 1) <> ":" Then
					i = InStr(rp2path, ",")
					If i > 0 Then
						fmt = (fmt).Replace(Left(rp2path, i - 1), Mid(rp2path, i + 1))
					End If
				End If
			Loop
			objFile.Close
			objFile = Nothing
		End If
		Console.WriteLine("SCRename.rp2 読み込み＆リネーム名置換終了")

		'使用不可文字置換
		str2 = ""
		If Mid(CmdArgs(1), 2, 1) = ":" And Mid(fmt, 2, 1) = ":" Then
			str2 = Left(fmt, 2)
			fmt = Mid(fmt, 3)
		End If
		For i = 1 To Len(char6)
			fmt = (fmt).Replace(Mid(char6, i, 1), Mid(char7, i, 1))
		Next
		fmt = str2 + fmt
		Console.WriteLine("使用不可文字置換終了")

		'不要空白削除
		i = 2
		Do While i <= Len(fmt)
			i = InStr(i, fmt, "\")
			If i < 1 Then
				Exit Do
			End If
			For j = i - 1 To 1 Step -1
				If Mid(fmt, j, 1) <> " " Then
					Exit For
				End If
			Next
			If j < i - 1 Then
				fmt = Left(fmt, j) + Mid(fmt, i)
			End If
			i = i + 1
		Loop
		If (opt And 8) = 0 Then
			fmt = (fmt).Trim()
			i = 1
			Do While i <= Len(fmt)
				str2 = Mid(fmt, i, 1)
				If str2 = " " Or str2 = "　" Then
					For j = i + 1 To Len(fmt)
						str2 = Mid(str1, j, 1)
						If str2 <> " " And str2 <> "　" Then
							Exit For
						End If
					Next
					str2 = Mid(fmt, i - 1, 1)
					If str2 = ":" Or str2 = "\" Then
						i = i - 1
					End If
					fmt = Left(fmt, i) + Mid(fmt, j)
				End If
				i = i + 1
			Loop
		End If
		Console.WriteLine("不要空白削除終了")

		'フルパス生成
		i = 0
		If Left(fmt, 2) = "\\" Then
			i = InStr(4, fmt, "\")
		Else
			If Left(fmt, 1) = "\" And Mid(CmdArgs(0), 2, 1) = ":" Then
				fmt = Left(CmdArgs(0), 2) + fmt
				i = 3
			ElseIf Mid(fmt, 2, 1) <> ":" Then
				outputfile = rpath + fmt
				i = Len(rpath)
			End If
		End If

		'256文字以上ファイルパス削除
		j = 255 - Len(ext)
		If Len(outputfile) > j Then
			msg = "ファイルパスが256文字以上のため切り詰めます。"
			ErrGuidance("", msg)
			outputfile = Left(outputfile, j)
		End If

		'フォルダ作成
		If (opt And 1) = 0 Then
			Do
				i = InStr(i + 2, outputfile, "\")
				If i > 1 Then
					str2 = Left(outputfile, i - 1)
					If objFSO.FolderExists(str2) = False Then
						objFSO.CreateFolder(str2)
					End If
				End If
			Loop While i > 0
			If Err.Number <> 0 Then
				msg = "フォルダ " + str2 + " を作成できませんでした。"
				ErrGuidance(CmdArgs(0), msg)
				Environment.Exit(1)
			End If
		End If

		'移動およびリネーム
		outputfile = outputfile + ext
		If (opt And 1) = 0 Then
			If objFSO.FileExists(outputfile) Then
				Console.WriteLine(CmdArgs(0))
				msg = outputfile + " はすでに存在しています。"
				ErrGuidance(CmdArgs(0), msg)
				Environment.Exit(1)
			End If
			objFSO.MoveFile(CmdArgs(0), outputfile)
			If Err.Number <> 0 Then
				msg = "ファイル処理のエラーのため完了できませんでした。"
				ErrGuidance(CmdArgs(0), msg)
				Environment.Exit(1)
			End If
		End If
		msg = "処理が完了しました。"
		ErrGuidance(outputfile, msg)
		Environment.Exit(0)
	End Sub
End Module
