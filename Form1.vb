Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Xml
Imports ICSharpCode.SharpZipLib.Zip
Imports System.Threading

Public Class Form1

    Private hs2Path As String = Path.GetDirectoryName(Path.GetDirectoryName(Application.ExecutablePath)) ' HS2パス
    Private modsPath As String = Path.Combine(hs2Path, "mods\") ' modsパス
    'Private hs2Path As String = "N:\illusion\HoneySelect2" ' HS2パスデバッグ
    'Private modsPath As String = "N:\illusion\HoneySelect2\mods\" ' modsパスデバッグ
    Private dbPath As String = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "modlist.db") ' DBパス
    Private dbBuPath As String = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "InMemoryDB_backup.db") ' DBバックアップパス
    Private cts As CancellationTokenSource = Nothing ' スレッドキャンセレーショントークン
    Private scanFullMode As Boolean = False 'スキャンモードフラグ
    Private sw As New Stopwatch() ' デバッグ用ストップウオッチ


    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        sw.Start()

        If Not LoadLanguageData() Then ' リソース化した外部変数読み込み
            Environment.Exit(0)
            Exit Sub
        Else
            logAdd("[Done] Language.xml has been loaded.")
            logAdd($"----------------------------------------------------------------------------------")
        End If

        WindowInitialize() ' ウィンドウ位置初期化読み込み
        InitializeEmptyDataTable() ' ダミーテーブル初期化

        Me.Refresh() ' 一度画面を更新して、すべてのコントロールが描画されることを保証
        picBox.AllowDrop = True ' ドロップ許可

        ' 起動チェック ハニセレexeファイルとmodsフォルダ
        If Not File.Exists(Path.Combine(hs2Path, "HoneySelect2.exe")) Then
            If Not File.Exists(Path.Combine(hs2Path, "AI-Syoujyo.exe")) Then
                MessageBox.Show($"HoneySelect2.exe is not found.{Environment.NewLine}Copy the HS2FND folder into the HoneySelect 2 folder and run it.")
                Environment.Exit(0)
                Exit Sub
            End If
        ElseIf Not Directory.Exists(modsPath) Then
            MessageBox.Show($"Mods folder {modsPath} is not found.")
            Environment.Exit(0)
            Exit Sub
        End If

        ' SQL dbファイルが存在しなければ新規作成
        If Not File.Exists(dbPath) Then
            InitialSQL()
            logAdd("[Now] SQL : Create a new database file...")
            logAdd($"----------------------------------------------------------------------------------")
            scanFullMode = True
        End If

        ' SQL dbをメモリにコピー
        DbLoadToMemory()

        ' イニシャルルーチン起動
        cts = New CancellationTokenSource()
        Dim longProcessThread As New Thread(Sub() InitialProcess(cts.Token, scanFullMode))
        longProcessThread.Start()
    End Sub

    ' イニシャルルーチン--------------------------------------------------------------
    Private Async Sub InitialProcess(token As CancellationToken, scanFullMode As Boolean)
        ' 画面サイズ変更を受け付けるために、マルチスレッドで実行する
        Control.CheckForIllegalCrossThreadCalls = False

        Dim fileCheckResult As (deleteList As String(), updateList As String(), existList As String()) = (New String() {}, New String() {}, New String() {})
        If Not scanFullMode Then
            ' DBをチェックして存在しないファイルと更新ファイル、既存ファイルのリストを作成
            logAdd("[Now] SQL :Checking the database...")
            fileCheckResult = Await GetDeleteAndUpdateRecordsAsync()
            If fileCheckResult.deleteList.Count = 0 And fileCheckResult.updateList.Count = 0 Then
                logAdd($"[Done] SQL : All files in DB description exist.")
            Else
                If fileCheckResult.updateList.Count > 0 Then
                    logAdd($"[Done] SQL : {fileCheckResult.updateList.Count} New file(s) found.")
                End If
                If fileCheckResult.deleteList.Count > 0 Then
                    ' 不要なレコードは削除
                    logAdd($"[Done] SQL : {fileCheckResult.deleteList.Count} Deleted or moved file(s) detected in database.")
                    Dim errMes As String = Await SqlDeleteRecordsAsync(fileCheckResult.deleteList)
                    If Not fileCheckResult.deleteList.Count = 0 Then
                        logAdd($"[Done] SQL : {fileCheckResult.deleteList.Count} Records deleted.")
                    Else
                        logAdd($"[Error] SQL : {errMes}")
                    End If
                End If
            End If
        End If

        ' modファイルリスト一覧取得
        logAdd("[Now] Listing all .zipmod files...")
        Dim zipModList As List(Of String) = Await GetZipModFilesAsync(modsPath)
        If zipModList.Count = 0 Then
            logAdd("[Error] No .zipmod files found in the specified directory.")
            Exit Sub
        End If
        logAdd($"[Done] {zipModList.Count} .zipmod files found.")

        ' 初回起動またはリスキャンモードでなければDBに残ったリストを取得
        Dim validModList As New List(Of String())()
        If Not scanFullMode Then
            logAdd("[Now] SQL : Getting mod list from database...")
            validModList = Await GetValidModListAsync()
            logAdd($"[Done] SQL: {validModList.Count} Items list retrieved from the database.")
        End If

        ' mod個別アイテムリストを生成
        logAdd($"[Now] Working on an exhaustive mod items list...")
        Dim fullModList As List(Of String()) = Await GetModListAsync(zipModList, validModList, token)

        If token.IsCancellationRequested Then ' キャンセルされた場合の処理
            logAdd("[Notice] Operation is cancelled.")
            Exit Sub
        End If

        ' 書き込み対象があればDB書き込み実行
        If fullModList.Count = 0 Then
            logAdd("[Done] SQL : Database is already up to date.")
        Else
            logAdd($"[Done] {fullModList.Count} New mod items listed.")
            If scanFullMode Then
                logAdd("[Now] SQL : Cleaning the database.")
                Await ClearModsTableAsync() ' Rescanモードなら一旦DB初期化
            End If
            logAdd($"[Now] SQL : Writing {fullModList.Count} records in the database...")
            Dim errMes = Await SqlBulkInsertAsync(fullModList, 500)
            If Not String.IsNullOrEmpty(errMes) Then
                logAdd(errMes)
            Else
                logAdd("[Done] SQL : Database up-to-date.")
            End If
        End If

        ' データベースのバックアップ
        Await BackupSqlDbAsync()

        ' DBから全データ取得して検索スタンバイ
        logAdd("[Now] SQL : Processing query...")
        Dim dT As DataTable
        dT = Await GetDataFromDbAsync()
        Await ShowResultAsync(dT)
        logAdd("[Done] SQL : Get all data from database.")
        Await SearchGuiResestAsync() ' 表示リセット
        Await Task.Run(Function() IsHs2RunningAsync()) ' HS2起動チェック
        cts = Nothing
    End Sub


    '======================================================================================
    ' zipmodファイルリスト取得--------------------------------------------------------------
    Private Async Function GetZipModFilesAsync(ByVal folderPath As String) As Task(Of List(Of String))
        Dim fileList As New List(Of String)
        Try
            Await Task.Run(Sub()
                               Dim files() As String = Directory.GetFiles(folderPath, "*.*", IO.SearchOption.AllDirectories)
                               For Each file As String In files
                                   Dim ext As String = Path.GetExtension(file)
                                   If ext.ToLower() = ".zipmod" OrElse ext.ToLower() = ".zi_mod" OrElse ext.ToLower() = ".zip" OrElse ext.ToLower() = ".zi_" Then
                                       fileList.Add(file)
                                   End If
                               Next
                           End Sub)
        Catch ex As Exception
            MessageBox.Show("Cannot access the file. It is being used by another application. Please close the other application and try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            logAdd($"[Error] Cannot access the file. {folderPath}")
            Return fileList
        End Try

        Return fileList
    End Function

    ' 個別modリスト生成--------------------------------------------------------------
    Private Async Function GetModListAsync(ByVal zipModList As List(Of String), validModList As List(Of String()), token As CancellationToken) As Task(Of List(Of String()))
        Dim fullItemList As New List(Of String())
        Dim totalCount As Integer = 0
        Dim zipCount As Integer = 0
        For Each modFilePath As String In zipModList

            If token.IsCancellationRequested Then Return fullItemList ' キャンセル処理

            'validModListの1列目を検索して一致するものがあれば、次のループへスキップ
            If validModList.Any(Function(x) x(1) = modFilePath) Then
                zipCount += 1
                'logAdd($"[Debug] Skip. {Path.GetFileName(modFilePath)}")
                Continue For
            End If

            Dim inModList As List(Of String) = Await GetCSVandXMLFilesAsync(modFilePath)
            Dim inImageList As List(Of String) = Await GetImageFilesAsync(modFilePath)

            If inModList.Count = 0 Then
                logAdd($"[Error] No .csv or .xml files found. {Path.GetFileName(modFilePath)}")
            Else
                inModList.Sort(Function(x, y) If(x.EndsWith(".xml") And y.EndsWith(".csv"), -1, If(y.EndsWith(".xml") And x.EndsWith(".csv"), 1, 0)))

                'logAdd($"[Debug] inModList : {Environment.NewLine}{Path.GetFileName(modFilePath)}{Environment.NewLine}{String.Join(Environment.NewLine, inModList)}")

                Dim inManifest As New Dictionary(Of String, String)()
                Dim xmlCount As Integer = 0

                For Each inFile In inModList
                    Dim csvLineList As New List(Of String())
                    Dim itemArr As String() = Nothing

                    If inFile.EndsWith("org.xml", StringComparison.OrdinalIgnoreCase) Or inFile.EndsWith("orig.xml", StringComparison.OrdinalIgnoreCase) Then
                        'logAdd($"[Notice] Skipping manifest-org.xml & {Path.GetFileName(modFilePath)}")
                        Continue For
                    ElseIf Regex.IsMatch(inFile, ".*\.xml$") AndAlso Not Regex.IsMatch(inFile, ".*\.g\.xml$") Then
                        ' XMLファイルの処理
                        inManifest = Await ReadManifestAsync(modFilePath, inFile)
                        'logAdd($"{String.Join(Environment.NewLine, inManifest)}")
                        If String.IsNullOrEmpty(inManifest("guid")) Then
                            logAdd($"[Error] Guid in manifest.xml not found. {Path.GetFileName(modFilePath)}")
                            Exit For
                        End If
                        ' ファイル更新日時とサイズ、有効無効をManifest変数に追加
                        Dim fileInfoGet As New FileInfo(modFilePath)
                        inManifest("inFileDate") = CStr(fileInfoGet.LastWriteTime)
                        inManifest("inFileSize") = CStr(fileInfoGet.Length)
                        If Path.GetExtension(modFilePath).ToLower() = ".zipmod" Or Path.GetExtension(modFilePath).ToLower() = ".zip" Then
                            inManifest("inFileEnable") = "1"
                        Else
                            inManifest("inFileEnable") = "0"
                        End If
                        xmlCount += 1
                    ElseIf inFile.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) Then
                        If Regex.IsMatch(inFile, "^(?=.*[Cc]ategory)|(?=.*[Gg]roup)|(?=.*[Bb]one)[^\\]+$") Then ' スタジオアイテムの設定CSVは飛ばす
                            'logAdd($"[Debug] Skip. {inFile} {Path.GetFileName(modFilePath)}")
                        ElseIf Regex.IsMatch(inFile, "^(?=.*map.*)(?!.*name|kPlug)[^\\]+$") Then ' マップの付属要素などは飛ばす
                            'logAdd($"[Debug] Skip. {inFile} : {Path.GetFileName(modFilePath)}")
                        Else
                            ' CSVファイルの処理
                            Dim inCsv As List(Of String())
                            inCsv = Await ReadModCsvAsync(modFilePath, inFile, inImageList)
                            'logAdd($"Oh")
                            'logAdd($"{String.Join(vbCrLf, inCsv.Select(Function(arr) String.Join(",", arr)))}")
                            For Each csvLine As String() In inCsv
                                itemArr = {
                                        totalCount.ToString(), modFilePath, Path.GetFileName(modFilePath),
                                        inManifest("inFileSize"), inManifest("inFileDate"),
                                        inManifest("guid"), inManifest("name"), inManifest("version"),
                                        inManifest("author"), inManifest("website"), inManifest("description"),
                                        csvLine(0), csvLine(1), csvLine(2), csvLine(3), inManifest("inFileEnable")
                                    }
                                totalCount += 1
                                fullItemList.Add(itemArr)
                                'logAdd($"{String.Join(Environment.NewLine, itemArr)}")
                                csvLineList.Add(itemArr)
                            Next
                        End If
                    Else
                        logAdd($"[Debug] Unknown file. {inFile} in {Path.GetFileName(modFilePath)}")
                    End If

                    If xmlCount > 1 Then
                        logAdd($"[Notice] There are multiple xml. {inFile} in {Path.GetFileName(modFilePath)}")
                    ElseIf Not inModList.Any(Function(s) s.Contains(".csv", StringComparison.OrdinalIgnoreCase)) Then ' manifestしかないファイル
                        If Not inManifest("KK_UncensorSelector/body/displayName") Is Nothing Then 'アンセンサーセレクター用はこっちに入る
                            inManifest("mod_name") = inManifest("KK_UncensorSelector/body/displayName")
                        Else
                            inManifest("mod_name") = inManifest("name")
                        End If
                        itemArr = {
                            totalCount.ToString(), modFilePath, Path.GetFileName(modFilePath),
                            inManifest("inFileSize"), inManifest("inFileDate"),
                            inManifest("guid"), inManifest("name"), inManifest("version"),
                            inManifest("author"), inManifest("website"), inManifest("description"),
                            "-", inManifest("mod_name"), "Manifest Only", "", inManifest("inFileEnable")
                        }
                        totalCount += 1
                        'logAdd($"[Debug] .xml only. {String.Join(Environment.NewLine, itemArr)}")
                        fullItemList.Add(itemArr)
                    End If
                Next
            End If
            zipCount += 1
            Await Task.Run(Sub() msgArea.BeginInvoke(Sub() msgPut(zipCount.ToString() & " / " & zipModList.Count.ToString() & " Files loading...")))
        Next
        msgPut("Searching...")
        Return fullItemList
    End Function

    ' zip内のCSVとXMLを検索--------------------------------------------------------------
    Private Async Function GetCSVandXMLFilesAsync(ByVal path As String) As Task(Of List(Of String))
        Dim fileList As New List(Of String)
        If Not Await CheckZipFileAsync(path) Then Return fileList ' Zipファイルチェック
        Try
            Await Task.Run(Sub()
                               Using archive As New ZipFile(path)
                                   For Each entry As ZipEntry In archive
                                       If entry.IsFile AndAlso (entry.Name.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) OrElse entry.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) Then
                                           fileList.Add(entry.Name)
                                       End If
                                   Next
                               End Using
                           End Sub)
        Catch ex As IOException
            MessageBox.Show("Cannot access the file. It is being used by another application. Please close the other application and try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            logAdd($"[Error] Cannot access the file. {path}")
            Return fileList
        End Try
        Return fileList
    End Function

    ' zip内の画像を検索--------------------------------------------------------------
    Private Async Function GetImageFilesAsync(ByVal path As String) As Task(Of List(Of String))
        Dim fileList As New List(Of String)
        If Not Await CheckZipFileAsync(path) Then Return fileList ' Zipファイルチェック
        Try
            Await Task.Run(Sub()
                               Using archive As New ZipFile(path)
                                   For Each entry As ZipEntry In archive
                                       If entry.IsFile AndAlso (entry.Name.EndsWith(".png", StringComparison.OrdinalIgnoreCase) OrElse entry.Name.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)) Then
                                           fileList.Add(entry.Name)
                                       End If
                                   Next
                               End Using
                           End Sub)
        Catch ex As IOException
            MessageBox.Show("Cannot access the file. It is being used by another application. Please close the other application and try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            logAdd($"[Error] Cannot access the file. {path}")
            Return fileList
        End Try
        Return fileList
    End Function

    ' manifest読み取り--------------------------------------------------------------
    Private Async Function ReadManifestAsync(ByVal zipFilePath As String, ByVal manifestFilePath As String) As Task(Of Dictionary(Of String, String))
        Dim dict As New Dictionary(Of String, String) From {
                {"guid", Nothing},
                {"name", Nothing},
                {"version", Nothing},
                {"author", Nothing},
                {"description", Nothing},
                {"website", Nothing},
                {"KK_UncensorSelector/body/displayName", Nothing}
            }
        If Not Await CheckZipFileAsync(zipFilePath) Then Return dict ' Zipファイルチェック
        Try
            Await Task.Run(Sub()
                               Using fileStream As FileStream = File.OpenRead(zipFilePath)
                                   Using zipStream As New ZipInputStream(fileStream)
                                       Dim manifestEntry As ZipEntry = zipStream.GetNextEntry()
                                       While manifestEntry IsNot Nothing
                                           If manifestEntry.Name.Equals(manifestFilePath) Then
                                               Using stream As New StreamReader(zipStream)
                                                   Dim doc As New XmlDocument()
                                                   doc.Load(stream)
                                                   Dim root As XmlNode = doc.SelectSingleNode("manifest")
                                                   Dim maniTags() As String = {"guid", "name", "version", "author", "description", "website", "KK_UncensorSelector/body/displayName"}
                                                   For Each maniTag In maniTags
                                                       Dim node As XmlNode = root.SelectSingleNode(maniTag)
                                                       Dim value As String = If(node IsNot Nothing AndAlso Not String.IsNullOrEmpty(node.InnerText), node.InnerText, Nothing)
                                                       dict(maniTag) = value
                                                   Next
                                               End Using
                                               Exit While
                                           End If
                                           manifestEntry = zipStream.GetNextEntry()
                                       End While
                                   End Using
                               End Using
                           End Sub)
        Catch ex As Exception
            MessageBox.Show("An error occurred while trying to read the manifest file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            logAdd($"[Error] An error occurred while trying to read the manifest file. {zipFilePath}")
            Return dict
        End Try

        Return dict
    End Function



    ' csv読み取り--------------------------------------------------------------
    Private Async Function ReadModCsvAsync(ByVal zipPath As String, ByVal csvPath As String, ByVal inImageList As List(Of String)) As Task(Of List(Of String()))
        Dim modCatNo As String = ""
        Dim foundModCount As Integer = 0
        Dim itemTempTable As New List(Of String())
        If Not Await CheckZipFileAsync(zipPath) Then Return itemTempTable ' Zipファイルチェック
        Try
            Await Task.Run(Sub()
                               Using archive As New ZipFile(zipPath)
                                   Dim csvEntry As ZipEntry = archive.GetEntry(csvPath)
                                   Using csvStream As Stream = archive.GetInputStream(csvEntry)
                                       Using reader As New StreamReader(csvStream)
                                           Dim isFirstLine As Boolean = True
                                           Dim foundIdRow As Boolean = False
                                           Dim foundHeadSkin As Boolean = False
                                           Dim foundMap As Boolean = False
                                           Dim foundAnime As Boolean = False
                                           Dim foundBgm As Boolean = False
                                           Dim foundLight As Boolean = False
                                           Dim foundProbe As Boolean = False
                                           Dim foundItem As Boolean = False
                                           Dim nameColumn As Integer = -1
                                           Dim thumColumn As Integer = -1
                                           Dim counter As Integer = 1
                                           While Not reader.EndOfStream
                                               Dim fields As String() = reader.ReadLine().Split(","c)

                                               ' 空の行である場合はスキップ
                                               If fields(0).Trim().Length = 0 Then Continue While

                                               ' タブしかない場合はスキップ
                                               If Regex.IsMatch(fields(0), "^\t+$") Then Continue While

                                               If isFirstLine Then
                                                   'logAdd($"[Debug] Now. {csvPath} {Path.GetFileName(zipPath)}")

                                                   If Regex.IsMatch(csvPath, "^.*[Mm]ap[^\\]+[Nn]ame[^\\]+$") Then ' 本編マップ
                                                       itemTempTable.Add({fields(1), If(fields.Length > 2, fields(2), ""), "Map (in Game)", ""}) '一行しかないはずなのでそのまま追加
                                                       foundModCount += 1
                                                       'logAdd($"[Debug] HS2 Map Detected.  {Path.GetFileName(zipPath)} {fields(2)}") ' 
                                                   ElseIf Regex.IsMatch(csvPath, "^.*[Mm]ap_[^\\]+$") Or Regex.IsMatch(fields(0), "^(?i)map.*") Then ' 通常スタジオマップ(kPlug)3行目から。
                                                       foundMap = True
                                                       modCatNo = "Map"
                                                       'logAdd($"[Debug] kPlug Map Detected. {Path.GetFileName(zipPath)} {csvPath}")
                                                   ElseIf Regex.IsMatch(csvPath, "^.*[Aa]nime_[^\\]+$") Or Regex.IsMatch(fields(0), "^(?i)Anime.*") Then ' アニメーション3行目から。
                                                       foundAnime = True
                                                       modCatNo = "Anime"
                                                       'logAdd($"[Debug] Animation Detected. {Path.GetFileName(zipPath)} {csvPath}")
                                                   ElseIf Regex.IsMatch(csvPath, "^.*(?i)bgm_[^\\]+$") Then ' BGM2行目から。
                                                       foundBgm = True
                                                       modCatNo = "BGM"
                                                       'logAdd($"[Debug] BGM Detected. {Path.GetFileName(zipPath)} {csvPath}")
                                                   ElseIf Regex.IsMatch(csvPath, "^.*(?i)light_[^\\]+$") Then ' Light1行目から。
                                                       foundLight = True
                                                       modCatNo = "Light"
                                                       If Integer.TryParse(fields(0), Nothing) AndAlso fields.Length > 1 AndAlso String.IsNullOrEmpty(fields(1)) = False Then 'Light
                                                           itemTempTable.Add({fields(0), If(fields.Length > 1, fields(1), ""), modCatNo, ""})
                                                           foundModCount += 1
                                                       End If
                                                       'logAdd($"[Debug] Light Detected. {Path.GetFileName(zipPath)} {csvPath}")
                                                   ElseIf Regex.IsMatch(csvPath, "^.*(?i)Probe_[^\\]+$") Then ' Probe1行目から。
                                                       foundProbe = True
                                                       modCatNo = "Probe"
                                                       'logAdd($"[Debug] Probe Detected. {Path.GetFileName(zipPath)} {csvPath}")
                                                       If Integer.TryParse(fields(0), Nothing) AndAlso fields.Length > 1 AndAlso String.IsNullOrEmpty(fields(1)) = False Then 'Light
                                                           itemTempTable.Add({fields(0), If(fields.Length > 1, fields(1), ""), modCatNo, ""})
                                                           foundModCount += 1
                                                       End If

                                                   ElseIf Regex.IsMatch(csvPath, "^.*[Ii]tem.*[Ll]ist[^\\]+$") Then ' スタジオアイテム
                                                       foundItem = True
                                                       modCatNo = "Studio Item"

                                                   ElseIf Integer.TryParse(fields(0), Nothing) Then ' 通常アイテム
                                                       If fields(0) = "211" OrElse fields(0) = "111" Then foundHeadSkin = True
                                                       modCatNo = ConvertCategory(fields(0))

                                                   ElseIf fields(0) = "id" OrElse fields(0) = "ID" Then ' 不正ヘッダ対策。今のところヒットはしない
                                                       foundIdRow = True
                                                       nameColumn = Array.IndexOf(fields, "name", StringComparison.OrdinalIgnoreCase)
                                                       If nameColumn = -1 Then nameColumn = 3
                                                       modCatNo = ConvertCategory(fields(1))
                                                       logAdd($"[Notice] Category cell is not integer. [{modCatNo}] {Path.GetFileName(zipPath)} {csvPath}")

                                                   End If
                                                   isFirstLine = False

                                               ElseIf Not foundIdRow AndAlso (fields(0) = "ID" OrElse fields(0) = "id") Then ' 通常アイテムID行
                                                   foundIdRow = True
                                                   nameColumn = Array.IndexOf(fields, "Name", StringComparison.OrdinalIgnoreCase)
                                                   thumColumn = Array.FindIndex(fields, Function(s) Not s.Contains("AB") AndAlso Regex.IsMatch(s, "Thu|thu"))
                                                   If nameColumn = -1 AndAlso fields.Length > 3 Then nameColumn = 3
                                                   If thumColumn = -1 Then thumColumn = 0
                                                   If foundHeadSkin AndAlso fields.Length > 4 Then nameColumn = 4
                                                   If foundHeadSkin AndAlso fields.Length > 10 Then thumColumn = 10
                                                   'logAdd($"[Debug] ID Found. {Path.GetFileName(zipPath)} {csvPath}")

                                               ElseIf foundIdRow Then ' 通常アイテム表示行
                                                   Dim thumValue As String
                                                   If fields.Length > 7 AndAlso thumColumn > 7 AndAlso ImageIsFound(fields(thumColumn), inImageList) Then
                                                       thumValue = fields(thumColumn)
                                                   Else
                                                       thumValue = ""
                                                   End If
                                                   Try
                                                       itemTempTable.Add({fields(0), fields(nameColumn), modCatNo, thumValue})
                                                       foundModCount += 1
                                                   Catch ex As Exception
                                                       MessageBox.Show($"Failed to load column. {Path.GetFileName(zipPath)} {csvPath} {fields(0)} {nameColumn} {thumColumn}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                   End Try

                                               ElseIf foundMap AndAlso fields.Length > 1 AndAlso String.IsNullOrEmpty(fields(1)) = False Then 'スタジオマップ
                                                   itemTempTable.Add({fields(0), If(fields.Length > 1, fields(1), ""), modCatNo, ""})
                                                   foundModCount += 1
                                                   'logAdd($"[Debug] kPlug Map Detected. {Path.GetFileName(zipPath)} {csvPath}")

                                               ElseIf foundAnime AndAlso Integer.TryParse(fields(0), Nothing) AndAlso fields.Length > 4 AndAlso String.IsNullOrEmpty(fields(4)) = False Then 'アニメーション
                                                   itemTempTable.Add({fields(0), If(fields.Length > 4, fields(4), ""), modCatNo, ""})
                                                   foundModCount += 1

                                               ElseIf foundBgm AndAlso Integer.TryParse(fields(0), Nothing) AndAlso fields.Length > 1 AndAlso String.IsNullOrEmpty(fields(1)) = False Then 'BGM
                                                   itemTempTable.Add({fields(0), If(fields.Length > 1, fields(1), ""), modCatNo, ""})
                                                   foundModCount += 1

                                               ElseIf foundLight AndAlso Integer.TryParse(fields(0), Nothing) AndAlso fields.Length > 1 AndAlso String.IsNullOrEmpty(fields(1)) = False Then 'Light
                                                   itemTempTable.Add({fields(0), If(fields.Length > 1, fields(1), ""), modCatNo, ""})
                                                   foundModCount += 1

                                               ElseIf foundProbe AndAlso Integer.TryParse(fields(0), Nothing) AndAlso fields.Length > 1 AndAlso String.IsNullOrEmpty(fields(1)) = False Then 'Probe
                                                   itemTempTable.Add({fields(0), If(fields.Length > 1, fields(1), ""), modCatNo, ""})
                                                   foundModCount += 1

                                               ElseIf foundItem AndAlso Integer.TryParse(fields(0), Nothing) AndAlso fields.Length > 3 AndAlso String.IsNullOrEmpty(fields(3)) = False Then ' スタジオアイテム
                                                   itemTempTable.Add({fields(0), If(fields.Length > 3, fields(3), ""), modCatNo, ""})
                                                   foundModCount += 1
                                               Else
                                                   'logAdd($"[Debug] unknown row : {fields(0)} {Path.GetFileName(zipPath)} {csvPath}")
                                               End If
                                               counter += 1
                                           End While
                                       End Using
                                   End Using
                               End Using
                               If foundModCount = 0 Then logAdd($"[Notice] unknown csv : {Path.GetFileName(zipPath)} {csvPath}")
                           End Sub)
        Catch ex As IOException
            MessageBox.Show("Cannot access the file. It is being used by another application. Please close the other application and try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            logAdd($"[Error] Cannot access the file. {zipPath}")
            Return itemTempTable
        End Try
        '多次元配列に変換する
        Return itemTempTable.Select(Function(itemTemp) {itemTemp(0), itemTemp(1), itemTemp(2), itemTemp(3)}).ToList()
    End Function

    ' csv記載の画像が含まれているか
    Private Function ImageIsFound(field As String, inImageList As List(Of String)) As Boolean
        Dim jpgFile As String = $"{field}.jpg"
        Dim pngFile As String = $"{field}.png"

        ' リストに.jpgまたは.pngファイルが含まれているかチェック
        If inImageList.Any(Function(x) x.Contains(jpgFile)) OrElse inImageList.Any(Function(x) x.Contains(pngFile)) Then
            Return True
        Else
            Return False
        End If
    End Function

    ' ウィンドウ初期化--------------------------------------------------------------
    Private Sub WindowInitialize()
        ' 言語名差し替え
        For Each kvp As KeyValuePair(Of String, String) In guiName
            Dim ctrl As Control = Controls.Find(kvp.Key, True).FirstOrDefault()
            If TypeOf ctrl Is Button Then
                Dim btn As Button = DirectCast(ctrl, Button)
                btn.Text = kvp.Value
            End If
        Next

        If Not My.Settings.WindowPosition.X = -1 Then
            ' My.Settingsからウィンドウの位置とサイズを取得して再現する
            Me.Location = My.Settings.WindowPosition
            Me.Size = My.Settings.WindowSize
            SplitCont1.SplitterDistance = My.Settings.Split1
            SplitCont2.SplitterDistance = My.Settings.Split2
            SplitCont3.SplitterDistance = My.Settings.Split3
        End If
        picBox.Image = My.Resources.PicDef
    End Sub

    ' 終了処理--------------------------------------------------------------
    Private Async Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' ウィンドウの位置とサイズをMy.Settingsに保存する
        Await Task.Run(Sub() Me.BeginInvoke(Sub() My.Settings.WindowPosition = Me.Location))
        Await Task.Run(Sub() Me.BeginInvoke(Sub() My.Settings.WindowSize = Me.Size))
        Await Task.Run(Sub() SplitCont1.BeginInvoke(Sub() My.Settings.Split1 = SplitCont1.SplitterDistance))
        Await Task.Run(Sub() SplitCont2.BeginInvoke(Sub() My.Settings.Split2 = SplitCont2.SplitterDistance))
        Await Task.Run(Sub() SplitCont3.BeginInvoke(Sub() My.Settings.Split3 = SplitCont3.SplitterDistance))
        My.Settings.Save()

        ' ローカルにデータベースをバックアップ
        Await BackupSqlDbAsync()

        ' メインスレッドで処理中の場合、キャンセルする
        If cts IsNot Nothing Then
            logAdd("[Notice] Process is cancelled. Please wait...")
            cts.Cancel()
            'cts.Dispose() ' CancellationTokenSourceを破棄する
            Thread.Sleep(1000)
        End If

        ' 遅延モード時の更新リストが空でなければ確認
        If delayRenameList.Count > 0 Then
            Dim confirmResult As DialogResult = MessageBox.Show("Enable/Disable settings made while Honeyselect2 is running will be discarded. Do you want to exit?", "Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)
            If confirmResult = DialogResult.Cancel Then
                e.Cancel = True ' 終了をキャンセルする
                Return
            End If
        End If
    End Sub

End Class