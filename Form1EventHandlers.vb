Imports System.IO
Imports System.Threading
Imports System.Data.SQLite
Imports System.Text.RegularExpressions
Imports ICSharpCode.SharpZipLib.Zip

Partial Public Class Form1

    ' Rescan DBボタン--------------------------------------------------------------
    Private Async Sub BtnRescan_ClickAsync(sender As Object, e As EventArgs) Handles btnRescan.Click
        If cts IsNot Nothing Then Return ' メインスレッドで処理中の場合、実行しない
        sw.Start() '計測用
        ButtonColorResest(Me) ' ボタンをクリア
        Await Task.Run(Sub() labelCat.BeginInvoke(Sub() labelCat.Text = "Category")) ' ラベルをクリア
        Await SearchGuiStartAsync() ' 表示を検索中モードにする
        Await Task.Run(Sub() labelKey.BeginInvoke(Sub() labelKey.Text = "Keyword Search     - Rebuilding the database. -"))

        ' DBを再スキャン
        Await Task.Run(Sub() logAdd($"----------------------------------------------------------------------------------"))
        Await Task.Run(Sub() logAdd($"[Now] Re-scanning the data and rebuilding the database."))
        Await Task.Run(Sub() logAdd($"----------------------------------------------------------------------------------"))
        scanFullMode = True 'スキャンモードフラグ
        cts = New CancellationTokenSource()
        Dim longProcessThread As New Thread(Sub() InitialProcess(cts.Token, scanFullMode))
        longProcessThread.Start()

    End Sub

    ' Clearボタン--------------------------------------------------------------
    Private Async Sub BtnClr_ClickAsync(sender As Object, e As EventArgs) Handles btnClr.Click
        Await Task.Run(Sub() TextBox1.BeginInvoke(Sub() TextBox1.Text = String.Empty)) ' テキストボックスをクリア
        If cts IsNot Nothing Then Return ' メインスレッドで処理中の場合、実行しない

        sw.Start() '計測用
        Await SearchGuiStartAsync() ' 表示を検索中モードにする
        Await Task.Run(Sub() DataGrid.BeginInvoke(Sub() bindingSource.Filter = ""))
        Await SearchGuiResestAsync() ' 表示リセット
        Await Task.Run(Sub() msgPut($"{bindingSource.Count} Records found."))
        Await Task.Run(Sub() TextBox1.BeginInvoke(Sub()
                                                      TextBox1.ForeColor = Color.Gray
                                                      TextBox1.Text = "Enter text and press return."
                                                  End Sub))
    End Sub

    ' Resetボタン--------------------------------------------------------------
    Private Async Sub BtnRsdt_ClickAsync(sender As Object, e As EventArgs) Handles btnRst.Click
        If cts IsNot Nothing Then Return ' メインスレッドで処理中の場合、実行しない

        sw.Start() '計測用
        ButtonColorResest(Me) ' ボタン色リセット
        Await Task.Run(Sub() TextBox1.BeginInvoke(Sub() TextBox1.Text = String.Empty)) ' テキストボックスをクリア
        Await SearchGuiStartAsync() ' 表示を検索中モードにする
        Await Task.Run(Sub() labelCat.BeginInvoke(Sub() labelCat.Text = "Category")) ' ラベルをクリア
        Dim allData = Await GetDataFromDbAsync() ' SQL完全初期化
        Await Task.Run(Sub() DataGrid.BeginInvoke(Sub() bindingSource.Filter = ""))
        Await ShowResultAsync(allData)
        Await Task.Run(Sub() TextBox1.BeginInvoke(Sub()
                                                      TextBox1.ForeColor = Color.Gray
                                                      TextBox1.Text = "Enter text and press return."
                                                  End Sub))
        Await SearchGuiResestAsync() ' 表示リセット

    End Sub

    ' 重複ID比較--------------------------------------------------------------
    Private Async Sub BtnId_ClickAsync(sender As Object, e As EventArgs) Handles btnId.Click
        If cts IsNot Nothing Then Return ' メインスレッドで処理中の場合、実行しない
        ButtonColorResest(Me) ' ボタン色クリア
        Dim clickedButton = DirectCast(sender, Button)
        Await Task.Run(Sub() clickedButton.BeginInvoke(Sub() clickedButton.BackColor = Color.FromArgb(224, 255, 255))) ' ボタン色変更
        Await Task.Run(Sub() TextBox1.BeginInvoke(Sub() TextBox1.Text = String.Empty)) ' テキストボックスをクリア
        Await SearchGuiStartAsync() ' 表示を検索中モードにする
        Await Task.Run(Sub() labelCat.BeginInvoke(Sub() labelCat.Text = "Category")) ' ラベルをクリア

        ' DBから全てのデータを取得
        Dim allData = Await GetDataFromDbAsync()

        ' DataViewを作成して、xml_guid列でグループ化し、file_fullpath列の値が異なるグループを取得
        Dim query =
            From row In allData.AsEnumerable()
            Group row By guid = row.Field(Of String)("xml_guid") Into rows = Group
            Where rows.Select(Function(row) row.Field(Of String)("file_fullpath")).Distinct().Count() > 1
            Select rows.CopyToDataTable()

        Dim filteredData As New DataTable
        For Each dt In query
            filteredData.Merge(dt)
        Next

        ' 表示する
        If filteredData IsNot Nothing AndAlso filteredData.Rows.Count > 0 Then
            Await Task.Run(Sub() DataGrid.BeginInvoke(Sub() bindingSource.Filter = "")) ' 一旦フィルタークリア
            Await ShowResultAsync(filteredData)
            Await Task.Run(Sub() msgPut($"{bindingSource.Count} Records found."))
        Else
            Await Task.Run(Sub() msgPut("No matching rows found."))
        End If
        Await SearchGuiResestAsync() ' 検索表示用画面色リセット
        Await Task.Run(Sub() TextBox1.BeginInvoke(Sub()
                                                      TextBox1.ForeColor = Color.Gray
                                                      TextBox1.Text = "Enter text and press return."
                                                  End Sub))
    End Sub

    ' TexBoxサーチ--------------------------------------------------------------

    Private Async Sub TextBox1_GotFocusAsync(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        If TextBox1.Text = "Enter text and press return." Then
            Await Task.Run(Sub()
                               TextBox1.BeginInvoke(Sub()
                                                        TextBox1.ForeColor = Color.Black
                                                        TextBox1.Text = ""
                                                    End Sub)
                           End Sub)
        End If
    End Sub

    Private Async Sub TextBox1_LostFocusAsync(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        If TextBox1.Text = "" Then
            Await Task.Run(Sub()
                               TextBox1.BeginInvoke(Sub()
                                                        TextBox1.ForeColor = Color.Gray
                                                        TextBox1.Text = "Enter text and press return."
                                                    End Sub)
                           End Sub)
        End If
    End Sub


    Private Async Sub TextBox1_TextChangedAsync(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If cts IsNot Nothing Then Return ' メインルーチンでスレッド処理中の場合、実行しない

        sw.Start() '計測用
        If e.KeyCode = Keys.Return Then
            e.Handled = True
            e.SuppressKeyPress = True

            Dim searchText As String = Nothing ' 検索文字列

            Await SearchGuiStartAsync() ' 表示を検索中モードにする

            ' テキストを空白トリムして取得
            Await Task.Run(Sub() searchText = TextBox1.Text.Trim())

            ' 空の場合は全検索
            If String.IsNullOrEmpty(searchText) Then
                Await Task.Run(Sub() DataGrid.BeginInvoke(Sub() bindingSource.Filter = "")) ' 一旦フィルタークリア
                Await ShowResultAsync(_allPage)
                Await SearchGuiResestAsync() ' 表示リセット
                Return
            End If

            ' 検索タスクを開始
            Await ShowFilteredResultAsync(searchText)

            Await SearchGuiResestAsync() ' 表示リセット
        End If
    End Sub

    Private Async Function ShowFilteredResultAsync(searchText As String) As Task

        Await Task.Run(Sub()
                           DataGrid.BeginInvoke(Sub()
                                                    ' 勝手にDataSource=listとしたときにデータをバインドさせる処理を取りやめる
                                                    bindingSource.SuspendBinding()

                                                    ' DataSource = listとすると、リストデータをDataGridViewに反映できるが、
                                                    ' そのまま入力すると自動的にリスト生成が反映されてしまうので、
                                                    ' それを停止させる
                                                    bindingSource.RaiseListChangedEvents = False

                                                    ' 一旦別のbsへ
                                                    Dim tempSource As New BindingSource

                                                    tempSource.DataSource = _allPage

                                                    ' フィルタを設定
                                                    tempSource.Filter = $"mod_name Like '%{searchText.Replace("'", "''")}%' OR " +
                                                           $"xml_name Like '%{searchText.Replace("'", "''")}%' OR " +
                                                           $"xml_guid Like '%{searchText.Replace("'", "''")}%' OR " +
                                                           $"xml_author Like '%{searchText.Replace("'", "''")}%' OR " +
                                                           $"file_name Like '%{searchText.Replace("'", "''")}%'"

                                                    ' データを入力する
                                                    'bindingSource.DataSource = viewData
                                                    'DataGrid.DataSource = tempSource
                                                    bindingSource = tempSource

                                                    DataGrid.BeginInvoke(Sub() DataGrid.DataSource = bindingSource)

                                                    ' データ一覧の変更をtrueとする
                                                    bindingSource.RaiseListChangedEvents = True

                                                    ' データバインドを継続させる
                                                    ' このデータバインドを逐一行わせるのが処理を遅くさせている要因である
                                                    bindingSource.ResumeBinding()

                                                    ' バインドされた結果の表示を更新する
                                                    ' - ここでfalseに指定することで、データ構造に変化がないことを通知する
                                                    ' (trueにするとデータ構造ごと反映する必要が出てくるので重くなる)
                                                    ' - ListChangedEventsは、このタイミングで初めて行われるので無駄な処理がなくなる
                                                    bindingSource.ResetBindings(False)

                                                    'DataGrid.DataSource = bindingSource
                                                    If bindingSource.Count = 0 Then
                                                        msgPut("0 Records found.")
                                                    Else
                                                        msgPut($"{bindingSource.Count} Records found.")

                                                    End If

                                                End Sub)
                       End Sub)


    End Function

    ' Open Folderボタン--------------------------------------------------------------
    Private Async Sub BtnOpen_ClickAsync(sender As Object, e As EventArgs) Handles btnOpen.Click
        If cts IsNot Nothing Then Return ' メインスレッドで処理中の場合、実行しない

        'DataGridの選択行のfile_fullpathを取得
        Dim selectedRow = DataGrid.CurrentRow

        If selectedRow Is Nothing Then
            MessageBox.Show("No row selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim filePath As String = selectedRow.Cells("file_fullpath").Value?.ToString()

        If String.IsNullOrEmpty(filePath) Then
            MessageBox.Show("File path is empty.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        'ファイル名を削除したフォルダのパスを取得
        Dim folderPath As String = Path.GetDirectoryName(filePath)

        'フォルダが存在するか確認して、存在しなければエラーメッセージを表示
        If Not Directory.Exists(folderPath) Then
            MessageBox.Show("Folder does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        'フォルダを非同期で開く
        Await Task.Run(Sub()
                           Try
                               Process.Start("explorer.exe", $"/select,""{filePath}""")
                           Catch ex As Exception
                               MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                           End Try
                       End Sub)
    End Sub


    ' カテゴリ検索ボタン--------------------------------------------------------------
    Private Async Sub CategoryButton_ClickAsyn(sender As Object, e As EventArgs) Handles btnCat1.Click, btnCat2.Click, btnCat3.Click, btnCat4.Click, btnCat5.Click, btnCat6.Click, btnCat7.Click, btnCat8.Click, btnCat9.Click, btnCat10.Click, btnCat11.Click, btnCat12.Click, btnCat13.Click, btnCat14.Click, btnCat15.Click, btnCat16.Click, btnCat17.Click, btnCat18.Click, btnCat19.Click, btnCat20.Click, btnCat21.Click, btnCat22.Click, btnCat23.Click, btnCat24.Click, btnCat25.Click, btnCat26.Click, btnCat27.Click, btnCat28.Click, btnCat29.Click, btnCat30.Click, btnCat31.Click, btnCat32.Click, btnCat33.Click, btnCat34.Click, btnCat35.Click, btnCat36.Click, btnCat37.Click, btnCat38.Click, btnCat39.Click, btnCat40.Click, btnCat41.Click, btnCat42.Click, btnCat43.Click, btnCat44.Click, btnCat45.Click, btnCat46.Click, btnCat47.Click, btnCat48.Click, btnCat49.Click, btnCat50.Click, btnCat51.Click, btnCat52.Click, btnCat54.Click, btnCat55.Click, btnCat56.Click
        If cts IsNot Nothing Then Return ' メインスレッドで処理中の場合、実行しない

        sw.Start() '計測用
        Await SearchGuiStartAsync() ' 表示を検索中モードにする

        'クリックされたボタンのタグデータを取得
        Dim tagData As String = If(DirectCast(sender, Button).Tag IsNot Nothing, DirectCast(sender, Button).Tag.ToString(), "")

        ' ボタンの背景色を変更
        ButtonColorResest(Me)
        Dim clickedButton = DirectCast(sender, Button)
        Await Task.Run(Sub() clickedButton.BeginInvoke(Sub() clickedButton.BackColor = Color.FromArgb(224, 255, 255)))

        'クリックされたボタンのテキストを取得してラベルに表示
        Dim buttonText As String = DirectCast(sender, Button).Text.Trim()
        Await Task.Run(Sub() labelCat.BeginInvoke(Sub() labelCat.Text = $"Category - {CType(sender, Button).Parent.Text} {buttonText}"))

        'SQLiteデータベースに接続してデータを取得
        Dim query As String = $"SELECT mod_enable, mod_name, mod_category, mod_id, xml_name, xml_version, xml_author, xml_guid, file_name, file_size, file_date, xml_website, xml_description, mod_thum, file_fullpath FROM mods WHERE mod_category='{ConvertCategory(tagData)}' ORDER BY file_fullpath"
        Dim _allPage = Await GetDataFromDbAsync(query)

        'ShowSearchResultsメソッドに_allPageを渡す
        Await ShowResultAsync(_allPage)

        Await SearchGuiResestAsync() ' 表示リセット
    End Sub

    ' Othersカテゴリ検索--------------------------------------------------------------
    Private Async Sub BtnCat53_Click(sender As Object, e As EventArgs) Handles btnCat53.Click
        If cts IsNot Nothing Then Return ' メインスレッドで処理中の場合、実行しない

        sw.Start() '計測用
        Await SearchGuiStartAsync() ' 表示を検索中モードにする

        ' ボタンの背景色を変更
        ButtonColorResest(Me)
        Dim clickedButton = DirectCast(sender, Button)
        Await Task.Run(Sub() clickedButton.BeginInvoke(Sub() clickedButton.BackColor = Color.FromArgb(224, 255, 255)))

        ' Dictionaryに含まれるカテゴリをクエリし、mod_categoryが一致しないレコードを取得する
        ' exclusionList の初期化
        Dim exclusionList As New List(Of String)

        ' キーが指定された値の場合は exclusionList に値を追加
        For Each category In categories
            Select Case category.Key
                Case "110", "111", "112", "121", "131", "132", "133", "140", "141", "144", "147", "210", "211", "212", "231", "232", "233", "240", "241", "242", "243", "244", "245", "246", "247", "300", "301", "302", "303", "313", "314", "315", "316", "317", "318", "319", "320", "322", "323", "334", "335", "348", "351", "352", "353", "354", "355", "356", "357", "358", "359", "360", "361", "362", "363", "Map", "Studio Item"
                    exclusionList.Add("'" & category.Value & "'")
            End Select
        Next

        ' SQL クエリ文の作成してコマンドを投げる
        Dim sql As String = "SELECT mod_enable, mod_name, mod_category, mod_id, xml_name, xml_version, xml_author, xml_guid, file_name, file_size, file_date, xml_website, xml_description, mod_thum, file_fullpath FROM mods WHERE mod_category NOT IN (" & String.Join(",", exclusionList) & ")"

        Dim resultTable As DataTable = Await GetDataFromDbAsync(sql)

        'ShowSearchResultsメソッドに結果を渡す
        Await ShowResultAsync(resultTable)

        Await SearchGuiResestAsync() ' 表示リセット

    End Sub

    ' ZIPMOD内ビュワー--------------------------------------------------------------
    Private lastCustomThumbnail As String = Nothing ' 最後に表示したカスタムサムネイルパス

    Private Async Sub DataGrid_SelectionChanged(sender As Object, e As EventArgs) Handles DataGrid.SelectionChanged
        If cts IsNot Nothing Then Return ' メインスレッドで処理中の場合、実行しない

        picBox.Image = Nothing ' 画像クリア
        lastCustomThumbnail = Nothing ' 最後に表示したカスタムサムネイルパス

        ' 選択された行のfile_fullpathを取得
        Dim filePath = TryCast(DataGrid.CurrentRow?.Cells("file_fullpath")?.Value, String)
        If String.IsNullOrWhiteSpace(filePath) Then Return
        Dim modId As String = DataGrid.CurrentRow?.Cells("mod_id").Value.ToString()
        Dim modName As String = DataGrid.CurrentRow?.Cells("mod_name").Value.ToString()

        ' ファイルが存在するか確認
        If Not File.Exists(filePath) Then
            'logAdd($"[Error] File not found. {filePath}")
            Return
        End If

        ' ラベルに表示
        viewLabel.Text = $"Item(s) contained in {Path.GetFileName(filePath)}"

        ' データベースから情報を取得して表示
        Dim command As New SQLiteCommand("SELECT mod_name, mod_category, xml_guid, xml_author, mod_thum FROM mods WHERE file_fullpath = @file_fullpath", memoryConnection)
        command.Parameters.AddWithValue("@file_fullpath", filePath)
        Dim reader = Await command.ExecuteReaderAsync()

        Dim table As New DataTable()
        table.Load(reader)
        If table.Rows.Count > 0 Then
            DataGrid2.DataSource = table
        Else
            logAdd($"[Error] SQL : No matching records found in zipmod. {filePath}")
        End If

        ' カスタムサムネイルがあれば表示
        Dim thumbPath = Regex.Replace(filePath, "^.*\\mods\\(.*)\..+$", $"{Application.StartupPath}thumbnail\$1_{modId}_{modName}")
        For Each ext In {".png", ".jpg", ".bmp", ".gif", ".jpeg", ".tiff", ".ico"}
            Dim findoutPath As String = thumbPath + ext
            If File.Exists(findoutPath) Then
                picBox.ImageLocation = findoutPath
                lastCustomThumbnail = findoutPath
                Return
            End If
        Next

        If Not Await CheckZipFileAsync(filePath) Then Return ' Zipファイルチェック

        ' Zipファイルの中身を読み込む
        Dim entryPaths As List(Of String) = Nothing
        Using fileStream As FileStream = File.OpenRead(filePath)
            Using zipStream As New ZipInputStream(fileStream)
                Dim entry As ZipEntry = zipStream.GetNextEntry()
                entryPaths = New List(Of String)()
                While entry IsNot Nothing
                    If Not entry.IsDirectory Then
                        entryPaths.Add(entry.Name)
                    End If
                    entry = zipStream.GetNextEntry()
                End While
            End Using
        End Using

        ' サムネ指定ファイルがあればZipファイルから画像を取得
        Dim modThum = TryCast(DataGrid.CurrentRow?.Cells("mod_thum")?.Value, String)
        Dim imgPath = entryPaths.FirstOrDefault(Function(path) Regex.IsMatch(path, $"\b{modThum}\.(png|jpg|PNG|JPG)\b", RegexOptions.IgnoreCase))
        If Not String.IsNullOrWhiteSpace(modThum) AndAlso Not String.IsNullOrWhiteSpace(imgPath) Then
            'logAdd($"[Debug] Image found. : {imgPath}")
            Using fileStream As FileStream = File.OpenRead(filePath)
                Using zipStream As New ZipInputStream(fileStream)
                    Dim entry As ZipEntry = zipStream.GetNextEntry()
                    While entry IsNot Nothing
                        If Not entry.IsDirectory AndAlso entry.Name.Equals(imgPath) Then
                            picBox.Image = Image.FromStream(zipStream)
                            Return
                        End If
                        entry = zipStream.GetNextEntry()
                    End While
                End Using
            End Using
        End If

        ' 指定画像がなければ1つ目の画像を読み込んで表示
        Dim imageEntryPath = entryPaths.FirstOrDefault(Function(path) path.EndsWith(".png") OrElse path.EndsWith(".jpg"))
        If imageEntryPath IsNot Nothing Then
            Using fileStream As FileStream = File.OpenRead(filePath)
                Using zipStream As New ZipInputStream(fileStream)
                    Dim entry As ZipEntry = zipStream.GetNextEntry()
                    While entry IsNot Nothing
                        If Not entry.IsDirectory AndAlso entry.Name.Equals(imageEntryPath) Then
                            picBox.Image = Image.FromStream(zipStream)
                            'logAdd($"[Info] The image has been automatically loaded. Please note that it may differ from the actual in-game screen. {Path.GetFileName(filePath)} : {imageEntryPath}")
                            Return
                        End If
                        entry = zipStream.GetNextEntry()
                    End While
                End Using
            End Using
        Else
            ' 画像がどこにも無い場合
            picBox.Image = My.Resources.PicDef
        End If

    End Sub

    ' 画像ドロップ--------------------------------------------------------------
    Private Async Sub PicBox_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles picBox.DragDrop
        If cts IsNot Nothing Then Return ' メインスレッドで処理中の場合、実行しない

        ' ドロップされたファイルの取得とチェック
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files As String() = DirectCast(e.Data.GetData(DataFormats.FileDrop), String())
            If files.Length = 1 Then
                Dim fileStr As String = files(0)
                Dim fileExtension As String = Path.GetExtension(fileStr).ToLower()

                ' ドロップされたファイルが画像ファイルでない場合は終了
                If {".png", ".jpg", ".bmp", ".gif", ".jpeg", ".tiff", ".ico"}.Contains(fileExtension) Then
                    ' DataGridのfile_fullpathカラムからサムネイルファイルのパスを生成
                    Dim selectedRowIndex As Integer = DataGrid.CurrentRow.Index
                    Dim fullPath As String = DataGrid.Rows(selectedRowIndex).Cells("file_fullpath").Value.ToString()
                    Dim modId As String = DataGrid.Rows(selectedRowIndex).Cells("mod_id").Value.ToString()
                    Dim modName As String = DataGrid.Rows(selectedRowIndex).Cells("mod_name").Value.ToString()

                    ' サムネイルファイルの出力先パスを生成
                    Dim thumbPath = Regex.Replace(fullPath, "^.*\\mods\\(.*)\..+$", $"{Application.StartupPath}\thumbnail\$1_{modId}_{modName}{fileExtension}")
                    'logAdd(thumbPath)

                    Try
                        ' 出力先フォルダが存在しない場合は作成
                        Dim destinationFolderPath As String = Path.GetDirectoryName(thumbPath)
                        If Not Directory.Exists(destinationFolderPath) Then
                            Directory.CreateDirectory(destinationFolderPath)
                        End If

                        ' ファイルをコピー
                        Await Task.Run(Sub() File.Copy(fileStr, thumbPath, True))

                        ' ピクチャーボックスに表示
                        picBox.ImageLocation = thumbPath
                        lastCustomThumbnail = thumbPath
                    Catch ex As Exception
                        ' エラーが発生した場合は、エラーメッセージを表示する
                        MessageBox.Show("An error occurred while copying the file." & vbCrLf & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                Else
                    ' ドロップされたファイルが画像ファイルでない場合は、エラーメッセージを表示する
                    MessageBox.Show("Please drop an image file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                ' 複数ファイルがドロップされた場合は、エラーメッセージを表示する
                MessageBox.Show("Please drop a single file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    ' 画像削除--------------------------------------------------------------
    Private Sub PicBox_DoubleClick(sender As Object, e As EventArgs) Handles picBox.DoubleClick
        If cts IsNot Nothing Then Return ' メインスレッドで処理中の場合、実行しない
        ' 既に設定されているサムネイルがない場合は何もしない
        If lastCustomThumbnail Is Nothing Then Return

        ' ファイルの存在を確認し、ダイアログを表示する
        If File.Exists(lastCustomThumbnail) Then
            Dim result = MessageBox.Show("Are you sure you want to delete the thumbnail?" & vbCrLf & vbCrLf & lastCustomThumbnail, "Delete Thumbnail", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If result = DialogResult.Yes Then
                Try
                    ' ファイルを削除し、picBoxにデフォルトの画像を設定する
                    File.Delete(lastCustomThumbnail)
                    picBox.Image = My.Resources.PicDef
                    lastCustomThumbnail = Nothing ' 変数を初期化する
                Catch ex As Exception
                    ' エラー処理
                    MessageBox.Show("Failed to delete thumbnail.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            ' ファイルが存在しない場合はエラーメッセージを表示する
            MessageBox.Show("Thumbnail file not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    ' ファイル削除--------------------------------------------------------------
    Private Async Sub BtnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If cts IsNot Nothing Then Return ' メインスレッドで処理中の場合、実行しない

        If Await Task.Run(Function() IsHs2RunningAsync()) Then
            MessageBox.Show("HoneySelect2 is running. Please try again later.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        '現在選択されている行を取得して削除実行
        Dim selectedRows = DataGrid.SelectedRows
        DataGrid.ClearSelection()
        If selectedRows.Count > 0 Then
            Dim selectedRow = selectedRows(0)
            Dim fullPath = selectedRow.Cells("file_fullpath").Value.ToString()

            If File.Exists(fullPath) Then
                Dim result = MessageBox.Show("The file will be moved to the Recycle Bin. Do you really want to delete it?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
                If result = DialogResult.Yes Then
                    FileIO.FileSystem.DeleteFile(fullPath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    Await DeleteRecordAsync(fullPath)
                    Dim rowsToDelete = _allPage.Select("file_fullpath = '" & fullPath.Replace("'", "''") & "'")
                    For Each row In rowsToDelete
                        row.Delete()
                    Next
                    Await ShowResultAsync(_allPage)
                End If
            Else
                MessageBox.Show("The file does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            MessageBox.Show("Please select a row to delete.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Async Function DeleteRecordAsync(fullPath As String) As Task
        Dim connectionString = $"Data Source={dbPath}"
        Using connection As New SQLiteConnection(connectionString)
            Await connection.OpenAsync()
            Using transaction = connection.BeginTransaction()
                Try
                    Dim deleteCommand = connection.CreateCommand()
                    deleteCommand.Transaction = transaction
                    deleteCommand.CommandText = "DELETE FROM mods WHERE file_fullpath = $fullPath"
                    deleteCommand.Parameters.AddWithValue("$fullPath", fullPath)
                    Await deleteCommand.ExecuteNonQueryAsync()
                    Await transaction.CommitAsync()
                Catch ex As Exception
                    transaction.Rollback()
                    MessageBox.Show("Failed to delete record from database." & vbCrLf & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        End Using
    End Function

    '　Enable/Disable処理--------------------------------------------------------------
    Private Async Sub EnableDisable_Click(sender As Object, e As EventArgs) Handles btnEnable.Click, btnDisable.Click
        If cts IsNot Nothing Then Return ' メインスレッドで処理中の場合、実行しない
        If Await Task.Run(Function() IsHs2RunningAsync()) Then delayMode = True '起動中はリネームのみ

        ' 選択された行を取得
        Dim selectedRows = DataGrid.SelectedRows
        If selectedRows.Count = 0 Then Return

        ' 選択された行のmod_enableカラムが正しいか確認
        Dim enableColumn = DataGrid.Columns("mod_enable").Index
        Dim selectedRow = selectedRows(0)
        If (sender Is btnEnable AndAlso CBool(selectedRow.Cells(enableColumn).Value)) OrElse (sender Is btnDisable AndAlso Not CBool(selectedRow.Cells(enableColumn).Value)) Then
            Return
        End If

        ' 選択された行のファイルフルパスを取得し、HS2起動中でなければファイルが存在するか確認する
        Dim fileFullPath As String = DataGrid.CurrentRow.Cells("file_fullpath").Value.ToString()
        If Not delayMode And File.Exists(fileFullPath) = False Then
            logAdd($"[Error] File is not found. {fileFullPath}")
            Return
        End If

        ' 新しいファイル名を設定
        Dim newFileFullPath As String
        If Path.GetExtension(fileFullPath).ToLower() = ".zip" Or Path.GetExtension(fileFullPath).ToLower() = ".zi_" Then
            newFileFullPath = If(sender Is btnEnable, Path.ChangeExtension(fileFullPath, ".zip"), Path.ChangeExtension(fileFullPath, "zi_"))
        Else
            newFileFullPath = If(sender Is btnEnable, Path.ChangeExtension(fileFullPath, ".zipmod"), Path.ChangeExtension(fileFullPath, "zi_mod"))
        End If

        ' データグリッドの値を変更
        For Each rowToUpdate As DataGridViewRow In DataGrid.Rows.Cast(Of DataGridViewRow)().Where(Function(row) row.Cells("file_fullpath").Value.ToString() = fileFullPath).ToList()
            rowToUpdate.Cells("file_fullpath").Value = newFileFullPath
            rowToUpdate.Cells("mod_enable").Value = sender Is btnEnable
            rowToUpdate.DefaultCellStyle.ForeColor = If(sender Is btnEnable, SystemColors.WindowText, Color.Gray)
        Next

        ' データベースに接続し、クエリを実行する
        Dim modEnable As String = If(sender Is btnEnable, "1", "0")
        Dim query As String = $"UPDATE mods SET file_fullpath = @newFileFullPath, mod_enable = '{modEnable}' WHERE file_fullpath = @fileFullPath"
        Await SqlUpdateAsync(query, fileFullPath, newFileFullPath)
        Await GetDataFromDbAsync()

        ' ローカルにデータベースをバックアップ
        'Await BackupSqlDbAsync()

        If Not delayMode Then
            ' ファイルをリネーム
            File.Move(fileFullPath, newFileFullPath)
            logAdd($"[Done] Renamed to {Path.GetFileName(newFileFullPath)}")
        Else
            ' 遅延モード時はリストへ入れるだけ
            delayRenameList.Add((fileFullPath, newFileFullPath))
            logAdd($"[Notice] Renaming will be executed after HoneySelect is closed. {Path.GetFileName(fileFullPath)} to {Path.GetExtension(newFileFullPath)}")
        End If
    End Sub

    ' 遅延モード時のファイルリネーム設定反映--------------------------------------------------------------
    Private Async Function EnableUpdateDeleyAsync() As Task
        Await Task.Run(Sub()
                           For Each item In delayRenameList

                               If Not File.Exists(item.fileFullPath) Then Continue For
                               File.Move(item.fileFullPath, item.newFileFullPath)
                               logAdd($"[Done] Renamed to {Path.GetFileName(item.newFileFullPath)}")
                           Next
                           logAdd($"[Done] {delayRenameList.Count} item(s) have been renamed.")
                       End Sub)
    End Function


    ' ボタンカラー初期化メソッド--------------------------------------------------------------
    Private Function ButtonColorResest(parent As Control) As List(Of Control)
        Dim result As New List(Of Control)

        For Each control In parent.Controls
            result.Add(CType(control, Control))
            result.AddRange(ButtonColorResest(CType(control, Control)))
        Next

        For Each control In result 'ボタンクリア
            If TypeOf control Is Button Then
                Dim button = DirectCast(control, Button)
                If button.BackColor = Color.FromArgb(224, 255, 255) Then
                    button.Invoke(Sub() button.BackColor = SystemColors.ControlLight)
                End If
            End If
        Next

        Return result
    End Function

    ' サーチ表示画面メソッド--------------------------------------------------------------
    Private Async Function SearchGuiStartAsync() As Task
        Await Task.Run(Sub() msgArea.BeginInvoke(Sub() msgPut("Searching...")))
        Await Task.Run(Sub() labelKey.BeginInvoke(Sub() labelKey.Text = "Keyword Search     - Searching -"))
        Await Task.Run(Sub() TextBox1.BeginInvoke(Sub() TextBox1.BackColor = Color.FromArgb(218, 237, 255)))
        Await Task.Run(Sub() DataGrid.BeginInvoke(Sub() DataGrid.DefaultCellStyle.BackColor = Color.FromArgb(218, 237, 255)))
    End Function
    ' 表示初期化メソッド--------------------------------------------------------------
    Private Async Function SearchGuiResestAsync() As Task
        Await Task.Run(Sub() TextBox1.BeginInvoke(Sub() TextBox1.BackColor = SystemColors.Window))
        Await Task.Run(Sub() DataGrid.BeginInvoke(Sub() DataGrid.DefaultCellStyle.BackColor = SystemColors.Window))
        Await Task.Run(Sub() labelKey.BeginInvoke(Sub() labelKey.Text = "Keyword Search"))
    End Function

End Class
