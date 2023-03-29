Imports System.Text
Imports System.Data.SQLite
Imports System.IO
Imports System.Text.RegularExpressions

Partial Public Class Form1

    Private memoryConnection As SQLiteConnection ' メモリ上のSQLコネクション

    ' SQLテーブル初期化--------------------------------------------------------------
    Public Sub InitialSQL()
        ' SQLiteに接続するための変数を宣言
        Dim connectionString As String = $"Data Source={dbPath};"
        Using conn As New SQLiteConnection(connectionString)
            ' データベースファイルへのパスを指定して接続
            conn.Open()

            ' modsテーブルを作成するSQL文
            Dim createTableQuery As String = "CREATE TABLE IF NOT EXISTS mods (id INTEGER PRIMARY KEY AUTOINCREMENT"
            For i As Integer = 1 To sqlIndex.Length - 1
                createTableQuery &= $", {sqlIndex(i)} TEXT NOT NULL"
            Next
            createTableQuery &= ");"

            ' modsテーブルを作成
            Dim command As New SQLiteCommand(createTableQuery, conn)
            command.ExecuteNonQuery()

            ' インデックスを作成するSQL文
            Dim createIndexQuery As String = $"CREATE INDEX IF NOT EXISTS idx_mods ON mods ({String.Join(", ", sqlIndex.Skip(1))});"

            ' インデックスを作成
            command = New SQLiteCommand(createIndexQuery, conn)
            command.ExecuteNonQuery()

            ' データベース接続を閉じる
            conn.Close()
        End Using
    End Sub

    ' SQL dbファイルをメモリにロード--------------------------------------------------------------
    Public Sub DbLoadToMemory()
        Dim fileConnection As New SQLiteConnection($"Data Source={dbPath}")
        fileConnection.Open()

        ' メモリ上のDBファイルを作成
        memoryConnection = New SQLiteConnection("Data Source=:memory:")
        memoryConnection.Open()

        Try
            ' SQL dbをメモリにコピー
            fileConnection.BackupDatabase(memoryConnection, "main", "main", -1, Nothing, 0)
            logAdd("[Done] SQL : Database loaded into memory successfully.")
        Catch ex As Exception
            logAdd($"[Error] SQL : Database load failed. {ex.Message}")
        End Try

        fileConnection.Close()
    End Sub

    ' ローカルにデータベースファイルをバックアップ
    Public Async Function BackupSqlDbAsync() As Task
        Await Task.Run(Sub()
                           File.Copy(dbPath, dbBuPath, True)
                           ' Task.Delay(500)
                           ' ローカルファイルにバックアップ
                           Using backupConnection As New SQLiteConnection($"Data Source={dbBuPath}")
                               backupConnection.Open()
                               memoryConnection.BackupDatabase(backupConnection, "main", "main", -1, Nothing, 0)
                           End Using

                           Try
                               ' バックアップファイルを前のDBに上書き
                               File.Move(dbBuPath, dbPath, True)
                               logAdd("[Done] SQL : Database backup completed.")
                               File.Delete(dbBuPath)
                           Catch ex As Exception
                               logAdd($"[Error] SQL : Database backup failed. {ex.Message}")
                           End Try
                           ' いったん閉じてまた開く
                           memoryConnection.Close()
                           DbLoadToMemory()
                       End Sub)
    End Function


    ' SQL dbを照合して存在しないファイル＆新しいファイルをチェック--------------------------------------------------------------
    Private Async Function GetDeleteAndUpdateRecordsAsync() As Task(Of (deleteList As String(), updateList As String(), existList As String()))
        Dim deleteOn As Boolean = False ' フラグ: 削除フラグ
        Dim prevFile As String = "" ' 前回処理したファイルのフルパス
        Dim deleteList As New List(Of String)() ' 削除対象のレコードのIDのリスト
        Dim updateList As New List(Of String)() ' 更新対象のファイルパスのリスト
        Dim existList As New List(Of String)() ' 上書きスキップ対象のファイルパスのリスト
        Dim dbAllList As New List(Of String)() ' db検索結果

        ' SQLite接続してmodsテーブルからID、フルパス、ファイルサイズ、更新日時のレコードを取得
        Await Task.Run(Sub()
                           Dim command As New SQLiteCommand("SELECT id, file_fullpath, file_size, file_date FROM mods", memoryConnection)
                           Dim reader As SQLiteDataReader = command.ExecuteReader()
                           While reader.Read()
                               Dim line As String = $"{reader.GetInt32(0)},{reader.GetString(1)},{reader.GetString(2)},{reader.GetString(3)}"
                               dbAllList.Add(line)
                           End While
                       End Sub)

        For Each line As String In dbAllList ' 全レコードを処理
            Dim arrStr As String() = line.Split(","c)

            ' 前回と同じファイルのレコードである場合
            If arrStr(1) = prevFile AndAlso deleteOn Then
                deleteList.Add(arrStr(0)) ' 削除対象のレコードのIDをリストに追加する
                Continue For
            ElseIf arrStr(1) = prevFile AndAlso Not deleteOn Then
                'logAdd($"[OK] {arrStr(1)}")
                Continue For
            End If
            Await Task.Run(Sub()
                               ' ファイルが存在しない場合、削除対象のレコードのIDをリストに追加してフラグを立てる
                               If Not File.Exists(Regex.Replace(arrStr(1), "^(.*\\)(mods\\)(.*)$", hs2Path & "\$2$3")) Then
                                   deleteList.Add(arrStr(0))
                                   deleteOn = True
                                   'logAdd($"[Caution] File is not found. {arrStr(1)}")
                               Else
                                   ' ファイルが存在する場合、サイズと日付を比較して更新ファイルのパスをリストに追加する
                                   Dim fileDateLocal As String = CStr(File.GetLastWriteTime(arrStr(1)))
                                   Dim fileSizeLocal As String = CStr(New FileInfo(arrStr(1)).Length)
                                   If arrStr(2) <> fileSizeLocal OrElse arrStr(3) <> fileDateLocal Then
                                       updateList.Add(arrStr(1))
                                       logAdd($"[Update] {arrStr(1)}")
                                   Else
                                       existList.Add(arrStr(1))
                                       'logAdd($"[OK] {arrStr(1)}")
                                   End If
                                   deleteOn = False ' 削除フラグを解除する
                               End If
                               prevFile = arrStr(1) ' 前回処理したファイルの更新
                           End Sub)
        Next
        Return (deleteList.ToArray(), updateList.ToArray(), existList.ToArray())
    End Function

    ' SQL全取得1(再書き込み用リスト取得)--------------------------------------------------------------
    Public Async Function GetValidModListAsync() As Task(Of List(Of String()))
        Dim validModList As New List(Of String())
        Await Task.Run(Sub()
                           Dim commandText As String = "SELECT * FROM mods"
                           Using command As New SQLiteCommand(commandText, memoryConnection)
                               Using reader As SQLiteDataReader = command.ExecuteReader()
                                   Dim modData As New List(Of String)()
                                   Dim fieldCount As Integer = reader.FieldCount
                                   While reader.Read()
                                       modData.Clear()
                                       For i As Integer = 0 To fieldCount - 1
                                           modData.Add(reader(i).ToString())
                                       Next
                                       validModList.Add(modData.ToArray())
                                   End While
                               End Using
                           End Using
                       End Sub)
        Return validModList
    End Function


    ' SQL全取得2(表示用)--------------------------------------------------------------
    Private Async Function GetDataFromDbAsync(Optional ByVal sqlQuery As String = "") As Task(Of DataTable)

        ' 引数がなければ全取得
        Dim query As String
        If sqlQuery = "" Then
            query = "SELECT mod_enable, mod_name, mod_category, mod_id, xml_name, xml_version, xml_author, xml_guid, file_name, file_size, file_date, xml_website, xml_description, mod_thum, file_fullpath FROM mods ORDER BY file_fullpath"
        Else
            query = sqlQuery
        End If

        Dim dataTable As New DataTable
        ' データベースからデータを取得
        Await Task.Run(Sub()
                           Using command As New SQLiteCommand(query, memoryConnection)
                               Using reader As SQLiteDataReader = command.ExecuteReader()
                                   ' DataTableのデータ型を明示設定（処理速度を考慮）
                                   dataTable.Columns.Add("mod_enable", GetType(Boolean))
                                   dataTable.Columns.Add("mod_name", GetType(String))
                                   dataTable.Columns.Add("mod_category", GetType(String))
                                   dataTable.Columns.Add("mod_id", GetType(String))
                                   dataTable.Columns.Add("xml_name", GetType(String))
                                   dataTable.Columns.Add("xml_version", GetType(String))
                                   dataTable.Columns.Add("xml_author", GetType(String))
                                   dataTable.Columns.Add("xml_guid", GetType(String))
                                   dataTable.Columns.Add("file_name", GetType(String))
                                   dataTable.Columns.Add("file_size", GetType(Long))
                                   dataTable.Columns.Add("file_date", GetType(String))
                                   dataTable.Columns.Add("xml_website", GetType(String))
                                   dataTable.Columns.Add("xml_description", GetType(String))
                                   dataTable.Columns.Add("mod_thum", GetType(String))
                                   dataTable.Columns.Add("file_fullpath", GetType(String))
                                   ' データを取得して DataTable に追加
                                   While reader.Read()
                                       Dim row As DataRow = dataTable.NewRow()
                                       row("mod_enable") = If(String.Equals(reader.GetString(0), "1"), True, False) ' Boolean型に変換 reader.GetBoolean(0) Microsoftではこっち
                                       row("mod_name") = reader.GetString(1)
                                       row("mod_category") = reader.GetString(2)
                                       row("mod_id") = reader.GetString(3) ' Long型に変換
                                       row("xml_name") = reader.GetString(4)
                                       row("xml_version") = reader.GetString(5)
                                       row("xml_author") = reader.GetString(6)
                                       row("xml_guid") = reader.GetString(7)
                                       row("file_name") = reader.GetString(8)
                                       row("file_size") = Convert.ToInt64(reader.GetString(9)) ' Long型に変換
                                       row("file_date") = reader.GetString(10)
                                       row("xml_website") = reader.GetString(11)
                                       row("xml_description") = reader.GetString(12)
                                       row("mod_thum") = reader.GetString(13)
                                       row("file_fullpath") = reader.GetString(14)

                                       dataTable.Rows.Add(row)
                                       'msgPut(dataTable.Rows.Count.ToString & " Records found.")
                                   End While
                               End Using
                           End Using
                       End Sub)
        Return dataTable
    End Function

    ' SQL削除実行(引数あり)--------------------------------------------------------------
    Private Async Function SqlDeleteRecordsAsync(ByVal deleteList As String()) As Task(Of String)
        Dim errorMessage As String = ""

        ' トランザクションを開始
        Using transaction As SQLiteTransaction = memoryConnection.BeginTransaction()
            Try
                ' レコードを削除
                For Each id As String In deleteList
                    Using command As New SQLiteCommand($"DELETE FROM mods WHERE id = '{id}'", memoryConnection, transaction)
                        Await command.ExecuteNonQueryAsync()
                    End Using
                Next
                ' トランザクションをコミット
                transaction.Commit()
            Catch ex As Exception
                ' エラーが発生した場合は、トランザクションをロールバック
                transaction.Rollback()
                errorMessage = ex.Message
            End Try
        End Using

        Return errorMessage
    End Function

    ' SQLテーブル全削除--------------------------------------------------------------
    Private Async Function ClearModsTableAsync() As Task
        Dim sql As String = "DELETE FROM mods;"

        Using transaction As SQLiteTransaction = memoryConnection.BeginTransaction()
            Try
                Dim command As New SQLiteCommand(sql, memoryConnection, transaction)
                Await command.ExecuteNonQueryAsync()
                transaction.Commit()
            Catch ex As Exception
                transaction.Rollback()
                Throw
            End Try
        End Using
    End Function

    ' SQL書き込み実行--------------------------------------------------------------
    Private Async Function SqlBulkInsertAsync(ByVal writeList As List(Of String()), ByVal batchSize As Integer) As Task(Of String)
        Try
            ' トランザクションを開始
            Using transaction = memoryConnection.BeginTransaction()
                ' writeListを分割してバルクインサート
                For i As Integer = 0 To writeList.Count - 1 Step batchSize
                    ' バルクインサート用のSQL文を作成
                    Dim insertQuery = New StringBuilder("INSERT INTO mods (")
                    For j As Integer = 1 To sqlIndex.Length - 1
                        insertQuery.Append(sqlIndex(j))
                        If j < sqlIndex.Length - 1 Then insertQuery.Append(", ")
                    Next
                    insertQuery.Append(") VALUES ")

                    ' バルクインサートする値を設定
                    Dim endRecord As Integer = Math.Min(i + batchSize - 1, writeList.Count - 1)
                    For j As Integer = i To endRecord
                        insertQuery.Append("(")
                        For k As Integer = 1 To 15 ' writeListの2列目から16列目までを取得
                            insertQuery.Append($"'{If(writeList(j)(k) Is Nothing, "", writeList(j)(k).Replace("'", "''"))}'")
                            If k < 15 Then insertQuery.Append(", ")
                        Next
                        insertQuery.Append(")")

                        If j < endRecord Then insertQuery.Append(", ")
                    Next

                    ' バルクインサートを実行
                    Using cmd As New SQLiteCommand(insertQuery.ToString(), memoryConnection, transaction)
                        Await cmd.ExecuteNonQueryAsync()
                    End Using
                Next

                ' トランザクションをコミット
                transaction.Commit()
                Return "[Done] SQL : All records inserted successfully."
            End Using

        Catch ex As Exception
            Return $"[Error] SQL : inserting records: {ex.Message}{vbCrLf}{ex.StackTrace}"
        End Try
    End Function

    ' SQL更新（Enable/Disable変更）--------------------------------------------------------------
    Private Async Function SqlUpdateAsync(ByVal query As String, ByVal fileFullPath As String, ByVal newFileFullPath As String) As Task
        Try
            Using command As New SQLiteCommand(query, memoryConnection)
                command.Parameters.AddWithValue("@fileFullPath", fileFullPath)
                command.Parameters.AddWithValue("@newFileFullPath", newFileFullPath)
                Dim rowsUpdated As Integer = Await command.ExecuteNonQueryAsync()
                'logAdd($"[Done] SQL : {rowsUpdated} row(s) updated.")
            End Using
        Catch ex As Exception
            logAdd($"[Error] SQL : {ex.Message}")
        End Try
    End Function

    ' データ処理関係メソッド==============================================================
    Private ReadOnly _pageSize As Integer = 500 ' 1ページあたりの行数
    Private _allPage As DataTable ' ページング対象のデータテーブル
    Private _currentPage As Integer = 0 ' 現在のページ番号
    Private _totalPage As Integer ' 全ページ数
    Private _pageDataList As New List(Of DataTable) ' 分割データ
    Private bindingSource As New BindingSource() 'バインド用変数
    Private bindingSource2 As New BindingSource() 'バインド用変数

    ' データグリッド表示--------------------------------------------------------------
    Private Async Function ShowResultAsync(ByVal results As DataTable) As Task
        ' 渡されたデータが空の場合は表示をクリアして終了する
        If results.Rows.Count = 0 Then
            DataGrid.BeginInvoke(Sub()
                                     bindingSource.DataSource = emptyDataTable
                                     DataGrid.DataSource = bindingSource
                                     msgPut("0 Records found.")
                                     Return
                                 End Sub)
        End If
        ' ソートがある場合、resultsに反映する
        If _allPage IsNot Nothing AndAlso _allPage.Rows.Count > 0 AndAlso _allPage.Columns.Contains("SortColumn") Then
            Dim sortColumn As String = _allPage.Rows(0)("SortColumn").ToString()
            Dim sortOrder As String = _allPage.Rows(0)("SortOrder").ToString()
            If Not String.IsNullOrEmpty(sortColumn) AndAlso Not String.IsNullOrEmpty(sortOrder) Then
                results.DefaultView.Sort = $"{sortColumn} {sortOrder}"
                results = results.DefaultView.ToTable()
            End If
        End If

        ' 1000行以上だけ分割処理
        If results.Rows.Count > 1000 Then
            ' ページ分割処理
            _pageDataList.Clear()
            Dim pageCount As Integer = CInt(Math.Ceiling(results.Rows.Count / _pageSize))
            For i As Integer = 0 To pageCount - 1
                Dim startIndex As Integer = i * _pageSize
                Dim endIndex As Integer = Math.Min(startIndex + _pageSize - 1, results.Rows.Count - 1)
                Dim pageData As DataTable = results.AsEnumerable().Skip(startIndex).Take(endIndex - startIndex + 1).CopyToDataTable()
                _pageDataList.Add(pageData)
            Next
            _totalPage = _pageDataList.Count
        Else
            _totalPage = 1
            _pageDataList.Clear()
            _pageDataList.Add(results)
        End If

        ' データをバインドして新しいデータを表示
        _allPage = results
        _currentPage = 0

        DataGrid.BeginInvoke(Sub()
                                 ' 自動バインド停止
                                 'bindingSource.RaiseListChangedEvents = False
                                 bindingSource.SuspendBinding()

                                 bindingSource.DataSource = _pageDataList(0) ' データ入力
                                 'bindingSource.DataSource = _allPage ' データ入力
                                 bindingSource.RaiseListChangedEvents = True ' データ変更許可

                                 ' 更新反映
                                 bindingSource.ResumeBinding()
                                 bindingSource.ResetBindings(False)

                                 'DataGrid.FirstDisplayedScrollingRowIndex = 0 'スクロール位置リセット
                                 UpdatePaginationButtons()
                             End Sub)
        Await Task.Run(Sub() msgPut(If(_totalPage > 1, $"{_allPage.Rows.Count} Records found. (Page {_currentPage + 1} / {_totalPage})", $"{_allPage.Rows.Count} Records found.")))

        ' ページネーションボタンの生成

    End Function

    ' データグリッドページ遷移表示--------------------------------------------------------------
    Private Sub UpdateGridView()

        ' 現在のページのデータを取得
        Dim startIndex As Integer = _currentPage * _pageSize
        Dim endIndex As Integer = Math.Min(startIndex + _pageSize - 1, _allPage.Rows.Count - 1)
        Dim pageData As DataTable = _pageDataList(_currentPage)

        ' 自動バインド停止
        bindingSource.RaiseListChangedEvents = False
        bindingSource.SuspendBinding()

        DataGrid.BeginInvoke(Sub() bindingSource.DataSource = pageData) ' データ入力
        bindingSource.RaiseListChangedEvents = True ' データ変更許可

        ' 更新反映
        bindingSource.ResumeBinding()
        bindingSource.ResetBindings(False)

        msgPut($"{_allPage.Rows.Count} Records found. (Page {_currentPage + 1} / {_totalPage})")

        ' ページネーションボタンの生成
        DataGrid.BeginInvoke(Sub()
                                 ' ページネーションボタンの生成
                                 UpdatePaginationButtons()
                             End Sub)
    End Sub

    ' データグリッドのスクロールハンドラ--------------------------------------------------------------

    Private rowCounter As Integer = -1
    Private prevMove As Boolean = False

    Private Sub DataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGrid.RowsAdded
        ' データグリッドの描画が完了したらスクロールを実行する
        If prevMove AndAlso DataGrid.Rows.Count = rowCounter Then
            DataGrid.FirstDisplayedScrollingRowIndex = Math.Max(0, DataGrid.DisplayedRowCount(True) - 1)
            prevMove = False
        End If
    End Sub

    Private Sub DataGrid_Scroll(sender As Object, e As ScrollEventArgs) Handles DataGrid.Scroll
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll And bindingSource.Filter = Nothing Then
            If e.Type = ScrollEventType.SmallIncrement AndAlso DataGrid.DisplayedRowCount(False) + DataGrid.FirstDisplayedScrollingRowIndex >= DataGrid.RowCount Then
                ' スクロールが最下部に達した場合、次ページを表示
                If _currentPage >= _totalPage - 1 Then Return

                'logBox.BeginInvoke(Sub() msgPut($"{_allPage.Rows.Count} Records found. (Page {_currentPage} / {_totalPage}) Loading..."))
                _currentPage += 1
                UpdateGridView()
                DataGrid.CurrentCell = DataGrid.Rows(0).Cells(0)
            ElseIf e.Type = ScrollEventType.SmallDecrement AndAlso DataGrid.FirstDisplayedScrollingRowIndex = 0 Then
                ' スクロールが最上部に達した場合、前のページを表示
                If _currentPage <= 0 Then Return

                ' イベントを一時的に無効にして表示後、スクロール位置を最下部に設定 
                RemoveHandler DataGrid.Scroll, AddressOf DataGrid_Scroll
                rowCounter = DataGrid.RowCount
                _currentPage -= 1
                UpdateGridView()

                DataGrid.CurrentCell = DataGrid.Rows(DataGrid.Rows.Count - 1).Cells(0)
                prevMove = True
                'DataGrid.FirstDisplayedScrollingRowIndex = DataGrid.RowCount - 0
                AddHandler DataGrid.Scroll, AddressOf DataGrid_Scroll
            End If
        End If
    End Sub

    ' データグリッドのソートクリックハンドラ--------------------------------------------------------------
    Private SortDirection As Boolean = False
    Private Async Sub DataGrid_ColumnHeaderMouseClickAsync(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGrid.ColumnHeaderMouseClick
        ' ソート対象のカラムを取得
        Dim sortColumn As String = DataGrid.Columns(e.ColumnIndex).DataPropertyName

        ' ソート順を決定
        Dim sortOrderStr As String = "ASC"
        If DataGrid.Columns(e.ColumnIndex).HeaderCell.SortGlyphDirection = SortOrder.Descending Or SortDirection = False Then
            sortOrderStr = "ASC"
            SortDirection = True
        Else
            sortOrderStr = "DESC"
            SortDirection = False
        End If

        ' データをソート
        _allPage.DefaultView.Sort = $"{sortColumn} {sortOrderStr}"
        _allPage = _allPage.DefaultView.ToTable()

        ' 再度ページングを実行して画面を初期化
        Await ShowResultAsync(_allPage)
    End Sub

    ' ページのボタン遷移--------------------------------------------------------------
    Private Sub UpdatePaginationButtons()
        ' 既存のボタンを削除
        Panel1.Controls.Clear()

        ' FlowLayoutPanel の作成
        Dim flowLayoutPanelButtons As New FlowLayoutPanel With {
        .AutoSize = True,
        .FlowDirection = FlowDirection.LeftToRight,
        .WrapContents = False,
        .Dock = DockStyle.Right ' Dock プロパティを設定
    }
        ' 既存のボタンを削除
        flowLayoutPanelButtons.Controls.Clear()

        ' ボタン数を設定
        Dim minPage As Integer = Math.Max(0, Math.Max(_currentPage - 4, 0))
        Dim maxPage As Integer = Math.Min(_totalPage - 1, Math.Min(_currentPage + 5, _totalPage - 1))


        ' 開始ページ番号を調整
        If _currentPage <= 4 Then
            maxPage = Math.Min(9, _totalPage - 1)
        ElseIf _currentPage + 5 >= _totalPage Then
            minPage = Math.Max(0, _totalPage - 10)
        End If


        ' 最初のページに戻るボタン
        Dim firstButton As New Button With {
            .Text = "<<",
    .Size = New Size(30, 20),
    .Margin = New Padding(0, 0, 2, 0),
    .Font = New Font("Microsoft Sans Serif", 6) ' フォントサイズを6pxに設定
        }
        AddHandler firstButton.Click, AddressOf FirstButton_Click
        flowLayoutPanelButtons.Controls.Add(firstButton)

        ' 前のページに戻るボタン
        Dim prevButton As New Button With {
    .Text = "<",
    .Size = New Size(26, 20),
    .Margin = New Padding(2, 0, 2, 0),
    .Font = New Font("Microsoft Sans Serif", 6) ' フォントサイズを6pxに設定
}
        AddHandler prevButton.Click, AddressOf PrevButton_Click
        flowLayoutPanelButtons.Controls.Add(prevButton)

        ' ページ番号ボタン
        Dim startPage As Integer = Math.Max(1, _currentPage + 1 - 4)
        Dim endPage As Integer = Math.Min(_totalPage, _currentPage + 1 + 4)

        ' ページ番号ボタン
        For i As Integer = minPage To maxPage
            Dim pageButton As New Button With {
            .Size = New Size(26, 20),
            .Margin = New Padding(2, 0, 2, 0),
            .Font = New Font("Microsoft Sans Serif", 6),
            .Text = (i + 1).ToString()
        }

            ' 現在のページのボタンの背景色を変更
            If i = _currentPage Then
                pageButton.BackColor = Color.FromArgb(224, 255, 255)
            End If

            AddHandler pageButton.Click, AddressOf PageNumberButton_Click
            flowLayoutPanelButtons.Controls.Add(pageButton)
        Next

        ' 次のページに進むボタン
        Dim nextButton As New Button With {
    .Text = ">",
    .Size = New Size(26, 20),
    .Margin = New Padding(2, 0, 2, 0),
    .Font = New Font("Microsoft Sans Serif", 6) ' フォントサイズを6pxに設定
}
        AddHandler nextButton.Click, AddressOf NextButton_Click
        flowLayoutPanelButtons.Controls.Add(nextButton)

        ' 最後のページに進むボタン
        Dim lastButton As New Button With {
            .Text = ">>",
    .Size = New Size(32, 20),
    .Margin = New Padding(2, 0, 0, 0),
    .Font = New Font("Microsoft Sans Serif", 6) ' フォントサイズを6pxに設定
        }
        AddHandler lastButton.Click, AddressOf LastButton_Click
        flowLayoutPanelButtons.Controls.Add(lastButton)


        ' FlowLayoutPanel を Panel1 に追加
        Panel1.Controls.Add(flowLayoutPanelButtons)
    End Sub



    ' 最初のページに戻るボタンのクリックイベント
    Private Sub FirstButton_Click(sender As Object, e As EventArgs)
        _currentPage = 0
        UpdateGridView()
        UpdatePaginationButtons()
    End Sub

    ' 前のページに戻るボタンのクリックイベント
    Private Sub PrevButton_Click(sender As Object, e As EventArgs)
        If _currentPage > 0 Then
            _currentPage -= 1
            UpdateGridView()
            UpdatePaginationButtons()
        End If
    End Sub

    ' ページ番号ボタンのクリックイベント
    Private Sub PageNumberButton_Click(sender As Object, e As EventArgs)
        Dim button As Button = CType(sender, Button)
        Dim pageNumber As Integer = Integer.Parse(button.Text) - 1

        If _currentPage <> pageNumber Then
            _currentPage = pageNumber
            UpdateGridView()
            UpdatePaginationButtons()
        End If
    End Sub

    ' 次のページに進むボタンのクリックイベント
    Private Sub NextButton_Click(sender As Object, e As EventArgs)
        If _currentPage < _totalPage - 1 Then
            _currentPage += 1
            UpdateGridView()
            UpdatePaginationButtons()
        End If
    End Sub

    ' 最後のページに進むボタンのクリックイベント
    Private Sub LastButton_Click(sender As Object, e As EventArgs)
        _currentPage = _totalPage - 1
        UpdateGridView()
        UpdatePaginationButtons()
    End Sub




End Class