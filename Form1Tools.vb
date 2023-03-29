Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Xml
Imports ICSharpCode.SharpZipLib.Zip

Partial Public Class Form1

    ' ログ出力関数--------------------------------------------------------------
    Private Function logAdd(log As String) As String
        If Not logBox.IsDisposed Then
            Try
                logBox.BeginInvoke(Sub() logBox.AppendText(log & Environment.NewLine))
            Catch

            End Try
        End If
        Return log
    End Function

    ' メッセージ出力関数--------------------------------------------------------------
    Private Function msgPut(mes As String) As String
        If Not msgArea.IsDisposed Then
            Try
                msgArea.BeginInvoke(Sub() msgArea.Text = mes)
            Catch

            End Try
        End If
        Return mes
    End Function

    ' SQLインデックステーブル--------------------------------------------------------------
    Private sqlIndex As String() = {
        "id",
        "file_fullpath",
        "file_name",
        "file_size",
        "file_date",
        "xml_guid",
        "xml_name",
        "xml_version",
        "xml_author",
        "xml_website",
        "xml_description",
        "mod_id",
        "mod_name",
        "mod_category",
        "mod_thum",
        "mod_enable"
    }

    ' データグリッド初期化-------------------------------
    Private ReadOnly emptyDataTable As New DataTable()
    Private ReadOnly emptyDataTable2 As New DataTable()

    ' 仮想モードハンドラ
    Private Sub DataGridView1_CellValueNeeded(sender As Object, e As DataGridViewCellValueEventArgs) Handles DataGrid.CellValueNeeded
        ' 必要なセルの値を取得して、e.Valueに設定する
        If e.ColumnIndex = 0 Then
            e.Value = e.RowIndex ' 行番号を設定
        Else
            e.Value = "Data" ' ダミーデータを設定
        End If
    End Sub
    Private Sub InitializeEmptyDataTable()

        ' 仮想モードを有効にする
        DataGrid.VirtualMode = True
        'DataGrid.RowCount = 100 ' 表示する行数を設定

        ' ダブルバッファリングを有効にする
        DataGrid.GetType().InvokeMember("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, Nothing, DataGrid, New Object() {True})
        DataGrid2.GetType().InvokeMember("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, Nothing, DataGrid2, New Object() {True})

        emptyDataTable.Columns.Add("mod_enable", GetType(Boolean))
        emptyDataTable.Columns.Add("mod_name", GetType(String))
        emptyDataTable.Columns.Add("mod_category", GetType(String))
        emptyDataTable.Columns.Add("mod_id", GetType(String))
        emptyDataTable.Columns.Add("xml_name", GetType(String))
        emptyDataTable.Columns.Add("xml_version", GetType(String))
        emptyDataTable.Columns.Add("xml_author", GetType(String))
        emptyDataTable.Columns.Add("xml_guid", GetType(String))
        emptyDataTable.Columns.Add("file_name", GetType(String))
        emptyDataTable.Columns.Add("file_size", GetType(Long))
        emptyDataTable.Columns.Add("file_date", GetType(String))
        emptyDataTable.Columns.Add("xml_website", GetType(String))
        emptyDataTable.Columns.Add("xml_description", GetType(String))
        emptyDataTable.Columns.Add("mod_thum", GetType(String))
        emptyDataTable.Columns.Add("file_fullpath", GetType(String))

        bindingSource.DataSource = emptyDataTable
        DataGrid.DataSource = bindingSource

        DataGrid.Columns(0).HeaderText = ""
        DataGrid.Columns(1).HeaderText = "Name"
        DataGrid.Columns(2).HeaderText = "Category"
        DataGrid.Columns(3).HeaderText = "Mod id"
        DataGrid.Columns(4).HeaderText = "Mod name"
        DataGrid.Columns(5).HeaderText = "Ver."
        DataGrid.Columns(6).HeaderText = "Author"
        DataGrid.Columns(7).HeaderText = "Guid"
        DataGrid.Columns(8).HeaderText = "File name"
        DataGrid.Columns(9).HeaderText = "File size"
        DataGrid.Columns(10).HeaderText = "File date"
        DataGrid.Columns(11).HeaderText = "Website"
        DataGrid.Columns(12).HeaderText = "Description"
        DataGrid.Columns(13).HeaderText = "Thumbnail"
        DataGrid.Columns(14).HeaderText = "File fullpath"

        DataGrid.Columns(0).Width = 23
        DataGrid.Columns(1).Width = 190
        DataGrid.Columns(2).Width = 90
        DataGrid.Columns(3).Width = 70
        DataGrid.Columns(4).Width = 140
        DataGrid.Columns(5).Width = 40
        DataGrid.Columns(6).Width = 100
        DataGrid.Columns(7).Width = 180
        DataGrid.Columns(8).Width = 180
        DataGrid.Columns(9).Width = 72
        DataGrid.Columns(10).Width = 120
        DataGrid.Columns(11).Width = 160
        DataGrid.Columns(12).Width = 200
        DataGrid.Columns(13).Width = 160
        DataGrid.Columns(14).Width = 450

        emptyDataTable2.Columns.Add("mod_name", GetType(String))
        emptyDataTable2.Columns.Add("mod_category", GetType(String))
        emptyDataTable2.Columns.Add("xml_guid", GetType(String))
        emptyDataTable2.Columns.Add("xml_author", GetType(String))
        emptyDataTable2.Columns.Add("mod_thum", GetType(String))

        bindingSource2.DataSource = emptyDataTable2
        DataGrid2.DataSource = bindingSource2
        DataGrid2.Columns(0).HeaderText = "Name"
        DataGrid2.Columns(1).HeaderText = "Category"
        DataGrid2.Columns(2).HeaderText = "Guid"
        DataGrid2.Columns(3).HeaderText = "Author"
        DataGrid2.Columns(4).HeaderText = "Thumbnail"

    End Sub

    ' カテゴリーテーブル変換--------------------------------------------------------------
    Private categories As New Dictionary(Of String, String)
    Private guiName As New Dictionary(Of String, String)

    Private Function ConvertCategory(value As String) As String
        If categories.ContainsKey(value) Then
            Return categories(value)
        Else
            Return value ' 変換テーブルが見つからなかった場合は生データを返す
        End If
    End Function

    ' ランゲージファイル--------------------------------------------------------------
    Private Function LoadLanguageData() As Boolean
        Try
            Dim xmlDoc As New XmlDocument()
            Dim languageFilePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Language.xml")
            ' ファイルが存在しない場合、リソースからコピーする
            If Not File.Exists(languageFilePath) Then
                Dim lang As String = Thread.CurrentThread.CurrentUICulture.Name
                Dim resourceName As String = If(lang.StartsWith("ja"), "HS2FND.lang_ja.xml", "HS2FND.lang_en.xml")
                Dim resStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName)
                If resStream IsNot Nothing Then
                    Using fs As New FileStream(languageFilePath, FileMode.Create)
                        resStream.CopyTo(fs)
                        logAdd("[Done] Language file was copied from resources.")
                    End Using
                    Debug.Assert(resStream.Position = resStream.Length, "resStream was not fully read.")
                Else
                    logAdd("[Error] Language file was not found in resources.")
                End If
            End If

            xmlDoc.Load(languageFilePath)

            For Each categoryNode As XmlNode In xmlDoc.SelectNodes("//Categories/category")
                Dim key As String = categoryNode.Attributes("key").Value
                Dim value As String = categoryNode.InnerText
                categories.Add(key, value)
            Next

            For Each guiNode As XmlNode In xmlDoc.SelectNodes("//guiName/gui")
                Dim key As String = guiNode.Attributes("key").Value
                Dim value As String = guiNode.InnerText
                guiName.Add(key, value)
            Next

            Return True
        Catch ex As Exception
            MessageBox.Show("Failed to load language data. Error message: " & ex.Message, "Error")
            Return False
        End Try
    End Function

    ' ドロップされたデータがファイルの場合はコピーを許可する--------------------------------------------------------------
    Private Sub picBox_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles picBox.DragEnter
        ' ドロップされたデータがファイルの場合はコピーを許可する
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    ' Zipファイルチェッカー--------------------------------------------------------------
    Public Async Function CheckZipFileAsync(filePath As String) As Task(Of Boolean)
        Dim isValid As Boolean = False

        Try
            Await Task.Run(Sub()
                               Using archive As New ZipFile(File.OpenRead(filePath))
                                   For Each entry As ZipEntry In archive
                                       ' エントリーの情報にアクセスするだけでよい
                                       ' このループが最後まで回ることでZipファイルの検証は完了する
                                   Next
                                   isValid = True
                               End Using
                           End Sub)
        Catch ex As ICSharpCode.SharpZipLib.Zip.ZipException
            ' Zipファイルが不正な場合はFalseを返す
            logAdd($"[Error] Zip file is invalid and cannot be opened. {filePath}")

            Return False
        End Try
        ' Zipファイルが正常な場合はTrueを返す
        ' Zipファイルが正常な場合はTrueを返す
        Return isValid
    End Function

    ' HS2が起動中か判定する--------------------------------------------------------------
    Private shownNotice As Boolean = False ' 注意文を表示済みかどうか
    Private delayMode As Boolean = False '遅延モードフラグ
    Private delayRenameList As New List(Of (fileFullPath As String, newFileFullPath As String)) From {} ' 遅延モード用更新リスト
    Private WithEvents backgroundWorker As New System.ComponentModel.BackgroundWorker() ' バックグラウンドワーカー生成

    Private Async Function IsHs2RunningAsync() As Task(Of Boolean)
        Dim processNames As String() = {"HoneySelect2.exe", "StudioNEOV2.exe", "AI-Syoujyo.exe"}
        For Each processName As String In processNames
            Dim processes() As Process = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(processName))
            For Each process As Process In processes
                If process.MainModule.FileName.StartsWith(hs2Path, StringComparison.OrdinalIgnoreCase) Then
                    If Not shownNotice Then
                        logAdd($"----------------------------------------------------------------------------------")
                        logAdd($"[Notice] HoneySelect2 is running. ")
                        logAdd($"HS2FND has been switched to delay mode.")
                        logAdd($"Enable/Disable settings will take effect after HS2 is closed.")
                        logAdd($"----------------------------------------------------------------------------------")
                        shownNotice = True
                    End If
                    If Not delayMode Then
                        btnDelete.Enabled = False ' 削除ボタンを無効化する
                        Await Task.Run(Sub() backgroundWorker.RunWorkerAsync()) ' ワーカー起動
                        delayMode = True ' 遅延モードに移行
                    End If
                    Return True
                End If
            Next
        Next
        btnDelete.Enabled = True ' 削除ボタンを有効化する
        Return False
    End Function

    Private Async Sub backgroundWorker_DoWorkAsync(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles backgroundWorker.DoWork
        ' バックグラウンドで実行される処理をここに記述
        While Await IsHs2RunningAsync()
            System.Threading.Thread.Sleep(1000) ' 1秒待機
        End While
    End Sub

    Private Async Sub backgroundWorker_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundWorker.RunWorkerCompleted
        Try
            Dim hs2Running As Boolean = Await IsHs2RunningAsync()
            If hs2Running Then
                logAdd("[Now] HoneySelect2 is now Running.") 'タイマー起動確認ログ
            End If
            If Not hs2Running Then

                If delayRenameList.Count > 0 Then
                    logAdd("----------------------------------------------------------------------------------")
                    logAdd("[Now] HoneySelect2 is now closed. Delay mode has been turned off.")
                    logAdd("----------------------------------------------------------------------------------")
                    Await EnableUpdateDeleyAsync() '遅延リネーム実行

                Else
                    logAdd("[Now] HoneySelect2 is now closed. Delay mode has been turned off.")
                End If

                ' ウェイト
                Await Task.Delay(500)

                delayRenameList.Clear()
                btnDelete.Enabled = True ' 削除ボタンを有効化する
                delayMode = False ' 遅延モード解除
            End If
        Catch ex As Exception
            ' 例外処理
            logAdd("[Error] Error occurred: " & ex.Message)
        End Try
    End Sub

    ' データグリッドのenableがFalseの行に色を付ける--------------------------------------------------------------
    Private Sub DataGrid_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGrid.CellPainting
        If DataGrid.Rows.Count > 0 Then
            Dim columnIndex As Integer = _allPage.Columns("mod_enable").Ordinal
            Dim rows = DataGrid.Rows.Cast(Of DataGridViewRow)().
                                Where(Function(row) Not CBool(row.Cells(columnIndex).Value)).
                                ToArray()
            If rows.Length > 0 Then
                For Each row In rows
                    row.DefaultCellStyle.ForeColor = Color.Gray
                Next
            End If
        End If
        sw.Stop()
        'logAdd($"[Debug] Elapsed time: {sw.ElapsedMilliseconds} msec")
        sw.Reset()
    End Sub

    ' Long→MB表記変換--------------------------------------------------------------
    Private Function ConvertFileSizeToMB(ByVal fileSize As Long) As Double
        Return CDbl(fileSize) / (1024.0 * 1024.0)
    End Function

    ' ファイルサイズ列自動変換 & URLリンク化--------------------------------------------------------------
    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGrid.CellFormatting
        If DataGrid.Rows(e.RowIndex).DataBoundItem IsNot Nothing Then ' 仮想モード対応でNullチェックを追加
            If e.ColumnIndex = DataGrid.Columns("file_size").Index AndAlso e.RowIndex >= 0 Then

                Dim value = CType(e.Value, Long)
                e.Value = String.Format("{0:F2} MB", ConvertFileSizeToMB(value))
                e.FormattingApplied = True
            End If
        End If
    End Sub
    ' データテーブルエラー --------------------------------------------------------------
    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGrid.DataError
        ' エラーの発生行と列を表示する
        logAdd($"[Error] Data error at row {e.RowIndex}, column {e.ColumnIndex}")

        ' エラーの詳細を表示する
        logAdd(e.Exception.Message)
    End Sub

End Class
