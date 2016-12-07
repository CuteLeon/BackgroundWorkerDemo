Imports System.ComponentModel

Public Class Form1
    'ReportProgress 会花费大量时间


    Dim BGWDemo As BackgroundWorker

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        BGWDemo = New BackgroundWorker()
        BGWDemo.WorkerReportsProgress = True
        BGWDemo.WorkerSupportsCancellation = True

        AddHandler BGWDemo.DoWork, AddressOf DoWork
        AddHandler BGWDemo.ProgressChanged, AddressOf ProgressChanged
        AddHandler BGWDemo.RunWorkerCompleted, AddressOf RunWorkerCompleted
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        BGWDemo.RunWorkerAsync(100)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        BGWDemo.CancelAsync()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '后台任务正在执行时需要先取消任务
        If BGWDemo.IsBusy Then BGWDemo.CancelAsync()
        BGWDemo.Dispose()
    End Sub

    ''' <summary>
    ''' 异步任务
    ''' </summary>
    Private Sub DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
        'e.Cancel = False
        '不要在这里操作前台控件，因为代码运行在另一个线程了
        Dim BGWIns As BackgroundWorker = CType(sender, BackgroundWorker)
        Dim Count As Integer = 0
        For Index As Integer = 0 To 1000
            '使用 ReportProgress 更新进度会大量花费时间，所以每执行10次任务回报一次进度
            For IndexInside As Integer = 0 To 10
                Count += 1
            Next

            '在这里触发进度更新事件（userState可以帮助传出任意类型的数据）
            BGWIns.ReportProgress(Index, "任务完成了 " & Index & "%")
            '使用用户传入的参数当做等待超时
            Threading.Thread.Sleep(CInt(e.Argument))

            '需要手动检测是否取消任务（检测 CancellationPending 几乎不花费时间）
            If BGWIns.CancellationPending Then
                '需要手动设置为取消状态
                e.Cancel = True
                e.Result = "取消了任务"
                Exit Sub
            End If
        Next

        '任务完成，返回执行结果（Result 可以帮助传出任何类型的数据）
        e.Result = "任务结束"
    End Sub

    ''' <summary>
    ''' 任务进度更新
    ''' </summary>
    Private Sub ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs)
        'WorkLabel.Text = e.ProgressPercentage & "%"
        WorkLabel.Text = e.UserState.ToString
    End Sub

    ''' <summary>
    ''' 任务完成
    ''' </summary>
    Private Sub RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs)
        If e.Cancelled Then WorkLabel.Text = e.Result : Exit Sub
        If e.Error IsNot Nothing Then WorkLabel.Text = "任务出错：" & e.Error.HResult & vbCrLf & e.Error.Message : Exit Sub

        WorkLabel.Text = e.Result.ToString
    End Sub

End Class
