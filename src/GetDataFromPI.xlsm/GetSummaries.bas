Attribute VB_Name = "GetSummaries"
'   Copyright 2016 OSIsoft, LLC.
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'       http://www.apache.org/licenses/LICENSE-2.0
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.

Option Explicit

Sub GetSummaries()

On Error GoTo ErrH_Click
    
    ' Excelのシートの選択
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Set sourceBook = ActiveWorkbook
    Set sourceSheet = sourceBook.Sheets("サマリー")
    
    ' シートからデータの収集
    Dim serverName As String
    Dim tagName As String
    Dim startTime As String
    Dim endTime As String
    Dim duration As Long
    
    serverName = sourceSheet.Range("B1")
    tagName = sourceSheet.Range("B2")
    startTime = sourceSheet.Range("B3")
    endTime = sourceSheet.Range("B4")
    duration = DateDiff("s", CDate(startTime), CDate(endTime))
    
    'PI Serverに接続し、タグの定義の設定
    Dim myServer As PISDK.Server
    Dim myTag As PISDK.PIPoint
    
    Set myServer = Servers(serverName)
    'Explict Loginを使う場合は、ユーザー名とパスワードの設定は「("uid=piLoginDemo;pwd=!")」の様になります。
    myServer.Open
    
    Set myTag = myServer.PIPoints(tagName)
    
    ' サマリーを収集する
    Dim pdata As PIData
    Dim ipid2 As IPIData2, ipiCalc As IPICalculation
    ' NameValuesの型を使うために、PISDKCommon.dllの参照の追加が必要
    Dim nvsSum As NamedValues
    Set pdata = myTag.Data
    Set ipid2 = pdata                        ' get pointer to IPIData2 Interface
    
    Set nvsSum = ipid2.Summaries2(startTime, endTime, CStr(duration) + "s", ArchiveSummariesTypeConstants.asAll, CalculationBasisConstants.cbEventWeighted)
    
    
    ' 収集したサマリーをシートに書き込む
    Dim valsum As PIValues
    
    Set valsum = nvsSum("Maximum").value
    If valsum.Item(1).IsGood Then
        sourceSheet.Range("B5") = CDbl(valsum.Item(1))
    Else
        sourceSheet.Range("B5") = "Bad Maximum"
    End If
    
    Set valsum = nvsSum("Minimum").value
    If valsum.Item(1).IsGood Then
        sourceSheet.Range("B6") = CDbl(valsum.Item(1))
    Else
       sourceSheet.Range("B6") = "Bad Minimun"
    End If
    
    Set valsum = nvsSum("Average").value
    If valsum.Item(1).IsGood Then
        sourceSheet.Range("B7") = CDbl(valsum.Item(1))
    Else
        sourceSheet.Range("B7") = "Bad Average"
    End If
    
    ' オブジェクトと接続の処理
    If myServer.Connected Then
        'PISDK2014R2の前のバージョンでは、接続の処理に関する問題があるので、明示的に切断しないとお勧めします。
        'PISDK2014R2の以降のバージョンでは、下記の行のコメントを削除するとお勧めです
        myServer.Close
    End If
    
    Set myServer = Nothing
    Set myTag = Nothing
    
Exit_Click:
    Exit Sub
        
ErrH_Click:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_Click
End Sub
