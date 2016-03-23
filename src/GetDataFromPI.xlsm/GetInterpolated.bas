Attribute VB_Name = "GetInterpolated"
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

Sub GetInterpolated()

On Error GoTo ErrH_Click
    ' Excelのシートの選択
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Set sourceBook = ActiveWorkbook
    Set sourceSheet = sourceBook.Sheets("内挿値")
        
    '記入された内容の削除
    sourceSheet.Range("A7:B100").Select
    Selection.ClearContents
    
    ' シートからデータの収集
    Dim serverName As String
    Dim tagName As String
    Dim startTime As String
    Dim endTime As String
    Dim interval As String
    
    serverName = sourceSheet.Range("B1")
    tagName = sourceSheet.Range("B2")
    startTime = sourceSheet.Range("B3")
    endTime = sourceSheet.Range("B4")
    interval = sourceSheet.Range("B5")
    
    'PI Serverに接続し、タグの定義の設定
    Dim myServer As PISDK.Server
    Dim myTag As PISDK.PIPoint
    
    Set myServer = Servers(serverName)
    'Explict Loginを使う場合は、ユーザー名とパスワードの設定は「("uid=piLoginDemo;pwd=!")」の様になります。
    myServer.Open
    
    Set myTag = myServer.PIPoints(tagName)
    
    
    Dim piArchivedValues As PIValues
    Dim ipid2 As IPIData2
    Set ipid2 = myTag.Data
    Set piArchivedValues = ipid2.InterpolatedValues2(startTime, endTime, interval)
    
    ' TimeStampとValueを表示する。
    Dim i As Integer
    i = 7
    Dim piTempValue As piValue
    For Each piTempValue In piArchivedValues
        sourceSheet.Range("A" + CStr(i)) = piTempValue.timestamp.LocalDate
        sourceSheet.Range("B" + CStr(i)) = piTempValue.value
        i = i + 1
    Next piTempValue
    
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
