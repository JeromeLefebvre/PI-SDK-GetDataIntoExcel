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
    ' Excel�̃V�[�g�̑I��
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Set sourceBook = ActiveWorkbook
    Set sourceSheet = sourceBook.Sheets("���}�l")
        
    '�L�����ꂽ���e�̍폜
    sourceSheet.Range("A7:B100").Select
    Selection.ClearContents
    
    ' �V�[�g����f�[�^�̎��W
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
    
    'PI Server�ɐڑ����A�^�O�̒�`�̐ݒ�
    Dim myServer As PISDK.Server
    Dim myTag As PISDK.PIPoint
    
    Set myServer = Servers(serverName)
    'Explict Login���g���ꍇ�́A���[�U�[���ƃp�X���[�h�̐ݒ�́u("uid=piLoginDemo;pwd=!")�v�̗l�ɂȂ�܂��B
    myServer.Open
    
    Set myTag = myServer.PIPoints(tagName)
    
    
    Dim piArchivedValues As PIValues
    Dim ipid2 As IPIData2
    Set ipid2 = myTag.Data
    Set piArchivedValues = ipid2.InterpolatedValues2(startTime, endTime, interval)
    
    ' TimeStamp��Value��\������B
    Dim i As Integer
    i = 7
    Dim piTempValue As piValue
    For Each piTempValue In piArchivedValues
        sourceSheet.Range("A" + CStr(i)) = piTempValue.timestamp.LocalDate
        sourceSheet.Range("B" + CStr(i)) = piTempValue.value
        i = i + 1
    Next piTempValue
    
    ' �I�u�W�F�N�g�Ɛڑ��̏���
    If myServer.Connected Then
        'PISDK2014R2�̑O�̃o�[�W�����ł́A�ڑ��̏����Ɋւ����肪����̂ŁA�����I�ɐؒf���Ȃ��Ƃ����߂��܂��B
        'PISDK2014R2�̈ȍ~�̃o�[�W�����ł́A���L�̍s�̃R�����g���폜����Ƃ����߂ł�
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
