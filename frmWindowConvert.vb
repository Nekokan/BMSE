Option Strict Off
Option Explicit On

Imports System.Runtime.InteropServices

Friend Class frmWindowConvert
    Inherits System.Windows.Forms.Form

    Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationW" (<[In](), Out()> ByRef lpFileOp As SHFILEOPSTRUCT) As Integer

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> Private Structure SHFILEOPSTRUCT
        Dim hwnd As IntPtr
        Dim wFunc As Integer
        <MarshalAs(UnmanagedType.LPTStr)> Dim pFrom As String
        <MarshalAs(UnmanagedType.LPTStr)> Dim pTo As String
        Dim fFlags As Short
        Dim fAnyOperationsAborted As Integer
        Dim hNameMappings As Integer
        <MarshalAs(UnmanagedType.LPTStr)> Dim lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
    End Structure

    Private Const FO_MOVE As Integer = &H1
    Private Const FO_COPY As Integer = &H2
    Private Const FO_DELETE As Integer = &H3
    Private Const FO_RENAME As Integer = &H4

    Private Const FOF_MULTIDESTFILES As Short = &H1
    Private Const FOF_CONFIRMMOUSE As Short = &H2
    Private Const FOF_SILENT As Short = &H4 '  don't create progress/report
    Private Const FOF_RENAMEONCOLLISION As Short = &H8
    Private Const FOF_NOCONFIRMATION As Short = &H10 '  Don't prompt the user.
    Private Const FOF_WANTMAPPINGHANDLE As Short = &H20 '  Fill in SHFILEOPSTRUCT.hNameMappings
    Private Const FOF_ALLOWUNDO As Short = &H40
    Private Const FOF_FILESONLY As Short = &H80 '  on *.*, do only files
    Private Const FOF_SIMPLEPROGRESS As Short = &H100 '  means don't show names of files
    Private Const FOF_NOCONFIRMMKDIR As Short = &H200 '  don't confirm making any needed dirs

    Private INDEX_MAX As Integer = IIf(frmMain._mnuOptionsBase62.Checked, 3843, 1295)

    Public Sub DeleteUnusedFile()

        Dim i As Integer
        Dim lngTemp As Integer
        Dim blnWAV(INDEX_MAX) As Boolean
        Dim blnBMP(INDEX_MAX) As Boolean
        Dim blnBGA(INDEX_MAX) As Boolean
        Dim strArray() As String
        Dim strLogArray() As String

        ReDim strLogArray(0)

        For i = 0 To UBound(g_Obj) - 1

            With g_Obj(i)

                Select Case .intCh

                    Case Is > 36 * 36

                        blnWAV(.sngValue) = True

                    Case 1 * 36 + 1 To 2 * 36 + 9

                        If .intAtt <> 3 Then blnWAV(.sngValue) = True

                    Case 4, 6, 7

                        blnBMP(.sngValue) = True
                        blnBGA(.sngValue) = True

                End Select

            End With

        Next i

        For i = 0 To UBound(blnBGA)

            If blnBGA(i) Then

                If Len(g_strBGA(i)) Then

                    strArray = Split(g_strBGA(i), " ")

                    lngTemp = modInput.strToNum(strArray(0))

                    If lngTemp >= 0 And lngTemp <= UBound(blnBMP) Then

                        blnBMP(lngTemp) = True

                    End If

                End If

            End If

        Next i

        For i = 1 To INDEX_MAX

            If Not blnWAV(i) Then

                If Len(g_strWAV(i)) Then

                    strLogArray(UBound(strLogArray)) = modInput.strFromNum(modMain.CMD_LOG.LIST_DEL) & "1" & modInput.strFromNum(i) & g_strWAV(i)
                    ReDim Preserve strLogArray(UBound(strLogArray) + 1)

                    g_strWAV(i) = ""

                End If

            End If

            If Not blnBMP(i) Then

                If Len(g_strBMP(i)) Then

                    strLogArray(UBound(strLogArray)) = modInput.strFromNum(modMain.CMD_LOG.LIST_DEL) & "2" & modInput.strFromNum(i) & g_strBMP(i)
                    ReDim Preserve strLogArray(UBound(strLogArray) + 1)

                    g_strBMP(i) = ""

                End If

            End If

            If Not blnBGA(i) Then

                If Len(g_strBGA(i)) Then

                    strLogArray(UBound(strLogArray)) = modInput.strFromNum(modMain.CMD_LOG.LIST_DEL) & "3" & modInput.strFromNum(i) & g_strBGA(i)
                    ReDim Preserve strLogArray(UBound(strLogArray) + 1)

                    g_strBGA(i) = ""

                End If

            End If

        Next i

        If UBound(strLogArray) Then

            Call g_InputLog.AddData(Join(strLogArray, modLog.getSeparator))

        End If

    End Sub

    Public Sub DeleteFile()
        Dim i As Integer
        Dim j As Integer
        Dim lngTemp As Integer
        Dim strArray() As String
        Dim strList() As String
        Dim strDeleteList() As String
        Dim strTemp As String
        Dim sh As SHFILEOPSTRUCT = New SHFILEOPSTRUCT

        If Len(Trim(txtExtension.Text)) = 0 Then Exit Sub

        strArray = Split(txtExtension.Text, ",")

        For i = 0 To UBound(strArray)

            strArray(i) = UCase(strArray(i))

        Next i

        ReDim strList(INDEX_MAX + INDEX_MAX + 1)

        For i = 1 To INDEX_MAX

            If Len(g_strWAV(i)) Then

                strList(lngTemp) = UCase(g_strWAV(i))
                lngTemp = lngTemp + 1

            End If

            If Len(g_strBMP(i)) Then

                strList(lngTemp) = UCase(g_strBMP(i))
                lngTemp = lngTemp + 1

            End If

        Next i

        If Len(frmMain.txtMissBMP.Text) Then

            strList(lngTemp) = UCase(Trim(frmMain.txtMissBMP.Text))
            lngTemp = lngTemp + 1

        End If

        If Len(frmMain.txtStageFile.Text) Then

            strList(lngTemp) = UCase(Trim(frmMain.txtStageFile.Text))
            lngTemp = lngTemp + 1

        End If

        If lngTemp = 0 Then Exit Sub

        ReDim Preserve strList(lngTemp - 1)

        ReDim strDeleteList(0)
        lngTemp = 0

        'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
        strTemp = Dir(g_BMS.strDir & "*.*", FileAttribute.Normal)

        Do While strTemp <> vbNullString

            For i = 0 To UBound(strArray)

                If strArray(i) = UCase(Mid(strTemp, InStrRev(strTemp, ".") + 1)) Then

                    For j = 0 To UBound(strList)

                        If UCase(strTemp) = strList(j) Then

                            Exit For

                        End If

                    Next j

                    If j = UBound(strList) + 1 Then

                        ReDim Preserve strDeleteList(lngTemp)
                        strDeleteList(lngTemp) = g_BMS.strDir & strTemp
                        lngTemp = lngTemp + 1

                    End If

                End If

            Next i

            'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
            strTemp = Dir()

        Loop

        If lngTemp <> 0 Then

            If chkFileRecycle.CheckState Then

                For i = 0 To UBound(strDeleteList)

                    Call lngDeleteFile(strDeleteList(i))

                Next i

            Else

                With sh

                    .hwnd = Me.Handle
                    .wFunc = FO_DELETE
                    .pFrom = Join(strDeleteList, Chr(0))
                    .pTo = vbNullString
                    .fFlags = FOF_SILENT Or FOF_ALLOWUNDO

                End With

                Call SHFileOperation(sh)

            End If

            Call MsgBox(g_Message(modMain.Message.MSG_DELETE_FILE) & vbCrLf & vbCrLf & Replace(Join(strDeleteList, vbCrLf), g_BMS.strDir, ""), MsgBoxStyle.Information, g_strAppTitle)

        End If

    End Sub

    Public Sub ListAlign()
        Dim i As Integer
        'Dim blnUseOldFormat As Boolean
        Dim intTemp As Integer
        Dim lngTemp As Integer
        Dim lngWAV As Integer
        Dim lngBMP As Integer
        Dim lngArray(INDEX_MAX) As Integer
        Dim lngArrayWAV(INDEX_MAX) As Integer
        Dim lngArrayBMP(INDEX_MAX) As Integer
        Dim strArray() As String
        Dim strLogArray() As String

        '共通の処理とか
        For i = 0 To UBound(g_Obj) - 1 '空オブジェ対策

            With g_Obj(i)

                Select Case .intCh

                    Case 1 * 36 + 1 To 2 * 36 + 9, Is > 36 * 36

                        If Len(g_strWAV(.sngValue)) = 0 And .intAtt <> 3 Then

                            g_strWAV(.sngValue) = "***"

                        End If

                    Case 4, 6, 7

                        If Len(g_strBMP(.sngValue)) = 0 Then

                            g_strBMP(.sngValue) = "***"

                        End If

                        If Len(g_strBGA(.sngValue)) = 0 Then

                            g_strBGA(.sngValue) = "***"

                        End If

                End Select

            End With

        Next i

        If frmMain.cboLNObj.SelectedIndex > 0 Then g_strWAV(frmMain.cboLNObj.SelectedIndex) = "***LNOBJ"

        For i = 1 To UBound(g_strWAV)

            If Len(g_strWAV(i)) Then

                lngWAV = lngWAV + 1

            End If

            If Len(g_strBMP(i)) <> 0 Or Len(g_strBGA(i)) <> 0 Then

                lngBMP = lngBMP + 1

            End If

            lngArrayWAV(i) = i
            lngArrayBMP(i) = i
            lngArray(i) = i

        Next i

        If (lngWAV < 256 And lngBMP < 256) And (lngWAV > 0 Or lngBMP > 0) Then

            'If MsgBox(g_Message(Message.MSG_ALIGN_LIST), vbYesNo + vbInformation, g_strAppTitle) = vbNo Then

            'blnUseOldFormat = True

            'Else

            'blnUseOldFormat = False

            'End If

            If chkUseOldFormat.Checked Then

                'blnUseOldFormat = True
                frmMain._mnuOptionsBase16.Checked = True
                frmMain._mnuOptionsBase36.Checked = False
                frmMain._mnuOptionsBase62.Checked = False

            End If

        ElseIf (lngWAV < 1295 And lngBMP < 1295) And (lngWAV > 0 Or lngBMP > 0) Then

            If chkUseOldFormat.Checked Then

                'blnUseOldFormat = False
                frmMain._mnuOptionsBase16.Checked = False
                frmMain._mnuOptionsBase36.Checked = True
                frmMain._mnuOptionsBase62.Checked = False

            End If

        Else

            frmMain._mnuOptionsBase16.Checked = False
            frmMain._mnuOptionsBase36.Checked = False
            frmMain._mnuOptionsBase62.Checked = True

        End If

        'ファイル名ソート
        Dim strDummy(INDEX_MAX) As String
        If chkSortByName.CheckState Then

            Call strQSort(g_strWAV, strDummy, lngArrayWAV, 1, UBound(g_strWAV))
            Call strQSort(g_strBMP, g_strBGA, lngArrayBMP, 1, UBound(g_strBMP))

        End If

        ReDim strLogArray(0)
        strLogArray(0) = ""

        'ここからWAV
        If lngWAV Then

            For i = 0 To UBound(lngArray)

                lngArray(i) = lngArrayWAV(i)

            Next i

            '255以下ならまず後ろに整列する
            'If blnUseOldFormat Then
            If frmMain._mnuOptionsBase16.Checked Then

                lngTemp = 1295

                For i = UBound(g_strWAV) To 1 Step -1

                    If Len(g_strWAV(i)) Then

                        If i <> lngTemp Then

                            'g_strWAV(lngTemp) = g_strWAV(i)
                            'g_strWAV(i) = ""
                            Call swapString(g_strWAV, lngTemp, i)

                            'strTemp = lngArray(i)
                            'lngArray(i) = lngArray(lngTemp)
                            'lngArray(lngTemp) = val(strTemp)
                            Call swapValue(lngArray, lngTemp, i)

                        End If

                        lngTemp = lngTemp - 1

                    End If

                Next i

            End If

            If frmMain._mnuOptionsBase36.Checked Then

                lngTemp = 3843

                For i = UBound(g_strWAV) To 1 Step -1

                    If Len(g_strWAV(i)) Then

                        If i <> lngTemp Then

                            'g_strWAV(lngTemp) = g_strWAV(i)
                            'g_strWAV(i) = ""
                            Call swapString(g_strWAV, lngTemp, i)

                            'strTemp = lngArray(i)
                            'lngArray(i) = lngArray(lngTemp)
                            'lngArray(lngTemp) = val(strTemp)
                            Call swapValue(lngArray, lngTemp, i)

                        End If

                        lngTemp = lngTemp - 1

                    End If

                Next i

            End If

            lngTemp = 1
            intTemp = 1

            For i = 1 To UBound(g_strWAV)

                If Len(g_strWAV(i)) Then

                    'If blnUseOldFormat Then

                    'intTemp = modInput.strToNum(Hex$(lngTemp))

                    'Else

                    'intTemp = lngTemp

                    'End If

                    intTemp = frmMain.lngFromLong(lngTemp)

                    If intTemp <> i Then

                        'g_strWAV(intTemp) = g_strWAV(i)
                        'g_strWAV(i) = ""
                        Call swapString(g_strWAV, intTemp, i)

                        'strTemp = lngArray(intTemp)
                        'lngArray(intTemp) = lngArray(i)
                        'lngArray(i) = val(strTemp)
                        Call swapValue(lngArray, intTemp, i)

                        lngWAV = 1

                    End If

                    lngTemp = lngTemp + 1

                End If

            Next i

            If lngWAV Then

                For i = 1 To UBound(lngArray)

                    If Len(g_strWAV(i)) <> 0 Then

                        lngArrayWAV(lngArray(i)) = i

                        strLogArray(UBound(strLogArray)) = "1" & modInput.strFromNum(lngArray(i)) & modInput.strFromNum(i)
                        ReDim Preserve strLogArray(UBound(strLogArray) + 1)

                    End If

                Next i

            End If

        End If

        'ここからBMP/BGA
        If lngBMP Then

            lngBMP = 0

            For i = 0 To UBound(lngArray)

                lngArray(i) = lngArrayBMP(i)

            Next i

            '255以下ならまず後ろに整列する
            'If blnUseOldFormat Then
            If frmMain._mnuOptionsBase16.Checked Then

                lngTemp = 1295

                For i = UBound(g_strBMP) To 1 Step -1

                    If Len(g_strBMP(i)) <> 0 Or Len(g_strBGA(i)) <> 0 Then

                        If i <> lngTemp Then

                            'g_strBMP(lngTemp) = g_strBMP(i)
                            'g_strBMP(i) = ""
                            'g_strBGA(lngTemp) = g_strBGA(i)
                            'g_strBGA(i) = ""
                            Call swapString(g_strBMP, lngTemp, i)
                            Call swapString(g_strBGA, lngTemp, i)

                            'strTemp = lngArray(i)
                            'lngArray(i) = lngArray(lngTemp)
                            'lngArray(lngTemp) = Val(strTemp)
                            Call swapValue(lngArray, i, lngTemp)

                        End If

                        lngTemp = lngTemp - 1

                    End If

                Next i

            End If

            If frmMain._mnuOptionsBase36.Checked Then

                lngTemp = 3843

                For i = UBound(g_strBMP) To 1 Step -1

                    If Len(g_strBMP(i)) <> 0 Or Len(g_strBGA(i)) <> 0 Then

                        If i <> lngTemp Then

                            'g_strBMP(lngTemp) = g_strBMP(i)
                            'g_strBMP(i) = ""
                            'g_strBGA(lngTemp) = g_strBGA(i)
                            'g_strBGA(i) = ""
                            Call swapString(g_strBMP, lngTemp, i)
                            Call swapString(g_strBGA, lngTemp, i)

                            'strTemp = lngArray(i)
                            'lngArray(i) = lngArray(lngTemp)
                            'lngArray(lngTemp) = Val(strTemp)
                            Call swapValue(lngArray, i, lngTemp)

                        End If

                        lngTemp = lngTemp - 1

                    End If

                Next i

            End If

            lngTemp = 1
            intTemp = 1

            For i = 1 To UBound(g_strBMP)

                If Len(g_strBMP(i)) <> 0 Or Len(g_strBGA(i)) <> 0 Then

                    'If blnUseOldFormat Then

                    'intTemp = modInput.strToNum(Hex$(lngTemp))

                    'Else

                    'intTemp = lngTemp

                    'End If

                    intTemp = frmMain.lngFromLong(lngTemp)

                    If intTemp <> i Then

                        'g_strBMP(intTemp) = g_strBMP(i)
                        'g_strBMP(i) = ""
                        'g_strBGA(intTemp) = g_strBGA(i)
                        'g_strBGA(i) = ""
                        Call swapString(g_strBMP, intTemp, i)
                        Call swapString(g_strBGA, intTemp, i)

                        'strTemp = lngArray(intTemp)
                        'lngArray(intTemp) = lngArray(i)
                        'lngArray(i) = val(strTemp)
                        Call swapValue(lngArray, intTemp, i)

                        lngBMP = 1

                    End If

                    lngTemp = lngTemp + 1

                End If

            Next i

            If lngBMP Then

                For i = 1 To UBound(lngArray)

                    If Len(g_strBMP(i)) <> 0 Or Len(g_strBGA(i)) <> 0 Then

                        lngArrayBMP(lngArray(i)) = i

                        strLogArray(UBound(strLogArray)) = "2" & modInput.strFromNum(lngArray(i)) & modInput.strFromNum(i)
                        ReDim Preserve strLogArray(UBound(strLogArray) + 1)

                    End If

                Next i

                'BGAの方も直すよー
                For i = 0 To UBound(g_strBGA)

                    If Len(g_strBGA(i)) Then

                        strArray = Split(g_strBGA(i), " ")

                        If UBound(strArray) Then

                            strArray(0) = modInput.strFromNum(lngArrayBMP(modInput.strToNum(strArray(0))), 2)
                            g_strBGA(i) = Join(strArray, " ")

                        End If

                    End If

                Next i

            End If

        End If

        '後の処理
        For i = 0 To UBound(g_strWAV)

            If g_strWAV(i) = "***" Then
                g_strWAV(i) = ""
            ElseIf g_strWAV(i) = "***LNOBJ" Then
                g_strWAV(i) = ""
                g_BMS.intLNObj = i
            End If
            If g_strBMP(i) = "***" Then g_strBMP(i) = ""
            If g_strBGA(i) = "***" Then g_strBGA(i) = ""

        Next i

        If lngWAV <> 0 Or lngBMP <> 0 Then

            Call g_InputLog.AddData(modInput.strFromNum(modMain.CMD_LOG.LIST_ALIGN) & Join(strLogArray, ""))

            'Call RefreshList

            For i = 0 To UBound(g_Obj) - 1

                With g_Obj(i)

                    Select Case .intCh

                        Case 1 * 36 + 1 To 2 * 36 + 9, Is > 36 * 36

                            If .intAtt <> 3 Then .sngValue = lngArrayWAV(.sngValue)

                        Case 4, 6, 7

                            .sngValue = lngArrayBMP(.sngValue)

                    End Select

                End With

            Next i

            frmMain.picMain.Refresh()

        End If

    End Sub

    Private Sub strQSort(ByRef strArray1() As String, ByRef strArray2() As String, ByRef lngArray() As Integer, ByVal lngLeft As Integer, ByVal lngRight As Integer)

        Dim i As Integer
        Dim j As Integer

        If lngLeft >= lngRight Then Exit Sub

        i = lngLeft + 1
        j = lngRight

        Do While i <= j

            Do While i <= j

                If StrComp(strArray1(i), strArray1(lngLeft)) > 0 Then
                    Exit Do
                End If

                i = i + 1

            Loop

            Do While i <= j

                If StrComp(strArray1(j), strArray1(lngLeft)) < 0 Then
                    Exit Do
                End If

                j = j - 1

            Loop

            If i >= j Then Exit Do

            Call swapString(strArray1, j, i)
            Call swapString(strArray2, j, i)
            Call swapValue(lngArray, j, i)

            i = i + 1
            j = j - 1

        Loop

        Call swapString(strArray1, j, lngLeft)
        Call swapString(strArray2, j, lngLeft)
        Call swapValue(lngArray, j, lngLeft)

        Call strQSort(strArray1, strArray2, lngArray, lngLeft, j - 1)
        Call strQSort(strArray1, strArray2, lngArray, j + 1, lngRight)

    End Sub

    Private Sub swapString(ByRef strArray() As String, ByVal i As Integer, ByVal j As Integer)

        Dim str_Renamed As String

        str_Renamed = strArray(i)
        strArray(i) = strArray(j)
        strArray(j) = str_Renamed

    End Sub

    Private Sub swapValue(ByRef lngArray() As Integer, ByVal i As Integer, ByVal j As Integer)

        Dim value As Integer

        value = lngArray(i)
        lngArray(i) = lngArray(j)
        lngArray(j) = value

    End Sub

    Public Sub FileNameConvert()
        If (frmMain._mnuOptionsBase62.Checked) Then Exit Sub

        Dim i As Integer
        Dim j As Integer
        Dim strArray() As String
        Dim strTemp As String
        Dim intTemp As Integer
        Dim blnTemp As Boolean
        Dim blnWAV(INDEX_MAX) As Boolean
        Dim blnBMP(INDEX_MAX) As Boolean
        Dim lngTemp As Integer
        Dim strNameFrom() As String
        Dim strNameTo() As String
        Dim sh As SHFILEOPSTRUCT = New SHFILEOPSTRUCT

        ReDim strNameFrom(0)
        ReDim strNameTo(0)
        lngTemp = 0

        Call mciSendString("close PREVIEW", vbNullString, 0, 0)

        For i = 1 To INDEX_MAX

            If Not blnWAV(i) Then

                'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
                If Len(g_strWAV(i)) <> 0 And Dir(g_BMS.strDir & g_strWAV(i), FileAttribute.Normal) <> vbNullString Then

                    strArray = Split(g_strWAV(i), ".")
                    strTemp = strFromNum(i) & "." & strArray(UBound(strArray))

                    'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
                    If Dir(g_BMS.strDir & strTemp, FileAttribute.Normal) = vbNullString Then

                        blnTemp = True '変換するよ

                    Else

                        blnTemp = False 'しないよ

                    End If

                    For j = i + 1 To UBound(g_strWAV)

                        If Not blnWAV(j) Then

                            If g_strWAV(i) = g_strWAV(j) Then

                                If blnTemp Then g_strWAV(j) = strTemp

                                blnWAV(j) = True

                            End If

                        End If

                    Next j

                    If blnTemp Then

                        'Name g_BMS.strDir & g_strWAV(i) As g_BMS.strDir & strTemp

                        ReDim Preserve strNameFrom(lngTemp)
                        ReDim Preserve strNameTo(lngTemp)

                        strNameFrom(lngTemp) = g_BMS.strDir & g_strWAV(i)
                        strNameTo(lngTemp) = g_BMS.strDir & strTemp

                        lngTemp = lngTemp + 1

                        g_strWAV(i) = strTemp

                        intTemp = 1

                    End If

                    blnWAV(i) = True

                End If

            End If

            If Not blnBMP(i) Then

                'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
                If Len(g_strBMP(i)) <> 0 And Dir(g_BMS.strDir & g_strBMP(i), FileAttribute.Normal) <> vbNullString Then

                    strArray = Split(g_strBMP(i), ".")
                    strTemp = strFromNum(i) & "." & strArray(UBound(strArray))

                    'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
                    If Dir(g_BMS.strDir & strTemp, FileAttribute.Normal) = vbNullString Then

                        blnTemp = True '変換するよ

                    Else

                        blnTemp = False 'しないよ

                    End If

                    For j = i + 1 To UBound(g_strBMP)

                        If Not blnBMP(j) Then

                            If g_strBMP(i) = g_strBMP(j) Then

                                If blnTemp Then g_strBMP(j) = strTemp

                                blnBMP(j) = True

                            End If

                        End If

                    Next j

                    If blnTemp Then

                        'Name g_BMS.strDir & g_strBMP(i) As g_BMS.strDir & strTemp

                        ReDim Preserve strNameFrom(lngTemp)
                        ReDim Preserve strNameTo(lngTemp)

                        strNameFrom(lngTemp) = g_BMS.strDir & g_strBMP(i)
                        strNameTo(lngTemp) = g_BMS.strDir & strTemp

                        lngTemp = lngTemp + 1

                        g_strBMP(i) = strTemp

                        intTemp = 1

                    End If

                    blnBMP(i) = True

                End If

            End If

        Next i

        If Len(Trim(frmMain.txtMissBMP.Text)) Then

            'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
            If Dir(g_BMS.strDir & frmMain.txtMissBMP.Text, FileAttribute.Normal) <> vbNullString Then

                strArray = Split(frmMain.txtMissBMP.Text, ".")
                strTemp = "00." & strArray(UBound(strArray))

                'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
                If Dir(g_BMS.strDir & strTemp, FileAttribute.Normal) = vbNullString Then

                    'Name g_BMS.strDir & frmMain.txtMissBMP.Text As g_BMS.strDir & strTemp

                    ReDim Preserve strNameFrom(lngTemp)
                    ReDim Preserve strNameTo(lngTemp)

                    strNameFrom(lngTemp) = g_BMS.strDir & frmMain.txtMissBMP.Text
                    strNameTo(lngTemp) = g_BMS.strDir & strTemp

                    lngTemp = lngTemp + 1

                    frmMain.txtMissBMP.Text = strTemp

                    intTemp = 1

                End If

            End If

        End If

        If lngTemp Then

            With sh

                .hwnd = Me.Handle
                .wFunc = FO_MOVE
                .pFrom = Join(strNameFrom, Chr(0))
                .pTo = Join(strNameTo, Chr(0))
                .fFlags = FOF_SILENT Or FOF_ALLOWUNDO Or FOF_MULTIDESTFILES

            End With

            Call SHFileOperation(sh)

        End If

        If intTemp Then Call frmMain.SaveChanges()

    End Sub

    'UPGRADE_WARNING: イベント chkDeleteFile.CheckStateChanged は、フォームが初期化されたときに発生します。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"' をクリックしてください。
    Private Sub chkDeleteFile_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDeleteFile.CheckStateChanged

        If chkDeleteFile.CheckState Then

            lblExtension.Enabled = True
            txtExtension.Enabled = True
            chkFileRecycle.Enabled = True

        Else

            lblExtension.Enabled = False
            txtExtension.Enabled = False
            chkFileRecycle.Enabled = False

        End If

    End Sub

    'UPGRADE_WARNING: イベント chkListAlign.CheckStateChanged は、フォームが初期化されたときに発生します。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"' をクリックしてください。
    Private Sub chkListAlign_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkListAlign.CheckStateChanged

        If chkListAlign.CheckState Then

            chkUseOldFormat.Enabled = True
            chkSortByName.Enabled = True

        Else

            chkUseOldFormat.Enabled = False
            chkSortByName.Enabled = False

        End If

    End Sub

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click

        Call Me.Close()

    End Sub

    Private Sub cmdDecide_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDecide.Click

        If chkDeleteUnusedFile.CheckState = 0 And chkDeleteFile.CheckState = 0 And chkListAlign.CheckState = 0 And chkFileNameConvert.CheckState = 0 Then Exit Sub

        'If chkFileNameConvert.Value Then

        'If MsgBox(g_Message(Message.MSG_CONFIRM), vbYesNo + vbInformation, g_strAppTitle) = vbNo Then Exit Sub

        'End If

        Me.Enabled = False

        If chkDeleteUnusedFile.CheckState Then Call DeleteUnusedFile()

        If chkDeleteFile.CheckState Then Call DeleteFile()

        If chkListAlign.CheckState Then Call ListAlign()

        If chkFileNameConvert.CheckState Then Call FileNameConvert()

        Call frmMain.RefreshList()

        Me.Enabled = True

        Call Me.Close()

    End Sub

    'UPGRADE_WARNING: Form イベント frmWindowConvert.Activate には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
    Private Sub frmWindowConvert_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        With txtExtension

            Call .SetBounds(lblExtension.Left + lblExtension.Width + 4, .Top, Me.ClientRectangle.Width - lblExtension.Left + lblExtension.Width - 12, .Height)

        End With

        If Len(g_BMS.strDir) = 0 Then

            chkDeleteFile.Enabled = False
            chkFileNameConvert.Enabled = IIf(frmMain._mnuOptionsBase62.Checked, False, True) '62進数のファイル名はWindowsで不可

        Else

            chkDeleteFile.Enabled = True
            chkFileNameConvert.Enabled = IIf(frmMain._mnuOptionsBase62.Checked, False, True) '62進数のファイル名はWindowsで不可

        End If

        chkDeleteUnusedFile.CheckState = System.Windows.Forms.CheckState.Checked
        chkDeleteFile.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkFileRecycle.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkListAlign.CheckState = System.Windows.Forms.CheckState.Checked
        chkUseOldFormat.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkFileNameConvert.CheckState = System.Windows.Forms.CheckState.Unchecked

        Call cmdDecide.Focus()

    End Sub

    Private Sub frmWindowConvert_FormClosed(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        e.Cancel = True

        Call Me.Hide()

        Call frmMain.picMain.Focus()
    End Sub
End Class