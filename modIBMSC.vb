Option Strict Off
Option Explicit On
Module modIBMSC

    ' iBMSC Column Index
    Private Enum ColumnIndex
        niMeasure = 0
        niSCROLL = 1
        niBPM = 2
        niSTOP = 3
        niS1 = 4

        niA1 = 5
        niA2 = 6
        niA3 = 7
        niA4 = 8
        niA5 = 9
        niA6 = 10
        niA7 = 11
        niA8 = 12
        niA9 = 13
        niAA = 14
        niAB = 15
        niAC = 16
        niAD = 17
        niAE = 18
        niAF = 19
        niAG = 20
        niAH = 21
        niAI = 22
        niAJ = 23
        niAK = 24
        niAL = 25
        niAM = 26
        niAN = 27
        niAO = 28
        niAP = 29
        niAQ = 30
        niS2 = 31

        niD1 = 32
        niD2 = 33
        niD3 = 34
        niD4 = 35
        niD5 = 36
        niD6 = 37
        niD7 = 38
        niD8 = 39
        niD9 = 40
        niDA = 41
        niDB = 42
        niDC = 43
        niDD = 44
        niDE = 45
        niDF = 46
        niDG = 47
        niDH = 48
        niDI = 49
        niDJ = 50
        niDK = 51
        niDL = 52
        niDM = 53
        niDN = 54
        niDO = 55
        niDP = 56
        niDQ = 57
        niS3 = 58

        niBGA = 59
        niLAYER = 60
        niPOOR = 61
        niS4 = 62
        niB = 63
    End Enum

    Public Function IBMSCColumnIndexToChannel(ByVal Index As Integer) As Integer

        Dim Result As Integer = 0

        If frmMain.cboDispSC1P.SelectedIndex = 0 Then '1P側はスクラッチの位置で区別する必要がある；左SCの場合

            Select Case Index

                Case ColumnIndex.niBPM : Return OBJ_CH.CH_EXBPM
                Case ColumnIndex.niSTOP : Return OBJ_CH.CH_STOP
                Case ColumnIndex.niSCROLL : Return OBJ_CH.CH_SCROLL

                Case ColumnIndex.niA1 : Return OBJ_CH.CH_1P_SC
                Case ColumnIndex.niA3 : Return OBJ_CH.CH_1P_KEY1
                Case ColumnIndex.niA4 : Return OBJ_CH.CH_1P_KEY2
                Case ColumnIndex.niA5 : Return OBJ_CH.CH_1P_KEY3
                Case ColumnIndex.niA6 : Return OBJ_CH.CH_1P_KEY4
                Case ColumnIndex.niA7 : Return OBJ_CH.CH_1P_KEY5
                Case ColumnIndex.niA8 : Return OBJ_CH.CH_1P_KEY6
                Case ColumnIndex.niA9 : Return OBJ_CH.CH_1P_KEY7

                Case ColumnIndex.niD1 : Return OBJ_CH.CH_2P_KEY1
                Case ColumnIndex.niD2 : Return OBJ_CH.CH_2P_KEY2
                Case ColumnIndex.niD3 : Return OBJ_CH.CH_2P_KEY3
                Case ColumnIndex.niD4 : Return OBJ_CH.CH_2P_KEY4
                Case ColumnIndex.niD5 : Return OBJ_CH.CH_2P_KEY5
                Case ColumnIndex.niD6 : Return OBJ_CH.CH_2P_KEY6
                Case ColumnIndex.niD7 : Return OBJ_CH.CH_2P_KEY7
                Case ColumnIndex.niDP : Return OBJ_CH.CH_2P_SC

                Case ColumnIndex.niBGA : Return OBJ_CH.CH_BGA
                Case ColumnIndex.niLAYER : Return OBJ_CH.CH_LAYER
                Case ColumnIndex.niPOOR : Return OBJ_CH.CH_POOR

                Case Is >= ColumnIndex.niB : Return 36 ^ 2 + 1 + (Index - ColumnIndex.niB)

            End Select

        Else '右SCの場合

            Select Case Index

                Case ColumnIndex.niBPM : Return OBJ_CH.CH_EXBPM
                Case ColumnIndex.niSTOP : Return OBJ_CH.CH_STOP
                Case ColumnIndex.niSCROLL : Return OBJ_CH.CH_SCROLL

                Case ColumnIndex.niA1 : Return OBJ_CH.CH_1P_KEY1
                Case ColumnIndex.niA2 : Return OBJ_CH.CH_1P_KEY2
                Case ColumnIndex.niA3 : Return OBJ_CH.CH_1P_KEY3
                Case ColumnIndex.niA4 : Return OBJ_CH.CH_1P_KEY4
                Case ColumnIndex.niA5 : Return OBJ_CH.CH_1P_KEY5
                Case ColumnIndex.niA6 : Return OBJ_CH.CH_1P_KEY6
                Case ColumnIndex.niA7 : Return OBJ_CH.CH_1P_KEY7
                Case ColumnIndex.niA8 : Return OBJ_CH.CH_1P_SC

                Case ColumnIndex.niD1 : Return OBJ_CH.CH_2P_KEY1
                Case ColumnIndex.niD2 : Return OBJ_CH.CH_2P_KEY2
                Case ColumnIndex.niD3 : Return OBJ_CH.CH_2P_KEY3
                Case ColumnIndex.niD4 : Return OBJ_CH.CH_2P_KEY4
                Case ColumnIndex.niD5 : Return OBJ_CH.CH_2P_KEY5
                Case ColumnIndex.niD6 : Return OBJ_CH.CH_2P_KEY6
                Case ColumnIndex.niD7 : Return OBJ_CH.CH_2P_KEY7
                Case ColumnIndex.niDP : Return OBJ_CH.CH_2P_SC

                Case ColumnIndex.niBGA : Return OBJ_CH.CH_BGA
                Case ColumnIndex.niLAYER : Return OBJ_CH.CH_LAYER
                Case ColumnIndex.niPOOR : Return OBJ_CH.CH_POOR

                Case Is >= ColumnIndex.niB : Return 36 ^ 2 + 1 + (Index - ColumnIndex.niB)

            End Select

        End If

        Return Result

    End Function

End Module
