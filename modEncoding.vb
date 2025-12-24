Option Strict Off
Option Explicit On

Imports System.IO
Imports System.Text

Module modEncoding

    '必要に応じて変える
    Public defaultEncoding As Encoding = Encoding.GetEncoding("SHIFT-JIS")

    Public InputEncoding As Encoding = Encoding.Default
    Public ForceReloadEncoding As Encoding = Nothing

    ''' <summary>BOMを調べて、文字コードを判別する。</summary>
    ''' <param name="bytes">文字コードを調べるByte列。</param>
    ''' <returns>BOMが見つかった時は、対応するEncodingオブジェクト。
    ''' 見つからなかった時は、Nothing。</returns>
    Private Function EncodingFromBOM(ByVal bytes As Byte()) As Encoding

        Dim bomUTF8 As Byte() = {&HEF, &HBB, &HBF}
        Dim bomUTF16LE As Byte() = {&HFF, &HFE}
        Dim bomUTF16BE As Byte() = {&HFE, &HFF}
        Dim bomUTF32LE As Byte() = {&HFF, &HFE, &H0, &H0}
        Dim bomUTF32BE As Byte() = {&H0, &H0, &HFE, &HFF}

        If bytes.Length >= 4 Then

            bytes = bytes.Take(4).ToArray()

            If Enumerable.SequenceEqual(bytes, bomUTF32BE) Then
                Return New System.Text.UTF32Encoding(True, True) 'UTF-32BE
            ElseIf Enumerable.SequenceEqual(bytes, bomUTF32LE) Then
                Return New System.Text.UTF32Encoding(False, True) 'UTF-32LE
            End If

        End If

        If bytes.Length >= 3 Then

            bytes = bytes.Take(3).ToArray()

            If Enumerable.SequenceEqual(bytes, bomUTF8) Then
                Return New UTF8Encoding(True, True) 'UTF-8
            End If

        End If

        If bytes.Length >= 2 Then

            bytes = bytes.Take(2).ToArray()

            If Enumerable.SequenceEqual(bytes, bomUTF16BE) Then
                Return New UnicodeEncoding(True, True) 'UTF-16BE
            ElseIf Enumerable.SequenceEqual(bytes, bomUTF16LE) Then
                Return New UnicodeEncoding(False, True) 'UTF-16LE
            End If

        End If

        Return Nothing

    End Function

    ''' <summary>文字化けを調べて、文字コードを判別する。</summary>
    ''' <param name="bytes">文字コードを調べるByte列。</param>
    ''' <returns>登録されている文字コードごとに文字化けをカウント、
    ''' 閾値以下なら、そのEncodingオブジェクト。
    ''' Encodingオブジェクトが見つからなかったら、Nothing。</returns>
    Private Function EncodingFromGarbled(ByRef bytes As Byte()) As Encoding

        Dim EncodingsToTry As New List(Of Encoding) From {
            New UnicodeEncoding(True, True), 'UTF-16BE
            New UnicodeEncoding(False, True), 'UTF-16LE
            New UTF32Encoding(True, True), 'UTF-32BE
            New UTF32Encoding(False, True) 'UTF-32LE
            }

        For Each encoding As Encoding In EncodingsToTry
            Try
                Dim decodedString As String = encoding.GetString(bytes)

                If Not (decodedString.Contains(vbCrLf) OrElse decodedString.Contains(vbCr) OrElse decodedString.Contains(vbLf) OrElse decodedString.Contains("#")) Then
                    Continue For ' BMSファイルに改行コードと"#"がないことはありえない
                End If

                If decodedString IsNot Nothing AndAlso Len(decodedString) > 1 Then
                    decodedString = Left(decodedString, LenB(decodedString) - 1) '不完全かもしれない最後の文字をカット
                Else
                    Continue For
                End If

                Dim newbytes As Byte() = bytes.Take(encoding.GetByteCount(decodedString)).ToArray() ' 文字列のバイト数で比較する

                '文字列を再度バイナリ化して比較、一致すれば文字化けが存在しないため正しいエンコーディングと推測できる。
                If Enumerable.SequenceEqual(newbytes, encoding.GetBytes(decodedString)) Then
                    Return encoding
                End If
            Catch ex As Exception
                ' その他のエラー
            End Try
        Next

        Return Nothing

    End Function

    ''' <summary>
    ''' 指定されたバイト配列がUTF-8のエンコーディングルールに違反していないかを判定。
    ''' </summary>
    ''' <param name="bytes">判定するバイト配列。</param>
    ''' <returns>違反がない場合は True、違反がある場合は False。</returns>
    Private Function IsValidUTF8(ByRef bytes As Byte(), Optional size As Integer = 1024 * 64) As Boolean
        If bytes.Length < 7 Then Return False ' コードの簡略化のため7バイト未満の場合は検証しない（4バイトでBMSは成立しない:ValidなTextかもしれないがBMSではない）

        Dim i As Integer = 3 'BOMかもしれない先頭3バイトをスキップ
        While i <= Math.Min(UBound(bytes), size - 1) - 3　'終端判定の省略のため サイズ - 3 までを検証
            If (bytes(i) And &B10000000) = &B0 Then
                '1 byte patturn (ASCII)
                i += 1
            ElseIf (bytes(i) And &B11100000) = &B11000000 Then
                If (bytes(i + 1) And &B11000000) = &B10000000 Then
                    '2 byte patturn
                    i += 2
                Else
                    Return False
                End If
            ElseIf (bytes(i) And &B11110000) = &B11100000 Then
                If (bytes(i + 1) And &B11000000) = &B10000000 Then
                    If (bytes(i + 2) And &B11000000) = &B10000000 Then
                        '3 Bytes patturn
                        i += 3
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            ElseIf (bytes(i) And &B11111000) = &B11110000 Then
                If (bytes(i + 1) And &B11000000) = &B10000000 Then
                    If (bytes(i + 2) And &B11000000) = &B10000000 Then
                        If (bytes(i + 3) And &B11000000) = &B10000000 Then
                            '4 Bytes patturn
                            i += 4
                        Else
                            Return False
                        End If
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End While

        Return True

    End Function

    ''' <summary>
    ''' 指定されたバイト配列がUTF-16BEのエンコーディングルールに違反していないかを判定。
    ''' </summary>
    ''' <param name="bytes">判定するバイト配列。</param>
    ''' <returns>違反がない場合は True、違反がある場合は False。</returns>
    Private Function IsValidUTF16be(ByRef bytes As Byte(), Optional size As Integer = 1024 * 64) As Boolean
        If bytes.Length < 6 Then Return False 'コードの簡略化のため6バイト未満の場合は検証しない（6バイトでBMSは成立しない:ValidなTextかもしれないがBMSではない）
        If bytes(3) = &H0 OrElse bytes(5) = &H0 Then Return False　'偶数バイトが0x00ならUTF-16LEの可能性が高い。

        Dim i As Integer = 2 'BOMかもしれない先頭2バイトをスキップ
        While i <= Math.Min(UBound(bytes), size - 1) - 3　'終端判定の省略のため サイズ - 3 までを検証
            If (bytes(i) And &B11111000) <> &B11011000 Then
                '2 byte patturn
                i += 2
            ElseIf (bytes(i) And &B11111100) = &B11011000 Then '上位サロゲートペア 0xD800から0xDBFF
                If (bytes(i + 2) And &B11111100) = &B11011100 Then '下位サロゲートペア 0xDC00から0xDFFF
                    i += 4
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End While

        Return True

    End Function

    ''' <summary>
    ''' 指定されたバイト配列がUTF-16LEのエンコーディングルールに違反していないかを判定。
    ''' </summary>
    ''' <param name="bytes">判定するバイト配列。</param>
    ''' <returns>違反がない場合は True、違反がある場合は False。</returns>
    Private Function IsValidUTF16le(ByRef bytes As Byte(), Optional size As Integer = 1024 * 64) As Boolean
        If bytes.Length < 6 Then Return False 'コードの簡略化のため6バイト未満の場合は検証しない（4バイトでBMSは成立しない:ValidなTextかもしれないがBMSではない）
        If bytes(2) = &H0 OrElse bytes(4) = &H0 Then Return False '奇数バイトが0x00ならUTF-16BEの可能性が高い。

        Dim i As Integer = 2 'BOMかもしれない先頭2バイトをスキップ
        While i + 1 <= Math.Min(UBound(bytes), size - 1) - 3　'終端判定の省略のため サイズ - 3 までを検証
            If (bytes(i + 1) And &B11111000) <> &B11011000 Then
                '2 byte patturn
                i += 2
            ElseIf (bytes(i + 1) And &B11111100) = &B11011000 Then
                '上位サロゲートペア 0xD800から0xDBFF
                If (bytes(i + 3) And &B11111100) = &B11011100 Then
                    '下位サロゲートペア
                    i += 4
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End While

        Return True

    End Function

    ''' <summary>
    ''' 指定されたバイト配列がSHIFT-JISのエンコーディングルールに違反していないかを判定。
    ''' </summary>
    ''' <param name="bytes">判定するバイト配列。</param>
    ''' <returns>違反がない場合は True、違反がある場合は False。</returns>
    Private Function IsValidShiftJIS(ByRef bytes() As Byte, Optional size As Integer = 64) As Boolean
        If bytes.Length < 2 Then Return False 'コードの簡略化のため2バイト未満の場合は検証しない（2バイトでBMSは成立しない::ValidなTextかもしれないがBMSではない）

        Dim i As Integer = 0
        While i <= Math.Min(UBound(bytes), size - 1) - 1　'終端判定の省略のため サイズ - 1 までを検証
            If (bytes(i) <= &H7F) OrElse (bytes(i) >= &HA1 AndAlso bytes(i) <= &HDF) Then '1バイト文字の範囲判定 ASCII(<=0x7F) または 半角カナ(0xA1から0xDF)
                i += 1
                Continue While
            ElseIf (bytes(i) >= &H81 AndAlso bytes(i) <= &H9F) OrElse (bytes(i) >= &HE0 AndAlso bytes(i) <= &HEF) Then
                If (bytes(i + 1) >= &H40 AndAlso bytes(i + 1) <= &H7E) OrElse (bytes(i + 1) >= &H80 AndAlso bytes(i + 1) <= &HFC) Then '2バイト目の範囲判定
                    '有効な2バイトシーケンス
                    i += 2
                    Continue While
                Else
                    '2バイト目が範囲外 -> Shift-JISのルール違反
                    Return False
                End If
            Else
                '上記のどの範囲にも含まれないバイト値が見つかった場合
                'Shift-JISの有効な開始バイトではないため、ルール違反とする
                Return False
            End If
        End While

        '全バイトがルールに違反しなかった場合、Shift-JISである可能性が高い
        Return True

    End Function

    ''' <summary>
    ''' 指定されたバイト配列がEUC-KRのエンコーディングルールに違反していないかを判定。
    ''' </summary>
    ''' <param name="bytes">判定するバイト配列。</param>
    ''' <returns>ルール違反がない場合は True、違反がある場合は False。</returns>
    Private Function IsValidEUCKR(ByRef bytes() As Byte, Optional size As Integer = 1024 * 64) As Boolean
        If bytes.Length < 2 Then Return False 'コードの簡略化のため2バイト未満の場合は検証しない（2バイトでBMSは成立しない:ValidなTextかもしれないがBMSではない）
        Dim i As Integer = 0

        While i <= Math.Min(UBound(bytes), size - 1) - 1 '終端判定の省略のため サイズ - 1 までを検証
            If bytes(i) <= &H7F Then ' 1バイト文字の範囲 (ASCII: 0x00 - 0x7F)
                i += 1
                Continue While
            ElseIf bytes(i) >= &H81 AndAlso bytes(i) <= &HFE Then ' 2バイト文字の可能性がある範囲 (1バイト目: 0x81 から 0xFE)
                ' 2バイト目の範囲チェック
                If bytes(i + 1) >= &H81 AndAlso bytes(i + 1) <= &HFE Then ' EUC-KR/CP949の2バイト目は 0x81 から 0xFE の範囲内である必要がある
                    '有効な2バイトシーケンス
                    i += 2
                    Continue While
                Else
                    '2バイト目が範囲外 -> EUC-KRのルール違反
                    Return False
                End If
            Else
                '上記のどの範囲にも含まれないバイト値が見つかった場合
                Return False
            End If
        End While

        ' 全バイトがルールに違反しなかった場合、EUC-KRである可能性が高い
        Return True

    End Function

    ''' <summary>
    ''' Byte列が各エンコーディングルールに違反していないを調べて、エンコーディングを判別する。
    ''' </summary>
    ''' <param name="bytes">文字コードを調べるByte列。</param>
    ''' <returns>エンコーディングルールに違反していなければ、そのEncodingオブジェクト。
    ''' Encodingオブジェクトが見つからなかったら、Nothing。</returns>
    Private Function EncodingFromValidation(ByRef bytes As Byte(), Optional size As Integer = 1024 * 64) As Encoding
        If IsValidShiftJIS(bytes, size) Then
            Return Encoding.GetEncoding("SHIFT-JIS")
        ElseIf IsValidEUCKR(bytes, size) Then
            Return Encoding.GetEncoding("EUC-KR")
        ElseIf IsValidUTF8(bytes, size) Then
            Return New UTF8Encoding(True, False) ' UTF-8
        ElseIf IsValidUTF16be(bytes, size) Then
            Return New UnicodeEncoding(True, True) 'UTF-16BE
        ElseIf IsValidUTF16le(bytes, size) Then
            Return New UnicodeEncoding(False, True) 'UTF-16LE
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' 指定されたパスのファイルのエンコーディングを自動判定する。
    ''' </summary>
    ''' <param name="path">判定するファイルのフルパス。</param>
    ''' <param name="size">取得するサイズ。デフォルトは4KB</param>
    ''' <returns>判定できたならばそのEncoding、見つからなかったらデフォルト値。</returns>
    Public Function DetectEncoding(path As String, Optional size As Integer = 1024 * 64) As Encoding
        Dim bytes() As Byte
        DetectEncoding = Nothing

        If FileIO.FileSystem.FileExists(path) Then
            Using stream As New FileStream(path, FileMode.Open, FileAccess.Read)
                Dim binary As New BinaryReader(stream)
                bytes = binary.ReadBytes(size) '64KB as the default
            End Using

            DetectEncoding = EncodingFromBOM(bytes)
            If DetectEncoding IsNot Nothing Then
                Return DetectEncoding
            End If

            DetectEncoding = EncodingFromValidation(bytes, size)
            If DetectEncoding IsNot Nothing Then
                Return DetectEncoding
            End If

            DetectEncoding = EncodingFromGarbled(bytes)
            If DetectEncoding IsNot Nothing Then
                Return DetectEncoding
            End If

            Return defaultEncoding
        End If
    End Function

End Module
