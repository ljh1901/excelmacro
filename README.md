'excelmacro'

Option Explicit
Sub SplitLegalData()

    '====================================================
    ' [1] 변수 선언
    '====================================================
    Dim ws As Worksheet          ' 원본 시트
    Dim outWs As Worksheet       ' 결과 시트

    Dim lastRow As Long          ' 원본 마지막 행
    Dim outRow As Long           ' 결과 시트 출력 행

    Dim i As Long                ' 원본 행 반복용
    Dim j As Long                ' 셀 내부 줄 반복용

    Dim seqNo As String          ' 순번
    Dim instCd As String         ' 기관코드
    Dim instNm As String         ' 기관명
    Dim executive As String      ' 임원
    Dim duty As String           ' 책무

    Dim raw As String            ' 법령및내규명 원본 문자열
    Dim txt As String            ' 줄 단위 문자열

    Dim arr As Variant           ' Split 결과
    Dim currentType As String    ' 현재 코드값 (01: 법령, 02: 내규)

    '====================================================
    ' [2] 원본 시트 지정
    '====================================================
    Set ws = ActiveSheet

    '====================================================
    ' [3] 기존 RESULT 시트 삭제
    '====================================================
    Application.DisplayAlerts = False

    On Error Resume Next
    Worksheets("RESULT").Delete
    On Error GoTo 0

    Application.DisplayAlerts = True

    '====================================================
    ' [4] 결과 시트 생성
    '====================================================
    Set outWs = Worksheets.Add
    outWs.Name = "RESULT"

    '====================================================
    ' [5] 헤더 생성
    '====================================================
    outWs.Range("A2:H2").Value = Array( _
        "순번", _
        "기관코드", _
        "기관명", _
        "임원", _
        "책무", _
        "법령및내규명", _
        "법령구분", _
        "법령코드")

    outRow = 2

    '====================================================
    ' [6] 마지막 행 조회(F열 기준)
    '====================================================
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row

    '====================================================
    ' [7] 원본 행 반복
    '====================================================
    For i = 2 To lastRow

        seqNo = Trim(CStr(ws.Cells(i, 1).Value))
        instCd = Trim(CStr(ws.Cells(i, 2).Value))
        instNm = Trim(CStr(ws.Cells(i, 3).Value))
        executive = Trim(CStr(ws.Cells(i, 4).Value))
        duty = Trim(CStr(ws.Cells(i, 5).Value))
        raw = CStr(ws.Cells(i, 6).Value)

        If raw <> "" Then

            '================================================
            ' [8] 개행 문자 통일
            '================================================
            raw = Replace(raw, vbCrLf, vbLf)
            raw = Replace(raw, vbCr, vbLf)

            ' 줄 단위로 분리
            arr = Split(raw, vbLf)

            ' 현재 상태 초기화
            currentType = ""

            '================================================
            ' [9] 셀 내부 줄 반복
            '================================================
            For j = 0 To UBound(arr)
                txt = Trim(arr(j))
                    ' 엔터 제거
                txt = Replace(txt, Chr(10), " ")
                txt = Replace(txt, Chr(13), " ")
                If txt <> "" Then

                    '============================================
                    ' [10] 특수문자 제거
                    '============================================
                    txt = Replace(txt, "「", "")
                    txt = Replace(txt, "」", "")
                    txt = Replace(txt, "[", "")
                    txt = Replace(txt, "]", "")
                    txt = Replace(txt, "<", "")
                    txt = Replace(txt, ">", "")

                    txt = Trim(txt)

                    '============================================
                    ' [11] 관련법령 / 관련내규 구간 판별
                    '============================================
                    Select Case txt

                        Case "관련법령"

                            currentType = "01"

                        Case "관련내규"

                            currentType = "02"

                        Case Else

                            '====================================
                            ' [12] 결과 데이터 생성
                            '====================================
                            outWs.Cells(outRow, 1).Value = seqNo
                            outWs.Cells(outRow, 2).Value = instCd
                            outWs.Cells(outRow, 3).Value = instNm
                            outWs.Cells(outRow, 4).Value = executive
                            outWs.Cells(outRow, 5).Value = duty
                            outWs.Cells(outRow, 6).Value = txt

                            '====================================
                            ' [13] 구분 및 코드 설정
                            '====================================
                            If currentType = "" Then
                                outWs.Cells(outRow, 7).Value = "99"
                                outWs.Cells(outRow, 8).Value = "99"
                            Else
                                outWs.Cells(outRow, 7).Value = currentType
                                outWs.Cells(outRow, 8).Value = currentType
                            End If
                            
                            ' 다음 행으로 이동
                            outRow = outRow + 1
                    End Select
                End If
            Next j
        End If
    Next i
End Sub

