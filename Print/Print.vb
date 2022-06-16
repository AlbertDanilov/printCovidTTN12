Imports System.Drawing
Imports System.IO
Imports DevExpress.Spreadsheet
Imports DOC

Public Class Print
    Private Shared mainPath As String = KnowFolders.GetKnowFolder(KnowFolder.Downloads)
    Private Shared docTemplateFileNamePath As String = $"{AppDomain.CurrentDomain.BaseDirectory}\Ковид.xlsx"

    Public Shared Sub PrintExcel(kodapt As Integer, av_id As Integer, docItems As DOCS)
        Dim docFile As Byte() = Nothing
        'Dim json As String = File.ReadAllText(jsonDocItems)
        'Dim pv As noPaperService_common.Entities.EcpSignData_pv = JsonConvert.DeserializeObject(Of noPaperService_common.Entities.EcpSignData_pv)(json)

        'Dim docFileNamePathExtension As String = $"{mainPath}\Отчеты\"
        Dim docFileNamePathExtension As String = $"C:\APT_TTN_TORG12\"

        'Directory.CreateDirectory(docFileNamePathExtension)
        'Directory.CreateDirectory(docFileNamePathExtension & $"{kodapt}\")
        'Directory.CreateDirectory(docFileNamePathExtension & $"{kodapt}\" & $"{av_id}\")

        'Dim di As DirectoryInfo = New DirectoryInfo(docFileNamePathExtension & $"{kodapt}\" & $"{av_id}\")

        'For Each file As FileInfo In di.GetFiles()
        '    file.Delete()
        'Next


        'Dim docFileName As String = $"Ковид {docItems.av_id} от {Date.Now:dd.MM.yyyy [HH.mm.ss.ffff]}.xlsx"
        Dim docFileName As String = $"{docItems.doc_nom.Replace("/", " ")} от {Date.Now:dd.MM.yyyy [HH.mm.ss.ffff]}.xlsx"

        docFileNamePathExtension &= $"{kodapt}\" & $"{av_id}\" & docFileName

        Using wb As New Workbook()
            wb.LoadDocument(docTemplateFileNamePath)
            Dim ws As Worksheet = wb.Worksheets(0)
            Dim rowIndexPaste As Integer = 38
            Dim rowIndexFormat As Integer = 47
            Dim rowIndexSum As Integer = 9
            Dim listPage As Integer = 1

            'Dim pageBreak As Integer = 55
            Dim pageLenght As Integer = 56

            Dim pageLenghtSum As Integer = 58
            'Dim pageLenghtRow As Integer = 79

            Dim allSumOtpBnds As Decimal = 0
            Dim allSumOtpNds As Decimal = 0
            'Dim allSumRoznNds As Decimal = 0
            'Dim allSumNdsRozn As Decimal = 0
            ''Dim ndsSumOpt As Decimal = 0
            'Dim ndsSumRozn As Decimal = 0
            Dim sumToString As New DOC.SumToString

            Dim ks As String
            Dim rs As Long

            Dim zayTypeS As String = String.Empty
            Dim osnName As String = String.Empty
            Dim prim As String = String.Empty

            wb.Unit = DevExpress.Office.DocumentUnit.Point
            wb.BeginUpdate()

            'Dim rn As Cell = "DATE1"
            Try
                'If PV.pv_work_program_id = CSKLAD.c_WORK_PROG_ROZN Then
                '    zayTypeS = "Сводная заявка № "
                '    prim = ""
                'ElseIf PV.pv_work_program_id = CSKLAD.c_WORK_PROG_RODSERT Then
                '    zayTypeS = "Заявка № "
                '    prim = PV.pv_zay_lpu
                'ElseIf PV.pv_work_program_id = CSKLAD.c_WORK_PROG_ONLS Then
                '    zayTypeS = "Заявка № "
                '    prim = ""
                'ElseIf PV.pv_work_program_id = CSKLAD.c_WORK_PROG_7NOZ Then
                '    zayTypeS = "Заявка № "
                '    prim = ""
                'ElseIf PV.pv_work_program_id = CSKLAD.c_WORK_PROG_SPEC_PROG Then
                '    zayTypeS = ""
                '    prim = PV.pv_zay_lpu
                'ElseIf PV.pv_work_program_id = CSKLAD.c_WORK_PROG_10ST Then
                '    zayTypeS = "Заявка № "
                '    If PV.pv_sklad_iname = "МЗ РФ 3" Then
                '        prim = "Гос. контракт № 12-216 от 14.08.2012 г."
                '    Else
                '        prim = PV.pv_zay_lpu
                '    End If
                'Else
                '    zayTypeS = "Заявка № "
                '    prim = PV.pv_zay_lpu
                'End If

                'If PV.pv_zay_zname IsNot String.Empty Then
                '    osnName = zayTypeS & PV.pv_zay_zname & " от " & PV.pv_zay_cdate.Value.ToString("dd.MM.yyyy")
                'Else
                '    osnName = PV.pv_reason
                'End If

                Dim k = 1
                Dim rng As CellRange
                Dim listRng As New List(Of String)

                Dim sgtinCount As Long = 0

                ws.Range("A1").Value = docItems.sender
                ws.Range("K5").Value = docItems.recipient_printname
                ws.Range("K11").Value = docItems.sender
                ws.Range("AG23").Value = docItems.doc_nom

                Dim dt_str As String = Date.Now.ToString("dd.MM.yyyy")

                ws.Range("AW23").Value = dt_str
                ws.Range("BI23").Value = dt_str
                ws.Range("X33").Value = $"Дата отгрузки: {dt_str}"

                If docItems.pv_sklad_name.ToLower.Contains("регион") Then
                    ws.Range("I23").Value = "Оплата МЗ РТ"
                ElseIf docItems.pv_sklad_name.ToLower.Contains("федерал") Then
                    ws.Range("I23").Value = "Оплата МЗ РФ"
                Else
                    ws.Range("I23").Value = "Оплата МЗ"
                End If

                For Each i As DOC_SPEC In docItems.ds_list
                    sgtinCount += i.ts_sgtin_cnt.Value

                    ws.Range($"A{rowIndexPaste}").Value = k
                    ws.Range($"D{rowIndexPaste}").Value = i.ts_shifr
                    ws.Range($"D{rowIndexPaste + 3}").Value = $"{i.ts_sert}, {i.ts_sert_date_s.Value:dd.MM.yyyy}"
                    ws.Range($"W{rowIndexPaste}").Value = $"{i.ts_p_tn} {i.ts_p_fv_doz} {i.ts_p_proizv}"
                    ws.Range($"W{rowIndexPaste + 5}").Value = $"{i.ts_seria}"
                    ws.Range($"W{rowIndexPaste + 7}").Value = $"{i.ts_sgod.Value}"
                    ws.Range($"AJ{rowIndexPaste + 5}").Value = $"{i.pvs_kol_tov.Value}"
                    ws.Range($"AJ{rowIndexPaste + 7}").Value = $"{i.ts_ed_shortname}"
                    ws.Range($"AR{rowIndexPaste + 5}").Value = $"{i.ts_temp_regim}"

                    Dim s = 0
                    s = i.ts_nds_i_val + 100
                    s /= 100

                    Dim ndsCenaOpt = i.ts_ocena_nds - i.ts_ocena_nds / s
                    Dim ndsSumOpt = i.pvs_psum_nds - i.pvs_psum_bnds / s

                    ndsCenaOpt = Decimal.Round(ndsCenaOpt, 2, MidpointRounding.AwayFromZero)
                    ndsSumOpt = Decimal.Round(ndsSumOpt, 2, MidpointRounding.AwayFromZero)

                    allSumOtpBnds += i.pvs_psum_bnds
                    allSumOtpNds += ndsSumOpt

                    ws.Range($"BG{rowIndexPaste + 5}").Value = $"{i.ts_pcena_bnds}"
                    ws.Range($"BG{rowIndexPaste + 7}").Value = $"{ndsCenaOpt}"
                    ws.Range($"BX{rowIndexPaste + 5}").Value = $"{i.pvs_psum_bnds}"
                    ws.Range($"BX{rowIndexPaste + 7}").Value = $"{ndsSumOpt}"

                    If k < docItems.ds_list.Count Then
                        Dim temprowIndexPaste = rowIndexPaste
                        temprowIndexPaste += rowIndexSum

                        If temprowIndexPaste + rowIndexSum > pageLenght Then
                            ws.Rows.Insert(rowIndexFormat, 1) 'смещаем вниз на одну позицию, чтобы добавить пустую строку
                            rowIndexFormat += 1
                            rowIndexPaste += 1
                            ws.Rows.Insert(rowIndexFormat, 3)
                            ws.Range($"A{rowIndexFormat}").CopyFrom(ws.Range("A35:CL37"), PasteSpecial.All)
                            rowIndexFormat += 3
                            rowIndexPaste += 3
                            'rowIndexPaste += rowIndexSum
                        End If

                        ws.Rows.Insert(rowIndexFormat, rowIndexSum)
                        ws.Range($"A{rowIndexFormat}").CopyFrom(ws.Range($"A38:CL46"), PasteSpecial.Formats)
                        rowIndexFormat += rowIndexSum
                        rowIndexPaste += rowIndexSum

                        k += 1

                        If rowIndexPaste + rowIndexSum > pageLenght Then
                            pageLenght += pageLenghtSum
                            ws.HorizontalPageBreaks.Add(rowIndexPaste - 5) ' разрыв страницы, если превышает определенную длину
                            listRng.Add($"CL{rowIndexPaste - 4}")
                            'ws.Range($"CL{rowIndexPaste - 5}").Value = $"ТТН № {docItems.doc_nom} лист {listPage} из {listPage}"
                        End If
                    End If
                Next

                Dim list As Short = 1
                listPage += listRng.Count

                For Each cRng As String In listRng
                    ws.Range(cRng).Font.Size = 14
                    ws.Range(cRng).Alignment.Vertical = SpreadsheetVerticalAlignment.Center
                    ws.Range(cRng).Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right
                    ws.Range(cRng).Value = $"ТТН № {docItems.doc_nom} лист {list} из {listPage}"
                    list += 1
                Next

                If allSumOtpBnds.ToString.Replace(",", ".").Contains(".") Then
                    ks = allSumOtpBnds.ToString.Replace(",", ".").Split(".")(1)
                    rs = CLng(allSumOtpBnds.ToString.Replace(",", ".").Split(".")(0))
                Else
                    ks = "0"
                    rs = CLng(allSumOtpBnds)
                End If

                If ks.Length = 1 Then
                    ks &= "0"
                End If

                ws.Range("SUM_OTP").Value = allSumOtpBnds
                Dim allOtpText As String = sumToString.sum_to_string(rs, CByte(ks))
                ws.Range("SUM_OTP_TEXT").Value = allOtpText
                ws.Range("ITOGO_OTP_TEXT").Value = allOtpText

                If allSumOtpNds.ToString.Replace(",", ".").Contains(".") Then
                    ks = allSumOtpNds.ToString.Replace(",", ".").Split(".")(1)
                    rs = CLng(allSumOtpNds.ToString.Replace(",", ".").Split(".")(0))
                Else
                    ks = "0"
                    rs = CLng(allSumOtpNds)
                End If

                If ks.Length = 1 Then
                    ks &= "0"
                End If

                ws.Range("SUM_NDS").Value = allSumOtpNds
                ws.Range("SUM_NDS_TEXT").Value = sumToString.sum_to_string(rs, CByte(ks))

                ws.Range("END_LIST").Value = $"ТТН № {docItems.doc_nom} лист {listPage} из {listPage}"

                If sgtinCount > 0I Then
                    rng = ws.Range("BY18:CL19")
                    rng.Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thin)
                    rng.Value = "Маркировка"
                End If

                rng = ws.Range("BY21:CL22")
                rng.Value = docItems.work_porgram
                rng.FillColor = Color.Gray
                rng.Font.Color = Color.White

                ws.Range("ITOGO").Value = $"ИТОГО ПО ТТН № {docItems.doc_nom}" ' ОТ {PV.pv_otr_date.Value:dd.MM.yyyy} отгр {PV.pv_otg_date.Value:dd.MM.yyyy}"

                Dim listName = sumToString.sum_to_string2(listPage)
                Dim index = listName.IndexOf("лист")

                ws.Range("COUNT_LIST").Value = listName.Substring(0, index - 1)
                ws.Range("NAME_LIST").Value = listName.Substring(index - 1)
                ws.Range("COUNT_POS").Value = sumToString.sum_to_string(k, 0, False)

                ws.Range("DATE1").Value = dt_str
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                'ws.DeleteCells(ws.Range("END_LIST") - 1, DeleteMode.EntireRow)
                ws.Rows(rowIndexFormat - 1).Delete()
                wb.EndUpdate()
            End Try

            wb.Calculate()

            wb.SaveDocument(docFileNamePathExtension, DocumentFormat.OpenXml)
        End Using
    End Sub

End Class
