Imports iTextSharp.text.pdf
Imports System.Data.OleDb
Imports System.IO
Imports System

Public Class clsBorangCP5012016
    Private Const pdfSubFormName = "topmostSubform[0]."

    Dim pdfForm As New clsPDFMaker
    Dim pdfFormFields As AcroFields
    Dim datHandler As New clsDataHandler("")

#Region "CStor"

    Public Sub New()

        datHandler = New clsDataHandler(pdfForm.GetFormType)
        pdfFormFields = pdfForm.GetStamper.AcroFields
        Page1()
        pdfForm.OpenFile()
        pdfForm.CloseStamper()
    End Sub
#End Region

    Private Sub Page1()

        Dim pdfFieldPath As String = ""
        Dim strTempString As String = ""
        Dim dr As OleDbDataReader = Nothing


        Try
            ' === Part Slip === '
            pdfFieldPath = pdfSubFormName & "Page1[0]."
            dr = datHandler.GetDataReader("Select * from taxp_profile where (tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read() Then
                If Not IsDBNull(dr("tp_ref_no_prefix")) And Not IsDBNull(dr("tp_ref_no1")) And Not IsDBNull(dr("tp_ref_no2")) And Not IsDBNull(dr("tp_ref_no3")) Then
                    If Not String.IsNullOrEmpty(dr("tp_ref_no_prefix").ToString & dr("tp_ref_no1").ToString & dr("tp_ref_no2").ToString & dr("tp_ref_no3").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Slip1", dr("tp_ref_no_prefix") & dr("tp_ref_no1") & dr("tp_ref_no2") & dr("tp_ref_no3"))
                    End If
                End If

                'danny ------------------------------------------
                If Not IsDBNull(dr("TP_CURR_ADD_LINE1")) Then
                    If Not IsDBNull(dr("TP_CURR_ADD_LINE2")) Then
                        If Not IsDBNull(dr("TP_CURR_ADD_LINE3")) Then
                            pdfFormFields.SetField(pdfFieldPath & "Address1", dr("TP_CURR_ADD_LINE1").ToString)
                            pdfFormFields.SetField(pdfFieldPath & "Address2", dr("TP_CURR_ADD_LINE2").ToString)
                            pdfFormFields.SetField(pdfFieldPath & "Address3", dr("TP_CURR_ADD_LINE3").ToString)
                            pdfFormFields.SetField(pdfFieldPath & "Address4", dr("TP_CURR_POSTCODE").ToString + dr("TP_CURR_CITY").ToString)
                            pdfFormFields.SetField(pdfFieldPath & "Address5", dr("TP_CURR_STATE").ToString)
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "Address1", dr("TP_CURR_ADD_LINE1").ToString)
                            pdfFormFields.SetField(pdfFieldPath & "Address2", dr("TP_CURR_ADD_LINE2").ToString)
                            pdfFormFields.SetField(pdfFieldPath & "Address3", dr("TP_CURR_POSTCODE").ToString + dr("TP_CURR_CITY").ToString)
                            pdfFormFields.SetField(pdfFieldPath & "Address4", dr("TP_CURR_STATE").ToString)
                            pdfFormFields.SetField(pdfFieldPath & "Address5", "")
                        End If
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "Address1", dr("TP_CURR_ADD_LINE1").ToString)
                        pdfFormFields.SetField(pdfFieldPath & "Address2", dr("TP_CURR_POSTCODE").ToString + dr("TP_CURR_CITY").ToString)
                        pdfFormFields.SetField(pdfFieldPath & "Address3", dr("TP_CURR_STATE").ToString)
                        pdfFormFields.SetField(pdfFieldPath & "Address4", "")
                        pdfFormFields.SetField(pdfFieldPath & "Address5", "")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "Address1", "")
                    pdfFormFields.SetField(pdfFieldPath & "Address2", "")
                    pdfFormFields.SetField(pdfFieldPath & "Address3", "")
                    pdfFormFields.SetField(pdfFieldPath & "Address4", "")
                    pdfFormFields.SetField(pdfFieldPath & "Address5", "")

                End If
                'end ---------------------------------------------
                If Not IsDBNull(dr("TP_NAME")) Then
                    If Not String.IsNullOrEmpty(dr("TP_NAME").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Name", dr("TP_NAME").ToString)
                    End If
                End If

                'If Not IsDBNull(dr("TP_CURR_ADD_LINE1")) Then
                '    If Not String.IsNullOrEmpty(dr("TP_CURR_ADD_LINE1").ToString) Then
                '        strTempString = strTempString + dr("TP_CURR_ADD_LINE1").ToString
                '    End If
                'End If

                'If Not IsDBNull(dr("TP_CURR_ADD_LINE2")) Then
                '    If Not String.IsNullOrEmpty(dr("TP_CURR_ADD_LINE2").ToString) Then
                '        If Right(Trim(dr("TP_CURR_ADD_LINE2")), 1) = "," Then
                '            strTempString = strTempString + " " + dr("TP_CURR_ADD_LINE2").ToString
                '        Else
                '            strTempString = strTempString + ", " + dr("TP_CURR_ADD_LINE2").ToString
                '        End If
                '    End If
                'End If

                'If Not IsDBNull(dr("TP_CURR_ADD_LINE3")) Then
                '    If Not String.IsNullOrEmpty(dr("TP_CURR_ADD_LINE3").ToString) Then
                '        If Right(Trim(dr("TP_CURR_ADD_LINE3")), 1) = "," Then
                '            strTempString = strTempString + " " + dr("TP_CURR_ADD_LINE3").ToString
                '        Else
                '            strTempString = strTempString + ", " + dr("TP_CURR_ADD_LINE3").ToString
                '        End If
                '    End If
                'End If
                'If Not IsDBNull(dr("TP_CURR_POSTCODE")) Then
                '    If Not String.IsNullOrEmpty(dr("TP_CURR_POSTCODE").ToString) Then
                '        strTempString = strTempString + Environment.NewLine + dr("TP_CURR_POSTCODE").ToString
                '    End If
                'End If
                'If Not IsDBNull(dr("TP_CURR_CITY")) Then
                '    If Not String.IsNullOrEmpty(dr("TP_CURR_CITY").ToString) Then
                '        strTempString = strTempString + " " + dr("TP_CURR_CITY").ToString + Environment.NewLine
                '    End If
                'End If
                'If Not IsDBNull(dr("TP_CURR_STATE")) Then
                '    If Not String.IsNullOrEmpty(dr("TP_CURR_STATE").ToString) Then
                '        strTempString = strTempString + dr("TP_CURR_STATE").ToString
                '    End If
                'End If
                'If Not String.IsNullOrEmpty(strTempString) Then
                '    pdfFormFields.SetField(pdfFieldPath & "Slip3", strTempString.ToString.ToUpper)
                'End If

                strTempString = ""
                If Not IsDBNull(dr("TP_IC_NEW_1")) Then
                    If Not String.IsNullOrEmpty(dr("TP_IC_NEW_1").ToString) Then
                        strTempString = strTempString + dr("TP_IC_NEW_1").ToString
                    End If
                End If
                If Not IsDBNull(dr("TP_IC_NEW_2")) Then
                    If Not String.IsNullOrEmpty(dr("TP_IC_NEW_2").ToString) Then
                        strTempString = strTempString + dr("TP_IC_NEW_2").ToString
                    End If
                End If
                If Not IsDBNull(dr("TP_IC_NEW_3")) Then
                    If Not String.IsNullOrEmpty(dr("TP_IC_NEW_3").ToString) Then
                        strTempString = strTempString + dr("TP_IC_NEW_3").ToString
                    End If
                End If
                If Not String.IsNullOrEmpty(strTempString) Then
                    pdfFormFields.SetField(pdfFieldPath & "IC", strTempString)
                End If
            End If
            dr.Close()

            dr = datHandler.GetDataReader("Select TC_BALANCE_TAX_PAYABLE from tax_computation where tc_ref_no= '" & pdfForm.GetRefNo & "' and tc_ya= '" & pdfForm.GetYA & "'")
            If dr.Read() Then
                If Not IsDBNull("TC_BALANCE_TAX_PAYABLE") Then
                    pdfFormFields.SetField(pdfFieldPath & "Amount", FormatFloatingAmount(dr("TC_BALANCE_TAX_PAYABLE").ToString, True))
                End If
            End If
            dr.Close()

            pdfFormFields.SetField(pdfFieldPath & "Slip5", "")
            pdfFormFields.SetField(pdfFieldPath & "Slip6", "")
            pdfFormFields.SetField(pdfFieldPath & "Slip7", "")
            pdfFormFields.SetField(pdfFieldPath & "Slip8", "")


            'Nama dan Ruj
            dr = datHandler.GetDataReader("select tp_name from taxp_profile where " _
                        & "(tp_ref_no1 + tp_ref_no2 + tp_ref_no3)= '" & pdfForm.GetRefNo & "'")
            If dr.Read Then
                If Not IsDBNull(dr("tp_name")) Then
                    If Not String.IsNullOrEmpty(dr("tp_name").ToString) Then
                        pdfFormFields.SetField(pdfFieldPath & "Nama11", dr("tp_name").ToString.ToUpper)
                    End If
                End If
                If Not String.IsNullOrEmpty(pdfForm.GetRefNo) Then
                    pdfFormFields.SetField(pdfFieldPath & "Ruj11", pdfForm.GetRefNo)
                End If
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub

    Protected Function FormatFloatingAmount(ByVal strTemp As String, ByVal intFloating As Boolean) As String

        If intFloating = True Then
            If Not strTemp = "" Then
                If CDbl(strTemp) > 0 Then
                    strTemp = strTemp.ToString.Replace(",", "").Replace(".", "")
                Else
                    strTemp = "000"
                End If
            End If
        Else
            If Not strTemp = "" Then
                If CDbl(strTemp) > 0 Then
                    strTemp = Math.Ceiling(CDbl(strTemp)).ToString.Replace(",", "")
                Else
                    strTemp = "0"
                End If
            End If
        End If
        Return strTemp

    End Function

End Class
