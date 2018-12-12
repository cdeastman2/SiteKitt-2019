'==========================================================
' LANG 					: VBS.net 
' NAME 					: TechTages
' AUTHOR				: Clifford Dinel Eastman ii	
' Mail					: CDEASTM4FUSD(at)gmail.com
' VERSION 				: alpha
' DATE 					: 11/25/2018
' Description			: print out cart,system and pickup tages
'						: 
' COMMENTS 				: 


' CONTRIBUTORS			:Paul , Solty 

' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
' ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
' LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
' CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
' SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
' INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
' CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
' ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
' POSSIBILITY OF SUCH DAMAGE.
'==========================================================


Imports Microsoft.Office.Interop
Imports System.Net.Dns


Module TechTags
    Dim verbos As Boolean = True
    Dim objWord, objDoc, objSelection ' Word objects"
    Dim MySite As String ' this is for the file system and "Site profiles folder 

    REM Cart tags and labels 
    Public Sub Print_CartTags(SitePrefix, txtItem, intfirstlabel, intlastLabel) ' Vertion 1.0 Aplha
        REM  Dim objWord, objDoc, objSelection ' Word objects"

        Dim txtCart_labble
        REM   Dim intNnumber As Integer 'old delete later 
        Dim intCurrent_Hardware_Item_number As Integer

        Dim Documnets_Name As String = SitePrefix & txtItem


        objWord = CreateObject("Word.Application")
        objWord.Caption = "Information Technology Cart Tag"


        If verbos = True Then
            objWord.Visible = True '<<<<<<<<< Debug point <<<<<<<<< 
        End If

        objDoc = objWord.Documents.Add()
        objSelection = objWord.Selection



        intCurrent_Hardware_Item_number = intfirstlabel

        Do While intCurrent_Hardware_Item_number <= intlastlabel

            'this routin makes all numbers two digits 
            If intCurrent_Hardware_Item_number <= 9 Then
                txtCart_labble = SitePrefix & txtItem & "0" & intCurrent_Hardware_Item_number
            Else
                txtCart_labble = SitePrefix & txtItem & intCurrent_Hardware_Item_number
            End If
            '***********************************************

            objSelection.ParagraphFormat.Alignment = 1
            objSelection.Font.Name = "IDAHC39M Code 39 Barcode"
            objSelection.Font.Bold = False
            objSelection.Font.Size = "14"
            objSelection.TypeText("*" & txtCart_labble & "*") 'need leading and trailing "*" for code 39 

            objSelection.Font.Name = "Arial"
            objSelection.TypeText("     ")
            objSelection.Font.Name = "IDAHC39M Code 39 Barcode"
            objSelection.TypeText("*" & txtCart_labble & "*") 'need leading and trailing "*" for code 39

            objSelection.TypeParagraph()
            objSelection.TypeParagraph()
            intCurrent_Hardware_Item_number = intCurrent_Hardware_Item_number + 1

        Loop



        If verbos = False Then
            'objDoc.PrintOut()
        End If

        objDoc.SaveAs(Documnets_Name & ".doc")

        If verbos = False Then
            'objWord.Quit
        End If





    End Sub

    Public Sub Print_Computer_Lables(SitePrefix, txtItem, intfirstlabel, intlastLabel) ' Vertion 1.0 Aplha
        REM Dim objWord, objDoc, objSelection ' Word objects"

        Dim txtCart_labble
        Dim intCurrent_Hardware_Item_number As Integer
        Dim Documnets_Name = SitePrefix & txtItem

        objWord = CreateObject("Word.Application")
        objWord.Caption = "Information Technology Computer"
        If verbos = True Then
            objWord.Visible = True '<<<<<<<<< Debug point <<<<<<<<< 
        End If



        objDoc = objWord.Documents.Add()
        objSelection = objWord.Selection



        intCurrent_Hardware_Item_number = 1
        Do While intCurrent_Hardware_Item_number <= intNnumber

            'this routin makes all numbers two digits 
            If intCurrent_Hardware_Item_number <= 9 Then
                txtCart_labble = SitePrefix & txtItem & "0" & intCurrent_Hardware_Item_number
            Else
                txtCart_labble = SitePrefix & txtItem & intCurrent_Hardware_Item_number
            End If
            '***********************************************

            objSelection.ParagraphFormat.Alignment = 1
            objSelection.Font.Name = "IDAHC39M Code 39 Barcode"
            objSelection.Font.Bold = False
            objSelection.Font.Size = "14"
            objSelection.TypeText("*" & txtCart_labble & "*")

            objSelection.Font.Name = "Arial"
            objSelection.TypeText("     ")
            objSelection.Font.Name = "IDAHC39M Code 39 Barcode"
            objSelection.TypeText("*" & txtCart_labble & "*")

            objSelection.TypeParagraph()
            objSelection.TypeParagraph()
            intCurrent_Hardware_Item_number = intCurrent_Hardware_Item_number + 1

        Loop

        'objDoc.PrintOut()
        objDoc.SaveAs(Documnets_Name & ".doc")
        'objWord.Quit

    End Sub
    REM ***************************************************************************************

    REM Servise tage **********************************************************************
    Public Sub Print_ServiceTag(txtTech, txtIncident, txtSite, txtLocation, txtInvoTrackingNumber, txtNotes) ' Vertion 1.0 Aplha

        ' Word objects"
        Dim objWord
        Dim objDoc
        Dim objSelection
        Dim The_date As String = Today





        objWord = CreateObject("Word.Application")
        objWord.Caption = ("Information Technology Service Request")
        objWord.Visible = True
        objDoc = objWord.Documents.Add()
        objSelection = objWord.Selection

        REM Document headers
        objSelection.ParagraphFormat.Alignment = 1
        objSelection.Font.Name = "Times New Roman"
        objSelection.Font.Size = "24"
        objSelection.TypeText("Information Technology Service Tag")
        objSelection.TypeParagraph()
        objSelection.TypeParagraph()
        REM ************************************************************



        objSelection.ParagraphFormat.Alignment = 0
        objSelection.Font.Size = "14"
        objSelection.TypeText("Date: ")
        objSelection.Font.Bold = True
        objSelection.TypeText(The_date)

        objSelection.Font.Bold = False
        objSelection.TypeParagraph()

        objSelection.Font.Size = "14"
        objSelection.TypeText("Tech: ")
        objSelection.Font.Bold = True
        objSelection.TypeText(txtTech)
        objSelection.Font.Bold = False
        objSelection.TypeParagraph()

        objSelection.Font.Size = "14"
        objSelection.TypeText("Incident#: ")
        objSelection.Font.Name = "IDAHC39M Code 39 Barcode"
        objSelection.Font.Bold = False
        objSelection.Font.Size = "14"
        objSelection.TypeText("*" & txtIncident & "*")
        objSelection.Font.Bold = False
        objSelection.Font.Name = "Arial"
        objSelection.TypeParagraph()

        objSelection.Font.Name = "Times New Roman"
        objSelection.Font.Size = "14"
        objSelection.TypeText("Site: ")
        objSelection.Font.Bold = True
        objSelection.TypeText(txtSite)
        objSelection.Font.Bold = False
        objSelection.TypeParagraph()

        objSelection.Font.Size = "14"
        objSelection.TypeText("Location: ")
        objSelection.Font.Bold = True
        objSelection.TypeText(txtLocation)
        objSelection.Font.Bold = False

        objSelection.TypeParagraph()



        objSelection.Font.Size = "14"
        objSelection.TypeText("Serial#/DPN: ")
        objSelection.Font.Name = "IDAHC39M Code 39 Barcode"
        objSelection.Font.Bold = False
        objSelection.Font.Size = "14"
        objSelection.TypeText("*" & txtInvoTrackingNumber & "*")

        objSelection.Font.Bold = False
        objSelection.Font.Name = "Arial"
        objSelection.TypeParagraph()


        objSelection.ParagraphFormat.Alignment = 1
        objSelection.Font.Name = "Arial"
        objSelection.Font.Size = "10"
        objSelection.TypeText("Please E-Mail Any budget information to ")
        objSelection.Font.Bold = True
        objSelection.TypeText("IT.helpdesk@frensounified.org")
        objSelection.Font.Bold = False
        objSelection.TypeParagraph()


        If txtNotes <> "" Then
            objSelection.ParagraphFormat.Alignment = 0
            objSelection.Font.Bold = False
            objSelection.Font.Size = "14"
            objSelection.TypeText("========= Note =============================")
            objSelection.TypeParagraph()
            objSelection.Font.Bold = True
            objSelection.TypeText(txtNotes)
            objSelection.Font.Bold = False
            objSelection.TypeParagraph()

        End If


        'objDoc.PrintOut()
        objDoc.SaveAs(txtIncident & ".doc")
        'objWord.Quit

    End Sub

    Function Get_My_name() ' Vertion 1.0 Aplha
        Dim objSysInfo
        Dim strUser
        Dim objUser
        Dim My_NAme_IS(1)
        Get_My_name = "*Clifford* "

        Try
            objSysInfo = CreateObject("ADSystemInfo")
            strUser = objSysInfo.UserName
            objUser = GetObject("LDAP://" & strUser)
            My_NAme_IS(0) = objUser.givenName
            My_NAme_IS(1) = objUser.SN
            Get_My_name = My_NAme_IS(0)


        Catch ex As Exception
            REM  MsgBox(ex.Message)

        End Try


    End Function
    REM ***************************************************************************************

End Module
