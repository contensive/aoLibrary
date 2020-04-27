Option Explicit On
Option Strict On

Imports Contensive.BaseClasses

Namespace Controllers
    Public NotInheritable Class genericController
        Private Sub New()
        End Sub
        '
        '====================================================================================================
        ''' <summary>
        ''' if date is invalid, set to minValue
        ''' </summary>
        ''' <param name="srcDate"></param>
        ''' <returns></returns>
        Public Shared Function encodeMinDate(srcDate As DateTime) As DateTime
            Dim returnDate As DateTime = srcDate
            If srcDate < New DateTime(1900, 1, 1) Then
                returnDate = DateTime.MinValue
            End If
            Return returnDate
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' if valid date, return the short date, else return blank string 
        ''' </summary>
        ''' <param name="srcDate"></param>
        ''' <returns></returns>
        Public Shared Function getShortDateString(srcDate As DateTime) As String
            Dim returnString As String = ""
            Dim workingDate As DateTime = encodeMinDate(srcDate)
            If Not isDateEmpty(srcDate) Then
                returnString = workingDate.ToShortDateString()
            End If
            Return returnString
        End Function
        '
        '====================================================================================================
        Public Shared Function isDateEmpty(srcDate As DateTime) As Boolean
            Return (srcDate < New DateTime(1900, 1, 1))
        End Function
        '
        '====================================================================================================
        Public Shared Function getSortOrderFromInteger(id As Integer) As String
            Return id.ToString().PadLeft(7, "0"c)
        End Function
        '
        '====================================================================================================
        Public Shared Function getDateForHtmlInput(source As DateTime) As String
            If isDateEmpty(source) Then
                Return ""
            Else
                Return source.Year.ToString() + "-" + source.Month.ToString().PadLeft(2, "0"c) + "-" + source.Day.ToString().PadLeft(2, "0"c)
            End If
        End Function
        '
        '====================================================================================================
        Public Shared Function convertToDosPath(sourcePath As String) As String
            Return sourcePath.Replace("/", "\")
        End Function
        '
        '====================================================================================================
        Public Shared Function convertToUnixPath(sourcePath As String) As String
            Return sourcePath.Replace("\", "/")
        End Function
        '
        Public Shared Function Main_getAddonOption(cp As CPBaseClass, requestName As String, ignore As String) As String
            Return cp.Doc.GetText(requestName)
        End Function
        '
        Public Shared Function getBoolean_Main_getAddonOption(cp As CPBaseClass, requestName As String, ignore As String) As Boolean
            Return cp.Doc.GetBoolean(requestName)
        End Function
        '
        Public Shared Sub Main_testpoint(cp As CPBaseClass, message As String)
            cp.Site.TestPoint(message)
        End Sub
        '
        Public Shared Function KmaEncodeSQLNumber(cp As CPBaseClass, src As Integer) As String
            Return cp.Db.EncodeSQLNumber(src)
        End Function
        '
        Public Shared Function KmaEncodeSQLNumber(cp As CPBaseClass, src As Double) As String
            Return cp.Db.EncodeSQLNumber(src)
        End Function
        Public Shared Function KmaEncodeSQLText(cp As CPBaseClass, src As String) As String
            Return cp.Db.EncodeSQLText(src)
        End Function
        '
        Public Shared Function htmlDiv(innerHtml As String, Optional htmlClass As String = "", Optional htmlId As String = "") As String
            Dim result As String = "<div "
            result += If(String.IsNullOrEmpty(htmlClass), "", " class=""" & htmlClass & """")
            result += If(String.IsNullOrEmpty(htmlId), "", " id=""" & htmlId & """")
            Return result & ">" & innerHtml & "</div>"
        End Function
        '
        Public Shared Function htmlButton(value As String, Optional htmlClass As String = "", Optional htmlId As String = "", Optional onClick As String = "") As String
            Dim result As String = "<button name=""name"" value=""" & value & """"
            result += If(String.IsNullOrEmpty(htmlClass), "", " class=""" & htmlClass & """")
            result += If(String.IsNullOrEmpty(htmlId), "", " id=""" & htmlId & """")
            result += If(String.IsNullOrEmpty(onClick), "", " onClick=""" & onClick & """")
            Return result & ">"
        End Function
        '
        Public Shared Function htmlHidden(htmlName As String, htmlValue As String, Optional htmlClass As String = "", Optional htmlId As String = "") As String
            Dim result As String = "<input type=hidden name=""" & htmlName & """ value=""" & htmlValue & """"
            result += If(String.IsNullOrEmpty(htmlClass), "", " class=""" & htmlClass & """")
            result += If(String.IsNullOrEmpty(htmlId), "", " id=""" & htmlId & """")
            Return result & ">"
        End Function
        '
        Public Shared Function htmlHidden(htmlName As String, htmlValue As Integer, Optional htmlClass As String = "", Optional htmlId As String = "") As String
            Return htmlHidden(htmlName, htmlValue.ToString(), htmlClass, htmlId)
        End Function
        '
        Public Shared Function adminUrl(cp As CPBaseClass) As String
            Return cp.Site.GetText("adminurl")
        End Function
        '
        Public Shared Function kmaEncodeURL(cp As CPBaseClass, url As String) As String
            Return cp.Utils.EncodeUrl(url)
        End Function
        '
        Public Shared Function Main_GetPanel(src As String) As String
            Return htmlDiv(src)
        End Function
        '
        Public Shared Function kmaEncodeJavascript(cp As CPBaseClass, src As String) As String
            Return cp.Utils.EncodeJavascript(src)
        End Function
        '
        Public Shared Function kmaAddSpan(innerHtml As String, htmlClass As String) As String
            Return "<span class=""" & htmlClass & """>" & innerHtml & "</span>"
        End Function
        '
        Public Shared Function Main_GetFormInputCheckBox(htmlName As String, isChecked As Boolean) As String
            Return "<input type=checkbox name=""" & htmlName & """ value=1 " & If(isChecked, " checked", "") & ">"
        End Function
        '
        Public Shared Function Main_GetPanelInput(innerHtml As String) As String
            Return htmlDiv(innerHtml)
        End Function
        '
        Public Shared Function KmaEncodeMissingText(src As String, Optional ignore As String = "") As String
            Return src
        End Function
        '
        Public Shared Function KmaEncodeMissingBoolean(src As Boolean, Optional ignore As Boolean = False) As Boolean
            Return src
        End Function
        '
        Public Shared Function KmaEncodeMissingBoolean(src As Boolean, Optional ignore As String = "") As Boolean
            Return src
        End Function
        '
        Public Shared Function kmaEncodeHTML(cp As CPBaseClass, src As String) As String
            Return cp.Utils.EncodeHTML(src)
        End Function
        '

    End Class
End Namespace

