<%@ WebHandler Language="VB" Class="ShowFlash" %>

Imports System
Imports System.Web

Public Class ShowFlash : Implements IHttpHandler
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim requestedSource As String = HttpContext.Current.Request("source")
        Dim requestedHeight As Integer = HttpContext.Current.Request("height")
        Dim requestedWidth As Integer = HttpContext.Current.Request("width")
        Dim strCurrentPath As String = HttpContext.Current.Request.ApplicationPath
        Dim source As String
        If Right(strCurrentPath, 1) <> "/" Then strCurrentPath = strCurrentPath & "/"
        source = strCurrentPath & requestedSource
        
        Dim script As String = _
        "<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" " & _
        "codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" height=""" & requestedHeight & """ width=""" & requestedWidth & """>" & _
        "<param name=""quality"" value=""high""/>" & _
        "<param name=""movie"" value=""" & source & """/>" & _
        "<param name=""wmode"" value=""transparent""/>" & _
        "<embed src=""" & source & """ quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" wmode=""transparent"" width=""" & requestedWidth & """ height=""" & requestedHeight & """ />" & _
        "</object>"
        
        Dim sReturn As String = "document.write('" & script & "');"
        context.Response.ContentType = "text/plain"
        context.Response.Write(sReturn)
        
    End Sub
    
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return True
        End Get
    End Property

End Class