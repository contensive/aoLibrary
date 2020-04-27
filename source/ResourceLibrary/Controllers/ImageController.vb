
Imports System
Imports System.Drawing
Imports Contensive.BaseClasses

Namespace Controllers
    Public Class ImageEditController
        Implements IDisposable

        Private loaded As Boolean = False
        Private src As String = ""
        Private srcImage As System.Drawing.Image
        Private setWidth As Integer = 0
        Private setHeight As Integer = 0
        Protected disposed As Boolean = False

        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposed Then

                If disposing Then

                    If loaded Then
                        srcImage.Dispose()
                        srcImage = Nothing
                    End If
                End If
            End If

            Me.disposed = True
        End Sub

        Public Function load(cp As CPBaseClass, ByVal pathFilename As String) As Boolean
            Dim returnOk As Boolean = False

            Try
                If cp.File.fileExists(pathFilename) Then
                    src = pathFilename
                    srcImage = System.Drawing.Image.FromFile(cp.Site.PhysicalFilePath & pathFilename)
                    setWidth = srcImage.Width
                    setHeight = srcImage.Height
                    loaded = True
                End If
            Catch __unusedException1__ As Exception
            End Try
            Return returnOk
        End Function

        Public Function save(cp As CPBaseClass, ByVal pathFilename As String) As Boolean
            Dim returnOk As Boolean = False
            Try
                If loaded Then
                    If src = pathFilename Then
                        If (cp.File.fileExists(cp.Site.PhysicalFilePath & pathFilename)) Then
                            cp.File.DeleteVirtual(pathFilename)
                        End If
                    End If
                    Using imgOutput As Bitmap = New Bitmap(srcImage, setWidth, setHeight)
                        Dim imgFormat As System.Drawing.Imaging.ImageFormat = srcImage.RawFormat
                        imgOutput.Save(cp.Site.PhysicalFilePath & pathFilename, imgFormat)
                    End Using
                    returnOk = True
                End If

            Catch __unusedException1__ As Exception
            End Try

            Return returnOk
        End Function

        Public Property width As Integer
            Get
                Return setWidth
            End Get
            Set(ByVal value As Integer)
                setWidth = value
            End Set
        End Property

        Public Property height As Integer
            Get
                Return setWidth
            End Get
            Set(ByVal value As Integer)
                setHeight = value
            End Set
        End Property

        Public Sub dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overrides Sub Finalize()
            Dispose(False)
        End Sub
    End Class
End Namespace
