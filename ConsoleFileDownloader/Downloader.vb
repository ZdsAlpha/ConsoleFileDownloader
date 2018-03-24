Imports System.Net
Module Downloader
    Sub Main()
        Console.Write("Enter file name:")
        Dim Name As String = Console.ReadLine
        Console.WriteLine("Enter URL:")
        Dim URL As String = InputBox("Enter URL:", "File URL")
        Download(URL, Name)
        Console.ReadKey()
    End Sub
    Public Sub Download(URL As String, TargetFile As String)
        Dim Stopwatch As New Stopwatch
        Console.Clear()
        Console.WriteLine("Connecting...")
        Try
            Using FileStream As New IO.FileStream(TargetFile, IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite)
                Dim WebRequest As HttpWebRequest = HttpWebRequest.Create(URL)
                If FileStream.Length <> 0 Then WebRequest.AddRange(FileStream.Length)
                FileStream.Position = FileStream.Length
                Dim WebResponse = WebRequest.GetResponse
                Using Stream = WebResponse.GetResponseStream
                    Console.WriteLine("Connected!")
                    Console.WriteLine("Downloading...")
                    Dim Buffer(1024 * 64 - 1) As Byte
                    Dim DownloadedBytes As UInteger = FileStream.Length
                    Do
                        Try
                            Stopwatch.Restart()
                            Dim BytesRead = Stream.Read(Buffer, 0, Buffer.Length)
                            FileStream.Write(Buffer, 0, BytesRead)
                            Stopwatch.Stop()
                            DownloadedBytes += BytesRead
                            Console.Clear()
                            Console.WriteLine("Bytes Downloaded: " + DownloadedBytes.ToString + " Bytes")
                            Console.WriteLine("Download Speed: " + Int((BytesRead / 1024) / Stopwatch.Elapsed.TotalSeconds).ToString + " KB/s")
                            If BytesRead = 0 Then
                                Console.WriteLine("Downloading completed!")
                                Exit Do
                            End If
                        Catch ex As Exception
                            Console.WriteLine("Connection closed!")
                            Exit Do
                        End Try
                    Loop
                    Erase Buffer
                End Using
            End Using
        Catch ex As Exception
            Console.WriteLine("Failed to connect.")
        End Try
    End Sub
End Module
Namespace Networking
    Public Class BasicDownloader
        Implements IDisposable
        Private _source As IO.Stream
        Private _target As IO.Stream
        Private _contentlength As Long = -1
        Private _blocksize As Long = 1024 * 4
        Private _isdownloading As Boolean = False
        Private _downloadedbytes As Long = 0
        Private _bytes() As Byte
        Private _isdisposed As Boolean = False
        Private _isstarted As Boolean = False
        Private _stopwatch As New Stopwatch
        Public Event OnUpdate(ByRef StopDownloading As Boolean, ByRef DownloadedBytes As Long, ByRef CurrentDownloadedBytes As Long, ByRef Bytes() As Byte)
        Public Event OnFinished(ByRef ReDownload As Boolean, ByRef DownloadedBytes As Long, ByRef HttpStream As IO.Stream, ByRef TargetStream As IO.Stream)
        Public ReadOnly Property Source As IO.Stream
            Get
                Return _source
            End Get
        End Property
        Public ReadOnly Property Target As IO.Stream
            Get
                Return _target
            End Get
        End Property
        Public ReadOnly Property IsDownloading() As Boolean
            Get
                Return _isdownloading
            End Get
        End Property
        Public ReadOnly Property DownloadedBytes As Long
            Get
                Return _downloadedbytes
            End Get
        End Property
        Public ReadOnly Property IsDisposed As Boolean
            Get
                Return _isdisposed
            End Get
        End Property
        Public ReadOnly Property UsedBytes As Byte()
            Get
                Return _bytes
            End Get
        End Property
        Public ReadOnly Property IsRunning As Boolean
            Get
                Return _isstarted
            End Get
        End Property
        Public ReadOnly Property DownloadSpeed As Double
            Get
                Return _downloadedbytes / _stopwatch.ElapsedMilliseconds
            End Get
        End Property
        Public Property ContentLength As Long
            Get
                Return _contentlength
            End Get
            Set(value As Long)
                _contentlength = value
            End Set
        End Property
        Public Property BlockSize As Long
            Get
                Return _blocksize
            End Get
            Set(value As Long)
                _blocksize = value
            End Set
        End Property
        Public Sub Download()
            If _isdisposed = True Then Throw New Exception("The downloader is disposed.")
            If _isdownloading = False Then
                _downloadedbytes = 0
                _isdownloading = True
                _isstarted = True
                _stopwatch.Restart()
                Download(Source, Target, AddressOf Update, AddressOf Finished, BlockSize, ContentLength)
                If _stopwatch.IsRunning Then _stopwatch.Stop()
                _isdownloading = False
                _isstarted = False
                _bytes = Nothing
            Else
                Throw New Exception("Already downloading...")
            End If
        End Sub
        Public Function DownloadAsync() As Threading.Thread
            Dim Thread As New Threading.Thread(AddressOf Download)
            Thread.Start()
            Return Thread
        End Function
        Public Sub [Stop]()
            _isdownloading = False
        End Sub
        Public Sub Start()
            If _isdownloading = False Then Throw New Exception("Downloading is not started. Use Downloader.Download to start downloading.")
            _isstarted = True
            _stopwatch.Start()
        End Sub
        Public Sub Pause()
            If _isdownloading = False Then Throw New Exception("Downloading is not started. Use Downloader.Download to start downloading.")
            _isstarted = False
            _stopwatch.Stop()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _isdisposed = True Then Exit Sub
            If _isdownloading = False Then
                If _stopwatch.IsRunning = False Then _stopwatch.Stop()
                _stopwatch = Nothing
                _source.Dispose()
                _target.Dispose()
                _source = Nothing
                _target = Nothing
                _contentlength = Nothing
                _blocksize = Nothing
                _isdownloading = Nothing
                _downloadedbytes = Nothing
                _bytes = Nothing
                _isstarted = Nothing
                _isdisposed = True
            Else
                Throw New Exception("Downloading... The Downloader cannot be disposed. Use Downloader.Stop() to stop downloading.")
            End If
        End Sub
        Private Sub New(Source As IO.Stream, Target As IO.Stream)
            Me._source = Source
            Me._target = Target
        End Sub
        Private Sub New()
        End Sub
        Private Sub Update(ByRef StopDownloading As Boolean, ByRef DownloadedBytes As Long, ByRef CurrentDownloadedBytes As Long, ByRef Bytes() As Byte)
            _downloadedbytes = DownloadedBytes
            _bytes = Bytes
            RaiseEvent OnUpdate(StopDownloading, DownloadedBytes, CurrentDownloadedBytes, Bytes)
            If _isdownloading = False Then
                StopDownloading = True
            End If
            If _isstarted = False Then
                Dim spinner As New System.Threading.SpinWait
                Do
                    If _isstarted = False Then
                        If _isdownloading = False Then
                            StopDownloading = True
                            Exit Do
                        End If
                        spinner.SpinOnce()
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End Sub
        Private Sub Finished(ByRef ReDownload As Boolean, ByRef DownloadedBytes As Long, ByRef HttpStream As IO.Stream, ByRef TargetStream As IO.Stream)
            RaiseEvent OnFinished(ReDownload, DownloadedBytes, HttpStream, TargetStream)
        End Sub

        Public Shared Function Create(Source As IO.Stream, Target As IO.Stream) As BasicDownloader
            Return New BasicDownloader(Source, Target)
        End Function
        Public Shared Function Create(URL As String, Target As IO.Stream, Optional StartPosition As Long = -1, Optional EndPosition As Long = -1)
            Dim WebRequest As HttpWebRequest = HttpWebRequest.Create(URL)
            If StartPosition <> -1 Then
                If EndPosition = -1 Then
                    WebRequest.AddRange(StartPosition)
                Else
                    WebRequest.AddRange(StartPosition, EndPosition)
                End If
            End If
            Dim WebResponse = WebRequest.GetResponse
            Dim HttpStream As IO.Stream = WebResponse.GetResponseStream
            Return Create(HttpStream, Target)
        End Function
        Public Shared Function Create(URL As String, FileName As String, Optional FileStartPosition As Long = -1, Optional StartPosition As Long = -1, Optional EndPosition As Long = -1)
            Dim Target As New IO.FileStream(FileName, IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite)
            If FileStartPosition <> -1 Then
                Target.Position = FileStartPosition
            End If
            Return Create(URL, Target, StartPosition, EndPosition)
        End Function

        Public Shared Sub Download(SourceStream As IO.Stream, TargetStream As IO.Stream, Optional OnUpdate As OnUpdateDownloader = Nothing, Optional OnFinished As OnDownloaded = Nothing, Optional BlockSize As Long = 1024 * 4, Optional ContentLength As Long = -1)
            Dim DownloadedBytes As Long = 0
            Dim CurrentDownloadedBytes As Long = 0
            Dim Bytes(BlockSize - 1) As Byte
            Dim StopDownloading As Boolean = False
ReDownload:
            Do
                If ContentLength = -1 Then
                    CurrentDownloadedBytes = SourceStream.Read(Bytes, 0, Bytes.Length)
                Else
                    If ContentLength - DownloadedBytes <= BlockSize Then
                        CurrentDownloadedBytes = SourceStream.Read(Bytes, 0, ContentLength - DownloadedBytes)
                    Else
                        CurrentDownloadedBytes = SourceStream.Read(Bytes, 0, Bytes.Length)
                    End If
                End If
                DownloadedBytes = DownloadedBytes + CurrentDownloadedBytes
                If OnUpdate IsNot Nothing Then
                    OnUpdate.Invoke(StopDownloading, DownloadedBytes, CurrentDownloadedBytes, Bytes)
                    If StopDownloading = True Then Exit Do
                End If
                TargetStream.Write(Bytes, 0, CurrentDownloadedBytes)
                If CurrentDownloadedBytes = 0 Then Exit Do
            Loop
            If OnFinished IsNot Nothing Then
                Dim ReDownload As Boolean = False
                OnFinished.Invoke(ReDownload, DownloadedBytes, SourceStream, TargetStream)
                If ReDownload = True Then
                    StopDownloading = False
                    CurrentDownloadedBytes = 0
                    DownloadedBytes = 0
                    DownloadedBytes = 0
                    GoTo ReDownload
                End If
            End If
            StopDownloading = False
            CurrentDownloadedBytes = 0
            DownloadedBytes = 0
            Erase Bytes
        End Sub
        Public Shared Sub Download(URL As String, TargetStream As IO.Stream, Optional StartPosition As Long = -1, Optional EndPosition As Long = -1, Optional OnUpdate As OnUpdateDownloader = Nothing, Optional OnFinished As OnDownloaded = Nothing, Optional BlockSize As Long = 1024 * 4, Optional ContentLength As Long = -1)
            Dim WebRequest As HttpWebRequest = HttpWebRequest.Create(URL)
            If StartPosition <> -1 Then
                If EndPosition = -1 Then
                    WebRequest.AddRange(StartPosition)
                Else
                    WebRequest.AddRange(StartPosition, EndPosition)
                End If
            End If
            Dim WebResponse = WebRequest.GetResponse
            Dim HttpStream As IO.Stream = WebResponse.GetResponseStream
            Download(HttpStream, TargetStream, OnUpdate, OnFinished, BlockSize, ContentLength)
            HttpStream.Dispose()
            WebResponse.Dispose()
        End Sub
        Public Shared Sub Download(URL As String, FileName As String, Optional FileStartPosition As Long = -1, Optional StartPosition As Long = -1, Optional EndPosition As Long = -1, Optional OnUpdate As OnUpdateDownloader = Nothing, Optional OnFinished As OnDownloaded = Nothing, Optional BlockSize As Long = 1024 * 4, Optional ContentLength As Long = -1)
            Dim Target As New IO.FileStream(FileName, IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite)
            If FileStartPosition <> -1 Then
                Target.Position = FileStartPosition
            End If
            Download(URL, Target, StartPosition, EndPosition, OnUpdate, OnFinished, BlockSize, ContentLength)
        End Sub
        Public Shared Function Download(URL As String, Optional StartPosition As Long = -1, Optional EndPosition As Long = -1, Optional OnUpdate As OnUpdateDownloader = Nothing, Optional OnFinished As OnDownloaded = Nothing, Optional BlockSize As Long = 1024 * 4, Optional ContentLength As Long = -1) As Byte()
            Dim Target As New IO.MemoryStream
            Download(URL, Target, StartPosition, EndPosition, OnUpdate, OnFinished, BlockSize, ContentLength)
            Dim Bytes = Target.ToArray
            Target.Dispose()
            Return Bytes
        End Function
        Public Shared Function GetLength(URL As String) As Long
            Dim Request As HttpWebRequest = HttpWebRequest.Create(URL)
            Dim Response = Request.GetResponse
            Dim Length = Response.ContentLength
            Response.Dispose()
            Return Length
        End Function
        Public Delegate Sub OnUpdateDownloader(ByRef StopDownloading As Boolean, ByRef DownloadedBytes As Long, ByRef CurrentDownloadedBytes As Long, ByRef Bytes() As Byte)
        Public Delegate Sub OnDownloaded(ByRef ReDownload As Boolean, ByRef DownloadedBytes As Long, ByRef HttpStream As IO.Stream, ByRef TargetStream As IO.Stream)

        Public Class RangedDownloader
            Inherits BasicDownloader
            Implements IDisposable
            Private _url As String
            Private _start As Long = 0
            Private _end As Long = 0
            Private _response As HttpWebResponse
            Public ReadOnly Property URL As String
                Get
                    Return _url
                End Get
            End Property
            Public ReadOnly Property StartPosition As Long
                Get
                    Return _start
                End Get
            End Property
            Public ReadOnly Property EndPosition As Long
                Get
                    Return _end
                End Get
            End Property
            Public ReadOnly Property Position As Long
                Get
                    Return StartPosition + _downloadedbytes
                End Get
            End Property
            Sub New(URL As String, Target As IO.Stream, StartPosition As Long, EndPosition As Long)
                _start = StartPosition
                _end = EndPosition
                _url = URL
                Dim Request As HttpWebRequest = HttpWebRequest.Create(URL)
                Request.AddRange(StartPosition, EndPosition)
                Me._response = Request.GetResponse
                Dim ResponseStream = Me._response.GetResponseStream
                Me._source = ResponseStream
                Me._target = Target
            End Sub
            Sub New(URL As String, FileName As String, StartPosition As Long, EndPosition As Long)
                Dim Target = New IO.FileStream(FileName, IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite, IO.FileShare.ReadWrite)
                Target.Position = StartPosition
                _start = StartPosition
                _end = EndPosition
                _url = URL
                Dim Request As HttpWebRequest = HttpWebRequest.Create(URL)
                Request.Proxy = Nothing
                Request.AddRange(StartPosition, EndPosition)
                Me._response = Request.GetResponse
                Dim ResponseStream = Me._response.GetResponseStream
                Me._source = ResponseStream
                Me._target = Target
            End Sub
            Public Sub Renew(Target As IO.Stream)
                MyBase.Dispose()
                _response.Dispose()
                Dim Request As HttpWebRequest = HttpWebRequest.Create(URL)
                Request.AddRange(StartPosition, EndPosition)
                Me._response = Request.GetResponse
                Dim ResponseStream = Me._response.GetResponseStream
                Me._source = ResponseStream
                Me._target = Target
            End Sub
            Public Overloads Sub Dispose() Implements IDisposable.Dispose
                MyBase.Dispose()
                _response.Dispose()
                _start = Nothing
                _end = Nothing
            End Sub
        End Class
    End Class
End Namespace