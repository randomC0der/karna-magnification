Imports System.ComponentModel
Imports System.Runtime.InteropServices

' https://social.msdn.microsoft.com/Forums/en-US/e585f334-7007-4e84-b96f-09955bca8da5/screen-capture-implementation-using-magnification-api
Public Class Form1

    Public Shared Sub Main()
        Application.Run(New Form1)
    End Sub

    Public Const WS_OVERLAPPED = &H0
    Public Const WS_THICKFRAME = &H40000
    Public Const WS_BORDER = &H800000
    Public Const WS_POPUP = &H80000000
    Public Const WS_CHILD = &H40000000
    Public Const WS_VISIBLE = &H10000000
    Public Const WS_CLIPCHILDREN = &H2000000

    Public Const WS_EX_TOPMOST = &H8
    Public Const WS_EX_LAYERED = &H80000
    Public Const WS_EX_TRANSPARENT = &H20L
    Public Const WS_EX_TOOLWINDOW = &H80

    Public Const LWA_ALPHA = &H0

    Public Const MS_SHOWMAGNIFIEDCURSOR = &H10001

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode, EntryPoint:="CreateWindowExW")>
    Public Shared Function CreateWindowEx(ByVal dwExStyle As Integer, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hWndParent As IntPtr, ByVal hMenu As IntPtr, ByVal hInstance As IntPtr, ByVal lpParam As IntPtr) As IntPtr
    End Function

    Public Const SWP_NOREDRAW As Integer = &H8, SWP_FRAMECHANGED As Integer = &H20, SWP_NOCOPYBITS As Integer = &H100, SWP_NOOWNERZORDER As Integer = &H200, SWP_NOSENDCHANGING As Integer = &H400, SWP_NOREPOSITION As Integer = SWP_NOOWNERZORDER, SWP_DEFERERASE As Integer = &H2000, SWP_ASYNCWINDOWPOS As Integer = &H4000, SWP_HIDEWINDOW As Integer = &H80

    <DllImport("User32.dll", SetLastError:=True, ExactSpelling:=True, CharSet:=CharSet.Auto)>
    Public Shared Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal flags As Integer) As Boolean
    End Function

    Public Delegate Function WindowProc(ByVal hWnd As IntPtr, ByVal msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr

    <StructLayout(LayoutKind.Sequential)>
    Public Structure WNDCLASSEX
        Public cbSize As UInteger
        Public style As UInteger
        <MarshalAs(UnmanagedType.FunctionPtr)>
        Public lpfnWndProc As WindowProc
        Public cbClsExtra As Integer
        Public cbWndExtra As Integer
        Public hInstance As IntPtr
        Public hIcon As IntPtr
        Public hCursor As IntPtr
        Public hbrBackground As IntPtr
        Public lpszMenuName As String
        Public lpszClassName As String
        Public hIconSm As IntPtr
    End Structure

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function RegisterClassEx(<[In]> ByRef lpwcx As WNDCLASSEX) As Short
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function DefWindowProc(ByVal hWnd As IntPtr, ByVal uMsg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function SetLayeredWindowAttributes(ByVal hwnd As IntPtr, ByVal crKey As UInteger, ByVal bAlpha As Byte, ByVal dwFlags As UInteger) As Boolean
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function GetDesktopWindow() As IntPtr
    End Function


    <DllImport("Magnification.dll")>
    Public Shared Function MagInitialize() As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure

    <DllImport("Magnification.dll")>
    Public Shared Function MagSetWindowSource(hWnd As IntPtr, ByVal rect As IntPtr) As Boolean
    End Function

    <DllImport("Magnification.dll")>
    Public Shared Function MagSetWindowSource(hWnd As IntPtr, <[In], MarshalAs(UnmanagedType.Struct)> rect As RECT) As Boolean
    End Function

    <DllImport("Magnification.dll")>
    Public Shared Function MagSetWindowFilterList(ByVal hwnd As IntPtr, ByVal dwFilterMode As Integer, ByVal count As Integer, ByRef pHWND As IntPtr) As Boolean
    End Function

    <DllImport("Magnification.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function MagSetImageScalingCallback(ByVal hwnd As IntPtr, MagImageScalingCallback As MagImageScalingCallback) As Boolean
    End Function

    '<UnmanagedFunctionPointer(CallingConvention.StdCall)>
    'Public Delegate Function MagImageScalingCallback(ByVal hwnd As IntPtr, srcdata As IntPtr, ByRef srcheader As MAGIMAGEHEADER, ByRef destdata As Byte(), ByRef destheader As MAGIMAGEHEADER,
    '                                                 ByRef unclipped As RECT, ByRef clipped As RECT, dirty As IntPtr) As IntPtr

    Public Delegate Function MagImageScalingCallback(hwnd As IntPtr, srcdata As IntPtr, ByRef srcheader As MAGIMAGEHEADER, destdata As IntPtr, destheader As MAGIMAGEHEADER,
                                                     unclipped As IntPtr, clipped As IntPtr, dirty As IntPtr) As IntPtr

    '<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=8)>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure MAGIMAGEHEADER
        Public width As Integer
        Public height As Integer
        Public format As Guid
        Public stride As Integer
        Public offset As Integer
        Public cbSize As Integer
    End Structure

    Public Shared GUID_WICPixelFormat32bppRGBA As New Guid("F5C7AD2D-6A8D-43DD-A7A8-A29935261AE9")

    <StructLayout(LayoutKind.Sequential)>
    Public Structure BITMAPINFOHEADER
        <MarshalAs(UnmanagedType.I4)>
        Public biSize As Integer
        <MarshalAs(UnmanagedType.I4)>
        Public biWidth As Integer
        <MarshalAs(UnmanagedType.I4)>
        Public biHeight As Integer
        <MarshalAs(UnmanagedType.I2)>
        Public biPlanes As Short
        <MarshalAs(UnmanagedType.I2)>
        Public biBitCount As Short
        <MarshalAs(UnmanagedType.I4)>
        Public biCompression As Integer
        <MarshalAs(UnmanagedType.I4)>
        Public biSizeImage As Integer
        <MarshalAs(UnmanagedType.I4)>
        Public biXPelsPerMeter As Integer
        <MarshalAs(UnmanagedType.I4)>
        Public biYPelsPerMeter As Integer
        <MarshalAs(UnmanagedType.I4)>
        Public biClrUsed As Integer
        <MarshalAs(UnmanagedType.I4)>
        Public biClrImportant As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure BITMAPINFO
        <MarshalAs(UnmanagedType.Struct, SizeConst:=40)>
        Public bmiHeader As BITMAPINFOHEADER
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=1024)>
        Public bmiColors As Integer()
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure BITMAPINFO_FLAT
        Public biSize As Integer
        Public biWidth As Integer
        Public biHeight As Integer
        Public biPlanes As Short
        Public biBitCount As Short
        Public biCompression As Integer
        Public biSizeImage As Integer
        Public biXPelsPerMeter As Integer
        Public biYPelsPerMeter As Integer
        Public biClrUsed As Integer
        Public biClrImportant As Integer
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=1024)>
        Public bmiColors As Byte()
    End Structure

    Public Const BI_RGB = 0
    Public Const BI_RLE8 = 1
    Public Const BI_RLE4 = 2
    Public Const BI_BITFIELDS = 3
    Public Const BI_JPEG = 4
    Public Const BI_PNG = 5

    Public Const DIB_RGB_COLORS = 0
    Public Const DIB_PAL_COLORS = 1

    Public Const CBM_INIT = &H4

    <DllImport("User32.dll", SetLastError:=True)>
    Private Shared Function GetDC(hWnd As IntPtr) As IntPtr
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Private Shared Function GetWindowDC(hWnd As IntPtr) As IntPtr
    End Function

    <DllImport("User32.dll", SetLastError:=True)>
    Private Shared Function ReleaseDC(ByVal hWnd As IntPtr, hdc As IntPtr) As IntPtr
    End Function

    <DllImport("Gdi32.dll", SetLastError:=True)>
    Public Shared Function DeleteDC(ByVal hDC As IntPtr) As Boolean
    End Function

    <DllImport("Gdi32.dll", SetLastError:=True)>
    Public Shared Function CreateCompatibleDC(ByVal hDC As IntPtr) As IntPtr
    End Function

    <DllImport("Gdi32.dll", SetLastError:=True)>
    Private Shared Function SelectObject(hDC As IntPtr, ByVal hObject As IntPtr) As IntPtr
    End Function

    <DllImport("Gdi32.dll", SetLastError:=True)>
    Public Shared Function DeleteObject(<[In]> hObject As IntPtr) As Boolean
    End Function

    Public Const SRCCOPY = &HCC0020

    <DllImport("Gdi32.dll", SetLastError:=True)>
    Public Shared Function BitBlt(hDCDest As IntPtr, XOriginDest As Integer, YOriginDest As Integer, WidthDest As Integer, HeightDest As Integer, hDCSrc As IntPtr, XOriginScr As Integer, YOriginSrc As Integer, dwRop As Integer) As Boolean
    End Function

    <DllImport("Gdi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function StretchDIBits(ByVal hdc As IntPtr, ByVal XDest As Integer, ByVal YDest As Integer, ByVal nDestWidth As Integer, ByVal nDestHeight As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal nSrcWidth As Integer, ByVal nSrcHeight As Integer, ByVal lpBits As IntPtr, ByRef lpBitsInfo As BITMAPINFO_FLAT, ByVal iUsage As Integer, ByVal dwRop As Integer) As Integer
    End Function

    <DllImport("Gdi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function CreateDIBitmap(ByVal hdc As IntPtr, ByRef lpbmih As BITMAPINFOHEADER, fdwInit As Integer, lpbInit As IntPtr, ByRef lpBitsInfo As BITMAPINFO, ByVal fuUsage As UInt32) As IntPtr
    End Function

    Public Const MW_FILTERMODE_EXCLUDE = 0
    Public Const MW_FILTERMODE_INCLUDE = 1

    Public Shared Function MagImageScaling(hwnd As IntPtr, srcdata As IntPtr, ByRef srcheader As MAGIMAGEHEADER, destdata As IntPtr, destheader As MAGIMAGEHEADER,
                                                     unclipped As IntPtr, clipped As IntPtr, dirty As IntPtr) As IntPtr
        Dim lpbmi As BITMAPINFO = New BITMAPINFO()
        lpbmi.bmiHeader.biSize = Marshal.SizeOf(lpbmi.bmiHeader)
        lpbmi.bmiHeader.biHeight = CInt(-srcheader.height)
        lpbmi.bmiHeader.biWidth = CInt(srcheader.width)
        lpbmi.bmiHeader.biSizeImage = CInt(srcheader.cbSize)
        lpbmi.bmiHeader.biPlanes = 1
        lpbmi.bmiHeader.biBitCount = 32
        lpbmi.bmiHeader.biCompression = BI_RGB
        Dim hDC As IntPtr = GetWindowDC(hwnd)
        Form1.hBitmap = CreateDIBitmap(hDC, lpbmi.bmiHeader, CBM_INIT, srcdata, lpbmi, DIB_RGB_COLORS)
        Form1.bCallbackDone = True
        DeleteDC(hDC)

        ' Test copy to screen
        'Dim hDCScreen As IntPtr = GetDC(IntPtr.Zero)
        'Dim hDCMem As IntPtr = CreateCompatibleDC(hDCScreen)
        'Dim hBitmapOld As IntPtr = SelectObject(hDCMem, Form1.hBitmap)
        'BitBlt(hDCScreen, 0, 0, CInt(srcheader.width), CInt(srcheader.height), hDCMem, 0, 0, SRCCOPY)
        'SelectObject(hDCMem, hBitmapOld)
        'DeleteObject(hDCMem)
        'ReleaseDC(IntPtr.Zero, hDCScreen)
    End Function

    Friend WithEvents Button1 As Button
    Dim hWndMag As IntPtr = IntPtr.Zero
    Dim hWndHost As IntPtr = IntPtr.Zero
    'Private Shared wndProc As WndProc = New WndProc(AddressOf HostWndProc)
    Dim nWidth As Integer = 0
    Dim nHeight As Integer = 0
    Dim sFilename As String = String.Empty
    Dim imageformat As System.Drawing.Imaging.ImageFormat
    Shared bCallbackDone As Boolean = False
    Shared hBitmap As IntPtr = IntPtr.Zero

    Public Shared Function HostWndProc(ByVal hWnd As IntPtr, ByVal msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        Return DefWindowProc(hWnd, msg, wParam, lParam)
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DLGTEMPLATE
        Public style As Integer
        Public dwExtendedStyle As Integer
        Public cdit As UShort
        Public x As Short
        Public y As Short
        Public cx As Short
        Public cy As Short
    End Structure

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        nWidth = Screen.PrimaryScreen.Bounds.Width
        nHeight = Screen.PrimaryScreen.Bounds.Height

        Static _wndProc As WindowProc = New WindowProc(AddressOf HostWndProc)
        Dim sWindowClassName As String = "MagnifierHost"
        Dim wcex As WNDCLASSEX = New WNDCLASSEX()
        wcex.cbSize = CUInt(Marshal.SizeOf(wcex))
        wcex.lpfnWndProc = _wndProc
        wcex.hInstance = IntPtr.Zero
        wcex.lpszClassName = sWindowClassName
        wcex.hCursor = IntPtr.Zero
        wcex.hbrBackground = IntPtr.Zero
        wcex.style = 0
        wcex.cbClsExtra = 0
        wcex.cbWndExtra = 0
        wcex.hIcon = IntPtr.Zero
        wcex.lpszMenuName = Nothing
        wcex.hIconSm = IntPtr.Zero
        If (RegisterClassEx(wcex) <> 0) Then
            hWndHost = CreateWindowEx(WS_EX_TOPMOST Or WS_EX_LAYERED Or WS_EX_TRANSPARENT, sWindowClassName, "Host Window", WS_POPUP Or WS_CLIPCHILDREN, 0, 0,
                                      0, 0, IntPtr.Zero, IntPtr.Zero, wcex.hInstance, IntPtr.Zero)
            If (hWndHost <> IntPtr.Zero) Then
                SetWindowPos(hWndHost, IntPtr.Zero, 0, 0, nWidth, nHeight, SWP_HIDEWINDOW)
                Dim bRet As Boolean = SetLayeredWindowAttributes(hWndHost, 0, &HFF, LWA_ALPHA)
                If (MagInitialize()) Then
                    hWndMag = CreateWindowEx(0, "Magnifier", "MagnifierWindow", WS_CHILD Or MS_SHOWMAGNIFIEDCURSOR Or WS_VISIBLE,
                    0, 0, nWidth, nHeight, hWndHost, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero)
                    If (hWndMag <> IntPtr.Zero) Then
                    Else
                        Throw New Win32Exception(Marshal.GetLastWin32Error())
                    End If
                    Static MagImageScalingProc As MagImageScalingCallback = New MagImageScalingCallback(AddressOf MagImageScaling)
                    If (MagSetImageScalingCallback(hWndMag, MagImageScalingProc) <> False) Then

                    End If
                End If
            End If
        End If

        Me.ClientSize = New System.Drawing.Size(260, 140)

        Me.Button1 = New Button()
        Me.Button1.Location = New System.Drawing.Point(92, 50)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 32)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Capture"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Controls.Add(Me.Button1)

        CenterToScreen()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim pWnd As IntPtr = Me.Handle
        If (MagSetWindowFilterList(hWndMag, MW_FILTERMODE_EXCLUDE, 1, pWnd) <> False) Then
            Dim sourceRect As RECT = New RECT()
            sourceRect.left = 0
            sourceRect.top = 0
            sourceRect.right = nWidth
            sourceRect.bottom = nHeight

            bCallbackDone = False
            If (MagSetWindowSource(hWndMag, sourceRect) = True) Then

                Cursor = Cursors.WaitCursor
                While (bCallbackDone = False)
                End While
                Cursor = Cursors.Default

                Dim saveFileDialog1 As New SaveFileDialog()
                saveFileDialog1.Filter = "JPeg Image|*.jpg|Bitmap Image|*.bmp|Gif Image|*.gif|Png Image|*.png"
                saveFileDialog1.Title = "Save Image File"
                saveFileDialog1.ShowDialog()
                If (saveFileDialog1.FileName <> "") Then
                    sFilename = saveFileDialog1.FileName
                    Select Case saveFileDialog1.FilterIndex
                        Case 1
                            imageformat = System.Drawing.Imaging.ImageFormat.Jpeg
                        Case 2
                            imageformat = System.Drawing.Imaging.ImageFormat.Bmp
                        Case 3
                            imageformat = System.Drawing.Imaging.ImageFormat.Gif
                        Case 4
                            imageformat = System.Drawing.Imaging.ImageFormat.Png
                    End Select
                    Dim img As Image = Image.FromHbitmap(hBitmap)
                    'img.Save("e:\testcapture.jpg")
                    img.Save(sFilename, imageformat)
                    DeleteObject(hBitmap)
                End If
            End If
        End If
    End Sub
End Class