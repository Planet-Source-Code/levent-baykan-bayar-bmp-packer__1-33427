Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type BITMAPINFOHEADER '40 bytes
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
End Type
 
 Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
 
 
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type BITMAPFILEHEADER
    bfType            As String * 2
    bfSize            As Long
    bfReserved1       As Integer
    bfReserved2       As Integer
    bfOhFileBits         As Long
End Type
'
'Bitmaps read the colors in this reverse order

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY ' Enough for 256 colors
End Type
Private Type Color
    B As Byte
    G As Byte
    R As Byte
End Type

Private Type Pack_Header
Magic As String * 6
fnumber As Integer
fname() As String
End Type

Public Type Display_Info
GFX As Long
OBJ As Long
BPP As Byte
FSize As Long
Wid As Integer
Hei As Integer
End Type

Private Type BMPF
fname As String * 48
BM As BITMAPFILEHEADER
BMI As BITMAPINFOHEADER
Palet As LOGPALETTE
data() As Byte
End Type

Dim FileHeader As BITMAPFILEHEADER
Dim InfoHeader As BITMAPINFOHEADER
Dim Color As Color
Dim OldSeek As Long

Dim x As Integer
Dim Y As Integer
Dim BmPalette As LOGPALETTE
Dim File As String
Dim PicBits() As Byte
Public MyBMP() As Display_Info


Private Sub Get_Single_BMP(Index As Integer)
'On Error Resume Next

        Get #1, , FileHeader 'Get File Header
        Get #1, , InfoHeader 'Get Info Header
  
  If FileHeader.bfType <> "BM" Then
    MsgBox "Invalid Bitmap Header!->" & Index & vbCr & FileHeader.bfType
    'It's not Windows Bitmap File!!!
   
    Exit Sub
  End If
   
  If InfoHeader.biCompression <> 0 Then
    MsgBox "This Bitmap has bitmap compression which is not supported by this program!"
    'If It's compressed skip ,because we don't have filter to decompress it...
    
    Exit Sub
  End If
  
  If InfoHeader.biBitCount <= 4 Then
    MsgBox "1-4 bit images are not supported", 16
    '1-4 bit images have
    'different file structure
    'work on them if you like,
    'but i find it useless since
    'no one likes 16 colour images any more...

    Exit Sub
  End If

  'Set MYBMP Structures,see Form1.List1_Click procedure
  'for why do we need these
  '-------------------------------------------------------
  MyBMP(Index).BPP = InfoHeader.biBitCount
  MyBMP(Index).FSize = FileHeader.bfSize
  
  MyBMP(Index).OBJ = CreateCompatibleBitmap(Form1.hdc, InfoHeader.biWidth, InfoHeader.biHeight)
  
  MyBMP(Index).GFX = CreateCompatibleDC(Form1.hdc)
  SelectObject MyBMP(Index).GFX, MyBMP(Index).OBJ
  
  MyBMP(Index).Wid = InfoHeader.biWidth
  MyBMP(Index).Hei = InfoHeader.biHeight
  '-------------------------------------------------------
  If InfoHeader.biClrUsed <= 256 And InfoHeader.biBitCount < 24 Then
  'If image is 8 bit we must create a palette of colours
  
  CreatePalette
  
 
  End If
    'Scale PictureBox
    Form1.Picture1.Height = InfoHeader.biHeight
    Form1.Picture1.Width = InfoHeader.biWidth
    '---------------------------------------------------
    'This does not work so,if you are good in memory management
    'and find a way faster than my ,please contact me...
    '-----------POINTER METHOD-------------------------
    'Dim pic() As Byte
    'Dim sa As SAFEARRAY2D
    'Dim bmp As BITMAP
    'Dim i As Integer
    'Dim j As Integer
    
    'bmp.bmBits = FileHeader.bfOhFileBits
    'bmp.bmBitsPixel = InfoHeader.biBitCount
    'bmp.bmHeight = InfoHeader.biHeight
    'bmp.bmPlanes = InfoHeader.biPlanes
    'bmp.bmType = 19778 'FileHeader.bfType
    'bmp.bmWidth = InfoHeader.biWidth
    'bmp.bmWidthBytes = 256
    
    'MsgBox _
    bmp.bmBitsPixel & vbCr & _
    bmp.bmBits & vbCr & _
    bmp.bmHeight & vbCr & _
    bmp.bmPlanes & vbCr & _
    bmp.bmType & vbCr & _
    bmp.bmWidth & vbCr & _
    bmp.bmWidthBytes
    
    
    'With sa
    '    .cbElements = 1
    '    .cDims = 2
    '    .Bounds(0).lLbound = 0
    '    .Bounds(0).cElements = bmp.bmHeight
    '    .Bounds(1).lLbound = 0
    '    .Bounds(1).cElements = bmp.bmWidthBytes
    '    .pvData = bmp.bmBits
    'End With
    'CopyMemory ByVal VarPtrArray(pic), VarPtr(sa), 4
    'For i = 0 To UBound(pic, 1)
    '    For j = 0 To UBound(pic, 2)
    '        pic(i, j) = 255 - pic(i, j)
    '    Next j
    'Next i
    '
    'CopyMemory ByVal VarPtrArray(pic), 0&, 4
    'MsgBox "ok"
    'Exit Sub
    
    '------------------END POINTER-------------------
    
    
    
    
     
    If InfoHeader.biBitCount = 8 Then
    'We got 8 bit
   
   
   
    Dim ClL As Byte
    Dim ColX As Long
    
    
   
    For Y = InfoHeader.biHeight - 1 To -1 Step -1
    
    
        
        For x = 0 To InfoHeader.biWidth - 1
            Get #1, , ClL 'get a byte
            ColX = RGB(BmPalette.palPalEntry(ClL).peRed, BmPalette.palPalEntry(ClL).peGreen, BmPalette.palPalEntry(ClL).peBlue)
            SetPixelV MyBMP(Index).GFX, x, Y, ColX
            'SetPixelV is a bit faster than SetPixel
        Next x

    Next Y
    Else
    '24 bit image no need of palette
    'We use a Color type to get every pixel color in RGB format
    For Y = InfoHeader.biHeight - 1 To -1 Step -1
        
        For x = 0 To InfoHeader.biWidth - 1
            Get #1, , Color
            SetPixelV MyBMP(Index).GFX, x, Y, RGB(Color.R, Color.G, Color.B)
        Next x

    Next Y
 End If
  
'Seek File Pointer to the next BitMap Header in file
Seek #1, OldSeek + FileHeader.bfSize
OldSeek = Seek(1)
End Sub

Public Sub OpenBMPPack(File As String, List As ListBox)
Dim Magic As String * 6
Dim fnumber As Integer
Dim i As Integer
Dim nLen As Integer
Dim AllData As BMPF
Dim CP As Long
Dim HHL As Pack_Header

Dim fname As String
If File = "" Or Dir(File) = "" Then Exit Sub
Open File For Binary As #1
Get #1, , HHL
'Get the header

If HHL.Magic <> "BMPACK" Then
'Its not our type of file...
MsgBox "Invalid BM_PACK file!", 16
Exit Sub
End If

Form1.Label2.Caption = "File Type  : " & HHL.Magic + vbCr + _
                       "File Number: " & HHL.fnumber
                        
Screen.MousePointer = 11
OldSeek = Seek(1)
ReDim MyBMP(HHL.fnumber - 1)


For i = 1 To HHL.fnumber

List.AddItem HHL.fname(i - 1)
'write the name of the file
Get_Single_BMP i - 1
'get the file..
DoEvents
Next
Screen.MousePointer = 0
Close #1
MsgBox "Completed.."
End Sub
Public Sub SaveBMPPack(File As String, List As ListBox)
On Error Resume Next
Dim MyDir
Dim i As Integer
Dim data() As Byte
Dim fn() As String
Dim j As Integer
Dim HH As Pack_Header
Dim FOLNAME As String


HH.Magic = "BMPACK"
List.Clear
If Dir(File) <> "" Then Kill File
If File = "" Then Exit Sub
    Open File For Binary As #1
    '-----------------------------
    'list all files in a directory
    
FOLNAME = InputBox("Choose a folder which contains bmps", , App.Path)
If FOLNAME = "" Then Exit Sub

    MyDir = Dir(FOLNAME + "\*.bmp")
    
    Do
    If MyDir = "" Then Exit Do
    ReDim Preserve HH.fname(j)
    
    List.AddItem Left(MyDir, Len(MyDir) - 4)
    HH.fname(j) = MyDir
    
    j = j + 1
    
    MyDir = Dir
    Loop
    '------------------------------
    
    
    HH.fnumber = List.ListCount
    
    'Put header into file
    Put #1, , HH
    
    'Simple binary read and write:
    'Open each file in listbox,read its content
    'write it into the main file
    '
    'may be faster with WriteFile and ReadFile routines
    'but i think its enough,for now
    
    For i = 1 To List.ListCount
    Open FOLNAME + "\" + List.List(i - 1) + ".bmp" For Binary As #2
    ReDim data(LOF(2) - 1)
    Get #2, , data
    Put #1, , data
    Close #2
    Next
    Close #1
    
    
MsgBox "Done!"
End Sub
Private Sub CreatePalette()
'Routine to create a palette for
'8 bit bmps

Dim i As Long
Dim BlueByte As Byte
Dim RedByte As Byte
Dim GreenByte As Byte
Dim AByte As Byte
Dim ClrU As Long
'24 bits do not need it so exit sub
If InfoHeader.biBitCount = 24 Then Exit Sub
  
  BmPalette.palVersion = &H300
  BmPalette.palNumEntries = InfoHeader.biClrUsed

If InfoHeader.biClrUsed <> 0 Then ClrU = InfoHeader.biClrUsed - 1
    For i = 0 To ClrU
    Get #1, , BlueByte
    Get #1, , GreenByte
    Get #1, , RedByte
    Get #1, , AByte 'line feed byte for 8 bit bmps
    BmPalette.palPalEntry(i).peBlue = BlueByte
    BmPalette.palPalEntry(i).peGreen = GreenByte
    BmPalette.palPalEntry(i).peRed = RedByte
    BmPalette.palPalEntry(i).peFlags = AByte
    Next
End Sub

