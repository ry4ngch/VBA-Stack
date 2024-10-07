Attribute VB_Name = "MultiPageBGColor"
' Adapted from: https://www.mrexcel.com/board/threads/can-a-userform-multipage-backcolor-be-changed.79069/page-4#posts
' Written by: Jaafar Tribak

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type uPicDesc
    size As Long
    Type As Long
    #If VBA7 Then
        hPic As LongPtr
        hPal As LongPtr
    #Else
       hPic As Long
       hPal As Long
    #End If
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


#If VBA7 Then
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirectAut Lib "oleAut32.dll" Alias "OleCreatePictureIndirect" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirectPro Lib "olepro32.dll" Alias "OleCreatePictureIndirect" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
    Private Declare PtrSafe Function CopyImage Lib "user32" (ByVal handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
    Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
    Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hdc As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
    Private Declare PtrSafe Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
  
    Private hdc As LongPtr, hMemDc As LongPtr, hMemBmp As LongPtr, hBrush As LongPtr, hCopy As LongPtr, AR() As LongPtr
  
#Else
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
    Private Declare Function OleCreatePictureIndirectAut Lib "oleAut32.dll" Alias "OleCreatePictureIndirect" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
    Private Declare Function OleCreatePictureIndirectPro Lib "olepro32.dll" Alias "OleCreatePictureIndirect" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
    Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
    Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
  
    Private hdc As Long, hMemDc As Long, hMemBmp As Long, hBrush As Long, hCopy As Long, AR() As Long
  
#End If


Private Const IMAGE_BITMAP = 0
Private Const PICTYPE_BITMAP = 1
Private Const LR_COPYRETURNORG = &H4
Private Const S_OK = 0

Private Sub SetBackColor(Page As MSForms.Page, BackColor As Long)

    #If VBA7 Then
        Dim hLib As LongPtr
    #Else
        Dim hLib As Long
    #End If

    Dim r As RECT
    Dim IID_IDispatch As GUID
    Dim uPicinfo As uPicDesc
    Dim iPic As IPicture
    Dim lRet As Long
    Static i As Integer
  

    Page.PictureSizeMode = fmPictureSizeModeStretch
  
    hdc = GetDC(0)
    SetRect r, 0, 0, 1, 1

    With r
        hMemBmp = CreateCompatibleBitmap(hdc, .Right - .Left, .Bottom - .Top)
    End With

    hMemDc = CreateCompatibleDC(hdc)
    DeleteObject SelectObject(hMemDc, hMemBmp)
    hBrush = CreateSolidBrush(BackColor)
    FillRect hMemDc, r, hBrush
    hCopy = CopyImage(hMemBmp, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)

    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    With uPicinfo
        .size = Len(uPicinfo)
        .Type = PICTYPE_BITMAP
        .hPic = hCopy
        .hPal = 0
    End With
  
    hLib = LoadLibrary("oleAut32.dll")
        If hLib Then
            lRet = OleCreatePictureIndirectAut(uPicinfo, IID_IDispatch, True, iPic)
        Else
            hLib = LoadLibrary("olepro32.dll")
            lRet = OleCreatePictureIndirectPro(uPicinfo, IID_IDispatch, True, iPic)
        End If
    FreeLibrary hLib
  
    If lRet = S_OK Then
        Set Page.Picture = iPic
    Else
        MsgBox "Unable to create Picture", vbCritical, "Error."
    End If

    DeleteObject hMemBmp
    DeleteObject hMemDc
    DeleteObject hBrush
    ReleaseDC 0, hdc

    ReDim Preserve AR(i)
    AR(i) = hCopy
    i = i + 1

End Sub

Private Sub DeleteResources()

    Dim Element As Variant
  
    For Each Element In AR
        DeleteObject Element
    Next

End Sub

