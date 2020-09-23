<div align="center">

## \_\_\_Fast Masks\!\_\_\_ \(Error fixed\)


</div>

### Description

Have you ever wanted to create your bitblt masks when the application starts without having the user wait 5 minutes, now u can!
 
### More Info
 
This code creates masks for bitblt during runtime really fast!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mathieu Chartier](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mathieu-chartier.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mathieu-chartier-fast-masks-error-fixed__1-50210/archive/master.zip)

### API Declarations

See the code below.


### Source Code

```
Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Sub CreateMask(hDestDC As Long, X As Long, Y As Long, nWidth As Long, nHeight As Long, hSrcDC As Long, XSrc As Long, YSrc As Long, TransColor As Long)
 Dim OrigColor As Long  ' Holds source original background color
 Dim DestBKColor As Long  ' Holds destination original background color
 Dim OrigTextColor As Long
 Dim hMaskBmp As Long
 Dim hMaskPrevBmp As Long
 Dim MaskDC As Long
 MaskDC = CreateCompatibleDC(hDestDC)
 hMaskBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&) '//Create a monochrome bitmap for our mask
 hMaskPrevBmp = SelectObject(MaskDC, hMaskBmp)
 OrigColor = SetBkColor(hSrcDC, TransColor)
  Call BitBlt(MaskDC, 0, 0, nWidth, nHeight, hSrcDC, XSrc, YSrc, vbSrcCopy) '//Copy hSrcDc into our mask bitmap
 SetBkColor hSrcDC, OrigColor '//Restore the original color
 DestBKColor = SetBkColor(hDestDC, vbWhite) '//All the white in our bitmap hasto be white
 OrigTextColor = SetTextColor(hDestDC, vbBlack)
  BitBlt hDestDC, X, Y, nWidth, nHeight, MaskDC, 0, 0, vbSrcCopy
 SetTextColor hDestDC, OrigTextColor
 SetBkColor hDestDC, DestBKColor '//Restore the original back color bak
 Call SelectObject(MaskDC, hMaskPrevBmp) 'Select our original bitmap bak
 Call DeleteObject(hMaskBmp) 'Delete our mask bitmap
 Call DeleteDC(MaskDC) 'Delete MaskDC
End Sub
```

