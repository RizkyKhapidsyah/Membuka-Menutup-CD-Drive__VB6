Attribute VB_Name = "Module1"
'Ketika Anda mengklik tombol Command1, CD-Drive akan 'terbuka. Untuk menutupnya kembali, klik tombol 'Command2.

Declare Function mciSendString Lib "winmm.dll" Alias _
"mciSendStringA" (ByVal lpstrCommand As String, ByVal _
lpstrReturnString As String, ByVal uReturnLength As _
Long, ByVal hwndCallback As Long) As Long


