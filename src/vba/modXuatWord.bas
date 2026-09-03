Attribute VB_Name = "modXuatWord"
Option Explicit

' The Word parser and export pipeline live in modWordExport.
' Keep this module as the thin UI-facing coordinator only.
Public Sub XuatVanBan()
    Call RunWordExport(False)
End Sub
