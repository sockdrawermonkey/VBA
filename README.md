Function GrabMFT()
    Dim file_path As String
    file_path = "J:\SessionData\2019\ChenMed_2018_2019\Technology\Automations\C.Grab_Inbound_MFT.bat"
    Call Shell(file_path, vbMinimizedNoFocus)
End Function

Function HRIS_OPS()
Shell "J:\SessionData\2019\ChenMed_2018_2019\Technology\Automations\HRIS_ to_OPS.bat", vbMinimizedNoFocus
End Function

