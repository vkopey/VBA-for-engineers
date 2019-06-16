'Процедура обробки події Click (натиск лівої кнопки миші)
Private Sub CommandButton1_Click()
CommandButton1.Enabled = True 'доступність
CommandButton1.Locked = False 'заблокованість
CommandButton1.Caption = "Click!" 'надпис
CommandButton1.AutoSize = True 'авторозмір
CommandButton1.Cancel = True 'асоціація з клавішею Esc
CommandButton1.Default = True 'асоціація з клавішею Enter
CommandButton1.Accelerator="A" 'клавіша-акселератор Alt-A
'фонова картинка
CommandButton1.Picture = LoadPicture("d:\WINDOWS\Паркет.bmp")
'позиція картинки
CommandButton1.PicturePosition = fmPicturePositionAboveLeft
End Sub

'Процедура обробки події DblClick (подвійний натиск лівої кнопки миші)
Private Sub CommandButton1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton1.Caption = "DblClick!" 'надпис
End Sub
