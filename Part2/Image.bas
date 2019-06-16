Private Sub UserForm_Initialize()
'з об'єктом Image1
With Image1
    .Picture = LoadPicture("d:\WINDOWS\Паркет.bmp") 'завантажити картинку
    .PictureSizeMode = fmPictureSizeModeClip 'розмір картинки
    .PictureTiling = True 'замостити
    .PictureAlignment = fmPictureAlignmentCenter 'вирівнювання
End With
End Sub
