Attribute VB_Name = "first"

Public Sub say_hello()
  MsgBox "Hello world!", , "title"

End Sub

Public Sub rng_B2()
  Range("B2").Value = "B2"

  Range("B3").Value = fn_hello()

End Sub

Private Function fn_hello()
  fn_hello = "hello miero modify"

End Function
