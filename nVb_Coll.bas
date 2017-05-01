Attribute VB_Name = "nVb_Coll"
Option Compare Database
Option Explicit

Function CollOy(Coll, OOy)
Erase OOy
Dim I
For Each I In Coll
    PushObj OOy, I
Next
CollOy = OOy
End Function
