Attribute VB_Name = "nVb_ObjColl"
Option Compare Database
Option Explicit

Function ObjCollAy(ObjColl, OAy)
Dim I
For Each I In ObjColl
    PushObj OAy, I
Next
ObjCollAy = OAy
End Function
