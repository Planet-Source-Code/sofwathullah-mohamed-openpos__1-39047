Attribute VB_Name = "Module2"
'// ===============================================================================//
'// Program: OpenPOS(Point of Sales)                                               //
'// Developed by: Sofwathullah Mohamed                                             //
'// Sofwath@Hotmail.Com                                                            //
'// You are free to use and modify this program as long as you give credit to the  //
'// original developer. Any comments or bugs report to sofwath@hotmail.com         //
'// Ver: 0.1                                                                       //
'// This Program is Still Under Development and Some of the Modules are Missing    //
'// ===============================================================================//

Global Const DEFSOURCE = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="
Public db As ADODB.Connection
Public Sub OpenDB()
        DataEnvironment1.Connection1 = DEFSOURCE & App.Path & "\setDB.mdb;"
End Sub
