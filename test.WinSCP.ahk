#Include WinSCP.ahk

/*
;~ ---------------------------------------------------
;~ Open Connection using plain FTP
;~ ---------------------------------------------------
FTPSession := new WinSCP
try
  FTPSession.OpenConnection("ftp://myserver.com","username","password")
} catch e
  msgbox % "Oops. . . Something went wrong``n" e.Message
*/

/*
;~ ---------------------------------------------------
;~ Open Connection using SSL FTP
;~ ---------------------------------------------------
FTPSession := new WinSCP
try
{
  FTPSession.Hostname		:= "ftp://myserver.com"
  FTPSession.Protocol 		:= WinSCPEnum.FtpProtocol.Ftp
  FTPSession.Secure 		:= WinSCPEnum.FtpSecure.ExplicitSsl
  FTPSession.User			:= "MyUserName"
  FTPSession.Password		:= "P@ssw0rd"
  FTPSession.Fingerprint    := "xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx" ;set to false to ignore server certificate
  FTPSession.OpenConnection()
} catch e
  msgbox % "Oops. . . Something went wrong``n" e.Message
*/

/*
;~ ---------------------------------------------------
;~ Upload File
;~ ---------------------------------------------------
FTPSession := new WinSCP
try
{
  FTPSession.OpenConnection("ftp://myserver.com","username","password")
  
  fName := "Windows10_InsiderPreview_x64_EN-US_10074.iso"
  fPath := "C:\temp"
  tPath := "/Win10beta/"
  if (!FTPSession.FileExists(tPath))
  	FTPSession.CreateDirectory(tPath)
  FTPSession.PutFiles(fPath "\" fName, tPath)
} catch e
  msgbox % "Oops. . . Something went wrong``n" e.Message
*/

/*
;~ ---------------------------------------------------
;~ Download File
;~ ---------------------------------------------------
FTPSession := new WinSCP
try
{
  FTPSession.OpenConnection("ftp://myserver.com","username","password")
  
  fName := "Windows10_InsiderPreview_x64_EN-US_10074.iso"
  lPath := "C:\temp"
  rPath := "/Win10beta/"
  if (FTPSession.FileExists(rPath "/" fName))
  	FTPSession.GetFiles(rPath "/" fName, lPath)
} catch e
  msgbox % "Oops. . . Something went wrong``n" e.Message  
*/

/*
;~ ---------------------------------------------------
;~ Get File Information
;~ ---------------------------------------------------
FTPSession := new WinSCP
try
{
  FTPSession.OpenConnection("ftp://myserver.com","username","password")
  
  FileCollection := t.ListDirectory("/")
  for file in FileCollection.Files {
	if (file.Name != "." && file.Name != "..")
	  msgbox % "Name: " file.Name "``nPermission: " file.FilePermissions.Octal "``nIsDir: " file.IsDirectory "``nFileType: " file.FileType "``nGroup: " file.Group "``nLastWriteTime: " file.LastWriteTime "``nLength: " file.Length "``nLength32: " file.Length32 "``nOwner: " file.Owner
} catch e
  msgbox % "Oops. . . Something went wrong``n" e.Message 
*/

/*
;~ ---------------------------------------------------
;~ Upload File with Progressbar
;~ ---------------------------------------------------
Gui, Font, S18 CDefault Bold, Arial
Gui, Add, Text, x17 y30 w450 h40 +Center vtxtTitle, Uploading… @ Speed
Gui, Add, GroupBox, x17 ys+35 w450 h190 , 
Gui, Font, , 
Gui, Add, Text, x42 ys+60 , Filename:
Gui, Add, Edit, x+40 w300 h20 vedtFileName Disabled, Edit
Gui, Add, Progress, x36 yp+30 cBlue BackgroundCCCCCC w400 h20 vproFileName, 1
Gui, Add, Text, x42 yp+30 w80 h20 , Overall
Gui, Add, Progress, x36 yp+30 w400 h20 cRed BackgroundCCCCCC vproOverall, 0
Gui, Add, Button, x222 yp+30 w100 h30 Disabled vbtnClose gbtnClose, Close
Gui, Add, Button, x+20 w100 h30 vbtnAbort gbtnAbort, Abort

session_FileTransferProgress(sender, e)
{
  RegExMatch(e.FileName, ".*\\(.+?)$", match)
  FileName        := match1
  CPS             := Round(e.CPS / 1024)
  FileProgress    := Round(e.FileProgress * 100)
  OverallProgress := Round(e.OverallProgress * 100)
  action          := (e.Side==0) ? "Uploading" : "Downloading"
	
  GuiControl,, txtTitle, % action " @ " CPS " kbps"
  GuiControl,, edtFileName, % FileName
  GuiControl,, proFileName, % FileProgress
  GuiControl,, proOverall, % OverallProgress
  if (OverallProgress==100)
    GuiControl, Enable, btnClose
	
  Gui, Show, , File Transfere
}

FTPSession := new WinSCP
try
{
  FTPSession.OpenConnection("ftp://myserver.com","username","password")
  
  fName := "Windows10_InsiderPreview_x64_EN-US_10074.iso"
  fPath := "C:\temp"
  tPath := "/Win10beta/"
  if (!FTPSession.FileExists(tPath))
  	FTPSession.CreateDirectory(tPath)
  FTPSession.PutFiles(fPath "\" fName, tPath)
} catch e
  msgbox % "Oops. . . Something went wrong``n" e.Message
*/