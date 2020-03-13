;~ -----------------------------------------------------------------------------
;~ Name:               WinSCP Class
;~ http://lipkau.github.io/WinSCP.ahk/
;~ -----------------------------------------------------------------------------
;~ Version:            1.1.0
;~ Date:               2015-06-20
;~ Author:             Lipkau, Oliver <oliver@lipkau.net>
;~                     https://github.com/lipkau
;~                     http://oliver.lipkau.net/blog
;~ -----------------------------------------------------------------------------
;~ TODO:
;~ 		- Try/Catch if dll is not registered?
;~ 		- Allow dll to be outside A_ScriptDir?
;~ 		- better documentation
;~ 		- more methods
;~ 		- more events
;~ 		- cleanup


;~ -----------------------------------------------------
;~  Enumerations
;~ -----------------------------------------------------
/*
Description
	Reproduce WinSCP Enumerations.

Scope
	Global
*/
class WinSCPEnum
{
	static FtpMode                 := {Passive:0, Active:1}
	static FtpSecure               := {None:0, Implicit:1, ExplicitTls:2, ExplicitSsl:3}
	static FtpProtocol             := {Sftp:0, Scp:1, Ftp:2}
	static TransferMode            := {Binary:0, Ascii:1, Automatic:2}
	static SynchronizationMode     := {Local:0, Remote:1, Both:2}
	static SynchronizationCriteria := {None:0, Time:1, Size:2, Either:3}
}

;~ -----------------------------------------------------
;~  Event Handlers
;~ -----------------------------------------------------
/*
Description
	Example on how to use the FileTransferred Event
*/
;~ session_FileTransferred(sender, e)
;~ {
	;~ MsgBox % e.FileName " => "  e.Destination
;~ }

/*
Description
	Example on how to use the Failed Event
*/
;~ session_Failed(sender, e)
;~ {
	;~ MsgBox % e.FileName " => "  e.Destination
;~ }

/*
Description
	Example on how to use the RemovalEventArgs Event
	http://winscp.net/eng/docs/library_removaleventargs
*/
;~ session_RemovalEventArgs(sender, e)
;~ {
	;~ if (e.Error)
		;~ MsgBox % "Failed to remove " e.FileName " => "  e.Error
	;~ else
		;~ MsgBox % "Removed " e.FileName
;~ }

/*
Description
	Example on how to use the FileTransferProgress Event to display the Progress using AHK GUI
*/
;~ session_FileTransferProgress(sender, e)
;~ {
	;~ RegExMatch(e.FileName, ".*\\(.+?)$", match)
	;~ FileName        := match1
	;~ CPS             := Round(e.CPS / 1024)
	;~ FileProgress    := Round(e.FileProgress * 100)
	;~ OverallProgress := Round(e.OverallProgress * 100)
	;~ action          := (e.Side==0) ? "Uploading" : "Downloading"
;~ }

;~ -----------------------------------------------------
;~  Class
;~ -----------------------------------------------------
class WinSCP
{
	;~ -----------------------------------------------------
	;~  Properties
	;~ -----------------------------------------------------
	Hostname                := ""
	Port                    := ""
	Secure                  := ""
	Protocol                := ""
	User                    := ""
	Password                := ""
	Mode                    := ""
	TransferMode            := ""
	SynchronizationMode     := ""
	SynchronizationCriteria := ""
	Fingerprint             := ""

	;http://winscp.net/eng/docs/library_session
	Session                 := ComObjCreate("WinSCP.Session")
	SessionOptions          := ComObjCreate("WinSCP.SessionOptions")
	TransferOptions         := ComObjCreate("WinSCP.TransferOptions")
	;http://winscp.net/eng/docs/library_filepermissions
	FilePermissions         := ComObjCreate("WinSCP.FilePermissions")

	/*
	Description
		Create a Property for Log Path. Class Session should be transparent in Class FTP.
	*/
	LogPath[]
	{
		Get
		{
			return this.Session.SessionLogPath
		}
		Set
		{
			return this.Session.SessionLogPath := value
		}
	}

	/*
	Description
		Create a Property for the HomePath. Class Session should be transparent in Class FTP.
	*/
	HomePath[]
	{
		Get
		{
			return this.Session.HomePath
		}
		Set
		{
			return this.Session.HomePath := value
		}
	}

	/*
	Description
		Create a Property for the Opened. Class Session should be transparent in Class FTP.
	*/
	Connected[]
	{
		Get
		{
			return this.Session.Opened
		}
	}

	;~ -----------------------------------------------------
	;~  Methods
	;~ -----------------------------------------------------
	/*
	Description
		Constructor - sets default values
	*/
	__New()
	{
		global WinSCPEnum ; WinSCP Enums

		;~ Set defaults
		this.Port                    := 21
		this.Secure                  := WinSCPEnum.FtpSecure.None
		this.Protocol                := WinSCPEnum.FtpProtocol.Ftp
		this.Mode                    := WinSCPEnum.FtpMode.Passive
		this.TransferMode            := WinSCPEnum.TransferMode.Binary
		this.SynchronizationMode     := WinSCPEnum.SynchronizationMode.Local
	    this.SynchronizationCriteria := WinSCPEnum.SynchronizationCriteria.Time
		this.Fingerprint             := false
		ComObjConnect(this.Session, "session_")
	}

	/*
	Description
		Destructor
	*/
	; __Delete()
	; {
	; 	this.Dispose()
	; 	this.CloseConnection()
	; }

	/*
	Description
		Open connection to server

	Input
		srv        : [string] Server DNS or IP
		uName      : [string] User Name
		pWord      : [string] Password
	*/
	OpenConnection(srv="", uName="", pWord="")
	{
		global WinSCPEnum ; WinSCP Enums

		this.SessionOptions.HostName := (srv)   ? srv   : this.Hostname
		this.SessionOptions.UserName := (uName) ? uName : this.User
		this.SessionOptions.Password := (pWord) ? pWord : this.Password

		;~ FTP Mode
		this.SessionOptions.FtpMode := this.Mode

		;~ FTP Protocol
		this.SessionOptions.Protocol := this.Protocol

		;~ FTP Security
		this.SessionOptions.FtpSecure := this.Secure

		;~ Fingerprint
		IsEncrypted := this.Secure
		Encryptions := WinSCPEnum.FtpSecure.Implicit "," WinSCPEnum.FtpSecure.ExplicitTls "," WinSCPEnum.FtpSecure.ExplicitSsl
		if IsEncrypted in %Encryptions%
		{
			if (StrLen(this.Fingerprint) > 1)
			{
				if (this.Protocol==WinSCPEnum.FtpProtocol.Scp || this.Protocol==WinSCPEnum.FtpProtocol.Sftp)
					this.SessionOptions.SshHostKeyFingerprint := this.Fingerprint
				else
					this.SessionOptions.TlsHostCertificateFingerprint := this.Fingerprint
			} else {
				if (this.Protocol==WinSCPEnum.FtpProtocol.Scp || this.Protocol==WinSCPEnum.FtpProtocol.Sftp)
					this.SessionOptions.GiveUpSecurityAndAcceptAnySshHostKey := true
				else
					this.SessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate := true
			}
		}

		this.Session.Open(this.SessionOptions)
	}

	/*
	Description
		Closes session.
		New session can be opened using Session.Open using the same instance of Session.
	*/
	CloseConnection()
	{
		try
			this.Session.Close()
	}

	/*
	Description
		If session was opened, closes it, terminates underlying WinSCP process, deletes XML log file and disposes object.
	*/
	Dispose()
	{
		try
			this.Session.Dispose()
	}

	/*
	Description:
		Lists the contents of specified remote directory.
		http://winscp.net/eng/docs/library_session_listdirectory

	Input:
		remotePath : [string] Full path to remote directory to be read.

	Output:
		[RemoteDirectoryInfo]
	*/
	ListDirectory(remotePath)
	{
		return this.Session.ListDirectory(remotePath)
	}

	/*
	Description:
		Creates remote directory.
		http://winscp.net/eng/docs/library_session_createdirectory

	Input:
		remotePath : [string] Full path to remote directory to create.

	Output:
		[void]
	*/
	CreateDirectory(remotePath)
	{
		return this.Session.CreateDirectory(remotePath)
	}

	/*
	Description:
		Checks for existence of remote file.
		http://winscp.net/eng/docs/library_session_fileexists

	Input:
		remotePath : [string] Full path to remote file. Note that you cannot use wild-cards here.

	Output:
		[bool] true if file exists, false otherwise.
	*/
	FileExists(remotePath)
	{
		return this.Session.FileExists(remotePath)
	}

	/*
	Description:
		Retrieves information about remote file.
		http://winscp.net/eng/docs/library_session_getfileinfo

	Input:
		remotePath : [string] Full path to remote file.

	Output:
		[RemoteFileInfo]
		http://winscp.net/eng/docs/library_remotefileinfo
	*/
	GetFileInfo(remotePath)
	{
		return this.Session.GetFileInfo(remotePath)
	}

	/*
	Description:
		Downloads one or more files from remote directory to local directory.
		http://winscp.net/eng/docs/library_session_getfiles

		This Method supports Wild-cards
		http://winscp.net/eng/docs/library_wildcard

	Input:
		remotePath : [string]          Full path to remote directory followed by slash and wild-card to
		                               select files or subdirectories to download. When wild-card is
							           omitted (path ends with slash), all files and subdirectories in
							           the remote directory are downloaded.
		localPath  : [string]          Full path to download the file to. When downloading multiple
		                               files, the filename in the path should be replaced with operation
							           mask or omitted (path ends with backslash).
		remove     : [bool]            When set to true, deletes source remote file(s) after transfer.
		                               Defaults to false.
		options    : [TransferOptions] Transfer options. Defaults to null, what is equivalent to new
		                               TransferOptions().

	Output:
		[TransferOperationResult]
		See also Capturing results of operations.
	*/
	GetFiles(remotePath, localPath, remove=false)
	{
		;~ Check
		;remove
		if remove not in 0,1
			throw "Invalid value for remove"

		return this.Session.GetFiles(remotePath, localPath, remove, this.TransferOptions)
	}

	/*
	Description:
		Uploads one or more files from local directory to remote directory.
		http://winscp.net/eng/docs/library_session_putfiles

		This Method supports Wild-cards
		http://winscp.net/eng/docs/library_wildcard

	Input:
		localPath  : [string]          Full path to local file or directory to upload. Filename in the
		                               path can be replaced with Windows wildcard1) to select multiple
					 		           files. When file name is omitted (path ends with backslash), all
							           files and subdirectories in the local directory are uploaded.
		remotePath : [string]          Full path to upload the file to. When uploading multiple files,
		                               the filename in the path should be replaced with operation mask
							           or omitted (path ends with slash).
		remove     : [bool]            When set to true, deletes source local file(s) after transfer.
		                               Defaults to false.
		options    : [TransferOptions] Transfer options. Defaults to null, what is equivalent to new
		                               TransferOptions().

	Output:
		[TransferOperationResult]
		See also Capturing results of operations.
	*/
	PutFiles(localPath, remotePath, remove:=false)
	{
		;~ Checks
		;localPath
		if (!FileExist(localPath))
			throw "Could not find " localPath

		;remove
		if remove not in 0,1
			throw "Invalid value for remove"

		return this.Session.PutFiles(localPath, remotePath, remove, this.TransferOptions)
	}

	/*
	Description:
		Moves remote file to another remote directory and/or renames remote file.
		http://winscp.net/eng/docs/library_session_movefile

	Input:
		sourcePath : [string] Full path to remote file to move/rename.
		targetPath : [string] Full path to new location/name to move/rename the file to.

	Output:
		[void]
	*/
	MoveFile(sourcePath, targetPath)
	{
		return this.Session.MoveFile(sourcePath, targetPath)
	}

	/*
	Description:
		Removes one or more remote files.
		http://winscp.net/eng/docs/library_session_removefiles

		This Method supports Wild-cards
		http://winscp.net/eng/docs/library_wildcard

	Input:
		remotePath : [string] Full path to remote directory followed by slash and
		                      wild-card to select files or subdirectories to remove.

	Output:
		[RemovalOperationResult]
		See also Capturing results of operations.
	*/
	RemoveFiles(remotePath)
	{
		return this.Session.RemoveFiles(remotePath)
	}

	/*
	Description:
		Synchronizes directories.
		http://winscp.net/eng/docs/library_session_synchronizedirectories

	Input:
		SynchronizationMode     : [enum]   Synchronization mode. Possible values are SynchronizationMode.Local,
		                                   SynchronizationMode.Remote and SynchronizationMode.Both.
		localPath               : [string] Full path to local directory.
		remotePath              : [string] Full path to remote directory.
		removeFiles             : [bool]   When set to true, deletes obsolete files. Cannot be used for
		                                   SynchronizationMode.Both.
		mirror                  : [bool]   When set to true, synchronizes in mirror mode (synchronizes also
		                                   older files). Cannot be used for SynchronizationMode.Both. Defaults
										   to false.
		SynchronizationCriteria : [enum]   Comparison criteria. Possible values are SynchronizationCriteria.None,
										   SynchronizationCriteria.Time (default), SynchronizationCriteria.Size
										   and SynchronizationCriteria.Either. For SynchronizationMode.Both
										   SynchronizationCriteria.Time can be used only.

	Output:
		[SynchronizationResult]
		See also Capturing results of operations.
	*/
	SynchronizeDirectories(SynchronizationMode, localPath, remotePath, removeFiles, mirror=false, SynchronizationCriteria=1)
	{
		global WinSCPEnum ; WinSCP Enums

		;~ Checks
		;Mode
		if SynchronizationMode not in % this.StringJoin(WinSCPEnum.SynchronizationMode,",")
			throw "Invalid SynchronizationMode"

		;localPath
		if (!FileExist(localPath))
			throw "Could not find " localPath

		;removeFiles
		if removeFiles not in 0,1
			throw "Invalid removeFiles"
		if (removeFiles && (removeFiles==WinSCPEnum.SynchronizationMode.Both))
			throw "Deletion of obsolete files cannot be used for SynchronizationMode.Both"

		;mirror
		if mirror not in 0,1
			throw "Invalid mirror"
		if (mirror && (removeFiles==WinSCPEnum.SynchronizationMode.Both))
			throw "Synchronization in mirror mode (synchronizes also older files) cannot be used for SynchronizationMode.Both"

		;SynchronizationCriteria
		if SynchronizationCriteria not in % this.StringJoin(WinSCPEnum.SynchronizationCriteria,",")
			throw "Invalid SynchronizationCriteria"

		;TransferOptions
		if (options && (ComObjType(options,"Name")=="_TransferOptions"))
			throw "Invalid TransferOptions"

		if (!this.TransferOptions.TransferMode)
			this.SetTransferOptions(, ,, ,, WinSCPEnum.TransferMode.Binary)

		this.Session.SynchronizeDirectories(SynchronizationMode, localPath, remotePath, removeFiles, mirror, SynchronizationCriteria, this.TransferOptions)
	}

	;~ -----------------------------------------------------
	;~  Helper Methods
	;~ -----------------------------------------------------
	/*
	Description
		Set TransferOptions
		http://winscp.net/eng/docs/library_transferoptions

	Input
		FileMask          : [string]                FileMask
		FilePermissions   : [FilePermissions]       Permissions to applied to a remote file (used for
		                                            uploads only). Use default null to keep default permissions.
		PreserveTimestamp : [bool]                  Preserve timestamps (set last write time of destination file
													to that of source file). Defaults to true.
                                                    When used with Session.SynchronizeDirectories, timestamps is
												    always preserved, disregarding property value, unless criteria
													parameter is SynchronizationCriteria.None or
													SynchronizationCriteria.Size.
		ResumeSupport     : [TransferResumeSupport] Configures automatic resume/transfer to temporary filename.
		                                            Read-only (set properties of returned TransferResumeSupport
													instance).
		SpeedLimit        : [int]                   Limit transfer speed (in KB/s).
		TransferMode      : [TransferMode]          Transfer mode. Possible values are TransferMode.Binary
		                                            (default), TransferMode. ASCII and TransferMode.Automatic
													(based on file extension).
	*/
	SetTransferOptions(FileMask="", FilePermissions="", PreserveTimestamp=true, ResumeSupport="", SpeedLimit="", TransferMode=0)
	{
		global WinSCPEnum ; WinSCP Enums

		;~ Check
		;FilePermissions
		if (FilePermissions && (ComObjType(permissions,"Name")=="_FilePermissions"))
			throw "Invalid FilePermissions"

		;PreserveTimestamp
		if PreserveTimestamp not in 0,1
			throw "Invalid value for PreserveTimestamp"

		;permissions
		if (ResumeSupport && (ComObjType(ResumeSupport,"Name")=="_TransferResumeSupport"))
			throw "Invalid ResumeSupport"

		;SpeedLimit
		if SpeedLimit is not Integer
			throw "Invalid SpeedLimit"

		;TransferMode
		if TransferMode not in % this.StringJoin(WinSCPEnum.TransferMode,",")
			throw "Invalid TransferMode"

		this.TransferOptions.FileMask 			:= (FileMask) ? FileMask : ""
		this.TransferOptions.FilePermissions 	:= (FilePermissions) ? FilePermissions: ""
		this.TransferOptions.PreserveTimestamp 	:= (PreserveTimestamp) ? PreserveTimestamp : ""
		this.TransferOptions.ResumeSupport	 	:= (ResumeSupport) ? ResumeSupport : ""
		this.TransferOptions.SpeedLimit	 		:= (SpeedLimit) ? SpeedLimit : ""
		this.TransferOptions.TransferMode  		:= (TransferMode) ? TransferMode : ""
	}

	/*
	Description
		Converts special characters in file path to make it unambiguous file mask/wild-card.
		http://winscp.net/eng/docs/library_session_escapefilemask

	Input
		fileMask : [string] File path to convert.
	*/
	EscapeFileMask(FileMask)
	{
		return this.Session.EscapeFileMask(FileMask)
	}

	/*
	Description
		Join Array to String

	Input
		array : [array]  Array that will be joined into string
		delim : [string] Character (set) that will be used between array elements in string

	Output
		[string]
	*/
	StringJoin(array, delim=";")
	{
		t := ""
		for key, value in array
			t .= value "" delim
		return SubStr(t,1, -StrLen(delim))
	}
}
