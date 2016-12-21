Imports System.Security.AccessControl
Imports System.Security.Principal
Imports System.IO.Path
Imports System.DirectoryServices
Imports System.IO

Public Class Form1
	Private ImageFolder As String
	Private Sid As String
	Private UserImage As Image
	Private Sub SetupPersonalImageFolder()
		Try
			Dim FolderAcl As New DirectorySecurity
			FolderAcl.AddAccessRule(New FileSystemAccessRule(My.User.Name, FileSystemRights.Modify, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))

			If IO.Directory.Exists(ImageFolder) Then
				Dim FolderInfo As IO.DirectoryInfo = New IO.DirectoryInfo(ImageFolder)
				FolderInfo.SetAccessControl(FolderAcl)
			Else
				IO.Directory.CreateDirectory(ImageFolder, FolderAcl)
			End If
		Catch ex As Exception
			'do nothing
		End Try
	End Sub

	Private Sub WriteFilesAndRegistry()
		Dim image_sizes As Int16() = New Int16() {32, 40, 48, 96, 192, 200, 240, 448}
		Dim ImageName As String
		Dim ImageFileName As String
		Dim i As Integer
		Dim RegKeyPath As String = String.Format("SOFTWARE\Microsoft\Windows\CurrentVersion\AccountPicture\Users\{0}", Sid)
		Dim RegKey As Microsoft.Win32.RegistryKey
		Dim RegValueName As String
		For i = 0 To image_sizes.Length - 1
			Try

				ImageName = String.Format("Image{0}.jpg", image_sizes(i))
				ImageFileName = IO.Path.Combine(ImageFolder, ImageName)
				UserImage.Save(ImageFileName)
				Dim RK As Microsoft.Win32.RegistryKey
				If Environment.Is64BitOperatingSystem Then
					RK = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64)
				Else
					RK = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32)
				End If

				RegKey = RK.OpenSubKey(RegKeyPath, True)
				RegValueName = String.Format("Image{0}", image_sizes(i))
				RegKey.SetValue(RegValueName, ImageFileName)
			Catch ex As Exception
				'do nothing
			End Try
		Next
	End Sub

	Private Sub WriteUserTile()
		Dim ImageName As String
		Dim ImageFileName As String
		ImageName = "UserImage.jpg"
		Try
			ImageFileName = IO.Path.Combine(GetTempPath, ImageName)
			UserImage.Save(ImageFileName)
			SetUserTile(System.Security.Principal.WindowsIdentity.GetCurrent().Name, 0, ImageFileName)
			IO.File.Delete(ImageFileName)
		Catch ex As Exception
			'do nothing
		End Try
	End Sub


	Private Sub GetUserInfos()
		Dim folder As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)
		Sid = DirectCast(DirectCast(My.User.CurrentPrincipal, System.Security.Principal.WindowsPrincipal).Identity, WindowsIdentity).User.ToString
		ImageFolder = IO.Path.Combine(folder, "AccountPictures", Sid)
		Dim ADSSearcher As DirectorySearcher = New DirectorySearcher()
		ADSSearcher.Filter = String.Format("(&(objectClass=user) (objectSid={0}))", Sid)
		Dim result As SearchResult = ADSSearcher.FindOne()
		If Not result Is Nothing Then
			Dim user As DirectoryEntry = New DirectoryEntry(result.Path)
			'		Dim photo As Byte() = user.Properties("jpegPhoto").Value
			Dim photo As Byte() = DirectCast(user.Properties("thumbnailPhoto").Value, Byte())
			If Not photo Is Nothing Then
				Dim ms As MemoryStream = New MemoryStream(photo)
				UserImage = Bitmap.FromStream(ms)
			End If
		End If
	End Sub


	Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
		Try
			GetUserInfos()
			If Not UserImage Is Nothing Then
				If (Environment.OSVersion.Version.Major > 6) Or ((Environment.OSVersion.Version.Major = 6) And (Environment.OSVersion.Version.Minor > 1)) Then
					' W8 or newer
					WriteFilesAndRegistry()
				Else
					WriteUserTile()
				End If
			End If
		Catch ex As Exception
			' do nothing
		End Try
		Close()
	End Sub




	<System.Runtime.InteropServices.DllImport("shell32.dll", EntryPoint:="#262", CharSet:=Runtime.InteropServices.CharSet.Unicode, PreserveSig:=False)>
	Shared Sub SetUserTile(ByVal strUserName As String, ByVal intWhatever As Integer, ByVal strPicPath As String)
	End Sub



End Class




