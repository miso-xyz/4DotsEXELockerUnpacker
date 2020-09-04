Imports System.Reflection
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
Imports System.Xml
Imports System.Globalization
Public Class Unpacker
    Dim s_password As String
    Dim s_name As String
    Dim s_folder As String
    Public s_length As New List(Of String)
    Public s_file As New List(Of String)
    Function GetDumpedPassword() As String
        Return s_password
    End Function

    Function GetLastOutputFolder() As String
        Return s_folder
    End Function

    Function GetLastOutputSize(ByVal value As Integer) As String
        Return s_length(value)
    End Function

    Function GetLastFileName(ByVal value As Integer) As String
        Return s_file(value)
    End Function

    Function GetDumpedName() As String
        Return s_name
    End Function

    Public Function InitUnpack(ByVal packedEXEPath As String, ByVal autoOutput As Boolean, ByVal dumpPassword As Boolean, ByVal outputPath As String, ByVal password As String, Optional ByVal onlyGetSize As Boolean = False) As Integer
        GetNameAndPassword(packedEXEPath)
        s_name.Replace(".exe", Nothing)
        If dumpPassword Then
            If autoOutput Then
                If onlyGetSize Then
                    UnpackFiles(packedEXEPath, My.Application.Info.DirectoryPath & "\4dotsEXEPacker_Extracted\" & s_name, s_password, True)
                Else
                    UnpackFiles(packedEXEPath, My.Application.Info.DirectoryPath & "\4dotsEXEPacker_Extracted\" & s_name, s_password)
                End If
            Else
                If onlyGetSize Then
                    UnpackFiles(packedEXEPath, outputPath, s_password, True)
                Else
                    UnpackFiles(packedEXEPath, outputPath, s_password)
                End If

            End If
        Else
            If autoOutput Then
                If onlyGetSize Then
                    UnpackFiles(packedEXEPath, My.Application.Info.DirectoryPath & "\4dotsEXEPacker_Extracted\" & s_name, password, True)
                Else
                    UnpackFiles(packedEXEPath, My.Application.Info.DirectoryPath & "\4dotsEXEPacker_Extracted\" & s_name, password)
                End If

            Else
                If onlyGetSize Then
                    UnpackFiles(packedEXEPath, outputPath, password, True)
                Else
                    UnpackFiles(packedEXEPath, outputPath, password)
                End If

            End If
        End If
        Return 1
    End Function
    Public Sub UnpackFiles(ByVal packedEXEPath As String, ByVal outputFolder As String, ByVal password As String, Optional ByVal onlyGetSize As Boolean = False)
        Dim executingAssembly As Assembly = Assembly.LoadFile(packedEXEPath)
        Dim manifestResourceNames As String() = executingAssembly.GetManifestResourceNames()
        Dim path As String = My.Application.Info.DirectoryPath & "\4dotsEXEPacker_Extracted\" & IO.Path.GetFileNameWithoutExtension(s_name) & "\"
        If File.Exists(path) Then
            File.Delete(path)
        End If
        For i As Integer = 0 To manifestResourceNames.Length - 1
            If manifestResourceNames(i).IndexOf("LockedDocument.rtf") >= 0 Then
                Using binaryReader As BinaryReader = New BinaryReader(executingAssembly.GetManifestResourceStream(manifestResourceNames(i)))
                    Using memoryStream As MemoryStream = New MemoryStream()
                        While True
                            Dim num As Long = 32768L
                            Dim buffer As Byte() = New Byte(num - 1) {}
                            Dim num2 As Integer = binaryReader.Read(buffer, 0, CInt(num))
                            If num2 <= 0 Then
                                Exit While
                            End If
                            memoryStream.Write(buffer, 0, num2)
                        End While
                        Dim bMessage As Byte() = memoryStream.ToArray()
                        Dim bytes As Byte() = DecryptBytes(bMessage, password)
                        If onlyGetSize = False Then
                            If Directory.Exists(My.Application.Info.DirectoryPath & "\4dotsEXEPacker_Extracted") = False Then
                                Directory.CreateDirectory(My.Application.Info.DirectoryPath & "\4dotsEXEPacker_Extracted")
                            End If
                            Directory.CreateDirectory(path)
                            File.Create(path & s_name).Close()
                            File.WriteAllBytes(path & s_name, bytes)
                            s_folder = path & IO.Path.GetFileNameWithoutExtension(s_name) & s_name
                            Return
                        End If
                        s_file.Add(s_name)
                        s_length.Add(GetSizeInMemory(bytes.LongCount))
                    End Using
                End Using
                Exit For
            End If
        Next
    End Sub

    Public Function GetSizeInMemory(ByVal size As Long) As String


        Dim sizes() As String = {"B", "KB", "MB", "GB", "TB"}

        Dim len As Double = Convert.ToDouble(size)
        Dim order As Integer = 0
        While len >= 1024D And order < sizes.Length - 1
            order += 1
            len /= 1024
        End While
        Return String.Format(CultureInfo.CurrentCulture, "{0:0.##} {1}", len, sizes(order))
    End Function


    Public Sub GetNameAndPassword(ByVal packedEXEPath As String)
        Dim executingAssembly As Assembly = Assembly.LoadFile(packedEXEPath)
        Dim manifestResourceNames As String() = executingAssembly.GetManifestResourceNames()
        For i As Integer = 0 To manifestResourceNames.Length - 1
            If manifestResourceNames(i).IndexOf("project.xml") >= 0 Then
                Using binaryReader As BinaryReader = New BinaryReader(executingAssembly.GetManifestResourceStream(manifestResourceNames(i)))
                    Using memoryStream As MemoryStream = New MemoryStream()
                        While True
                            Dim num As Long = 32768L
                            Dim buffer As Byte() = New Byte(num - 1) {}
                            Dim num2 As Integer = binaryReader.Read(buffer, 0, CInt(num))
                            If num2 <= 0 Then
                                Exit While
                            End If
                            memoryStream.Write(buffer, 0, num2)
                        End While
                        Dim text As String = Encoding.[Default].GetString(memoryStream.ToArray())
                        text = DecryptString(text, "4dotsSoftware012301230123")
                        Dim xmlDocument As XmlDocument = New XmlDocument()
                        xmlDocument.LoadXml(text)
                        Dim xmlNode As XmlNode = xmlDocument.SelectSingleNode("//Project")
                        s_name = xmlNode.Attributes.GetNamedItem("Name").Value
                        s_password = xmlNode.Attributes.GetNamedItem("Password").Value
                    End Using
                End Using
            End If
        Next
    End Sub
    Public Shared Function DecryptBytes(ByVal bMessage As Byte(), ByVal Passphrase As String) As Byte()
        Dim utf8Encoding As UTF8Encoding = New UTF8Encoding()
        Dim md5CryptoServiceProvider As MD5CryptoServiceProvider = New MD5CryptoServiceProvider()
        Dim key As Byte() = md5CryptoServiceProvider.ComputeHash(utf8Encoding.GetBytes(Passphrase))
        Dim tripleDESCryptoServiceProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
        tripleDESCryptoServiceProvider.Key = key
        tripleDESCryptoServiceProvider.Mode = CipherMode.ECB
        tripleDESCryptoServiceProvider.Padding = PaddingMode.PKCS7
        Dim result As Byte()
        Try
            Dim cryptoTransform As ICryptoTransform = tripleDESCryptoServiceProvider.CreateDecryptor()
            result = cryptoTransform.TransformFinalBlock(bMessage, 0, bMessage.Length)
        Finally
            tripleDESCryptoServiceProvider.Clear()
            md5CryptoServiceProvider.Clear()
        End Try
        Return result
    End Function

    Public Shared Function DecryptString(ByVal Message As String, ByVal Passphrase As String) As String
        Dim utf8Encoding As UTF8Encoding = New UTF8Encoding()
        Dim md5CryptoServiceProvider As MD5CryptoServiceProvider = New MD5CryptoServiceProvider()
        Dim key As Byte() = md5CryptoServiceProvider.ComputeHash(utf8Encoding.GetBytes(Passphrase))
        Dim tripleDESCryptoServiceProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
        tripleDESCryptoServiceProvider.Key = key
        tripleDESCryptoServiceProvider.Mode = CipherMode.ECB
        tripleDESCryptoServiceProvider.Padding = PaddingMode.PKCS7
        Dim array As Byte() = Convert.FromBase64String(Message)
        Dim bytes As Byte()
        Try
            Dim cryptoTransform As ICryptoTransform = tripleDESCryptoServiceProvider.CreateDecryptor()
            bytes = cryptoTransform.TransformFinalBlock(array, 0, array.Length)
        Finally
            tripleDESCryptoServiceProvider.Clear()
            md5CryptoServiceProvider.Clear()
        End Try
        Return utf8Encoding.GetString(bytes)
    End Function
End Class
