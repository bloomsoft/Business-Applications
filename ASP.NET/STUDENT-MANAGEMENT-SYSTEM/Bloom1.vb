Module Bloom1
    Const ServerFolder As String = "d:\hosting\whiteroseschoo"
    Const ServerName As String = "Data Source=p3swhsql-v05.shr.phx3.secureserver.net; Initial Catalog=zakria; User ID=zakria; Password='Abcdef123';"
    'Const ServerName As String = "workstation id=HRBLM;packet size=4096;integrated security=SSPI;data source=HRBLM;persist security info=False;initial catalog=zakria"
    'Const ServerFolder As String = "K:\"
    'Const ServerName As String = "C:\Inetpub\wwwroot\Whiterose\Access_Main.dsn"
    Public Function ConnectionString() As String
        Dim C As String
        'C = "workstation id=" & ServerName & ";packet size=4096;integrated security=SSPI;initial catalo" & _
        '"g=OTS;persist security info=False"
        'C = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & ServerName & ";User Id=admin; Password="
        C = ServerName
        Return C
    End Function
    Public Function ServerFolderPath() As String
        Dim C As String
        'C = "workstation id=" & ServerName & ";packet size=4096;integrated security=SSPI;initial catalo" & _
        '"g=OTS;persist security info=False"
        C = ServerFolder

        Return C
    End Function
End Module
