Imports Microsoft.Graph
Imports Microsoft.Identity.Client

Module Module1

    Private Const client_id As String = "{client id -- also known as application id}" '<-- enter the client_id guid here
    Private Const tenant_id As String = "{tenant id or name}" '<-- enter either your tenant id here
    Private authority As String = $"https://login.microsoftonline.com/{tenant_id}"

    Private _scopes As New List(Of String)
    Private ReadOnly Property scopes As List(Of String)
        Get
            If _scopes.Count = 0 Then
                _scopes.Add("User.read") '<-- add each scope you want to send as a seperate .add
            End If
            Return _scopes
        End Get
    End Property

    ''' <summary>
    ''' underlaying variable for the readonly property PCA which returns an instance of the PublicClientApplication
    ''' </summary>
    Private _pca As IPublicClientApplication = Nothing
    Private ReadOnly Property PCA As IPublicClientApplication
        Get
            If _pca Is Nothing Then
                _pca = PublicClientApplicationBuilder.Create(client_id).WithAdfsAuthority(authority).Build()
            End If
            Return _pca
        End Get
    End Property

    ''' <summary>
    ''' The underlaying variable for the Authentication Provider readonly property
    ''' </summary>
    Private _authProvider As InteractiveAuthenticationProvider = Nothing
    Private ReadOnly Property AuthProvider As InteractiveAuthenticationProvider
        Get
            If _authProvider Is Nothing Then
                _authProvider = New InteractiveAuthenticationProvider(PCA, scopes)
            End If
            Return _authProvider
        End Get
    End Property

    ''' <summary>
    ''' The underlaying variable for the graphClient readonly property
    ''' </summary>
    Private _graphClient As GraphServiceClient = Nothing
    Private ReadOnly Property GraphClient As GraphServiceClient
        Get
            If _graphClient Is Nothing Then
                _graphClient = New GraphServiceClient(AuthProvider)
            End If
            Return _graphClient
        End Get
    End Property


    Sub Main()

        Get_Me()


        Console.ReadKey()

    End Sub

    Private Async Sub Get_Me()
        Dim user As User

        'Using the Select to get the employeeId as the v1 endpoint does not automatically return that
        user = Await GraphClient.Me().Request().Select("displayName,employeeid").GetAsync()
        Console.WriteLine($"User = {user.DisplayName}, employeeid = {user.EmployeeId}")


    End Sub

End Module
