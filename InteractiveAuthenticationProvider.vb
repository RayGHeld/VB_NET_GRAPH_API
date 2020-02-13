Imports System.Net.Http
Imports Microsoft.Graph
Imports Microsoft.Identity.Client

''' <summary>
''' This is a custom class to implement the IAuthenticationProvider class.  You can implement the new preview Microsoft.Graph.Auth which has 
''' built in classes for the authentication providers:  https://github.com/microsoftgraph/msgraph-sdk-dotnet-auth
''' </summary>
Public Class InteractiveAuthenticationProvider
    Implements IAuthenticationProvider

    Private Property Pca As IPublicClientApplication
    Private Property Scopes As List(Of String)

    Private Sub New()
        'Intentionally left blank to prevent empty constructor
    End Sub

    ''' <summary>
    ''' The constructor for this custom implementation of the IAuthenticationProvider
    ''' </summary>
    ''' <param name="pca">The public client application -- in this example, I have pre-set this up prior to creating the auth provider</param>
    ''' <param name="scopes">The scopes for the request</param>
    Public Sub New(pca As IPublicClientApplication, scopes As List(Of String))
        Me.Pca = pca
        Me.Scopes = scopes
    End Sub

    ''' <summary>
    ''' This is the required implmentation of the AuthenticateRequestAsync Method for the IAuthenticationProvider interface
    ''' </summary>
    ''' <param name="request">The current graph request being made</param>
    ''' <returns></returns>
    Public Async Function AuthenticateRequestAsync(request As HttpRequestMessage) As Task Implements IAuthenticationProvider.AuthenticateRequestAsync
        Dim accounts As IEnumerable(Of IAccount)
        Dim result As AuthenticationResult = Nothing

        accounts = Await Pca.GetAccountsAsync()
        Dim interactionRequired As Boolean = False

        Try
            result = Await Pca.AcquireTokenSilent(Scopes, accounts.FirstOrDefault).ExecuteAsync()
        Catch ex1 As MsalUiRequiredException
            interactionRequired = True
        Catch ex2 As Exception
            Console.WriteLine($"Authentication error: {ex2.Message}")
        End Try

        If interactionRequired Then
            Try
                result = Await Pca.AcquireTokenInteractive(Scopes).ExecuteAsync()
            Catch ex As Exception
                Console.WriteLine($"Authentication error: {ex.Message}")
            End Try
        End If

        Console.WriteLine($"Access Token: {result.AccessToken}{Environment.NewLine}")
        Console.WriteLine($"Graph Request: {request.RequestUri}")
        'You must set the access token for the authorization of the current request
        request.Headers.Authorization = New Headers.AuthenticationHeaderValue("Bearer", result.AccessToken)
    End Function
End Class
