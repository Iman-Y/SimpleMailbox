using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;

class GraphHelper
{
    private static Settings? _settings;
    private static DeviceCodeCredential? _deviceCodeCredential;
    private static GraphServiceClient? _userClient;

    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
            settings.TenantId, settings.ClientId);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }

    public static async Task<string> GetUserTokenAsync()
    {
     
        _ = _deviceCodeCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

     
        _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

       
        var context = new TokenRequestContext(_settings.GraphUserScopes);
        var response = await _deviceCodeCredential.GetTokenAsync(context);
        return response.Token;
    }

    public static Task<User> GetUserAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
            .Request()
            .Select(u => new
            {
               
                u.DisplayName,
                u.Mail,
                u.UserPrincipalName
            })
            .GetAsync();
    }
    public static Task<IMailFolderMessagesCollectionPage> GetInboxAsync()
    {
      
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me

            .MailFolders["Inbox"]
            .Messages
            .Request()
            .Select(m => new
            {
  
                m.From,
                m.IsRead,
                m.ReceivedDateTime,
                m.Subject
            })

            .Top(50)

            .OrderBy("ReceivedDateTime DESC")
            .GetAsync();
    }

    public static async Task SendMailAsync(string subject, string body, string recipient)
    {

        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                Content = body,
                ContentType = BodyType.Text
            },
            ToRecipients = new Recipient[]
            {
            new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = recipient
                }
            }
            }
        };

        await _userClient.Me
            .SendMail(message)
            .Request()
            .PostAsync();
    }

}