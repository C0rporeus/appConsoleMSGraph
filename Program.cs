using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
using System.Text.Json;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;
//using Newtonsoft.Json;

namespace graphconsoleapp
{
  public class Program
  {
    private static object? _deltaLink = null;
    private static IUserDeltaCollectionPage? _previousPage = null;
    public static void Main(string[] args)
    {
      var config = LoadAppSettings();
      if (config == null)
      {
        Console.WriteLine("invalid appsettings.json file");
        return;
      }
      var userName = ReadUsername();
      var userPassword = ReadPassword();
      Console.WriteLine("All users in tenant:");
      CheckForUpdates(config, userName, userPassword);
      Console.WriteLine();
      while (true)
      {
        Console.WriteLine("... sleeping for 10s - press CTRL+C to terminate");
        System.Threading.Thread.Sleep(10 * 1000);
        Console.WriteLine("> Checking for new/updated users since last query...");
        CheckForUpdates(config, userName, userPassword);
      }
    }
    private static void CheckForUpdates(IConfigurationRoot config, string userName, SecureString userPassword)
    {
      var graphClient = GetAuthenticatedHTTPClient(config, userName, userPassword);

      // get a page of users
      var users = GetUsers(graphClient, _deltaLink);

      OutputUsers(users);

      // go through all of the pages so that we can get the delta link on the last page.
      while (users.NextPageRequest != null)
      {
        users = users.NextPageRequest.GetAsync().Result;
        OutputUsers(users);
      }
      object? deltaLink;

      if (users.AdditionalData.TryGetValue("@odata.deltaLink", out deltaLink))
      {
        _deltaLink = deltaLink;
      }
    }

    private static void OutputUsers(IUserDeltaCollectionPage users)
    {
      foreach (var user in users)
      {
        Console.WriteLine($"User: {user.Id}, {user.GivenName} {user.Surname}");
      }
    }
    private static IUserDeltaCollectionPage GetUsers(GraphServiceClient graphClient, object deltaLink)
    {
      IUserDeltaCollectionPage page;

      // IF this is the first request, then request all users
      //    and include Delta() to request a delta link to be included in the
      //    last page of data
      if (_previousPage == null || deltaLink == null)
      {
        page = graphClient.Users
                          .Delta()
                          .Request()
                          .Select("Id,GivenName,Surname")
                          .GetAsync()
                          .Result;
      }
      // ELSE, not the first page so get the next page of users
      else
      {
        _previousPage.InitializeNextPageRequest(graphClient, deltaLink.ToString());
        page = _previousPage.NextPageRequest.GetAsync().Result;
      }

      _previousPage = page;
      return page;
    }
    private static Message? GetMessageDetail(HttpClient client, string messageId, int defaultDelay = 2)
    {
      Message? messageDetail = null;
      string endpoint = "https://graph.microsoft.com/v1.0/me/message" + messageId;

      var clientResponse = client.GetAsync(endpoint).Result;
      var httpResponseTask = clientResponse.Content.ReadAsStringAsync();
      httpResponseTask.Wait();

      Console.WriteLine("...Response status code: {0}  ", clientResponse.StatusCode);

      if (clientResponse.StatusCode == HttpStatusCode.OK)
      {
        messageDetail = JsonSerializer.Deserialize<Message>(httpResponseTask.Result);
      }
      // ELSE IF request was throttled (429, aka: TooManyRequests)...
      else if (clientResponse.StatusCode == HttpStatusCode.TooManyRequests)
      {
        // get retry-after if provided; if not provided default to 2s
        var retryAfterDelay = defaultDelay;
        var retryAfter = clientResponse.Headers.RetryAfter;
        if (retryAfter != null && retryAfter.Delta.HasValue && (retryAfter.Delta.Value.Seconds > 0))
        {
          retryAfterDelay = retryAfter.Delta.Value.Seconds;
        }

        // wait for specified time as instructed by Microsoft Graph's Retry-After header,
        //    or fall back to default
        Console.WriteLine(">>>>>>>>>>>>> sleeping for {0} seconds...", retryAfterDelay);
        System.Threading.Thread.Sleep(retryAfterDelay * 1000);

        // call method again after waiting
        messageDetail = GetMessageDetail(client, messageId);
      }

      // rest to code
      return messageDetail;
    }

    public static GraphServiceClient? _graphClient;
    private static IConfigurationRoot? LoadAppSettings()
    {
      try
      {
        var config = new ConfigurationBuilder()
          .SetBasePath(System.IO.Directory.GetCurrentDirectory())
          .AddJsonFile("appsettings.json", false, true)
          .Build();
        if (string.IsNullOrEmpty(config["applicationId"]) ||
           //string.IsNullOrEmpty(config["applicationSecret"]) ||
           //string.IsNullOrEmpty(config["redirectUri"]) ||
           string.IsNullOrEmpty(config["tenantId"]))
        {
          return null;
        }
        return config;
      }
      catch (System.IO.FileNotFoundException)
      {
        return null;
      }
    }
    private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config, string userName, SecureString userPassword)
    {
      var clientId = config["applicationId"];
      //var clientSecret = config["applicationSecret"];
      //var redirectUri = config["redirectUri"];
      var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

      List<string> scopes = new List<string>();
      //scopes.Add("https://graph.microsoft.com/.default");
      scopes.Add("User.Read");
      scopes.Add("User.Read.All");
      var cca = PublicClientApplicationBuilder.Create(clientId)
                                              .WithAuthority(authority)
                                              .Build();
      return MsalAuthenticationProvider.GetInstance(cca, scopes.ToArray(), userName, userPassword);
    }
    private static HttpClient GetAuthenticatedHTTPClient(IConfigurationRoot config, string userName, SecureString userPassword)
    {
      var authenticationProvider = CreateAuthorizationProvider(config, userName, userPassword);
      var httpClient = new HttpClient(new AuthHandler(authenticationProvider, new HttpClientHandler()));
      return httpClient;
    }
    private static SecureString ReadPassword()
    {
      Console.WriteLine("Enter your password");
      SecureString password = new SecureString();
      while (true)
      {
        ConsoleKeyInfo c = Console.ReadKey(true);
        if (c.Key == ConsoleKey.Enter)
        {
          break;
        }
        password.AppendChar(c.KeyChar);
        Console.Write("*");
      }
      Console.WriteLine();
      return password;
    }
    private static string ReadUsername()
    {
      string? username;
      Console.WriteLine("Enter your username");
      username = Console.ReadLine();
      return username ?? "";
    }
  }
}