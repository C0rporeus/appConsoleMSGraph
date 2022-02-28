using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;

namespace graphconsoleapp
{
  public class Program
  {
    public static void Main(string[] args)
    {
      var config = LoadAppSettings();
      if (config == null)
      {
        Console.WriteLine("invalid appsettings.json file");
        return;
      }
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
      scopes.Add("Mail.Read");
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
    }
  }
}