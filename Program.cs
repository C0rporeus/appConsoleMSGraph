using System;
using System.Collections.Generic;
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
      var client = GetAuthenticatedGraphClient(config);

      /*var graphRequest = client
                         .Users.Request()
                         .Select(u => new {u.Id, u.DisplayName, u.Mail })
                         .Top(999);*/
      var graphRequest = client
                          .Groups.Request()
                          .Top(999)
                          .Expand("members");
      var result = graphRequest.GetAsync().Result;
      
      /* foreach(var user in result)
      {
        Console.WriteLine(user.Id + ": " + user.DisplayName + " <" + user.Mail + ">");
      } */
      foreach (var group in result)
      {
        Console.WriteLine("\n");
        Console.WriteLine(group.Id + ": " + group.DisplayName+"||||");
        Console.WriteLine("\n");
        foreach (var member in group.Members)
        {
          Console.WriteLine(" " + member.Id + ": " + ((Microsoft.Graph.User)member).DisplayName);
        }
      }
      Console.WriteLine("\nGraph Request:");
      Console.WriteLine(graphRequest.GetHttpRequestMessage().RequestUri);
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
        if(string.IsNullOrEmpty(config["applicationId"]) ||
           string.IsNullOrEmpty(config["applicationSecret"]) ||
           string.IsNullOrEmpty(config["redirectUri"]) ||
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
    private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
    {
      var clientId = config["applicationId"];
      var clientSecret = config["applicationSecret"];
      var redirectUri = config["redirectUri"];
      var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

      List<string> scopes = new List<string>();
      scopes.Add("https://graph.microsoft.com/.default");

      var cca = ConfidentialClientApplicationBuilder.Create(clientId)
        .WithAuthority(authority)
        .WithRedirectUri(redirectUri)
        .WithClientSecret(clientSecret)
        .Build();
      return new MsalAuthenticationProvider(cca, scopes.ToArray());
    }
    private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
    {
      var authenticationprovider = CreateAuthorizationProvider(config);
      _graphClient = new GraphServiceClient(authenticationprovider);
      return _graphClient;
    }
  }
}