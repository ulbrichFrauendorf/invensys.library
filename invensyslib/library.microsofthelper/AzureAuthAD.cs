using library.common;
using Microsoft.Identity.Client;
using System;
using System.Threading.Tasks;

namespace library.microsofthelper
{
	public class AzureAuthAD
	{
		private IPublicClientApplication PublicClientApp { get; set; }
		private IAccount UserAccount { get; set; }
		public string CurrentToken { get; set; }

		private void BuildPublicClientApplication()
		{
			if (PublicClientApp == null)
			{
				PublicClientApplicationOptions pcaOptions = new PublicClientApplicationOptions
				{
					ClientId = "ba9b49f9-c1f0-4a05-b025-b1942b2782b4",  //ConfigurationManager.AppSettings["appId"],
					TenantId = "3e32dd7c-41f6-492d-a1a3-c58eb02cf4f8", //ConfigurationManager.AppSettings["tenantId"]
					RedirectUri = @"http://localhost"
				};

				PublicClientApp = PublicClientApplicationBuilder
					.CreateWithApplicationOptions(pcaOptions).Build();
			}
		}

		public async Task SetOAuthAccessTokenAsync(string[] scopes)
		{
			try
			{
				BuildPublicClientApplication();

				if (scopes == null)
					return;

				AuthenticationResult authResult;

				authResult = UserAccount != null ? await PublicClientApp.AcquireTokenSilent(scopes, UserAccount).WithForceRefresh(true).ExecuteAsync() : await PublicClientApp.AcquireTokenInteractive(scopes).WithUseEmbeddedWebView(false).ExecuteAsync().ConfigureAwait(false);
				UserAccount = authResult.Account;
				CurrentToken = authResult.AccessToken;
				System.Collections.Generic.IEnumerable<IAccount> tokens = await PublicClientApp.GetAccountsAsync();
				return;
			}
			catch (MsalException ex)
			{
				throw new LocalSystemException("Error acquiring access token", ex);
			}
			catch (Exception ex)
			{
				throw new LocalSystemException("User Authentication Failed", ex);
			}
		}
	}

	//public class AzureConfiguration
	//{
	//	/// <summary>
	//	/// Authentication options
	//	/// </summary>
	//	public PublicClientApplicationOptions PublicClientApplicationOptions { get; set; }
	//	public static AzureConfiguration ReadFromJsonFile(string path)
	//	{
	//		// .NET configuration
	//		IConfigurationRoot Configuration;

	//		var builder = new ConfigurationBuilder().
	//		.AddJsonFile(path);

	//		Configuration = builder.Build();

	//		// Read the auth and graph endpoint config
	//		SampleConfiguration config = new SampleConfiguration()
	//		{
	//			PublicClientApplicationOptions = new PublicClientApplicationOptions()
	//		};
	//		Configuration.Bind("Authentication", config.PublicClientApplicationOptions);
	//		config.MicrosoftGraphBaseEndpoint = Configuration.GetValue<string>("WebAPI:MicrosoftGraphBaseEndpoint");
	//		return config;
	//	}
	//}
}