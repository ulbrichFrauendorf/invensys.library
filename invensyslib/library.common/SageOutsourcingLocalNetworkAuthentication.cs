using SimpleImpersonation;
using System;

namespace library.common
{
	public static class SageOutsourcingLocalNetworkAuthentication
	{
		public static bool AccessNetworkWithServiceAccount(Func<string, bool> AccessNetworkFunction, string saveFileName)
		{
			bool result;
			try
			{
				result = AccessNetworkFunction.Invoke(saveFileName);
			}
			catch
			{
				try
				{
					UserCredentials credentials = new UserCredentials("sagesl.za.adinternal.com", "za-pta-outsourcing-a", "P@ssw0rd");
					result = Impersonation.RunAsUser(credentials, LogonType.Interactive, () =>
					{
						return AccessNetworkFunction(saveFileName);
					});
				}
				catch
				{
					return false;
				}
			}

			return result;
		}
	}
}