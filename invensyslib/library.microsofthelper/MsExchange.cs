using library.common;
using Microsoft.Exchange.WebServices.Data;
using SimpleImpersonation;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace library.microsofthelper
{
	public class MsExchange
	{
		#region Properties

		private readonly Action<ItemId> MailProcessingMethod;
		public SearchFilter Searchfilters { get; set; }
		public ExchangeService ThisExchangeService { get; set; }
		private WinLogger Logger { get; set; }
		public string ThisMailbox { get; private set; }

		#endregion Properties

		#region Public Methods

		public MsExchange(Action<ItemId> mailProcessingMethod)
		{
			Logger = new WinLogger("ExchangeLog", "Ms Exchange");
			MailProcessingMethod = mailProcessingMethod;
		}

		public async Task<bool> ConnectExchange(string authToken)
		{
			bool success = true;
			try
			{
				ThisExchangeService = new ExchangeService
				{
					Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx"),
					Credentials = new OAuthCredentials(authToken)
				};

				Folder Inbox = Folder.Bind(ThisExchangeService, WellKnownFolderName.Inbox);
				AlternateId aiAlternateid = new AlternateId(IdFormat.EwsId, Inbox.Id.UniqueId, "mailbox@domain.com");
				AlternateIdBase aiResponse = ThisExchangeService.ConvertId(aiAlternateid, IdFormat.EwsId);
				ThisMailbox = ((AlternateId)aiResponse).Mailbox;
			}
			catch (Exception ex)
			{
				success = false;
				Logger.FireWindowsLog(WinLogger.ApplicationEventType.Error, "Exchange Operation Failed - " + ex);
			}

			//success = await EnumerateFoldersAsync();
			return success;
		}

		public async Task<bool> EnumerateFoldersAsync()
		{
			bool success = false;
			try
			{
				FindFoldersResults folders = ThisExchangeService.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(500));
				foreach (Folder folder in folders)
				{
					success = await RecurseFoldersAsync(folder).ConfigureAwait(false);
				}
			}
			catch (Exception ex)
			{
				success = false;
				Logger.FireWindowsLog(WinLogger.ApplicationEventType.Error, "Folder Enumeration Failed - " + ex);
			}

			//TODO
			//folders = ThisExchangeService.FindFolders(WellKnownFolderName.ArchiveMsgFolderRoot, new FolderView(500));
			//foreach (var folder in folders)
			//{
			//	success = await RecurseFoldersAsync(folder).ConfigureAwait(false);
			//}
			return success;
		}

		private async Task<bool> RecurseFoldersAsync(Folder Folder)
		{
			bool success = false;
			try
			{
				FindFoldersResults childFolders = Folder.FindFolders(new FolderView(500));
				if (childFolders.Count() > 0)
				{
					foreach (Folder childFolder in childFolders)
					{
						success = await RecurseFoldersAsync(childFolder);
					}
				}
				Debug.WriteLine(Folder.DisplayName);

				if (Folder.FolderClass == "IPF.Note")
					ProcessEmailMessages(Folder);

				success = true;
			}
			catch (Exception ex)
			{
				success = false;
				Logger.FireWindowsLog(WinLogger.ApplicationEventType.Error, "Folder Enumeration Failed - " + ex);
			}
			return success;
		}

		private void ProcessEmailMessages(Folder searchFolder)
		{
			try
			{
				FindItemsResults<Item> findResults;
				ItemView view = new ItemView(150, 0, OffsetBasePoint.Beginning)
				{
					PropertySet = new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Categories, ItemSchema.DateTimeReceived, ItemSchema.DateTimeSent),
					Traversal = ItemTraversal.Shallow
				};
				do
				{
					findResults = searchFolder.FindItems(Searchfilters, view);
					//FindItemsResults<EmailMessage> Emails;
					//var results = BatchUpdateCategories(findResults, searchFolder, "Sage Archived");
					for (int i = findResults.Count() - 1; i >= 0; i--)
					{
						EmailMessage item = findResults.Items[i] as EmailMessage;
						if (item == null)
							continue;

						if (!(item is EmailMessage))
							continue;

						try
						{
							MailProcessingMethod.Invoke(item.Id);
						}
						catch (Exception ex)
						{
							Logger.FireWindowsLog(WinLogger.ApplicationEventType.Error, "Mail Processing Error -->" + ex);
						}
					}

					//any more batches?
					if (findResults.NextPageOffset.HasValue)
					{
						view.Offset = findResults.NextPageOffset.Value;
					}
				}
				while (findResults.MoreAvailable);
			}
			catch (Exception ex)
			{
				Logger.FireWindowsLog(WinLogger.ApplicationEventType.Error, "Error retreiving search items -->" + ex);
			}
		}
		#endregion Public Methods
	}

	internal class TraceListener : ITraceListener
	{
		#region ITraceListener Members

		public void Trace(string traceType, string traceMessage) => CreateTextFile(traceType, traceMessage.ToString());

		#endregion ITraceListener Members

		private static void CreateTextFile(string fileName, string traceContent)
		{
			fileName = @"\\vip-out2012\Colleague_Admin\00AutoEmailApp\Logs\" + fileName;
			System.IO.File.WriteAllText(fileName + ".txt", traceContent);
		}
	}

	public class CommonMail
	{
		private WinLogger Logger { get; set; }
		private EmailMessage Email { get; set; }
		private string ThisMailbox { get; set; }

		public bool IsSentFromMe
		{
			get
			{
				try
				{
					return (Email.From.Address == ThisMailbox) ? true : false;
				}
				catch (Exception)
				{
					return false;
				}
			}
		}

		public List<string> EmailAdresses
		{
			get
			{
				List<string> adressList = new List<string>();
				if (!IsSentFromMe)
				{
					string addVal = (Email.From == null) ? string.Empty : Email.From.Address;
					adressList.Add(addVal);
					return adressList;
				}

				foreach (EmailAddress addr in Email.ToRecipients)
				{
					adressList.Add(addr.Address);
				}
				foreach (EmailAddress addr in Email.CcRecipients)
				{
					adressList.Add(addr.Address);
				}
				return adressList;
			}
		}

		public DateTime EmailDate => (!IsSentFromMe) ? Email.DateTimeReceived : Email.DateTimeSent;
		public string Subject => Email.Subject;

		public CommonMail(ExchangeService ewsClient, string mailboxName, ItemId emailItemId)
		{
			try
			{
				Logger = new WinLogger("ExchangeLog", "Ms Exchange");
				Email = EmailMessage.Bind(ewsClient, emailItemId, new PropertySet(
					BasePropertySet.FirstClassProperties,
					ItemSchema.Categories,
					ItemSchema.DateTimeReceived,
					ItemSchema.DateTimeSent,
					EmailMessageSchema.CcRecipients,
					EmailMessageSchema.ToRecipients,
					ItemSchema.Subject,
					ItemSchema.MimeContent,
					ItemSchema.TextBody,
					ItemSchema.Attachments));

				ThisMailbox = mailboxName;
			}
			catch
			{
				return;
			}
		}

		public string Category
		{
			get => Email.Categories.ToString();
			set
			{
				if (!Email.Categories.Contains(value))
				{
					Email.Categories.Add(value);
					Email.Update(ConflictResolutionMode.AutoResolve);
				}
			}
		}

		public void RemoveCategory(string category)
		{
			Email.Categories.Remove(category);
			Email.Update(ConflictResolutionMode.AutoResolve);
		}

		public bool Save(ExchangeService ewsClient, string saveFileName)
		{
			bool result;
			try
			{
				result = SaveEmailToNetwork(saveFileName);
			}
			catch
			{
				try
				{
					UserCredentials credentials = new UserCredentials("sagesl.za.adinternal.com", "za-pta-outsourcing-a", "P@ssw0rd");
					result = Impersonation.RunAsUser(credentials, LogonType.Interactive, () =>
					{
						return SaveEmailToNetwork(saveFileName);
					});
				}
				catch
				{
					return false;
				}
			}

			return result;
		}

		private bool SaveEmailToNetwork(string saveFileName)
		{
			using (FileStream fileStream = new FileStream(saveFileName, FileMode.Create))
			{
				try
				{
					fileStream.Write(Email.MimeContent.Content, 0, Email.MimeContent.Content.Length);
				}
				catch
				{
					fileStream.Write(Encoding.ASCII.GetBytes(Email.TextBody.Text), 0, (Email.TextBody.Text).Length);
				}

				fileStream.Close();
			}

			if (!File.Exists(saveFileName))
				return false;

			return true;
		}

		public void Delete() => Email.Delete(DeleteMode.HardDelete);
	}
}