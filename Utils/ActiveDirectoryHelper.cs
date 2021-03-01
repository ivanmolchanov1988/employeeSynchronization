using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Pkvs.DocsVision.Api.Entities;
 
public static class ActiveDirectoryHelper
{
	public static List<AdEmployee> GetEmployeesFromAd()
	{
		try
		{
			var allDomains = Forest.GetCurrentForest().Domains.Cast<Domain>();

			var allSearcher = allDomains.Select(domain =>
			{
				//var searcher = new DirectorySearcher(new DirectoryEntry("LDAP://" + domain.Name + "/OU=PKVS, DC=pkvs, DC=loc"))
				var searcher = new DirectorySearcher(new DirectoryEntry("LDAP://" + domain.Name))
				{
					Filter = String.Format(
						"(&(objectClass=user)(objectCategory=person))")
				};
				searcher.PropertiesToLoad.Add("samaccountname");
				searcher.PropertiesToLoad.Add("sn");
				searcher.PropertiesToLoad.Add("givenName");
				searcher.PropertiesToLoad.Add("mail");
				searcher.PropertiesToLoad.Add("middleName");
				return searcher;
			}
			);

			var employees = new List<AdEmployee>();
			var directoryEntriesFound =
			allSearcher.SelectMany(searcher =>
									searcher.FindAll()
									  .Cast<SearchResult>()
									 ).ToList();

			foreach (var value in directoryEntriesFound)
			{
				if (value.Properties["sn"] != null && value.Properties["sn"].Count > 0 &&
					value.Properties["samaccountname"] != null && value.Properties["samaccountname"].Count > 0 &&
					value.Properties["givenName"] != null && value.Properties["givenName"].Count > 0
					)
				{
					var surName = (String)value.Properties["sn"][0];
					var account = (String)value.Properties["samaccountname"][0];
					var firstName = (String)value.Properties["givenName"][0];
					var email = (value.Properties["mail"] != null && value.Properties["mail"].Count > 0) ? (String)value.Properties["mail"][0] : string.Empty;
					var middleName = (value.Properties["middleName"] != null && value.Properties["middleName"].Count > 0) ? (String)value.Properties["middleName"][0] : string.Empty;
					var employee = new AdEmployee { Account = account, SurName = surName, Name = firstName, Email = email, MiddleName = middleName };
					employees.Add(employee);
				}
			}

			return employees;
		}
		catch
		{
			return null;
		}
		
	}
}