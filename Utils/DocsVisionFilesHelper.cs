using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using DocsVision.Platform.ObjectManager.SystemCards;
using DocsVision.Platform.ObjectManager;
//using Pkvs.DocsVision.Api.Entities;
//using NLog;
using DocsVision.BackOffice.CardLib.CardDefs;
using DocsVision.Platform.ObjectModel;
using DocsVision.BackOffice.ObjectModel.Services;
using DocsVision.BackOffice.ObjectModel;

public class DocsVisionFilesHelper
{
	// Родительская папка (обозначена RootFolderGuid в app.config), в которой находяться папки пользователей
	private readonly string rootFolderGuid;
	private UserSession Session;
	// Guid карточки папок
	private const string FOLDER_CARD_TYPE = "{DA86FABF-4DD7-4A86-B6FF-C58C24D12DE2}";
	private FolderCard folderCard;
	private List<FoldersDelegate> foldersDelegate = new List<FoldersDelegate>();
	//private Logger logger;
	private ObjectContext objectContext;

	public class FoldersDelegate
	{
		public string value;
		public bool withSubfolders;
		public FoldersDelegate(string value, string withSubfolders)
		{
			this.value = value;
			this.withSubfolders = withSubfolders == "true" ? true : false;
		}
	}

	public DocsVisionFilesHelper(UserSession userSession/*, Logger logger*/) 
	{
		//this.logger = logger;
		this.rootFolderGuid = ConfigurationManager.AppSettings["RootFolderGuid"];
		this.Session = userSession;
		folderCard = (FolderCard)Session.CardManager.GetDictionary(new Guid(FOLDER_CARD_TYPE));
		setupFoldersDelegate();
	}

	public DocsVisionFilesHelper(UserSession userSession, ObjectContext objectContext/*, Logger logger*/)
	{
		this.Session = userSession;
		this.objectContext = objectContext;
		//this.logger = logger;
		this.rootFolderGuid = ConfigurationManager.AppSettings["RootFolderGuid"];            
		folderCard = (FolderCard)Session.CardManager.GetDictionary(new Guid(FOLDER_CARD_TYPE));
		setupFoldersDelegate();
	}

	// Папки, которые необходимо прикрепить в родительскую папку, обозначены Folder[от 1 до 10] в app.config
	private void setupFoldersDelegate()
	{
		for (int i = 0; i < 10; i++)
		{
			string value = ConfigurationManager.AppSettings["Folder" + i] as string;
			if (!string.IsNullOrEmpty(value))
			{
				List<string> list = value.Split(';').ToList();
				foldersDelegate.Add(new FoldersDelegate(list[0], list[1]));
			}
		}
	}

	public void UpdateEmployeeFolder(RowData employee)
	{            
		// Проверяем, задана ли Родительская папка 
		if (!isExistFolder(rootFolderGuid))
		{
			//logger.Info("Работа с папками завершена, поскольку не задан параметр RootFolderGuid родительской папки пользователей в App.config ");
			return;
		}

		// Узнаем какое должно быть имя папки пользователя
		string employeeFolderName = getEmployeeFolderName(employee);
		if (string.IsNullOrEmpty(employeeFolderName))
		{
			return;
		}

		// Проверяем, существует ли папка пользователя, если нет, пытаемся создать
		string employeeFolderGuid = findFolder(rootFolderGuid, employeeFolderName);
		if (string.IsNullOrEmpty(employeeFolderGuid))
		{
			employeeFolderGuid = createFolder(rootFolderGuid, employeeFolderName);

			if (string.IsNullOrEmpty(employeeFolderGuid))
			{
				return;
			}
			else
			{                   
				//logger.Info("Создана папка сотрудника");
				employee[RefStaff.Employees.PersonalFolder] = employeeFolderGuid;
				// Если мы создали папку пользователя, то должны назначить на нее права
				bool isSetRightsFoFolder = setRightsFoFolder(employee, employeeFolderGuid);
				if (!isSetRightsFoFolder)
				{
					//logger.Info("Права и ограничения на папку сотрудника созданы не были");
				}
			}
		}

		if (string.IsNullOrEmpty(employeeFolderGuid))
		{
			return;
		}

		// Устанавливаем созданную папку данному пользователю
		employee[RefStaff.Employees.PersonalFolder] = employeeFolderGuid;

		// Ищем папки и прикрепляем любые папки в родительскую
		UpdateDelegateFolders(employeeFolderGuid, employee);
	}

	private void UpdateDelegateFolders(string employeeFolderGuid, RowData employee)
	{
		foreach (var item in foldersDelegate)
		{
			CreateFolderDelegate(employeeFolderGuid, item.value, item.withSubfolders);
		}
	}

	private string CreateFolderDelegate(string parentFolderId, string childrenFolderId, bool withSubfolders)
	{

		if (string.IsNullOrEmpty(parentFolderId) || string.IsNullOrEmpty(childrenFolderId))
			return null;

		Folder childrenFolder = folderCard.GetFolder(new Guid(childrenFolderId));
		if (childrenFolder == null)
			return null;
	 
		// Если папка там уже создана, то ничего не делаем
		string insideFolderId = findFolder(parentFolderId.ToString(), childrenFolder.Name);
		if (insideFolderId != null)
			return null;

		Folder folderDelegate = folderCard.CreateFolder(new Guid(parentFolderId.ToString()), childrenFolder.Name);
		folderDelegate.Type = FolderTypes.Delegate;
		folderDelegate.RefId = new Guid(childrenFolder.Id.ToString());
		folderDelegate.DefaultViewId = childrenFolder.DefaultViewId;
		if (withSubfolders)
		{
			folderDelegate.Flags = FolderFlags.VirtualWithSubfolders;
		}
		return folderDelegate.Id.ToString();

	}

	private string getEmployeeFolderName(RowData employee)
	{
		if (employee == null)
		{
			return null;
		}

		string r = "";
		if (!String.IsNullOrEmpty(employee.GetString("FirstName")))
		{
			r += employee.GetString("FirstName")[0].ToString() + ".";
		}
		if (!String.IsNullOrEmpty(employee.GetString("MiddleName")))
		{
			r += employee.GetString("MiddleName")[0].ToString() + ".";
		}
		string res = string.Format("{0} {1}", employee.GetString("LastName"), r);
		if (!string.IsNullOrEmpty(res))
		{
			return res.Trim();
		}

		return null;
	}

	private string findFolder(string parentFolderId, string folderName)
	{
		if (string.IsNullOrEmpty(folderName) || string.IsNullOrEmpty(parentFolderId))
		{
			return null;
		}
		try
		{
			const string FOLDER_CARD_TYPE = "{DA86FABF-4DD7-4A86-B6FF-C58C24D12DE2}";
			FolderCard folderCard = (FolderCard)Session.CardManager.GetDictionary(new Guid(FOLDER_CARD_TYPE));

			Folder rootFolder = folderCard.GetFolder(new Guid(parentFolderId.ToString()));
			if (rootFolder != null)
			{
				foreach (var item in rootFolder.Folders.Where(s => s.Name.ToLower() == folderName.ToLower()))
				{
					return item.Id.ToString();
				}
			}
			return null;
		}
		catch
		{

			return null;
		}

		return null;
	}

	private bool isExistFolder(string folderId)
	{
		if (string.IsNullOrEmpty(folderId))
			return false;

		return folderCard.FolderExists(new Guid(folderId.ToString()));            
	}

	private string createFolder(string parentFolderId, string folderName)
	{
		if (string.IsNullOrEmpty(folderName) || string.IsNullOrEmpty(parentFolderId))
			return null;

		Folder folder = folderCard.CreateFolder(new Guid(parentFolderId.ToString()), folderName);
		if (folder != null)
			return folder.Id.ToString();

		return null;
	}

   private bool setRightsFoFolder(RowData newEmployeeDocsvision, string folderId)
	{ 
		if (folderId != null)
		{
			CardData refStaffData = Session.CardManager.GetDictionaryData(RefStaff.ID);
			SectionData unitsSection = refStaffData.Sections[RefStaff.Units.ID];
			
			//Права
			var accountName = newEmployeeDocsvision.GetString(RefStaff.Employees.AccountName);
			if (string.IsNullOrEmpty(accountName))
				return false;

			
			var profileCardId = Session.ProfileManager.GetProfileId(accountName);
			var profileCard = (UserProfileCard)Session.CardManager.GetCard(profileCardId);
			if (profileCard.DefaultFolderId == new Guid(folderId))
				return false;
			
			profileCard.DefaultFolderId = new Guid(folderId);

			var employee = objectContext.GetObject<StaffEmployee>(newEmployeeDocsvision.Id);
			if (employee.PersonalFolder == null)
				return false;

			var staffService = objectContext.GetService<IStaffService>();
			staffService.SetFoldersRights(employee.PersonalFolder, employee.AccountName);
		  

			//logger.Info("Права на папку созданы.");
			//Ограничения

			FolderCard folderCard = (FolderCard)Session.CardManager.GetDictionary(new Guid("DA86FABF-4DD7-4A86-B6FF-C58C24D12DE2"));
			
			Folder folder = folderCard.GetFolder(new Guid(folderId.ToString()));
			if ((folder.Restrictions & FolderRestrictions.Types) == 0)
				folder.Restrictions |= FolderRestrictions.Types;
		   

			//logger.Info("Ограничения на папку созданы.");
			//Разрешения
		   
			if ((folder.Restrictions & FolderRestrictions.Types) == 0)
			{
				folder.Restrictions = folder.Restrictions | FolderRestrictions.Types;
				folder.AllowedCardTypes.AddNew(new Guid("6D76D0A7-5434-40F2-912E-6370D33C3151"));
			}
		}
		return true;
	}
}