using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

using DocsVision.ApprovalDesigner.ObjectModel.Mapping;
using DocsVision.ApprovalDesigner.ObjectModel.Services;
using DocsVision.BackOffice.CardLib.CardDefs;
using DocsVision.BackOffice.ObjectModel;
using DocsVision.BackOffice.ObjectModel.Mapping;
using DocsVision.BackOffice.ObjectModel.Services;
using DocsVision.DocumentsManagement.ObjectModel.Services;
using DocsVision.Platform.Data.Metadata;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectModel;
using DocsVision.Platform.ObjectModel.Mapping;
using DocsVision.Platform.ObjectModel.Persistence;
using DocsVision.Platform.SystemCards.ObjectModel.Mapping;
using DocsVision.Platform.SystemCards.ObjectModel.Services;

using DocsVision.Platform.ObjectManager.SystemCards;
using DocsVision.Platform.ObjectManager.SearchModel;
using DocsVision.Platform.ObjectManager.Metadata;
using System.Configuration;



public class DocsvisionHelper
{
    private HttpApplicationState _application = null;
	public const string ACCOUNT_DOMAIN = "PKVS\\";
    private string LogEmptyInfo = "Операция не произведена, поскольку не переданы обязательные поля: {0}";

    public DocsvisionHelper(HttpApplicationState application)
    {
        _application = application;
    }

    //Сотрудники

    //im public string AddEmployee(string code, string familyName, string firstName, string middleName, bool? actual, string titleId,
	public object AddEmployee(string code, string familyName, string firstName, string middleName, bool? actual, string titleId,
                            string title, int routingType, string departmentCode, string email, string employeeChief)
    {
        try
        {
            string requestEmptyField = GetRequestEmployee(code, familyName, departmentCode, actual);
            if (!string.IsNullOrEmpty(requestEmptyField))
            {
                //logger.Info(requestEmptyField);
                //im return requestEmptyField;
				return true;
            }

            if (!actual.HasValue || !actual.Value)
			{
                //im return "Сотрудник неактуален, добавление не произведено.";
				return true;
			}
			
			//Поиск сотрудника ***
			//im RowData employeeDocsvision = GetExistObject(RefStaff.Employees.ID, RefStaff.Employees.IDCode, code);
			RowDataCollection employeeDocsvision = GetExistObject(RefStaff.Employees.ID, RefStaff.Employees.IDCode, code);
            //im if (employeeDocsvision == null)
			if (employeeDocsvision == null || employeeDocsvision.Count == 0)
			{
				//Поиск подразделения                    
				//im RowData department = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, departmentCode);
				RowDataCollection departmentColl = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, departmentCode);
				//im if (department == null)
				if (departmentColl == null || departmentColl.Count == 0 || departmentColl.Count > 1)
				{
					return "Не найдено подразделение.";
					//return false;
				}
				RowData department = departmentColl[0];

				CardData refStaffData = Session.CardManager.GetDictionaryData(RefStaff.ID);

				refStaffData.BeginUpdate();

				// Создание нового пользователя
				RowData newEmployeeDocsvision = CreateEmployee(department);
				

				// Обновление следующих полей: 1. Фамилия  2. Имя  3. Отчество 6. E-mail, 8. Маршрутизация, так же IDCode, displayName
				UpdateEmployeeSimple(newEmployeeDocsvision, refStaffData, code, familyName, firstName, middleName, email, routingType);

				// 7. Пользователь
				if (!string.IsNullOrEmpty(email))
					SetAccountName(newEmployeeDocsvision, familyName, firstName, middleName, email);

				//4. Должность
				UpdateEmployeePosition(newEmployeeDocsvision, refStaffData, titleId, title);

				//5. Руководитель
				UpdateEmployeeManager(newEmployeeDocsvision, refStaffData, employeeChief);

				refStaffData.EndUpdate();

				// 9. Папки  
				DocsVisionFilesHelper filesHelper = new DocsVisionFilesHelper(Session, ObjectContext/*, logger*/);
				filesHelper.UpdateEmployeeFolder(newEmployeeDocsvision);

				//logger.Info("Создание сотрудника завершено.");

				//Проверка наличия созданного сотрудника
				/*logger.Info(base.SearchEmployeeTest(refStaffData, employee1C) == true
								? "Наличие созданного сотрудника подтверждено!"
								: "Созданный сотрудник не подтвержден!");*/
				/*im
				return (SearchEmployeeTest(refStaffData, code)
								? "Наличие созданного сотрудника подтверждено!"
								: "Созданный сотрудник не подтвержден!");*/
				
				return (SearchEmployeeTest(refStaffData, code)
								? true
								: true); // создал, но не подтвердил
			}
			else
			{
				//im return "Сотрудник уже есть в справочнике";
				return true;
			}
        }
        catch (Exception ex)
        {
            //im return string.Format("Ошибка создания сотрудника {0}", ex.Message);
			return "Не удалось добавить сотрудника в справочник.";
        }
    }

    //im public string UpdateEmployee(string code, string familyName, string firstName, string middleName, bool? actual, string titleId,
	public object UpdateEmployee(string code, string familyName, string firstName, string middleName, bool? actual, string titleId,
                            string title, int routingType, string departmentCode, string email, string employeeChief)
    {
        string requestEmptyField = GetRequestEmployee(code, familyName, departmentCode, actual);
        if (!string.IsNullOrEmpty(requestEmptyField))
        {
            //logger.Info(requestEmptyField);
            //im return requestEmptyField;
			return true;
        }
		
        CardData refStaffData = Session.CardManager.GetDictionaryData(RefStaff.ID);

        // Поиск пользователя
		//var rowColl = new RowDataCollection<RowData>(); //im
		
        //im RowData employeeDocsvision = GetExistObject(RefStaff.Employees.ID, RefStaff.Employees.IDCode, code);
		RowDataCollection employeeDocsvisionColl = null;
		try{
			employeeDocsvisionColl = GetExistObject(RefStaff.Employees.ID, RefStaff.Employees.IDCode, code);
		}
		catch{}
		
        //im if (employeeDocsvision == null)
			/*
		if (employeeDocsvisionColl == null || employeeDocsvisionColl.Count == 0 || employeeDocsvisionColl.Count > 1)
        {
            //im var accountNameFromAd = GetAccountFromAD(familyName, firstName, middleName, email);
            //im employeeDocsvision = GetExistObject(RefStaff.Employees.ID, RefStaff.Employees.AccountName, accountNameFromAd);
			return true;
        }
		*/
		RowData employeeDocsvision = null;
		try{
			employeeDocsvision = employeeDocsvisionColl[0];
		}
		catch{}
		
        //im if (employeeDocsvision == null)
		if (employeeDocsvision == null)
        {
            //logger.Info("Сотрудник не найден, началось создание нового сотрудника.");
            return AddEmployee(code, familyName, firstName, middleName, actual, titleId, title, routingType,
                        departmentCode, email, employeeChief);
        }

        //Поиск подразделения
        //im RowData department = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, departmentCode);
		RowDataCollection departmentColl = null;
		try{
			departmentColl = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, departmentCode);
        }
		catch{}
		//im if (department == null)
		if (departmentColl == null || departmentColl.Count == 0 || departmentColl.Count > 1)
        {
            //logger.Info("Не найдено подразделение сотрудника.");
            //im return "Не найдено подразделение сотрудника.";
			return "Не найдено подразделение для сотрудника.";
        }
		RowData department = departmentColl[0];
		
        refStaffData.BeginUpdate();
		
        // Перемещение сотрудника в требуемое подразделение, если необходимо
        MoveExistEmployee(employeeDocsvision, department);
		
		
        // Обновление следующих полей: 1. Фамилия  2. Имя  3. Отчество 6. E-mail, 8. Маршрутизация, так же IDCode, displayName
        UpdateEmployeeSimple(employeeDocsvision, refStaffData, code, familyName, firstName, middleName, email, routingType);
		
        // Обновление статуса сотрудника
        SetEmployeeStatus(actual, employeeDocsvision);
		
        //logger.Info("Обновление базовых полей сотрудника завершено.");

        //4. Должность
        UpdateEmployeePosition(employeeDocsvision, refStaffData, titleId, title);
		
        //5. Руководитель
        UpdateEmployeeManager(employeeDocsvision, refStaffData, employeeChief);
		
		/*
		// 7. Пользователь
		if (email != null || email != string.Empty)
			SetAccountName(employeeDocsvision, familyName, firstName, middleName, email);
		*/
		
		refStaffData.EndUpdate();
		return true;
		
		// 9. Папки        
        //im DocsVisionFilesHelper filesHelper = new DocsVisionFilesHelper(Session, ObjectContext/*, logger*/);
        //im filesHelper.UpdateEmployeeFolder(employeeDocsvision);
        
		
        //im refStaffData.EndUpdate();
        //logger.Info("Обновление сотрудника завершено.");
        //im return "Обновление сотрудника завершено.";
		//im return true;
		
    }

    //im public string DeleteEmployee(string code)
	public bool DeleteEmployee(string code)
    {
		/*im 
        if (string.IsNullOrEmpty(code))
        {
            //logger.Info(string.Format(base.LogEmptyInfo, "Code"));
            return string.Format(LogEmptyInfo, "Code");
        }

        CardData refStaffData = Session.CardManager.GetDictionaryData(RefStaff.ID);
        // Поиск пользователя
        RowData employeeDocsvision = GetExistObject(RefStaff.Employees.ID, RefStaff.Employees.IDCode, code);

        if (employeeDocsvision != null)
        {
            refStaffData.BeginUpdate();
            employeeDocsvision[RefStaff.Employees.Status] = StaffEmployeeStatus.Discharged;
            employeeDocsvision[RefStaff.Employees.NotAvailable] = true;
            employeeDocsvision["NotSearchable"] = true;
            refStaffData.EndUpdate();

            return "Удаление произведено";
        }
        return "Сотрудник не найден";
		*/
		return false;
		
    }

    private string GetRequestEmployee(string code, string familyName, string departmentCode, bool? actual)
    {
        List<string> list = new List<string>();
        if (string.IsNullOrEmpty(code))
        {
            list.Add("Code");
        }
        if (string.IsNullOrEmpty(familyName))
        {
            list.Add("FamilyName");
        }
        if (string.IsNullOrEmpty(departmentCode))
        {
            list.Add("DepartmentCode");
        }
        /*if (string.IsNullOrEmpty(employee1C.EmployeeContractType))
        {
            list.Add("EmployeeContractType");
        }*/

        if (!actual.HasValue)
        {
            list.Add("Actual");
        }

        if (list.Count > 0)
        {
            return string.Format(LogEmptyInfo, string.Join(", ", list.ToArray()));
        }
        return null;
    }

    private void UpdateEmployeeSimple(RowData employeeDocsvision, CardData refStaffData,
                            string code, string familyName, string firstName, string middleName, string email, int routingType)
    {
        employeeDocsvision[RefStaff.Employees.LastName] = familyName;
        employeeDocsvision[RefStaff.Employees.FirstName] = firstName;
        employeeDocsvision[RefStaff.Employees.MiddleName] = middleName;
        employeeDocsvision[RefStaff.Employees.Email] = email;
        employeeDocsvision[RefStaff.Employees.IDCode] = code;
        if (string.IsNullOrEmpty(employeeDocsvision[RefStaff.Employees.RoutingType].ToString()) ||
            employeeDocsvision[RefStaff.Employees.RoutingType].ToString() != "2")
            employeeDocsvision[RefStaff.Employees.RoutingType] = routingType;

        //Отображаемое имя
        string displayName = GetDisplayName(familyName, firstName, middleName);
        employeeDocsvision[RefStaff.Employees.DisplayString] = displayName;
    }

    private string GetDisplayName(string familyName, string firstName, string middleName)
    {
        var firstNameChar = string.IsNullOrEmpty(firstName) ? string.Empty : (firstName[0].ToString() + ".");
        var firstMiddleNameChar = string.IsNullOrEmpty(middleName) ? string.Empty : middleName[0].ToString() + ".";
        var displayName = familyName + " " + firstNameChar + firstMiddleNameChar;
        return displayName;
    }

    private void SetAccountName(RowData employeeDocsvision, string familyName, string firstName, string middleName, string email)
    {
        string accountName = GetAccountFromAD(familyName, firstName, middleName, email);
        if (!string.IsNullOrEmpty(accountName))
		{	
			employeeDocsvision[RefStaff.Employees.AccountName] = accountName;
		}
		else
			return;
    }
	
	public string GetAccountName(string familyName, string firstName, string middleName, string email)
    {
        string accountName = GetAccountFromAD(familyName, firstName, middleName, email);
		if (!string.IsNullOrEmpty(accountName))
		{
			return accountName;
		}
		else
		{
			return null;
		}
    }

    private string GetAccountFromAD(string familyName, string firstName, string middleName, string email)
    {
        var adEmployees = ActiveDirectoryHelper.GetEmployeesFromAd();
        if (adEmployees == null)
        {
            //return string.Empty;
			return null;
        }
		/*im
        var foundEmployeeByCredentials =
            adEmployees.Where(x => (x.ConcatenatedCredentials == familyName.ToLower() + firstName.ToLower())).ToList();
        if (foundEmployeeByCredentials.Any())
        {
            if (foundEmployeeByCredentials.Count == 1)
            {
                var account = ACCOUNT_DOMAIN + foundEmployeeByCredentials[0].Account;
                return account;
            }
            else
            {
                var accountWithMiddleName = foundEmployeeByCredentials.Where(x => x.MiddleName == middleName).ToList();
                if (accountWithMiddleName.Any() && accountWithMiddleName.Count() == 1)
                {
                    var account = ACCOUNT_DOMAIN + accountWithMiddleName[0].Account;
                    return account;
                }
            }
        }
        else
        {
            if (email != null)
            {
                var foundEmployeeByEmail = adEmployees.FirstOrDefault(x => (x.Email.ToLower() == email.ToLower()));
                if (foundEmployeeByEmail != null)
                {
                    return ACCOUNT_DOMAIN + foundEmployeeByEmail.Account;
                }
            }
			else
			{
				return null;
			}
        }
		*/
		
		if (email != null)
		{
			var foundEmployeeByEmail = adEmployees.FirstOrDefault(x => (x.Email.ToLower() == email.ToLower()));
			if (foundEmployeeByEmail != null)
			{
				return ACCOUNT_DOMAIN + foundEmployeeByEmail.Account;
			}
		}
        //return string.Empty;
		return null;
    }

    private void UpdateEmployeePosition(RowData employeeDocsvision, CardData refStaffData, string titleId, string title)
    {
        SectionData positionSection = refStaffData.Sections[RefStaff.Positions.ID];
        SectionQuery sectionQueryPosition = Session.CreateSectionQuery();

        // Добавление условия поиска по дате создания документа
        sectionQueryPosition.ConditionGroup.Conditions.AddNew(RefStaff.Positions.Name, FieldType.Unistring, 
                                                                ConditionOperation.Equals, title);

        RowDataCollection positionCollection = positionSection.FindRows(sectionQueryPosition.GetXml());
        if (positionCollection.Count > 0)
        {
            //logger.Info("Должность сотрудника обновлена");
            employeeDocsvision[RefStaff.Employees.Position] = positionCollection[0].Id;
        }
        else
        {
            //logger.Info("Создана новая должность. Должность сотрудника обновлена");
            employeeDocsvision[RefStaff.Employees.Position] = AddPositionManual(title, titleId);
        }
    }

    private Guid AddPositionManual(string name, string code)
    {
        CardData refStaffData = Session.CardManager.GetDictionaryData(RefStaff.ID);
        SectionData positionSection = refStaffData.Sections[RefStaff.Positions.ID];

        var newRowDepartment = positionSection.Rows.AddNew();
        newRowDepartment[RefStaff.Positions.Name] = name;
        newRowDepartment[RefStaff.Positions.SyncTag] = code;

        return newRowDepartment.Id;
    }

    private void UpdateEmployeeManager(RowData employeeDocsvision, CardData refStaffData, string employeeChief)
    {
        if (employeeChief == null)
            // logger.Info("Руководитель сотрудника не найден.");
            return;

        SectionQuery sectionQueryEmployee = Session.CreateSectionQuery();
        sectionQueryEmployee.ConditionGroup.Conditions.AddNew(RefStaff.Employees.IDCode, FieldType.Unistring,
                                                                ConditionOperation.Equals, employeeChief);

        RowDataCollection employeeCollection = refStaffData.Sections[RefStaff.Employees.ID].FindRows(sectionQueryEmployee.GetXml());
        if (employeeCollection != null && employeeCollection.Count > 0)
        {
            employeeDocsvision[RefStaff.Units.Manager] = employeeCollection[0].Id;
            //logger.Info("Руководитель сотрудника обновлен.");
        }
        /*else
        {
            logger.Info("Руководитель сотрудника не найден.");

        }   */
    }

    private static void MoveExistEmployee(RowData employeeDocsvision, RowData department)
    {
        var newDepartmentCode = department[RefStaff.Units.Code];
        var existedDepartmentCode = employeeDocsvision.SubSection.ParentRow[RefStaff.Units.Code];
        if (!newDepartmentCode.Equals(existedDepartmentCode))
            employeeDocsvision.Move(Guid.Empty, department.Id);
    }

    private void SetEmployeeStatus(bool? actual, RowData employeeDocsvision)
    {
        if (actual.HasValue && !actual.Value)
            employeeDocsvision[RefStaff.Employees.Status] = StaffEmployeeStatus.Discharged;
        else
            employeeDocsvision[RefStaff.Employees.Status] = StaffEmployeeStatus.Active;
    }

    private bool SearchEmployeeTest(CardData refStaffData, string code)
    {
        SectionQuery sectionQueryTest = Session.CreateSectionQuery();
        sectionQueryTest.ConditionGroup.Conditions.AddNew(RefStaff.Employees.IDCode, FieldType.Unistring,
                                                      ConditionOperation.Equals, code);

        RowDataCollection employeeCollection = refStaffData.Sections[RefStaff.Employees.ID].FindRows(sectionQueryTest.GetXml());
        return (employeeCollection.Count > 0);
    }
	
	
	public RowData CreateEmployee(RowData department)
    {
        if (department == null)
            return null;

        RowDataCollection employeeRows = department.ChildSections[RefStaff.Employees.ID].Rows;
        return employeeRows.AddNew();
    }



    //Подразделения
    //im public string AddDepartment(string unitId, string unit, string unitFull, string ownerId, string chief, string curator)
	public object AddDepartment(string unitId, string unit, string unitFull, string ownerId, string chief, string curator)
    {
		
		try{
			
			if (string.IsNullOrEmpty(unit))
			{
				//logger.Info(string.Format(base.LogEmptyInfo, "Unit"));
				//return string.Format(LogEmptyInfo, "Unit");
				return true;
			}
			
			CardData refStaffData = Session.CardManager.GetDictionaryData(RefStaff.ID);

			SectionData unitsSection = refStaffData.Sections[RefStaff.Units.ID];
			
			//im RowData department = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, unitId);
			RowDataCollection departmentColl = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, unitId);
			if (departmentColl == null || departmentColl.Count == 0)
			{
				RowData unitRow = null;
				
				//Поиск по родителю

				if (!String.IsNullOrEmpty(ownerId))
				{
					//im unitRow = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, ownerId);
					RowDataCollection unitRowColl = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, ownerId);
					if (unitRowColl != null)
						unitRow = unitRowColl[0];
				}

				if (unitRow == null)
					unitRow = unitsSection.GetRow(new Guid(ConfigurationManager.AppSettings["RootDepartmentGuid"])); //Нужно брать по известному ID подразделения            

				RowDataCollection departmentRows = unitRow.ChildRows;
				
				refStaffData.BeginUpdate();
				
				var newRowDepartment = departmentRows.AddNew();
				var isFirstLoad = ConfigurationManager.AppSettings["FirstLoad"];
				if (isFirstLoad == "1")
				{
					newRowDepartment[RefStaff.Units.Name] = unit;
					newRowDepartment[RefStaff.Units.FullName] = unitFull;
				}
				else
				{
					newRowDepartment[RefStaff.Units.Name] = unit;
					newRowDepartment[RefStaff.Units.FullName] = unitFull;
				}
				newRowDepartment[RefStaff.Units.Code] = unitId;
				newRowDepartment[RefStaff.Units.Type] = StaffUnitType.Department;
				
				//Руководитель подразделения
				/*im
				UpdateDepartmentChief(newRowDepartment, refStaffData, chief);
				*/
				if (!String.IsNullOrEmpty(chief))
				{
					//im unitRow = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, ownerId);
					RowDataCollection unitRowColl = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, chief);
					if (unitRowColl != null)
					{
						unitRow = unitRowColl[0];
					}
				}
				//Куратор подразделения
				/*
				UpdateDepartmentCurator(newRowDepartment, refStaffData, curator);
				*/
				if (!String.IsNullOrEmpty(curator))
				{
					//im unitRow = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, ownerId);
					RowDataCollection unitRowColl = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, curator);
					if (unitRowColl != null)
						unitRow = unitRowColl[0];
				}
				
				refStaffData.EndUpdate();
				//return "Департамент добавлен";
				return true;
			}
			else{
				//return "Подразделение уже добавлено";
				return true;
			}
		}
		catch{
			CardData refStaffData = Session.CardManager.GetDictionaryData(RefStaff.ID);
			refStaffData.BeginUpdate();
			refStaffData.EndUpdate();
			return "Ошибка при попытке добавить новое подразделение. Необходимо перезапустить сервис!";
			//return false;
		}
    }

    //public string UpdateDepartment(string unitId, string unit, string unitFull, string ownerId, string chief, string curator)
	public object UpdateDepartment(string unitId, string unit, string unitFull, string ownerId, string chief, string curator)
    {
        if (string.IsNullOrEmpty(unit))
        {
            //logger.Info(string.Format(LogEmptyInfo, "Unit"));
            //return string.Format(LogEmptyInfo, "Unit");
			return true;
        }
		
        CardData refStaffData = Session.CardManager.GetDictionaryData(RefStaff.ID);
        SectionData unitsSection = refStaffData.Sections[RefStaff.Units.ID];
		
        //im RowData department = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, unitId);
		RowDataCollection departmentColl = GetExistObject(RefStaff.Units.ID, RefStaff.Units.Code, unitId);
		RowData department = null;
		try{
			department = departmentColl[0];
		}
		catch{}
		
        if (department == null)
            return AddDepartment(unitId, unit, unitFull, ownerId, chief, curator);
		
        refStaffData.BeginUpdate();
		
        //Родитель
		if (ownerId == unitId)
			return true;
        var existedParentDepartmentCode = department.ParentRow[RefStaff.Units.Code];
        if (!(existedParentDepartmentCode == null && ownerId == null) && !ownerId.Equals(existedParentDepartmentCode))
        {
            SectionQuery sectionQueryParentDepartment = Session.CreateSectionQuery();
            sectionQueryParentDepartment.ConditionGroup.Conditions.AddNew(RefStaff.Units.Code, FieldType.Unistring,
                                                                                ConditionOperation.Equals, ownerId);

            RowDataCollection departmentParentCollection = unitsSection.FindRows(sectionQueryParentDepartment.GetXml());
            if (departmentParentCollection != null && departmentParentCollection.Count > 0)            
                department.Move(departmentParentCollection[0].Id, Guid.Empty);            
        }

        refStaffData.EndUpdate();

        refStaffData.BeginUpdate();
		try{
			var isFirstLoad = ConfigurationManager.AppSettings["FirstLoad"];
			if (isFirstLoad == "1")
			{
				department[RefStaff.Units.Name] = unit;
				department[RefStaff.Units.FullName] = unitFull;
			}
			else
			{
				department[RefStaff.Units.Name] = unit;
				department[RefStaff.Units.FullName] = unitFull;
			}
			department[RefStaff.Units.Code] = unitId;
			
			//Руководитель подразделения
			if(chief != string.Empty)
			{
				try{
				UpdateDepartmentChief(department, refStaffData, chief);
				}
				catch{
					refStaffData.EndUpdate();
					return "Руководитель не найден.";
				}
			}
			
			//Куратор подразделения
			if(curator != string.Empty)
			{
				try{
				UpdateDepartmentCurator(department, refStaffData, curator);
				}
				catch{
					refStaffData.EndUpdate();
					return "Куратор не найден.";
				}
			}
			
			refStaffData.EndUpdate();
			
			//return "Департамент обновлен";
			return true;
		}
		catch{
			return "Ошибка при попытке обновить подразделение. Необходимо перезапустить сервис!";
		}
    }


	
    public string DeleteDepartment(string unitId)
    {
		/*
        if (string.IsNullOrEmpty(unitId))
        {
            //logger.Info(string.Format(base.LogEmptyInfo, "UnitId"));
            return string.Format(LogEmptyInfo, "UnitId");
        }

        CardData refStaffData = Session.CardManager.GetDictionaryData(RefStaff.ID);
        SectionData unitsSection = refStaffData.Sections[RefStaff.Units.ID];

        RowData department = GetExistDepartament(RefStaff.Units.ID, RefStaff.Units.Code, unitId);
        if (department != null)
        {
            refStaffData.BeginUpdate();
            unitsSection.DeleteRow(department.Id);
            refStaffData.EndUpdate();
        }*/
		
		return "Департамент хотели удалить";
    }

    private void UpdateDepartmentChief(RowData departmentDocsvision, CardData refStaffData, string chief)
    {
        if (chief != null)
        {
            //im RowData employee = GetExistObject(RefStaff.Employees.ID, RefStaff.Employees.IDCode, chief);
			RowDataCollection employeeCollection = GetExistObject(RefStaff.Employees.ID, RefStaff.Employees.IDCode, chief);
			RowData employee = employeeCollection[0];
            if (employee != null)
                departmentDocsvision[RefStaff.Units.Manager] = employee.Id;
        }
        else
        {
            //Поиск по руководителя по родителю
            var existedParentDepartment = departmentDocsvision.ParentRow;
            if (existedParentDepartment != null && existedParentDepartment[RefStaff.Units.Manager] != null)
                departmentDocsvision[RefStaff.Units.Manager] = existedParentDepartment[RefStaff.Units.Manager];
        }
    }

    public void UpdateDepartmentCurator(RowData departmentDocsvision, CardData refStaffData, string curator)
    {
        if (curator != null)
        {
            //im RowData employee = GetExistObject(RefStaff.Employees.ID, RefStaff.Employees.IDCode, curator);
			RowDataCollection employeeColl = GetExistObject(RefStaff.Employees.ID, RefStaff.Employees.IDCode, curator);
			RowData employee = employeeColl[0];
            if (employee != null)
                departmentDocsvision[RefStaff.Units.ContactPerson] = employee.Id;
        }
    }

    //Должности

    public string AddPosition(string titleId, string title)
    {
        if (string.IsNullOrEmpty(title) || string.IsNullOrEmpty(titleId))
        {
            //logger.Info(string.Format(base.LogEmptyInfo, "Title, TitleId"));
            return string.Format(LogEmptyInfo, "Title, TitleId");
        }

        CardData refStaffData = Session.CardManager.GetDictionaryData(RefStaff.ID);
        SectionData positionSection = refStaffData.Sections[RefStaff.Positions.ID];

        refStaffData.BeginUpdate();

        var newRowDepartment = positionSection.Rows.AddNew();
        newRowDepartment[RefStaff.Positions.Name] = title;
        newRowDepartment[RefStaff.Positions.SyncTag] = titleId;

        refStaffData.EndUpdate();

        return "Должность добавлена";
    }
    
    public string UpdatePosition(string titleId, string title)
    {
        if (string.IsNullOrEmpty(title) || string.IsNullOrEmpty(titleId))
        {
            //logger.Info(string.Format(base.LogEmptyInfo, "Title, TitleId"));
            return string.Format(LogEmptyInfo, "Title, TitleId");
        }

        CardData refStaffData = Session.CardManager.GetDictionaryData(RefStaff.ID);
        SectionData positionSection = refStaffData.Sections[RefStaff.Positions.ID];

        //im RowData position = GetExistObject(RefStaff.Positions.ID, RefStaff.Positions.Name, title);
		RowDataCollection positionColl = GetExistObject(RefStaff.Positions.ID, RefStaff.Positions.Name, title);
		RowData position = positionColl[0];
        if (position == null)        
            return AddPosition(titleId, title);        
        
        refStaffData.BeginUpdate();
        //position[RefStaff.Positions.Name] = title;
        position[RefStaff.Positions.SyncTag] = titleId;
        refStaffData.EndUpdate();

        return "Должность обновлена";        
    }

    public string DeletePosition(string titleId)
    {
        if (string.IsNullOrEmpty(titleId))
        {
            //logger.Info(string.Format(base.LogEmptyInfo, "TitleId"));
            return string.Format(LogEmptyInfo, "TitleId");
        }

        CardData refStaffData = Session.CardManager.GetDictionaryData(RefStaff.ID);
        SectionData unitsSection = refStaffData.Sections[RefStaff.Positions.ID];
        //im RowData position = GetExistObject(RefStaff.Positions.ID, RefStaff.Positions.SyncTag, titleId);
		RowDataCollection positionColl = GetExistObject(RefStaff.Positions.ID, RefStaff.Positions.SyncTag, titleId);
		RowData position = positionColl[0];
        if (position != null)
        {
            refStaffData.BeginUpdate();
            unitsSection.DeleteRow(position.Id);
            refStaffData.EndUpdate();
        }

        return "Должность удалена";
    }

    //im private RowData GetExistObject(Guid mode, string alias, string value)
	private RowDataCollection GetExistObject(Guid mode, string alias, string value)
    {
        if (string.IsNullOrEmpty(alias) || string.IsNullOrEmpty(value) || mode == null || mode == Guid.Empty)
        {
            return null;
        }
        try
        {
            CardData refStaffData = Session.CardManager.GetDictionaryData(RefStaff.ID);
            SectionData unitsSection = null;

            if (mode == RefStaff.Employees.ID)            
                unitsSection = refStaffData.Sections[RefStaff.Employees.ID];            
            else if (mode == RefStaff.Units.ID)            
                unitsSection = refStaffData.Sections[RefStaff.Units.ID];            
            else if (mode == RefStaff.Positions.ID)            
                unitsSection = refStaffData.Sections[RefStaff.Positions.ID];            

            if (unitsSection == null)            
                return null;            

            SectionQuery sectionQuery = Session.CreateSectionQuery();
            sectionQuery.ConditionGroup.Conditions.AddNew(alias, FieldType.Unistring, ConditionOperation.Equals, value);

            RowDataCollection employeeCollection = unitsSection.FindRows(sectionQuery.GetXml());
            if (employeeCollection != null && employeeCollection.Count > 0)
			{
				//im return employeeCollection[0];
				return employeeCollection;
			}
        }
        catch (Exception ex)
        {	
			return null;
		}

        return null;
    }

    #region Utils

    private UserSession _userSession = null;
    public UserSession Session
    {
        get
        {

            if (_userSession == null)
            {
                if (_application["DocsvisionSession"] != null && SessionIsLive(_application["DocsvisionSession"]))
                {
                    _userSession = (UserSession)_application["DocsvisionSession"];
                }
                else
                {
                    _userSession = CreateSession();

                    _application.Lock();
                    _application["DocsvisionSession"] = _userSession;
                    _application.UnLock();
                }
            }

            return _userSession;
        }
    }

    private bool SessionIsLive(object v)
    {
        if (v == null)
            return false;

        try
        {
            UserSession session = (UserSession)v;
            CardData staffData = session.CardManager.GetDictionaryData(RefStaff.ID);
            return (staffData != null);
        }
        catch
        {
            return false;
        }
    }

    private UserSession CreateSession()
    {
        SessionManager sessionManager = SessionManager.CreateInstance();
        sessionManager.Connect(
            CSXMLConfigManager.GetDVServerURL(),
            CSXMLConfigManager.GetDVDatabase(),
            CSXMLConfigManager.GetDVUser(),
            CSXMLConfigManager.GetDVPassword()
            );
        return sessionManager.CreateSession();
    }

    private ObjectContext _objectContext = null;
    public ObjectContext ObjectContext
    {
        get
        {
            if (_objectContext == null)
            {
                _objectContext = GetObjectContext(Session);
            }

            return _objectContext;
        }
    }

    private static ObjectContext GetObjectContext(UserSession session)
    {
        UserSession userSession = session;

        // Инициализация сервис-провайдера                            
        var sessionContainer = new System.ComponentModel.Design.ServiceContainer();
        sessionContainer.AddService(typeof(DocsVision.Platform.ObjectManager.UserSession), userSession);

        // Инициализация контекста объектов
        // В качестве контейнера может выступать компонент карточки, унаследованный от DocsVision.Platform.WinForms.CardControl
        ObjectContext objectContext = new ObjectContext(sessionContainer);

        // Получение сервис-реестра и регистрация фабрик преобразователей
        IObjectMapperFactoryRegistry mapperFactoryRegistry = objectContext.GetService<IObjectMapperFactoryRegistry>();
        mapperFactoryRegistry.RegisterFactory(typeof(SystemCardsMapperFactory));
        mapperFactoryRegistry.RegisterFactory(typeof(BackOfficeMapperFactory));
        mapperFactoryRegistry.RegisterFactory(typeof(ApprovalDesignerMapperFactory));

        // Получение сервис-реестра и регистрация фабрик сервисов
        IServiceFactoryRegistry serviceFactoryRegistry = objectContext.GetService<IServiceFactoryRegistry>();
        serviceFactoryRegistry.RegisterFactory(typeof(SystemCardsServiceFactory));
        serviceFactoryRegistry.RegisterFactory(typeof(BackOfficeServiceFactory));
        serviceFactoryRegistry.RegisterFactory(typeof(ApprovalDesignerServiceFactory));

        // Регистрация сервиса для работы с хранилищем Docsvision
        objectContext.AddService<IPersistentStore>(DocsVisionObjectFactory.CreatePersistentStore(new SessionProvider(userSession), null));

        // Регистрация поставщика метаданных карточек
        IMetadataProvider metadataProvider = DocsVisionObjectFactory.CreateMetadataProvider(userSession);
        objectContext.AddService<IMetadataManager>(DocsVisionObjectFactory.CreateMetadataManager(metadataProvider, userSession));
        objectContext.AddService<IMetadataProvider>(metadataProvider);

        return objectContext;
    }

    #endregion
}