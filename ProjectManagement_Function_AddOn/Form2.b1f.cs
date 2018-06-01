using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;
using System.Data;
using System.IO;
using Excel;
using SAPbouiCOM;
using System.Configuration;
using SAPbobsCOM;
using System.Globalization;
using System.Data.SqlClient;

namespace ProjectManagement_Function_AddOn
{
    [FormAttribute("ProjectManagement_Function_AddOn.Form2", "Form2.b1f")]
    class Form2 : UserFormBase
    {
        
        CompanyService oCompServ = null;
        ProjectManagementService pmgService = null;
        SAPbobsCOM.Company oCompany = null;
        SAPbouiCOM.Application oApp = null;
        SAPbouiCOM.Form oForm = null;
        DataSet ds_import = null;
        SqlConnection conn = null;
        public Form2()
        {
            //obj.SBO_Application.AppEvent += new _IApplicationEvents_AppEventEventHandler(SBO_AppEvent);
            //SAPbouiCOM.Framework.Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_path").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txt_pname").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("txt_bpname").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_25").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("txt_ctg").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Grid").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_path").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("bt_import").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("bt_octg").Specific));
            this.Button4.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button4_PressedAfter);
            this.Button5 = ((SAPbouiCOM.Button)(this.GetItem("bt_p").Specific));
            this.Button5.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button5_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {

        }

        private void OnCustomInitialize()
        {
            this.oApp = (SAPbouiCOM.Application)SAPbouiCOM.Framework.Application.SBO_Application;
            this.oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            //Create Connection SQL
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery("Select * from [@ADDONCFG]");
            if (oR_RecordSet.RecordCount > 0)
            {
                string uid = oR_RecordSet.Fields.Item("Code").Value.ToString();
                string pwd = oR_RecordSet.Fields.Item("Name").Value.ToString();
                conn = new SqlConnection(string.Format("Data Source={0}; Initial Catalog={1}; User id={2}; Password={3};", oCompany.Server, oCompany.CompanyDB, uid, pwd));
            }
            else
            {
                oApp.MessageBox("Can't connect DB !");
            }

        }

        private Folder Folder2;
        private Folder Folder3;
        private StaticText StaticText2;
        private EditText EditText0;
        private StaticText StaticText3;
        private StaticText StaticText4;
        private EditText EditText1;
        private EditText EditText2;
        private SAPbouiCOM.Button Button0;
        private Grid Grid0;
        private SAPbouiCOM.Button Button1;
        private StaticText StaticText5;
        private EditText EditText6;
        private SAPbouiCOM.Button Button4;
        private SAPbouiCOM.Button Button5;

        public DataSet readFile(string pathSource)
        {
            //string pathSource = HttpContext.Current.Server.MapPath("~/POS/" + fileName);
            using (var stream = File.Open(pathSource, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader = null;
                if (pathSource.Substring(pathSource.LastIndexOf('.')) == ".xls")
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (pathSource.Substring(pathSource.LastIndexOf('.')) == ".xlsx")
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                if (reader == null)
                    return null;
                reader.IsFirstRowAsColumnNames = true;
                DataSet ds = reader.AsDataSet();
                return ds;
            }
        }

        private string Create_Project(string pName, string pStart, string pSeries, string pType, string pCardCode, string pCardName, string pContact, string pEmployee, string pWithPhases, string pDueDate, string pClosingDate, string pFinancialProject, string pU_BPTH, string pU_PRJTYPE, string pU_PRJGROUP, string pU_CPHT1, string pU_CPHT2, string pU_DPBH, string pU_DPCP, string pU_CPNG, string pU_CPQLCT)
        {
            string New_ProjectNo = "-1";
            try
            {
                //Use UI API to create Project
                oApp.ActivateMenuItem("48897");
                var oform1 = oApp.Forms.ActiveForm;
                //Project Type
                if (pType == "E")
                {
                    var orad_extype = oform1.Items.Item("234000010").Specific;
                    ((SAPbouiCOM.OptionBtn)orad_extype).Selected = true;
                }
                else
                {
                    var orad_intype = oform1.Items.Item("234000011").Specific;
                    ((SAPbouiCOM.OptionBtn)orad_intype).Selected = true;
                }
                //BPCode
                if (!String.IsNullOrEmpty(pCardCode) && pCardCode != "NULL")
                {
                    var txt_bpcode = oform1.Items.Item("234000013").Specific;
                    ((SAPbouiCOM.EditText)txt_bpcode).Value = pCardCode;
                }
                //BPName
                if (!String.IsNullOrEmpty(pCardName) && pCardName != "NULL")
                {
                    var txt_bpname = oform1.Items.Item("234000016").Specific;
                    ((SAPbouiCOM.EditText)txt_bpname).Value = pCardName;
                }
                //Contact Person
                if (!String.IsNullOrEmpty(pContact) && pContact != "NULL")
                {
                    var cbo_contactperson = oform1.Items.Item("234000018").Specific;
                    ((SAPbouiCOM.ComboBox)cbo_contactperson).Select(pContact);
                }
                //Sale Employee
                if (!String.IsNullOrEmpty(pEmployee) && pEmployee != "-1" && pEmployee != "NULL")
                {
                    var cbo_employee = oform1.Items.Item("234000023").Specific;
                    ((SAPbouiCOM.ComboBox)cbo_employee).Select(pEmployee);
                }
                //Owner
                //string owner = oR_Project_Info.Fields.Item("OWNER").Value.ToString();
                //if (!String.IsNullOrEmpty(owner) && owner != "0")
                //{
                //    var txt_owner = oform1.Items.Item("234000025").Specific;
                //    ((SAPbouiCOM.EditText)txt_owner).Value = owner;
                //}
                //Bo phan thuc hien
                if (!String.IsNullOrEmpty(pU_BPTH))
                {
                    var cbo_u_bpth = oform1.Items.Item("U_BPTH").Specific;
                    ((SAPbouiCOM.ComboBox)cbo_u_bpth).Select(pU_BPTH);
                }
                //Project with Subproject
                if (pWithPhases == "Y")
                {
                    var chb_withpahse = oform1.Items.Item("234000027").Specific;
                    ((SAPbouiCOM.CheckBox)chb_withpahse).Checked = true;
                }
                //Project Name
                if (!String.IsNullOrEmpty(pName))
                {
                    var txt_pname = oform1.Items.Item("234000029").Specific;
                    ((SAPbouiCOM.EditText)txt_pname).Value = pName;
                }
                //Series
                if (!String.IsNullOrEmpty(pSeries))
                {
                    var cbo_series = oform1.Items.Item("234000031").Specific;
                    ((SAPbouiCOM.ComboBox)cbo_series).Select(pSeries);
                }
                //Status
                //string status = oR_Project_Info.Fields.Item("STATUS").Value.ToString();
                //if (!String.IsNullOrEmpty(status))
                //{
                //    var cbo_status = oform1.Items.Item("234000034").Specific;
                //    ((SAPbouiCOM.ComboBox)cbo_status).Select(status);
                //}
                //Start Date
                //string strdate = oR_Project_Info.Fields.Item("START").Value.ToString();
                if (!String.IsNullOrEmpty(pStart))
                {
                    DateTime startdate = DateTime.ParseExact(pStart, "yyyyMMdd", CultureInfo.InvariantCulture);
                    if (startdate > DateTime.MinValue)
                    {
                        var txt_strdate = oform1.Items.Item("234000036").Specific;
                        ((SAPbouiCOM.EditText)txt_strdate).Value = startdate.ToString("yyyyMMdd");
                    }
                }
                //Due Date
                if (!String.IsNullOrEmpty(pDueDate))
                {
                    DateTime duedate = DateTime.ParseExact(pDueDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                    if (duedate > DateTime.MinValue)
                    {
                        var txt_duedate = oform1.Items.Item("234000038").Specific;
                        ((SAPbouiCOM.EditText)txt_duedate).Value = duedate.ToString("yyyyMMdd");
                    }
                }
                //Closing Date
                if (!String.IsNullOrEmpty(pClosingDate))
                {
                    DateTime clodate = DateTime.ParseExact(pClosingDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                    if (clodate > DateTime.MinValue)
                    {
                        var txt_closedate = oform1.Items.Item("234000040").Specific;
                        ((SAPbouiCOM.EditText)txt_closedate).Value = clodate.ToString("yyyyMMdd");
                    }
                }
                //Financial Project
                if (!String.IsNullOrEmpty(pFinancialProject))
                {
                    var txt_finan_project = oform1.Items.Item("234000049").Specific;
                    ((SAPbouiCOM.EditText)txt_finan_project).Value = pFinancialProject;
                }
                //Overview Panel - Risk Level
                //string risk = oR_Project_Info.Fields.Item("RISK").Value.ToString();
                //if (!String.IsNullOrEmpty(risk))
                //{
                //    oform1.Freeze(true);
                //    oform1.PaneLevel = oform1.Items.Item("234000056").FromPane;//oForm.Items.Item("148").FromPane
                //    var cbo_risk = oform1.Items.Item("234000056").Specific;
                //    ((SAPbouiCOM.ComboBox)cbo_risk).Select(risk);
                //    oform1.Freeze(false);
                //}
                //Overview Panel - Industry
                //string industry = oR_Project_Info.Fields.Item("INDUSTRY").Value.ToString();
                //if (!String.IsNullOrEmpty(industry))
                //{
                //    oform1.Freeze(true);
                //    oform1.PaneLevel = oform1.Items.Item("234000058").FromPane;//oForm.Items.Item("148").FromPane
                //    var cbo_industry = oform1.Items.Item("234000058").Specific;
                //    ((SAPbouiCOM.ComboBox)cbo_industry).Select(industry);
                //    oform1.Freeze(false);
                //}
                //Overview Panel - Comments
                //string comments = oR_Project_Info.Fields.Item("REASON").Value.ToString();
                //if (!String.IsNullOrEmpty(comments))
                //{
                //    var txt_comments = oform1.Items.Item("234000060").Specific;
                //    ((SAPbouiCOM.EditText)txt_comments).Value = comments;
                //}
                //Remark Panel - Free Text
                //string freetext = oR_Project_Info.Fields.Item("Free_Text").Value.ToString();
                //if (!String.IsNullOrEmpty(freetext))
                //{
                //    oform1.Freeze(true);
                //    oform1.PaneLevel = oform1.Items.Item("234000167").FromPane;//oForm.Items.Item("148").FromPane
                //    var txt_freetext = oform1.Items.Item("234000167").Specific;
                //    ((SAPbouiCOM.EditText)txt_freetext).Value = freetext;
                //    oform1.Freeze(false);
                //}
                //U Project Type
                if (!String.IsNullOrEmpty(pU_PRJTYPE))
                {
                    var cbo_u_ptype = oform1.Items.Item("U_PRJTYPE").Specific;
                    ((SAPbouiCOM.ComboBox)cbo_u_ptype).Select(pU_PRJTYPE);
                }
                //U Project Group
                if (!String.IsNullOrEmpty(pU_PRJGROUP))
                {
                    var cbo_u_ptype = oform1.Items.Item("U_PRJGROUP").Specific;
                    ((SAPbouiCOM.ComboBox)cbo_u_ptype).Select(pU_PRJGROUP);
                }
                //Overview Panel - U_CPHT1
                if (!String.IsNullOrEmpty(pU_CPHT1))
                {
                    oform1.Freeze(true);
                    oform1.PaneLevel = oform1.Items.Item("U_CPHT1").FromPane;
                    var txt_u_cpht1 = oform1.Items.Item("U_CPHT1").Specific;
                    ((SAPbouiCOM.EditText)txt_u_cpht1).Value = pU_CPHT1.ToString();
                    oform1.Freeze(false);
                }
                //Overview Panel - U_CPHT2
                if (!String.IsNullOrEmpty(pU_CPHT2))
                {
                    oform1.Freeze(true);
                    oform1.PaneLevel = oform1.Items.Item("U_CPHT2").FromPane;
                    var txt_u_cpht2 = oform1.Items.Item("U_CPHT2").Specific;
                    ((SAPbouiCOM.EditText)txt_u_cpht2).Value = pU_CPHT2.ToString();
                    oform1.Freeze(false);
                }
                //Overview Panel - U_DPBH
                if (!String.IsNullOrEmpty(pU_DPBH))
                {
                    oform1.Freeze(true);
                    oform1.PaneLevel = oform1.Items.Item("U_DPBH").FromPane;
                    var txt_u_dpbh = oform1.Items.Item("U_DPBH").Specific;
                    ((SAPbouiCOM.EditText)txt_u_dpbh).Value = pU_DPBH.ToString();
                    oform1.Freeze(false);
                }
                //Overview Panel - U_DPCP
                if (!String.IsNullOrEmpty(pU_DPCP))
                {
                    oform1.Freeze(true);
                    oform1.PaneLevel = oform1.Items.Item("U_DPCP").FromPane;
                    var txt_u_dpcp = oform1.Items.Item("U_DPCP").Specific;
                    ((SAPbouiCOM.EditText)txt_u_dpcp).Value = pU_DPCP.ToString();
                    oform1.Freeze(false);
                }
                //Overview Panel - U_CPNG
                if (!String.IsNullOrEmpty(pU_CPNG))
                {
                    oform1.Freeze(true);
                    oform1.PaneLevel = oform1.Items.Item("U_CPNG").FromPane;
                    var txt_u_cpng = oform1.Items.Item("U_CPNG").Specific;
                    ((SAPbouiCOM.EditText)txt_u_cpng).Value = pU_CPNG.ToString();
                    oform1.Freeze(false);
                }
                //Overview Panel - U_CPQLCT
                if (!String.IsNullOrEmpty(pU_CPQLCT))
                {
                    oform1.Freeze(true);
                    oform1.PaneLevel = oform1.Items.Item("U_CPQLCT").FromPane;
                    var txt_u_cpqlct = oform1.Items.Item("U_CPQLCT").Specific;
                    ((SAPbouiCOM.EditText)txt_u_cpqlct).Value = pU_CPQLCT.ToString();
                    oform1.Freeze(false);
                }
                //Get New Project No
                var txt_projectno_new = oform1.Items.Item("234000032").Specific;
                New_ProjectNo = ((SAPbouiCOM.EditText)txt_projectno_new).Value;
                //Click Add Button
                oform1.Items.Item("1").Click();
                //Check Add Project Complete ?
                if (((SAPbouiCOM.EditText)oform1.Items.Item("234000032").Specific).Value != New_ProjectNo)
                {
                    New_ProjectNo = GetNewProjectNo(pName);
                    oform1.Close();
                }
                else
                {
                    New_ProjectNo = "-1";
                    oApp.MessageBox("Can't create project !");
                }
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                if (pmgService != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pmgService);
                }
            }
            return New_ProjectNo;
        }
        
        private string Create_Subproject(int pProjectNo, int pParentID, string pName, int pOwner, string pStartDate, string pEndDate, int pType, double pContribution, double pPlanedCost)
        {
            string New_SubprojectEntry = "-1";
            try
            {
                oCompServ = (CompanyService)oCompany.GetCompanyService();
                pmgService = (ProjectManagementService)oCompServ.GetBusinessService(ServiceTypes.ProjectManagementService);
                PM_SubprojectDocumentData subproject = (PM_SubprojectDocumentData)pmgService.GetDataInterface(ProjectManagementServiceDataInterfaces.pmsPM_SubprojectDocumentData);
                if (pParentID != 0)
                    subproject.ParentID = pParentID;
                subproject.ProjectID = pProjectNo;
                subproject.SubprojectName = pName;
                if (pOwner != 0)
                    subproject.Owner = pOwner;
                if (!string.IsNullOrEmpty(pStartDate))
                {
                    subproject.StartDate = DateTime.ParseExact(pStartDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                }
                if (!string.IsNullOrEmpty(pEndDate))
                {
                    subproject.SubprojectEndDate = DateTime.ParseExact(pEndDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                }
                if (pType != 0)
                    subproject.SubprojectType = pType;
                subproject.SubprojectContribution = pContribution;
                subproject.PlannedCost = pPlanedCost;
                subproject.AbsEntry = -1;
                SAPbobsCOM.PM_SubprojectDocumentParams subprojectParam = pmgService.AddSubproject(subproject);
                New_SubprojectEntry = subprojectParam.AbsEntry.ToString();
            }
            catch (Exception ex)
            {
                oApp.SetStatusBarMessage(ex.Message,BoMessageTime.bmt_Short,true);
            }
            finally
            {
                if (pmgService != null)
                {
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(pmgService);
                }
            }
            return New_SubprojectEntry;
        }
        
        private int GetAbsEntry(string pDocNum)
        {
            int result = -1;
            Recordset oR_Project_Info = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string str_sql = string.Format("Select AbsEntry from OPMG where DocNum={0}", pDocNum);
            oR_Project_Info.DoQuery(str_sql);
            if (oR_Project_Info.RecordCount > 0)
            int.TryParse(oR_Project_Info.Fields.Item("AbsEntry").Value.ToString(), out result);
            return result;
        }
        
        private string GetNewProjectNo(string pProjectName)
        {
            string result = "";
            Recordset oR_Project_Info = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string str_sql = string.Format("Select DocNum from OPMG where [Name] = N'{0}'", pProjectName);
            oR_Project_Info.DoQuery(str_sql);
            if (oR_Project_Info.RecordCount > 0)
                result = oR_Project_Info.Fields.Item("DocNum").Value.ToString();
            return result;
        }
        
        private List<int> Get_AbsEntry_Subproject_List(int pProjectID, int pSubprojectID = 0, int pDepthLevel = 0)
        {
            List<int> kq = new List<int>();
            string querystr = "";
            if (pSubprojectID > 0)
                querystr = string.Format("Select AbsEntry from OPHA where ProjectID={0} and ParentID={1} and Level={2}", pProjectID, pSubprojectID, pDepthLevel);
            else
            {
                if (pDepthLevel == 0)
                    querystr = string.Format("Select AbsEntry from OPHA where ProjectID={0} and ParentID is null and Level=0", pProjectID);
                else
                    querystr = string.Format("Select AbsEntry from OPHA where ProjectID={0} and Level={1}", pProjectID, pDepthLevel);
            }
            Recordset oR_Subproject_List = null;
            oR_Subproject_List = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oR_Subproject_List.DoQuery(querystr);
            while (!oR_Subproject_List.EoF)
            {
                kq.Add((int)oR_Subproject_List.Fields.Item("AbsEntry").Value);
                oR_Subproject_List.MoveNext();
            }
            return kq;
        }
        
        private int Get_New_AbsEntry_Subproject(int pProject_AbsEntry)
        {
            int kq = -1;
            Recordset oR_Project_Info = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string str_sql = string.Format("Select Max(AbsEntry) as New_AbsEntry from OPHA a where a.ProjectID={0}", pProject_AbsEntry);
            oR_Project_Info.DoQuery(str_sql);
            int.TryParse(oR_Project_Info.Fields.Item("New_AbsEntry").Value.ToString(), out kq);
            return kq;
        }
        private bool Check_FinancialProject(string pFinancialProject)
        {
            Recordset oR_FProject_Count = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string str_sql = string.Format("Select count(*) from OPRJ a where a.PRJCode='{0}'", pFinancialProject);
            oR_FProject_Count.DoQuery(str_sql);
            if (oR_FProject_Count.Fields.Item(0).Value.ToString() == "0") return false;
            else return true;
        }
        private int Update_UDF_SubProject(int pAbsEntry, string pU_001, string pU_002, string pU_003, string pU_KLDT, string pU_DG, string pU_TTBV, string pU_TTDT, string pRemark, string pU_DGHD, string pU_TTHD)
        {
            //Update truc tiep UDF vao DB - SAP B1 9.2 PL8 chua ho tro UserFields
            int kq = 0;
            string str_command = "Update OPHA set U_001=@U_001, U_002=@U_002, U_003=@U_003, U_KLDT=@U_KLDT, U_DG=@U_DG, U_TTBV=@U_TTBV, U_TTDT = @U_TTDT, U_REMARK = @U_REMARK, U_DGHD=@U_DGHD, U_TTHD=@U_TTHD  where AbsEntry=@AbsEntry";
            
            //string str_comand = string.Format("Update OPHA set U_001='{0}',U_002='{1}',U_003='{2}',U_KLDT='{3}',U_DG={4},U_TTBV={5}, U_TTDT = {6}, U_REMARK = N'{7}' where AbsEntry={8}",
            //    pU_001, pU_002, string.IsNullOrEmpty(pU_003) ? "0" : pU_003
            //    , string.IsNullOrEmpty(pU_KLDT) ? "0" : pU_KLDT
            //    , string.IsNullOrEmpty(pU_DG) ? "0" : pU_DG
            //    , string.IsNullOrEmpty(pU_TTBV) ? "0" : pU_TTBV
            //    , string.IsNullOrEmpty(pU_TTDT) ? "0" : pU_TTDT
            //    , pRemark.Length > 254 ? pRemark : ""
            //    , pAbsEntry);
            SqlCommand cmd = new SqlCommand(str_command, conn);
            try
            {
                cmd.Parameters.AddWithValue("@U_001", pU_001);
                cmd.Parameters.AddWithValue("@U_002", pU_002);
                //Khoi luong Ban ve
                double U_003 = 0;
                double.TryParse(pU_003, out U_003);
                cmd.Parameters.AddWithValue("@U_003", Math.Round(U_003, 2, MidpointRounding.AwayFromZero));
                //Khoi luong dau thau
                double U_KLDT = 0;
                double.TryParse(pU_KLDT, out U_KLDT);
                cmd.Parameters.AddWithValue("@U_KLDT", Math.Round(U_KLDT,2, MidpointRounding.AwayFromZero));
                //Don gia
                double U_DG = 0;
                double.TryParse(pU_DG, out U_DG);
                cmd.Parameters.AddWithValue("@U_DG", Math.Round(U_DG, 0, MidpointRounding.AwayFromZero));
                //Thanh tien Ban ve
                double U_TTBV = 0;
                double.TryParse(pU_TTBV, out U_TTBV);
                cmd.Parameters.AddWithValue("@U_TTBV", Math.Round(U_TTBV, 0, MidpointRounding.AwayFromZero));
                //Thanh tien Dau thau
                double U_TTDT = 0;
                double.TryParse(pU_TTDT, out U_TTDT);
                cmd.Parameters.AddWithValue("@U_TTDT", Math.Round(U_TTDT, 0, MidpointRounding.AwayFromZero));
                //Don gia Hop dong
                double U_DGHD = 0;
                double.TryParse(pU_DGHD, out U_DGHD);
                cmd.Parameters.AddWithValue("@U_DGHD", Math.Round(U_DGHD, 0, MidpointRounding.AwayFromZero));
                //Thanh tien Hop dong
                double U_TTHD = 0;
                double.TryParse(pU_TTHD, out U_TTHD);
                cmd.Parameters.AddWithValue("@U_TTHD", Math.Round(U_TTHD, 0, MidpointRounding.AwayFromZero));
                //Remarks
                cmd.Parameters.AddWithValue("@U_REMARK", pRemark.Length > 254 ? pRemark : "");
                //Subproject Key
                cmd.Parameters.AddWithValue("@AbsEntry", pAbsEntry);
                conn.Open();
                kq = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }
            return kq;
        }
        
        private bool CheckExistUniqueID(SAPbouiCOM.Form pForm, string pItemID)
        {
            if (oForm.DataSources.DataTables.Count > 0)
            {
                for (int i = 0; i < pForm.DataSources.DataTables.Count; i++)
                {
                    if (pForm.DataSources.DataTables.Item(i).UniqueID == pItemID)
                    {
                        return true;
                    }
                }
                return false;
            }
            else
            {
                return false;
            }
        }
        
        private SAPbouiCOM.DataTable Convert_SAP_DataTable(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_SubProject"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_SubProject");
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_SubProject");
            }

            //Add column to DataTable
            foreach (System.Data.DataColumn c in pDataTable.Columns)
            {
                oDT.Columns.Add(c.ColumnName, BoFieldsType.ft_Text);

            }
            //Add row to DataTable
            foreach (System.Data.DataRow r in pDataTable.Rows)
            {
                oDT.Rows.Add();
                foreach (System.Data.DataColumn c in pDataTable.Columns)
                {
                    oDT.SetValue(c.ColumnName, oDT.Rows.Count - 1, r[c.ColumnName].ToString());
                }
            }
            //oDT.Columns.Add("oDC_AbsEntry", BoFieldsType.ft_Integer);
            //oDT.Columns.Add("oDC_Owner", BoFieldsType.ft_Integer);
            //oDT.Columns.Add("oDC_Name", BoFieldsType.ft_Text);
            //oDT.Columns.Add("oDC_Start", BoFieldsType.ft_Text);
            //oDT.Columns.Add("oDC_Finished", BoFieldsType.ft_Text);
            //oDT.Columns.Add("oDC_ParentID", BoFieldsType.ft_Integer);
            //oDT.Columns.Add("oDC_TYP", BoFieldsType.ft_Integer);
            //oDT.Columns.Add("oDC_CONTRIB", BoFieldsType.ft_Integer);
            //oDT.Columns.Add("oDC_PLANNED", BoFieldsType.ft_Integer);
            //oDT.Columns.Add("oDC_Level", BoFieldsType.ft_Integer);
            //oDT.Columns.Add("oDC_U_002", BoFieldsType.ft_Text);
            //oDT.Columns.Add("oDC_U_003", BoFieldsType.ft_Text);
            //oDT.Columns.Add("oDC_U_KLDT", BoFieldsType.ft_Text);

            return oDT;
        }

        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            GetFileNameClass oGetFileName = new GetFileNameClass();
            oGetFileName.Filter = "Excel 97 - 2003 Workbook (*.xls)|*.xls|Excel Workbook (*.xlsx)|*.xlsx";
            oGetFileName.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            Thread threadGetExcelFile = new Thread(new ThreadStart(oGetFileName.GetFileName));
            threadGetExcelFile.SetApartmentState(ApartmentState.STA); // = ApartmentState.STA;
            try
            {
                threadGetExcelFile.Start();
                while (!threadGetExcelFile.IsAlive) ; // Wait for thread to get started
                Thread.Sleep(1);  // Wait a sec more
                threadGetExcelFile.Join();    // Wait for thread to end
                // Use file name as you will here
                string strValue = oGetFileName.FileName;
                EditText0.Value = strValue;
                if (!string.IsNullOrEmpty(strValue))
                {
                    //Doc file Excel
                    ds_import = readFile(strValue);
                    if (ds_import != null && ds_import.Tables.Count == 2)
                    {
                        System.Data.DataTable dt_project = ds_import.Tables[0];
                        System.Data.DataTable dt_subproject = ds_import.Tables[1];
                        //System.Data.DataTable dt_subproject_stage = ds_import.Tables[2];
                        if (dt_project.Rows.Count > 0)
                        {
                            EditText1.Value = dt_project.Rows[0]["NAME"].ToString();
                            EditText2.Value = dt_project.Rows[0]["CARDNAME"].ToString();
                        }
                        //SAPbouiCOM.Form oForm = oApp.Forms.ActiveForm;
                        Grid0.DataTable = Convert_SAP_DataTable(dt_subproject);
                        Grid0.AutoResizeColumns();
                    }
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.Message);
            }
            threadGetExcelFile = null;
            oGetFileName = null;
        }

        private void Button1_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (ds_import != null && ds_import.Tables.Count == 2)
                {
                    System.Data.DataTable dt_project = ds_import.Tables[0];
                    System.Data.DataTable dt_subproject = ds_import.Tables[1];
                    //System.Data.DataTable dt_subproject_stage = ds_import.Tables[2];
                    //Import Project
                    string P_Name = dt_project.Rows[0]["NAME"].ToString();
                    string P_Start = dt_project.Rows[0]["START"].ToString();
                    string P_Series = dt_project.Rows[0]["Series"].ToString();
                    string P_Type = dt_project.Rows[0]["TYP"].ToString();
                    string P_CardCode = dt_project.Rows[0]["CARDCODE"].ToString();
                    string P_CardName = dt_project.Rows[0]["CARDNAME"].ToString();
                    string P_Contact = dt_project.Rows[0]["CONTACT"].ToString();
                    string P_Employee = dt_project.Rows[0]["EMPLOYEE"].ToString();
                    string P_WithPhases = dt_project.Rows[0]["WithPhases"].ToString();
                    string P_DueDate = dt_project.Rows[0]["DUEDATE"].ToString();
                    string P_ClosingDate = dt_project.Rows[0]["CLOSING"].ToString();
                    string P_FinancialProject = dt_project.Rows[0]["FIPROJECT"].ToString();
                    string P_U_BPTH = dt_project.Rows[0]["U_BPTH"].ToString();
                    string P_U_PRJTYPE = dt_project.Rows[0]["U_PRJTYPE"].ToString();
                    string P_U_PRJGROUP = dt_project.Rows[0]["U_PRJGROUP"].ToString();
                    string P_U_CPHT1 = dt_project.Rows[0]["U_CPHT1"].ToString();
                    string P_U_CPHT2 = dt_project.Rows[0]["U_CPHT2"].ToString();
                    string P_U_DPBH = dt_project.Rows[0]["U_DPBH"].ToString();
                    string P_U_DPCP = dt_project.Rows[0]["U_DPCP"].ToString();
                    string P_U_CPNG = dt_project.Rows[0]["U_CPNG"].ToString();
                    string P_U_CPQLCT = dt_project.Rows[0]["U_CPQLCT"].ToString();
                    //Create Financial Project
                    if (!Check_FinancialProject(P_FinancialProject))
                    {
                        SAPbobsCOM.CompanyService oCmpSrv = null;
                        SAPbobsCOM.IProjectsService projectService = null;
                        SAPbobsCOM.Project project = null;
                        oCmpSrv = oCompany.GetCompanyService();
                        projectService = (SAPbobsCOM.IProjectsService)oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService);
                        project = (SAPbobsCOM.Project)projectService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProject);
                        project.Code = P_FinancialProject;
                        project.Name = P_Name;
                        project.ValidFrom = DateTime.ParseExact(P_Start, "yyyyMMdd", CultureInfo.InvariantCulture);
                        projectService.AddProject(project);
                    }
                    //Create Project
                    string NewProjectNo = Create_Project(P_Name, P_Start, P_Series, P_Type, P_CardCode, P_CardName, P_Contact, P_Employee, P_WithPhases, P_DueDate, P_ClosingDate, P_FinancialProject, P_U_BPTH, P_U_PRJTYPE, P_U_PRJGROUP, P_U_CPHT1, P_U_CPHT2, P_U_DPBH, P_U_DPCP, P_U_CPNG, P_U_CPQLCT);
                    //Import Subproject
                    //GetAbsEntry
                    if (!string.IsNullOrEmpty(NewProjectNo) && NewProjectNo != "-1")
                    {
                        Dictionary<int, int> mapping_sub = new Dictionary<int, int>();
                        if (dt_subproject.Rows.Count > 0)
                        {
                            int maxlevel_sub = -1;
                            int.TryParse(dt_subproject.Compute("Max(Level)", "").ToString(), out maxlevel_sub);
                            int AbsProject = GetAbsEntry(NewProjectNo);
                            for (int i = 0; i <= maxlevel_sub; i++)
                            {
                                oApp.StatusBar.SetText("Import SubProject Level: " + i.ToString(), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                                DataRow[] total_row = dt_subproject.Select("Level =" + i.ToString(), "AbsEntry ASC");
                                int tmp = 1;
                                foreach (DataRow r in total_row)
                                {
                                    int Typ = 0, owner = 0;
                                    int.TryParse(r["TYP"].ToString(), out Typ);
                                    int.TryParse(r["OWNER"].ToString(), out owner);
                                    double contribution = 0, plan = 0;
                                    double.TryParse(r["CONTRIB"].ToString(), out contribution);
                                    double.TryParse(r["PLANNED"].ToString(), out plan);
                                    int parentID = 0;
                                    int.TryParse(r["ParentID"].ToString(), out parentID);
                                    if (i == 0)
                                        Create_Subproject(AbsProject, 0, r["Name"].ToString().Length > 254 ? r["Name"].ToString().Substring(0, 250) + "..." : r["Name"].ToString(), owner, r["START"].ToString(), r["FINISHED"].ToString(), Typ, contribution, plan);
                                    else
                                        Create_Subproject(AbsProject, mapping_sub[parentID], r["Name"].ToString().Length > 254 ? r["Name"].ToString().Substring(0, 250) + "..." : r["Name"].ToString(), owner, r["START"].ToString(), r["FINISHED"].ToString(), Typ, contribution, plan);
                                    int NewSubprojectEntry = Get_New_AbsEntry_Subproject(AbsProject);
                                    Update_UDF_SubProject(NewSubprojectEntry, r["U_001"].ToString(), r["U_002"].ToString(), r["U_003"].ToString(), r["U_KLDT"].ToString(), r["U_DG"].ToString(), r["U_TTBV"].ToString(), r["U_TTDT"].ToString(), r["Name"].ToString(), r["U_DGHD"].ToString(), r["U_TTHD"].ToString());
                                    mapping_sub.Add(int.Parse(r["AbsEntry"].ToString()), NewSubprojectEntry);
                                    oApp.StatusBar.SetText("Importing SubProject Level " + i.ToString() + ": " + tmp.ToString() + "/" + total_row.Count().ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                    tmp++;
                                }
                            }
                            //Import Stage
                            //if (dt_subproject_stage.Rows.Count > 0)
                            //{
                            //foreach (DataRow r in dt_subproject_stage.Rows)
                            //{
                            //    oApp.StatusBar.SetText("Import Stage SubProject: " + mapping_sub[int.Parse(r["AbsEntry"].ToString())].ToString(), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                            //    Create_Stage_Subproject(mapping_sub[int.Parse(r["AbsEntry"].ToString())], r["START"].ToString(), r["CLOSE"].ToString(), r["DESC"].ToString());
                            //    int NewAbsEntry_SubProject = mapping_sub[int.Parse(r["AbsEntry"].ToString())];
                            //    int New_LineID = Get_New_LineID_Stage_SubProject(NewAbsEntry_SubProject);
                            //    if (New_LineID != -1)
                            //    {
                            //        //Get parameter UDF
                            //        string vItemNo = r["U_ITEMNO"].ToString();
                            //        string vItemName = r["U_ITEMNAME"].ToString();
                            //        string vLoaichiphi = r["U_LOAICHIPHI"].ToString();
                            //        string vMaThietBi = r["U_MATHIETBI"].ToString();
                            //        string vTenThietBi = r["U_TENTHIETBI"].ToString();
                            //        string vBusinessPartner = r["U_BUSINESSPARTNER"].ToString();
                            //        string vBusinessName = r["U_BUSINESSNAME"].ToString();
                            //        string vDVT = r["U_DVT"].ToString();
                            //        double vDinhmuc = 0;
                            //        double.TryParse(r["U_DINHMUC"].ToString(), out vDinhmuc);
                            //        double vHaohut = 0;
                            //        double.TryParse(r["U_HAOHUT"].ToString(), out vHaohut);
                            //        double vDGDAUTHAU = 0;
                            //        double.TryParse(r["U_DGDAUTHAU"].ToString(), out vDGDAUTHAU);
                            //        double vDGDUPHONG = 0;
                            //        double.TryParse(r["U_DGDUPHONG"].ToString(), out vDGDUPHONG);
                            //        double vTTDAUTHAU = 0;
                            //        double.TryParse(r["U_TTDAUTHAU"].ToString(), out vTTDAUTHAU);
                            //        double vDGNCC = 0;
                            //        double.TryParse(r["U_DGNCC"].ToString(), out vDGNCC);
                            //        double vDGNTP = 0;
                            //        double.TryParse(r["U_DGNTP"].ToString(), out vDGNTP);
                            //        double vDGDTC = 0;
                            //        double.TryParse(r["U_DGDTC"].ToString(), out vDGDTC);
                            //        double vDGVTP = 0;
                            //        double.TryParse(r["U_DGVTP"].ToString(), out vDGVTP);
                            //        double vDGVC = 0;
                            //        double.TryParse(r["U_DGVC"].ToString(), out vDGVC);
                            //        double vDGCN = 0;
                            //        double.TryParse(r["U_DGCN"].ToString(), out vDGCN);
                            //        double vDGK = 0;
                            //        double.TryParse(r["U_DGK"].ToString(), out vDGK);
                            //        double vDGMUABAN = 0;
                            //        double.TryParse(r["U_DGMUABAN"].ToString(), out vDGMUABAN);
                            //        double vDGTHUE = 0;
                            //        double.TryParse(r["U_DGTHUE"].ToString(), out vDGTHUE);
                            //        double vSLDUTRU = 0;
                            //        double.TryParse(r["U_SLDUTRU"].ToString(), out vSLDUTRU);
                            //        double vDGVH = 0;
                            //        double.TryParse(r["U_DGVH"].ToString(), out vDGVH);
                            //        double vDGVCTB = 0;
                            //        double.TryParse(r["U_DGVCTB"].ToString(), out vDGVCTB);
                            //        string vDVTTB = r["U_DVTTB"].ToString();
                            //        double vDGPRELIM = 0;
                            //        double.TryParse(r["U_DGPRELIM"].ToString(), out vDGPRELIM);
                            //        double vDGTB = 0;
                            //        double.TryParse(r["U_DGTB"].ToString(), out vDGTB);
                            //        double vDGDP = 0;
                            //        double.TryParse(r["U_DGDP"].ToString(), out vDGDP);
                            //        double vDGDP2 = 0;
                            //        double.TryParse(r["U_DGDP2"].ToString(), out vDGDP2);
                            //        //Update UDF Stage SubProject
                            //        Update_UDF_Stage_SubProject(NewAbsEntry_SubProject, New_LineID, vItemNo, vItemName, vLoaichiphi, vMaThietBi, vTenThietBi, vBusinessPartner, vBusinessName, vDVT, vDinhmuc, vHaohut, vDGDAUTHAU, vDGDUPHONG, vTTDAUTHAU, vDGNCC, vDGNTP, vDGDTC, vDGVTP, vDGVC, vDGCN, vDGK, vDGMUABAN, vDGTHUE, vSLDUTRU, vDGVH, vDGVCTB, vDVTTB, vDGPRELIM, vDGTB, vDGDP,vDGDP2);
                            //    }
                            //    oApp.StatusBar.SetText("Create Stage of SubProject: " + r["AbsEntry"].ToString(), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                            //}
                            //}
                        }
                        oApp.MessageBox("Import Complete with ProjectNo: " + NewProjectNo.ToString());
                    }
                    else
                    {
                        oApp.MessageBox("Import Subproject: Failed");
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(EditText0.Value))
                    {
                        oApp.MessageBox("File Format is not correct");
                    }
                    else
                    {
                        oApp.MessageBox("Please choose file to import");
                    }
                }
            }
            catch (Exception ex)
            {
                oApp.MessageBox("Import error: " + ex.Message);
            }
        }

        private void Button4_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            GetFileNameClass oGetFileName = new GetFileNameClass();
            oGetFileName.Filter = "Excel 97 - 2003 Workbook (*.xls)|*.xls|Excel Workbook (*.xlsx)|*.xlsx";
            oGetFileName.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            Thread threadGetExcelFile = new Thread(new ThreadStart(oGetFileName.GetFileName));
            threadGetExcelFile.SetApartmentState(ApartmentState.STA); // = ApartmentState.STA;
            try
            {
                threadGetExcelFile.Start();
                while (!threadGetExcelFile.IsAlive) ; // Wait for thread to get started
                Thread.Sleep(1);  // Wait a sec more
                threadGetExcelFile.Join();    // Wait for thread to end
                // Use file name as you will here
                string strValue = oGetFileName.FileName;
                EditText6.Value = strValue;
                if (!string.IsNullOrEmpty(strValue))
                {
                    //Doc file Excel
                    ds_import = readFile(strValue);
                    if (ds_import != null && ds_import.Tables.Count == 4)
                    {
                        this.Button5.Item.Enabled = true;
                    }
                    else
                    {
                        this.Button5.Item.Enabled = false;
                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Wrong Format !");
                    }
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.Message);
            }
            threadGetExcelFile = null;
            oGetFileName = null;
        }

        private void Button5_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.Framework.Application.SBO_Application.ActivateMenuItem("CTG");
            oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            try
            {
                //Fill Project
                if (ds_import.Tables[0].Rows.Count > 0)
                {
                    ((EditText)oForm.Items.Item("20_U_E").Specific).Value = ds_import.Tables[0].Rows[0][0].ToString();
                    ((EditText)oForm.Items.Item("22_U_E").Specific).Value = DateTime.Today.ToString("yyyyMMdd");
                }
                //Fill Table Dau Thau
                oForm.Freeze(true);
                oForm.PaneLevel = 1;
                SAPbouiCOM.Matrix oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;
                for (int i = 0; i < ds_import.Tables[0].Rows.Count; i++)
                {
                    oApp.SetStatusBarMessage(string.Format("Import DT: {0}",(i+1) +"/" +ds_import.Tables[0].Rows.Count), BoMessageTime.bmt_Short, false);
                    DataRow r = ds_import.Tables[0].Rows[i];
                    //Subproject Code
                    ((EditText)oMtx.Columns.Item(2).Cells.Item(i + 1).Specific).Value = r["U_001"].ToString();
                    //Vat tu
                    ((EditText)oMtx.Columns.Item(3).Cells.Item(i + 1).Specific).Value = r["U_ITEMNO"].ToString();
                    //Ten vat tu
                    ((EditText)oMtx.Columns.Item(4).Cells.Item(i + 1).Specific).Value = r["U_ITEMNAME"].ToString();
                    //Don vi tinh
                    ((EditText)oMtx.Columns.Item(5).Cells.Item(i + 1).Specific).Value = r["U_DVT"].ToString();
                    //Dinh muc
                    ((EditText)oMtx.Columns.Item(6).Cells.Item(i + 1).Specific).Value = r["U_DinhMuc"].ToString();
                    //Hao hut
                    ((EditText)oMtx.Columns.Item(7).Cells.Item(i + 1).Specific).Value = r["U_HAOHUT"].ToString();
                    //Don gia dau thau
                    ((EditText)oMtx.Columns.Item(8).Cells.Item(i + 1).Specific).Value = r["U_DGDAUTHAU"].ToString();
                    //Don gia du phong
                    ((EditText)oMtx.Columns.Item(9).Cells.Item(i + 1).Specific).Value = r["U_DGDUPHONG"].ToString();
                    //Thanh tien dau thau
                    ((EditText)oMtx.Columns.Item(10).Cells.Item(i + 1).Specific).Value = r["U_TTDAUTHAU"].ToString();
                    if (i < ds_import.Tables[0].Rows.Count - 1)
                    {
                        oMtx.AddRow();
                        ((EditText)oMtx.Columns.Item(1).Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                    }
                }
                oForm.Freeze(false);
                //Fill Table Thiet bi
                oForm.Freeze(true);
                oForm.PaneLevel = 2;
                oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("1_U_G").Specific;
                for (int i = 0; i < ds_import.Tables[1].Rows.Count; i++)
                {
                    oApp.SetStatusBarMessage(string.Format("Import TB: {0}", (i + 1) + "/" + ds_import.Tables[1].Rows.Count), BoMessageTime.bmt_Short, false);
                    DataRow r = ds_import.Tables[1].Rows[i];
                    if (!string.IsNullOrEmpty(r[0].ToString()))
                    {
                        //SubProject Code
                        ((EditText)oMtx.Columns.Item("C_1_2").Cells.Item(i + 1).Specific).Value = r["U_001"].ToString();
                        //Ma thiet bi
                        ((EditText)oMtx.Columns.Item("C_1_4").Cells.Item(i + 1).Specific).Value = r["U_MATHIETBI"].ToString();
                        //Ten thiet bi
                        ((EditText)oMtx.Columns.Item("C_1_5").Cells.Item(i + 1).Specific).Value = r["U_TENTHIETBI"].ToString();
                        //Don vi tinh
                        ((EditText)oMtx.Columns.Item("C_1_6").Cells.Item(i + 1).Specific).Value = r["U_DVTTB"].ToString();
                        //So luong du tru
                        ((EditText)oMtx.Columns.Item("C_1_7").Cells.Item(i + 1).Specific).Value = r["U_SLDUTRU"].ToString();
                        //So luong thue
                        ((EditText)oMtx.Columns.Item("C_1_8").Cells.Item(i + 1).Specific).Value = r["U_SLTHUE"].ToString();
                        //So luong van chuyen
                        ((EditText)oMtx.Columns.Item("C_1_9").Cells.Item(i + 1).Specific).Value = r["U_SLVANCHUYEN"].ToString();
                        //So luong van hanh
                        ((EditText)oMtx.Columns.Item("C_1_10").Cells.Item(i + 1).Specific).Value = r["U_SLVANHANH"].ToString();
                        //Don gia mua ban
                        ((EditText)oMtx.Columns.Item("C_1_11").Cells.Item(i + 1).Specific).Value = r["U_DGMUABAN"].ToString();
                        //Don gia thue
                        ((EditText)oMtx.Columns.Item("C_1_12").Cells.Item(i + 1).Specific).Value = r["U_DGTHUE"].ToString();
                        //Don gia van chuyen
                        ((EditText)oMtx.Columns.Item("C_1_13").Cells.Item(i + 1).Specific).Value = r["U_DGVCTB"].ToString();
                        //Don gia van hanh
                        ((EditText)oMtx.Columns.Item("C_1_14").Cells.Item(i + 1).Specific).Value = r["U_DGVH"].ToString();

                        //Gia tri mua ban
                        ((EditText)oMtx.Columns.Item("C_1_15").Cells.Item(i + 1).Specific).Value = r["U_GTMB"].ToString();
                        //Gia tri thue
                        ((EditText)oMtx.Columns.Item("C_1_16").Cells.Item(i + 1).Specific).Value = r["U_GTTHUE"].ToString();
                        //Gia tri van chuyen
                        ((EditText)oMtx.Columns.Item("C_1_17").Cells.Item(i + 1).Specific).Value = r["U_GTVANCHUYEN"].ToString();
                        //Gia tri van hanh
                        ((EditText)oMtx.Columns.Item("C_1_18").Cells.Item(i + 1).Specific).Value = r["U_GTVANHANH"].ToString();
                        //Ngay cap
                        if (!string.IsNullOrEmpty(r["U_NGAYCAP"].ToString()))
                            ((EditText)oMtx.Columns.Item("C_1_19").Cells.Item(i + 1).Specific).Value = DateTime.ParseExact(r["U_NGAYCAP"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture).ToString("yyyyMMdd");
                        //Ngay tra
                        if (!string.IsNullOrEmpty(r["U_NGAYTRA"].ToString()))
                            ((EditText)oMtx.Columns.Item("C_1_20").Cells.Item(i + 1).Specific).Value = DateTime.ParseExact(r["U_NGAYTRA"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture).ToString("yyyyMMdd");
                        //Don vi mua ban
                        //((EditText)oMtx.Columns.Item("C_1_16").Cells.Item(i + 1).Specific).Value = r["U_DVMUABAN"].ToString();
                        //Ten don vi mua ban
                        //((EditText)oMtx.Columns.Item("C_1_17").Cells.Item(i + 1).Specific).Value = r["U_TDVMUABAN"].ToString();
                        //Don vi thue
                        //((EditText)oMtx.Columns.Item("C_1_18").Cells.Item(i + 1).Specific).Value = r["U_DVTHUE"].ToString();
                        //Ten don vi thue
                        //((EditText)oMtx.Columns.Item("C_1_19").Cells.Item(i + 1).Specific).Value = r["U_TDVTHUE"].ToString();
                        //Don vi van chuyen
                        //((EditText)oMtx.Columns.Item("C_1_20").Cells.Item(i + 1).Specific).Value = r["U_DVVC"].ToString();
                        //Ten don vi van chuyen
                        //((EditText)oMtx.Columns.Item("C_1_21").Cells.Item(i + 1).Specific).Value = r["U_TDVVC"].ToString();
                        //Don vi van hanh
                        //((EditText)oMtx.Columns.Item("C_1_22").Cells.Item(i + 1).Specific).Value = r["U_DVVH"].ToString();
                        //Ten don vi van hanh
                        //((EditText)oMtx.Columns.Item("C_1_23").Cells.Item(i + 1).Specific).Value = r["U_TDVVH"].ToString();

                        if (i < ds_import.Tables[1].Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((EditText)oMtx.Columns.Item("C_1_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                }
                oForm.Freeze(false);

                //Fill Table CCM
                oForm.Freeze(true);
                oForm.PaneLevel = 3;
                oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("2_U_G").Specific;
                for (int i = 0; i < ds_import.Tables[2].Rows.Count; i++)
                {
                    oApp.SetStatusBarMessage(string.Format("Import CCM: {0}", (i + 1) + "/" + ds_import.Tables[2].Rows.Count), BoMessageTime.bmt_Short, false);
                    DataRow r = ds_import.Tables[2].Rows[i];
                    if (!string.IsNullOrEmpty(r[0].ToString()))
                    {
                        //SubProject Code
                        ((EditText)oMtx.Columns.Item("C_2_2").Cells.Item(i + 1).Specific).Value = r["U_001"].ToString();
                        //SubProject Name
                        ((EditText)oMtx.Columns.Item("C_2_3").Cells.Item(i + 1).Specific).Value = r["U_TENHM"].ToString();
                        //Loai chi phi
                        ((EditText)oMtx.Columns.Item("C_2_4").Cells.Item(i + 1).Specific).Value = r["U_LOAICHIPHI"].ToString();
                        //BP
                        //((EditText)oMtx.Columns.Item("C_2_4").Cells.Item(i + 1).Specific).Value = r["U_BUSINESSPARTNER"].ToString();
                        //BP Name
                        //((EditText)oMtx.Columns.Item("C_2_5").Cells.Item(i + 1).Specific).Value = r["U_BUSINESSNAME"].ToString();
                        //Don gia du phong
                        ((EditText)oMtx.Columns.Item("C_2_5").Cells.Item(i + 1).Specific).Value = r["U_DGDP"].ToString();
                        //Don gia du phong 2
                        ((EditText)oMtx.Columns.Item("C_2_6").Cells.Item(i + 1).Specific).Value = r["U_DGDP2"].ToString();
                        //Don gia Prelims
                        ((EditText)oMtx.Columns.Item("C_2_7").Cells.Item(i + 1).Specific).Value = r["U_DGPRELIM"].ToString();
                        //Don gia thiet bi
                        ((EditText)oMtx.Columns.Item("C_2_8").Cells.Item(i + 1).Specific).Value = r["U_DGTB"].ToString();
                        //Don gia khac
                        ((EditText)oMtx.Columns.Item("C_2_9").Cells.Item(i + 1).Specific).Value = r["U_DGK"].ToString();
                        //Don gia nha cung cap
                        ((EditText)oMtx.Columns.Item("C_2_10").Cells.Item(i + 1).Specific).Value = r["U_DGNCC"].ToString();
                        //Don gia nha thau phu
                        ((EditText)oMtx.Columns.Item("C_2_11").Cells.Item(i + 1).Specific).Value = r["U_DGNTP"].ToString();
                        //Don gia vat tu phu
                        ((EditText)oMtx.Columns.Item("C_2_12").Cells.Item(i + 1).Specific).Value = r["U_DGVTP"].ToString();
                        //Don gia van chuyen
                        ((EditText)oMtx.Columns.Item("C_2_13").Cells.Item(i + 1).Specific).Value = r["U_DGVC"].ToString();
                        //Don gia cong nhat
                        ((EditText)oMtx.Columns.Item("C_2_14").Cells.Item(i + 1).Specific).Value = r["U_DGCN"].ToString();
                        //Don gia doi thi cong
                        ((EditText)oMtx.Columns.Item("C_2_15").Cells.Item(i + 1).Specific).Value = r["U_DGDTC"].ToString();
                        if (i < ds_import.Tables[2].Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((EditText)oMtx.Columns.Item("C_2_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                }
                oForm.Freeze(false);

                //Fill Table BCH
                oForm.Freeze(true);
                oForm.PaneLevel = 4;
                oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("3_U_G").Specific;
                for (int i = 0; i < ds_import.Tables[3].Rows.Count; i++)
                {
                    oApp.SetStatusBarMessage(string.Format("Import BCH: {0}", (i + 1) + "/" + ds_import.Tables[3].Rows.Count), BoMessageTime.bmt_Short, false);
                    DataRow r = ds_import.Tables[3].Rows[i];
                    if (!string.IsNullOrEmpty(r[0].ToString()))
                    {
                        //Subproject Code
                        ((EditText)oMtx.Columns.Item("C_3_2").Cells.Item(i + 1).Specific).Value = r["U_001"].ToString();
                        //Tai khoan ke toan
                        ((EditText)oMtx.Columns.Item("C_3_3").Cells.Item(i + 1).Specific).Value = r["U_TKKT"].ToString();
                        //Ten tai khoan ke toan
                        ((EditText)oMtx.Columns.Item("C_3_4").Cells.Item(i + 1).Specific).Value = r["U_TTKKT"].ToString();
                        //Gia tri du phong
                        ((EditText)oMtx.Columns.Item("C_3_5").Cells.Item(i + 1).Specific).Value = r["U_GTDP"].ToString();
                        if (i < ds_import.Tables[3].Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((EditText)oMtx.Columns.Item("C_3_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                }
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }

        }

        #region Open File Dialog
        public class GetFileNameClass
        {
            [DllImport("user32.dll")]
            private static extern IntPtr GetForegroundWindow();

            OpenFileDialog _oFileDialog;

            // Properties
            public string FileName
            {
                get { return _oFileDialog.FileName; }
                set { _oFileDialog.FileName = value; }
            }

            public string Filter
            {
                get { return _oFileDialog.Filter; }
                set { _oFileDialog.Filter = value; }
            }

            public string InitialDirectory
            {
                get { return _oFileDialog.InitialDirectory; }
                set { _oFileDialog.InitialDirectory = value; }
            }

            // Constructor
            public GetFileNameClass()
            {
                _oFileDialog = new OpenFileDialog();
            }

            // Methods

            public void GetFileName()
            {
                IntPtr ptr = GetForegroundWindow();
                WindowWrapper oWindow = new WindowWrapper(ptr);
                if (_oFileDialog.ShowDialog(oWindow) != DialogResult.OK)
                {
                    _oFileDialog.FileName = string.Empty;
                }
                oWindow = null;
            } // End of GetFileName
        }
        public class WindowWrapper : System.Windows.Forms.IWin32Window
        {
            private IntPtr _hwnd;

            // Property
            public virtual IntPtr Handle
            {
                get { return _hwnd; }
            }

            // Constructor
            public WindowWrapper(IntPtr handle)
            {
                _hwnd = handle;
            }
        }
        #endregion

        #region Not use
        private void SetApplication()
        {
            SboGuiApi SboGuiApi = null;
            string sCon = ConfigurationManager.AppSettings["sCon"].ToString();
            SboGuiApi = new SboGuiApi();
            if (sCon == "")
            {
                if (Environment.GetCommandLineArgs().Length > 1)
                {
                    MessageBox.Show(Environment.GetCommandLineArgs().GetValue(0) + "|" + Environment.GetCommandLineArgs().GetValue(1));
                    sCon = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(0));

                }
                else
                {
                    sCon = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(0));
                    MessageBox.Show(sCon);
                }
            }
            SboGuiApi.AddonIdentifier = ConfigurationManager.AppSettings["AddonIdentifier"].ToString();
            SboGuiApi.Connect(sCon);
            oApp = SboGuiApi.GetApplication(-1);
        }
        private int Create_Stage_Subproject(int pAbsEntry_SubProject, string pStart, string pClose, string pDescription)
        {
            int rs = -1;
            try
            {
                oCompServ = (CompanyService)oCompany.GetCompanyService();
                pmgService = (ProjectManagementService)oCompServ.GetBusinessService(ServiceTypes.ProjectManagementService);
                PM_SubprojectDocumentParams subprojectpara = (PM_SubprojectDocumentParams)pmgService.GetDataInterface(ProjectManagementServiceDataInterfaces.pmsPM_SubprojectDocumentParams);
                subprojectpara.AbsEntry = pAbsEntry_SubProject;
                PM_SubprojectDocumentData subproject = pmgService.GetSubproject(subprojectpara);
                PMS_StageData sub_stg = subproject.PMS_StagesCollection.Add();
                sub_stg.Description = pDescription;
                if (!string.IsNullOrEmpty(pStart))
                    sub_stg.StartDate = DateTime.ParseExact(pStart, "yyyyMMdd", CultureInfo.InvariantCulture);
                if (!string.IsNullOrEmpty(pClose))
                    sub_stg.CloseDate = DateTime.ParseExact(pClose, "yyyyMMdd", CultureInfo.InvariantCulture);
                pmgService.UpdateSubproject(subproject);
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                if (pmgService != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pmgService);
                }
            }
            return rs;
        }
        private int Get_New_LineID_Stage_SubProject(int pSubProject_AbsEntry)
        {
            int kq = -1;
            Recordset oR_Project_Info = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string str_sql = string.Format("Select Max(LineID) as New_LineID from PHA1 a where a.AbsEntry={0}", pSubProject_AbsEntry);
            oR_Project_Info.DoQuery(str_sql);
            int.TryParse(oR_Project_Info.Fields.Item("New_LineID").Value.ToString(), out kq);
            return kq;
        }
        private int Update_UDF_Stage_SubProject(int pSubProject_AbsEntry, int pLineID, string pItemNo, string pItemName, string pLoaichiphi, string pMaThietBi, string pTenThietBi, string pBusinessPartner, string pBusinessName, string pDVT, double pDinhmuc, double pHaohut, double pDGDAUTHAU, double pDGDUPHONG, double pTTDAUTHAU, double pDGNCC, double pDGNTP, double pDGDTC, double pDGVTP, double pDGVC, double pDGCN, double pDGK, double pDGMUABAN, double pDGTHUE, double pSLDUTRU, double pDGVH, double pDGVCTB, string pDVTTB, double pDGPRELIM, double pDGTB, double pDGDP, double pDGDP2)
        {
            //Update truc tiep UDF vao DB - SAP B1 9.2 PL8 chua ho tro UserFields phan he PM
            int kq = 0;
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["DB"].ConnectionString);
            string str_comand = string.Format("Update PHA1 set U_ITEMNO='{0}',U_ITEMNAME='{1}',U_LOAICHIPHI='{2}',U_MATHIETBI='{3}',U_TENTHIETBI='{4}',U_BUSINESSPARTNER='{5}',U_BUSINESSNAME='{6}',U_DVT='{7}',U_DINHMUC={8},U_HAOHUT={9},U_DGDAUTHAU={10},U_DGDUPHONG={11},U_TTDAUTHAU={12},U_DGNCC={13},U_DGNTP={14},U_DGDTC={15},U_DGVTP={16},U_DGVC={17},U_DGCN={18},U_DGK={19},U_DGMUABAN={20},U_DGTHUE={21},U_SLDUTRU={22},U_DGVH={23},U_DGVCTB={24},U_DVTTB='{25}',U_DGPRELIM={26},U_DGTB={27},U_DGDP={28},U_DGDP2={29} where AbsEntry={30} and LineID={31}",
                pItemNo, pItemName, pLoaichiphi, pMaThietBi, pTenThietBi, pBusinessPartner, pBusinessName, pDVT, pDinhmuc, pHaohut, pDGDAUTHAU, pDGDUPHONG, pTTDAUTHAU, pDGNCC, pDGNTP, pDGDTC, pDGVTP, pDGVC, pDGCN, pDGK, pDGMUABAN, pDGTHUE, pSLDUTRU, pDGVH, pDGVCTB, pDVTTB, pDGPRELIM, pDGTB, pDGDP, pDGDP2, pSubProject_AbsEntry, pLineID);
            SqlCommand cmd = new SqlCommand(str_comand, conn);
            try
            {
                conn.Open();
                kq = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }
            return kq;
        }
        #endregion


    }
}
