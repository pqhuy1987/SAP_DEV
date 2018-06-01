using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using System.Windows.Forms;
using SAPbouiCOM;
using System.Configuration;
using System.Xml.Serialization;
using System.IO;
using System.Data.SqlClient;
namespace ProjectManagement_Function_AddOn
{
    [FormAttribute("ProjectManagement_Function_AddOn.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>

        CompanyService oCompServ = null;
        ProjectManagementService pmgService = null;
        SAPbobsCOM.Company oCompany = null;
        SAPbouiCOM.Application oApp = null;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.EditText EditText5;
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lb_pcode").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_pcode").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_info").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lb_ptype").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txt_ptype").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lb_pname").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("txt_pname").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("bt_clone").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("lb_bpname").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("txt_bpname").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("lb_pgrp").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("txt_pgroup").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("t_absentry").Specific));
            this.OnCustomInitialize();

        }
        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.SetApplication();
            this.oCompany = ((SAPbobsCOM.Company)(this.oApp.Company.GetDICompany()));
            //this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);
        }
        private void OnCustomInitialize()
        {
        }
        private void SetApplication()
        {
            //SboGuiApi SboGuiApi = null;
            //string sCon = ConfigurationManager.AppSettings["sCon"].ToString();
            //SboGuiApi = new SboGuiApi();
            //if (sCon == "")
            //{
            //    if (Environment.GetCommandLineArgs().Length > 1)
            //    {
            //        MessageBox.Show(Environment.GetCommandLineArgs().GetValue(0) + "|" + Environment.GetCommandLineArgs().GetValue(1));
            //        sCon = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(0));

            //    }
            //    else
            //    {
            //        sCon = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(0));
            //        MessageBox.Show(sCon);
            //    }
            //}
            //SboGuiApi.AddonIdentifier = ConfigurationManager.AppSettings["AddonIdentifier"].ToString();
            //SboGuiApi.Connect(sCon);
            //oApp = SboGuiApi.GetApplication(-1);
            oApp = SAPbouiCOM.Framework.Application.SBO_Application;
        }
        
        private int Get_Depth_Level_Subproject(int pProjectID)
        {
            int kq = 0;
            //Find Depth Level of Subproject in Project
            Recordset oR_Subproject_LV = null;
            oR_Subproject_LV = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oR_Subproject_LV.DoQuery("Select ISNULL(MAX(Level),0) as Depth_Level from OPHA where ProjectID=" + pProjectID.ToString());
            kq = (int)oR_Subproject_LV.Fields.Item("Depth_Level").Value;
            return kq;
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
        
        private int Get_Next_DocumentNo(int pSeries)
        {
            int kq = -1;
            string str_sql = string.Format("Select NextNumber from NNM1 where Series={0}", pSeries);
            Recordset oR_Next_DocNo = null;
            oR_Next_DocNo = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oR_Next_DocNo.DoQuery(str_sql);
            while (!oR_Next_DocNo.EoF)
            {
                kq = (int)oR_Next_DocNo.Fields.Item("NextNumber").Value;
                oR_Next_DocNo.MoveNext();
            }
            return kq;
        }

        private string Copy_Project(int pProjectNo)
        {
            string New_ProjectNo = "-1";
            //Get Project Info through DI API
            Recordset oR_Project_Info = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string str_sql = string.Format("Select * from OPMG where AbsEntry={0}", pProjectNo.ToString());
            oR_Project_Info.DoQuery(str_sql);
            if (oR_Project_Info.RecordCount == 1)
            {
                //Use UI API to create Project
                oApp.ActivateMenuItem("48897");
                var oform1 = oApp.Forms.ActiveForm;
                //Project Type
                string ptype = oR_Project_Info.Fields.Item("TYP").Value.ToString();
                if (ptype == "E")
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
                string bpcode = oR_Project_Info.Fields.Item("CARDCODE").Value.ToString();
                if (!String.IsNullOrEmpty(bpcode))
                {
                    var txt_bpcode = oform1.Items.Item("234000013").Specific;
                    ((SAPbouiCOM.EditText)txt_bpcode).Value = bpcode;
                }
                //BPName
                string bpname = oR_Project_Info.Fields.Item("CARDNAME").Value.ToString();
                if (!String.IsNullOrEmpty(bpname))
                {
                    var txt_bpname = oform1.Items.Item("234000016").Specific;
                    ((SAPbouiCOM.EditText)txt_bpname).Value = bpname;
                }
                //Contact Person
                string contact = oR_Project_Info.Fields.Item("CONTACT").Value.ToString();
                if (!String.IsNullOrEmpty(contact))
                {
                    var cbo_contactperson = oform1.Items.Item("234000018").Specific;
                    ((SAPbouiCOM.ComboBox)cbo_contactperson).Select(contact);
                }
                //Territory
                string territory = oR_Project_Info.Fields.Item("TERRITORY").Value.ToString();
                if (!String.IsNullOrEmpty(territory) && territory != "0")
                {
                    var txt_territory = oform1.Items.Item("234000020").Specific;
                    ((SAPbouiCOM.EditText)txt_territory).Value = territory;
                }
                //Sale Employee
                string employee = oR_Project_Info.Fields.Item("EMPLOYEE").Value.ToString();
                if (!String.IsNullOrEmpty(employee) && employee != "-1")
                {
                    var cbo_employee = oform1.Items.Item("234000023").Specific;
                    ((SAPbouiCOM.ComboBox)cbo_employee).Select(employee);
                }
                //Owner
                string owner = oR_Project_Info.Fields.Item("OWNER").Value.ToString();
                if (!String.IsNullOrEmpty(owner) && owner != "0")
                {
                    var txt_owner = oform1.Items.Item("234000025").Specific;
                    ((SAPbouiCOM.EditText)txt_owner).Value = owner;
                }
                //Bo phan thuc hien
                string u_bpth = oR_Project_Info.Fields.Item("U_BPTH").Value.ToString();
                if (!String.IsNullOrEmpty(u_bpth))
                {
                    var cbo_u_bpth = oform1.Items.Item("U_BPTH").Specific;
                    ((SAPbouiCOM.ComboBox)cbo_u_bpth).Select(u_bpth);
                }
                //Project with Subproject
                string withpahse = oR_Project_Info.Fields.Item("WithPhases").Value.ToString();
                if (withpahse == "Y")
                {
                    var chb_withpahse = oform1.Items.Item("234000027").Specific;
                    ((SAPbouiCOM.CheckBox)chb_withpahse).Checked = true;
                }
                //Project Name
                string pname = oR_Project_Info.Fields.Item("NAME").Value.ToString();
                if (!String.IsNullOrEmpty(pname))
                {
                    var txt_pname = oform1.Items.Item("234000029").Specific;
                    ((SAPbouiCOM.EditText)txt_pname).Value = pname + "_Clone_" + DateTime.Now.ToString("ddMMyyyy HHmmss");
                }
                //Series
                string series = oR_Project_Info.Fields.Item("Series").Value.ToString();
                if (!String.IsNullOrEmpty(series))
                {
                    var cbo_series = oform1.Items.Item("234000031").Specific;
                    ((SAPbouiCOM.ComboBox)cbo_series).Select(series);
                }
                //Status
                string status = oR_Project_Info.Fields.Item("STATUS").Value.ToString();
                if (!String.IsNullOrEmpty(status))
                {
                    var cbo_status = oform1.Items.Item("234000034").Specific;
                    ((SAPbouiCOM.ComboBox)cbo_status).Select(status);
                }
                //Start Date
                string strdate = oR_Project_Info.Fields.Item("START").Value.ToString();
                if (!String.IsNullOrEmpty(strdate))
                {
                    if (DateTime.Parse(strdate).Year > 1900)
                    {
                        var txt_strdate = oform1.Items.Item("234000036").Specific;
                        ((SAPbouiCOM.EditText)txt_strdate).Value = DateTime.Parse(strdate).ToString("yyyyMMdd");
                    }
                }
                //Due Date
                string duedate = oR_Project_Info.Fields.Item("DUEDATE").Value.ToString();
                if (!String.IsNullOrEmpty(duedate))
                {
                    if (DateTime.Parse(duedate).Year > 1900)
                    {
                        var txt_duedate = oform1.Items.Item("234000038").Specific;
                        ((SAPbouiCOM.EditText)txt_duedate).Value = DateTime.Parse(duedate).ToString("yyyyMMdd");
                    }
                }
                //Closing Date
                string closedate = oR_Project_Info.Fields.Item("CLOSING").Value.ToString();
                if (!String.IsNullOrEmpty(closedate))
                {
                    if (DateTime.Parse(closedate).Year > 1900)
                    {
                        var txt_closedate = oform1.Items.Item("234000040").Specific;
                        ((SAPbouiCOM.EditText)txt_closedate).Value = DateTime.Parse(closedate).ToString("yyyyMMdd");
                    }
                }
                //Financial Project
                string finan_project = oR_Project_Info.Fields.Item("FIPROJECT").Value.ToString();
                if (!String.IsNullOrEmpty(finan_project))
                {
                    //var txt_finan_project = oform1.Items.Item("234000049").Specific;
                    //((SAPbouiCOM.EditText)txt_finan_project).Value = finan_project;
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
                string freetext = oR_Project_Info.Fields.Item("Free_Text").Value.ToString();
                if (!String.IsNullOrEmpty(freetext))
                {
                    oform1.Freeze(true);
                    oform1.PaneLevel = oform1.Items.Item("234000167").FromPane;//oForm.Items.Item("148").FromPane
                    var txt_freetext = oform1.Items.Item("234000167").Specific;
                    ((SAPbouiCOM.EditText)txt_freetext).Value = freetext;
                    oform1.Freeze(false);
                }
                //U Project Type
                string u_ptype = oR_Project_Info.Fields.Item("U_PRJTYPE").Value.ToString();
                if (!String.IsNullOrEmpty(u_ptype))
                {
                    var cbo_u_ptype = oform1.Items.Item("U_PRJTYPE").Specific;
                    ((SAPbouiCOM.ComboBox)cbo_u_ptype).Select(u_ptype);
                }
                //U Project Group
                string u_pgrp = oR_Project_Info.Fields.Item("U_PRJGROUP").Value.ToString();
                if (!String.IsNullOrEmpty(u_pgrp))
                {
                    var cbo_u_ptype = oform1.Items.Item("U_PRJGROUP").Specific;
                    ((SAPbouiCOM.ComboBox)cbo_u_ptype).Select(u_pgrp);
                }
                //Overview Panel - U_CPHT1
                string u_cpht1 = oR_Project_Info.Fields.Item("U_CPHT1").Value.ToString();
                if (!String.IsNullOrEmpty(u_cpht1))
                {
                    oform1.Freeze(true);
                    oform1.PaneLevel = oform1.Items.Item("U_CPHT1").FromPane;
                    var txt_u_cpht1 = oform1.Items.Item("U_CPHT1").Specific;
                    ((SAPbouiCOM.EditText)txt_u_cpht1).Value = u_cpht1;
                    oform1.Freeze(false);
                }
                //Overview Panel - U_CPHT2
                string u_cpht2 = oR_Project_Info.Fields.Item("U_CPHT2").Value.ToString();
                if (!String.IsNullOrEmpty(u_cpht2))
                {
                    oform1.Freeze(true);
                    oform1.PaneLevel = oform1.Items.Item("U_CPHT2").FromPane;
                    var txt_u_cpht2 = oform1.Items.Item("U_CPHT2").Specific;
                    ((SAPbouiCOM.EditText)txt_u_cpht2).Value = u_cpht2;
                    oform1.Freeze(false);
                }
                //Overview Panel - U_DPBH
                string u_dpbh = oR_Project_Info.Fields.Item("U_DPBH").Value.ToString();
                if (!String.IsNullOrEmpty(u_dpbh))
                {
                    oform1.Freeze(true);
                    oform1.PaneLevel = oform1.Items.Item("U_DPBH").FromPane;
                    var txt_u_dpbh = oform1.Items.Item("U_DPBH").Specific;
                    ((SAPbouiCOM.EditText)txt_u_dpbh).Value = u_dpbh;
                    oform1.Freeze(false);
                }
                //Overview Panel - U_DPCP
                string u_dpcp = oR_Project_Info.Fields.Item("U_DPCP").Value.ToString();
                if (!String.IsNullOrEmpty(u_dpcp))
                {
                    oform1.Freeze(true);
                    oform1.PaneLevel = oform1.Items.Item("U_DPCP").FromPane;
                    var txt_u_dpcp = oform1.Items.Item("U_DPCP").Specific;
                    ((SAPbouiCOM.EditText)txt_u_dpcp).Value = u_dpcp;
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
                    oform1.Close();
                }
                else
                {
                    New_ProjectNo = "-1";
                    oApp.MessageBox("Can't copy project !");
                }
            }
            else
            {
                oApp.MessageBox("Không tìm thấy ProjectNo hoặc Trùng ProjectNo");
            }
            return New_ProjectNo;
        }

        private bool Copy_SubProject(int pProjectNo, int pProjectNo_Clone)
        {
            oCompServ = (CompanyService)oCompany.GetCompanyService();
            pmgService = (ProjectManagementService)oCompServ.GetBusinessService(ServiceTypes.ProjectManagementService);
            PM_ProjectDocumentParams projectParam = (PM_ProjectDocumentParams)pmgService.GetDataInterface(ProjectManagementServiceDataInterfaces.pmsPM_ProjectDocumentParams);
            projectParam.AbsEntry = pProjectNo;
            PM_ProjectDocumentData project = pmgService.GetProject(projectParam);
            try
            {
                List<int> subproject_tmp = new List<int>();
                List<int> subproject_clone = new List<int>();
                Dictionary<int, int> mapping_sub = new Dictionary<int, int>();
                int depth_lv_max = Get_Depth_Level_Subproject(pProjectNo);
                for (int depth_lv = 0; depth_lv <= depth_lv_max; depth_lv++)
                {
                    if (depth_lv == 0)
                    {
                        subproject_tmp = Get_AbsEntry_Subproject_List(pProjectNo);
                        foreach (int i in subproject_tmp)
                        {
                            PM_SubprojectDocumentParams tmp_subprojectpara = (PM_SubprojectDocumentParams)pmgService.GetDataInterface(ProjectManagementServiceDataInterfaces.pmsPM_SubprojectDocumentParams);
                            tmp_subprojectpara.AbsEntry = i;
                            PM_SubprojectDocumentData tmp_subproject = pmgService.GetSubproject(tmp_subprojectpara);
                            tmp_subproject.ProjectID = pProjectNo_Clone;
                            pmgService.AddSubproject(tmp_subproject);
                        }
                        subproject_clone = Get_AbsEntry_Subproject_List(pProjectNo_Clone);
                        for (int t = 0; t < subproject_tmp.Count; t++)
                        {
                            mapping_sub.Add(subproject_tmp.ToArray()[t], subproject_clone.ToArray()[t]);
                        }
                    }
                    else if (depth_lv > 0)
                    {
                        subproject_tmp = Get_AbsEntry_Subproject_List(pProjectNo, 0, depth_lv);
                        foreach (int i in subproject_tmp)
                        {
                            PM_SubprojectDocumentParams tmp_subprojectpara = (PM_SubprojectDocumentParams)pmgService.GetDataInterface(ProjectManagementServiceDataInterfaces.pmsPM_SubprojectDocumentParams);
                            tmp_subprojectpara.AbsEntry = i;
                            PM_SubprojectDocumentData tmp_subproject = pmgService.GetSubproject(tmp_subprojectpara);
                            tmp_subproject.ProjectID = pProjectNo_Clone;
                            tmp_subproject.ParentID = mapping_sub[tmp_subproject.ParentID];
                            pmgService.AddSubproject(tmp_subproject);
                        }
                        subproject_clone = Get_AbsEntry_Subproject_List(pProjectNo_Clone, 0, depth_lv);
                        for (int t = 0; t < subproject_tmp.Count; t++)
                        {
                            mapping_sub.Add(subproject_tmp.ToArray()[t], subproject_clone.ToArray()[t]);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
                return false;
            }
            return true;
        }

        private bool Copy_Stage(int pProjectNo, int pProjectNo_Clone)
        {
            //XmlSerializer serializer = new XmlSerializer(typeof(PM_StageData));
            oCompServ = (CompanyService)oCompany.GetCompanyService();
            pmgService = (ProjectManagementService)oCompServ.GetBusinessService(ServiceTypes.ProjectManagementService);
            PM_ProjectDocumentParams projectParam = (PM_ProjectDocumentParams)pmgService.GetDataInterface(ProjectManagementServiceDataInterfaces.pmsPM_ProjectDocumentParams);
            projectParam.AbsEntry = pProjectNo;
            PM_ProjectDocumentData project = pmgService.GetProject(projectParam);

            PM_ProjectDocumentParams projectCloneParam = (PM_ProjectDocumentParams)pmgService.GetDataInterface(ProjectManagementServiceDataInterfaces.pmsPM_ProjectDocumentParams);
            projectCloneParam.AbsEntry = pProjectNo_Clone;
            PM_ProjectDocumentData projectclone = pmgService.GetProject(projectCloneParam);

            foreach (PM_StageData sta_tmp in project.PM_StagesCollection)
            {
                PM_StageData tmp = projectclone.PM_StagesCollection.Add();
                tmp.StartDate = sta_tmp.StartDate;
                tmp.Description = sta_tmp.Description;
                pmgService.UpdateProject(projectclone);
            }
            //Khong con cach nao khac - Can thiep tho bao vao DB voi cac truong UDF cua Stage
            //SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["DB"].ConnectionString);
            //string str_comand = "Update A Set";
            //string UDF = ConfigurationManager.AppSettings["UDF_Project_Stage"].ToString();
            //if (!string.IsNullOrEmpty(UDF))
            //{
            //    foreach (string field in UDF.Split(','))
            //    {
            //        str_comand = str_comand + string.Format(" a.{0} = b.{1} ,", field, field);
            //    }
            //    str_comand = str_comand.Substring(0, str_comand.Length - 1);
            //    str_comand += string.Format("From PMG1 as a inner join (Select * from PMG1 where  AbsEntry = {0}) as b on a.LineID = b.LineID where a.AbsEntry = {1}", pProjectNo, pProjectNo_Clone);
            //    SqlCommand cmd = new SqlCommand(str_comand, conn);
            //    try
            //    {
            //        conn.Open();
            //        cmd.ExecuteNonQuery();
            //    }
            //    catch (Exception ex)
            //    {
            //        return false;
            //    }
            //    finally
            //    {
            //        conn.Close();
            //        cmd.Dispose();
            //    }
                
            //}
            return true;
        }

        private int GetAbsEntry(string pDocNum)
        {
            int result = -1;
            Recordset oR_Project_Info = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string str_sql = string.Format("Select AbsEntry from OPMG where DocNum={0}", pDocNum);
            oR_Project_Info.DoQuery(str_sql);
            int.TryParse(oR_Project_Info.Fields.Item("AbsEntry").Value.ToString(), out result);
            return result;
        }

        private Recordset Get_Project_Info(string pDocNum)
        {
            Recordset oR_Project_Info = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string str_sql = string.Format("Select a.*,(Select Descr from UFD1 where TableID='OPMG' and FieldID =1 and FldValue=a.U_BPTH) as U_BPTH_Des "
                + ",(Select Descr from UFD1 where TableID='OPMG' and FieldID =2 and FldValue=a.U_PRJGROUP) as U_PRJGROUP_Des "
                + ",(Select Descr from UFD1 where TableID='OPMG' and FieldID =3 and FldValue=a.U_PRJTYPE) as U_PRJTYPE_Des "
                + "from OPMG a where a.DocNum={0}", pDocNum);
            oR_Project_Info.DoQuery(str_sql);
            return oR_Project_Info;
        }

        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                Recordset oR_Project = Get_Project_Info(EditText0.Value);
                EditText1.Value = oR_Project.Fields.Item("U_PRJTYPE_Des").Value.ToString();
                EditText2.Value = oR_Project.Fields.Item("NAME").Value.ToString();
                EditText5.Value = oR_Project.Fields.Item("AbsEntry").Value.ToString();
                EditText3.Value = oR_Project.Fields.Item("CARDNAME").Value.ToString();
                EditText4.Value = oR_Project.Fields.Item("U_PRJGROUP_Des").Value.ToString();
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
        }

        private void Button1_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (!string.IsNullOrEmpty(EditText5.Value))
                {
                    int absentry = int.Parse(EditText5.Value.ToString());
                    string DocNum_Clone = Copy_Project(absentry);
                    if (DocNum_Clone != "-1")
                    {
                        int absentry_clone = GetAbsEntry(DocNum_Clone);
                        if (absentry_clone != -1)
                        {
                            Copy_SubProject(absentry, absentry_clone);
                            Copy_Stage(absentry, absentry_clone);
                            oApp.MessageBox("Copy successful with ProjectNo " + DocNum_Clone);
                        }
                    }
                }
                else
                {
                    oApp.MessageBox("ProjectNo can't be empty !");
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

        }
    }
}