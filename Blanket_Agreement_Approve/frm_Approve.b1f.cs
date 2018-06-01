using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;

namespace Blanket_Agreement_Approve
{
    [FormAttribute("Blanket_Agreement_Approve.frm_Approve", "frm_Approve.b1f")]
    class frm_Approve : UserFormBase
    {
        public frm_Approve()
        {
        }
       
        string Blanket_Agreement_No = "";
        string Blanket_Type = "";
        string AbsID = "";
        int LVL_Posting = 0;
        bool Fr_Authorise = false;
        SAPbouiCOM.Form Parent_Form= null;
        Dictionary<int, string> Post_Level;
        SAPbobsCOM.Company oCom;

        public frm_Approve(string p_Blanket_No, string p_Blanket_Type, SAPbouiCOM.Form oForm_Parent)
        {
            Blanket_Agreement_No = p_Blanket_No;
            Blanket_Type = p_Blanket_Type;
            Post_Level = new Dictionary<int, string>();
            oCom = ((SAPbobsCOM.Company)(Application.SBO_Application.Company.GetDICompany()));
            Get_Posting_Level();
            Get_User_Posting_Level();
            LVL_Posting = Check_Next_Level_Posting();
            int USR_Posting = Get_User_Posting_Level();
            if (LVL_Posting == USR_Posting && USR_Posting != -9)
            {
                Fr_Authorise = true;
                Parent_Form = oForm_Parent;
                Load_Data();
                if (LVL_Posting != 4)
                {
                    this.Button2.Item.Visible = false;
                }
            }
            else
            {
                Fr_Authorise = false;
                if (Post_Level.Count > 0)
                    Application.SBO_Application.MessageBox("Don't have rights to authorise LV" + LVL_Posting + " !");
                else
                    Application.SBO_Application.MessageBox("Already Approved or Canceled !");
                this.UIAPIRawForm.Close();
            }
        }
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lb_title").Specific));
            this.StaticText0.Item.FontSize = 16;
            this.StaticText0.Item.Height = 20;
            this.StaticText0.Item.ForeColor = 26879;
            this.StaticText0.Item.TextStyle = 1;
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_pno").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txt_fdcno").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("txt_bano").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("txt_comm").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_ap").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("bt_nap").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("bt_ip").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {
            //Set position for control
            this.GetItem("txt_pno").Left = ((this.GetItem("Item_3").Left + this.GetItem("Item_3").Width) + 5);
            this.GetItem("txt_fdcno").Left = ((this.GetItem("Item_3").Left + this.GetItem("Item_3").Width) + 5);
            this.GetItem("txt_bano").Left = ((this.GetItem("Item_3").Left + this.GetItem("Item_3").Width) + 5);
            this.GetItem("txt_comm").Left = ((this.GetItem("Item_3").Left + this.GetItem("Item_3").Width) + 5);
        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            
        }

        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;

        void Get_Posting_Level()
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string str_query = string.Format("Select U_Apprv1, U_Apprv2, U_Apprv3, U_Apprv4, U_Apprv5"
                                + " from OOAT where BpType = '{0}' and Status ='D' and Cancelled ='N' and Number ='{1}'", Blanket_Type, Blanket_Agreement_No);
            rs.DoQuery(str_query);
            if (rs.RecordCount > 0)
            {
                Post_Level.Add(1, rs.Fields.Item("U_Apprv1").Value.ToString());
                Post_Level.Add(2, rs.Fields.Item("U_Apprv2").Value.ToString());
                Post_Level.Add(3, rs.Fields.Item("U_Apprv3").Value.ToString());
                Post_Level.Add(4, rs.Fields.Item("U_Apprv4").Value.ToString());
                Post_Level.Add(5, rs.Fields.Item("U_Apprv5").Value.ToString());
            }
        }

        int Get_User_Posting_Level()
        {
            int result = -9;
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(string.Format("Select U_PostingLevel from OUSR where USER_CODE ='{0}'",oCom.UserName));
            int.TryParse(rs.Fields.Item("U_PostingLevel").Value.ToString(), out result);
            return result;
        }

        int Check_Next_Level_Posting()
        {
            int result = -9;
            if (Post_Level.Count > 0)
            {
                for (int i = 1; i <= Post_Level.Count; i++)
                {
                    if (i == 4 && Post_Level[i] == "2")
                    {
                        result = i;
                        break;
                    }
                    if (string.IsNullOrEmpty(Post_Level[i]))
                    {

                        result = i;
                        break;
                    }
                }
            }
            return result;
        }

        void Load_Data()
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(string.Format("Select AbsID from OOAT where Number ='{0}' and BpType = '{1}' and Status ='D' and Cancelled ='N'",Blanket_Agreement_No,Blanket_Type));
            AbsID = rs.Fields.Item("AbsID").Value.ToString();
            SAPbobsCOM.CompanyService oCompSer = oCom.GetCompanyService();
            SAPbobsCOM.BlanketAgreementsService oBAService = (SAPbobsCOM.BlanketAgreementsService)oCompSer.GetBusinessService(SAPbobsCOM.ServiceTypes.BlanketAgreementsService);
            SAPbobsCOM.BlanketAgreementParams oParams = (SAPbobsCOM.BlanketAgreementParams)oBAService.GetDataInterface(SAPbobsCOM.BlanketAgreementsServiceDataInterfaces.basBlanketAgreementParams);
            oParams.AgreementNo = int.Parse(AbsID);
            SAPbobsCOM.BlanketAgreement oBA = oBAService.GetBlanketAgreement(oParams);
            EditText1.Value = oBA.UserFields.Item("U_PRJ").Value.ToString();
            EditText2.Value = oBA.DocNum.ToString();
        }

        bool Check_Parent_Form_Closed(string formuid)
        {
            try
            {
                Application.SBO_Application.Forms.Item(formuid);
                return false;
            }
            catch
            {
                return true;
            }
        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //Approve
            if (Fr_Authorise)
            {
                SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oR_RecordSet.DoQuery("Select * from [@ADDONCFG]");
                string uid = oR_RecordSet.Fields.Item("Code").Value.ToString();
                string pwd = oR_RecordSet.Fields.Item("Name").Value.ToString();
                SqlConnection conn = new SqlConnection(string.Format("Data Source={0}; Initial Catalog={1}; User id={2}; Password={3};", oCom.Server, oCom.CompanyDB, uid, pwd));
                SqlCommand cmd = null;
                try
                {
                    cmd = new SqlCommand("Update_Blanket_Post_Level", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@AbsID", AbsID);
                    cmd.Parameters.AddWithValue("@Blanket_Type", Blanket_Type);
                    cmd.Parameters.AddWithValue("@Blanket_Level", LVL_Posting);
                    cmd.Parameters.AddWithValue("@Usr", oCom.UserName);
                    cmd.Parameters.AddWithValue("@Approve", 1);
                    cmd.Parameters.AddWithValue("@Usr_Comment", this.EditText3.Value.Trim());
                    conn.Open();
                    int row_count = cmd.ExecuteNonQuery();
                    if (row_count == 0)
                        Application.SBO_Application.StatusBar.SetText("Approve Failed !");
                    else
                    {
                        //Approve after level 5
                        if (LVL_Posting == 5)
                        {
                            SAPbobsCOM.CompanyService oCompSer = oCom.GetCompanyService();
                            SAPbobsCOM.BlanketAgreementsService oBAService = (SAPbobsCOM.BlanketAgreementsService)oCompSer.GetBusinessService(SAPbobsCOM.ServiceTypes.BlanketAgreementsService);
                            SAPbobsCOM.BlanketAgreementParams oParams = (SAPbobsCOM.BlanketAgreementParams)oBAService.GetDataInterface(SAPbobsCOM.BlanketAgreementsServiceDataInterfaces.basBlanketAgreementParams);
                            oParams.AgreementNo = int.Parse(AbsID);
                            SAPbobsCOM.BlanketAgreement oBA = oBAService.GetBlanketAgreement(oParams);
                            oBA.Status = SAPbobsCOM.BlanketAgreementStatusEnum.asApproved;
                            oBAService.UpdateBlanketAgreement(oBA);

                        }
                    }
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.MessageBox("Can't approve - Error: " + ex.Message);
                }
                finally
                {
                    conn.Close();
                    cmd.Dispose();
                    if (!Check_Parent_Form_Closed(Parent_Form.UniqueID))
                    {
                        Parent_Form.Select();
                        Application.SBO_Application.ActivateMenuItem("1304");
                    }
                    this.UIAPIRawForm.Close();
                }

            }
        }

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //Not Approve
            if (Fr_Authorise)
            {
                SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oR_RecordSet.DoQuery("Select * from [@ADDONCFG]");
                string uid = oR_RecordSet.Fields.Item("Code").Value.ToString();
                string pwd = oR_RecordSet.Fields.Item("Name").Value.ToString();
                SqlConnection conn = new SqlConnection(string.Format("Data Source={0}; Initial Catalog={1}; User id={2}; Password={3};", oCom.Server, oCom.CompanyDB, uid, pwd));
                SqlCommand cmd = null;
                try
                {
                    cmd = new SqlCommand("Update_Blanket_Post_Level", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@AbsID", AbsID);
                    cmd.Parameters.AddWithValue("@Blanket_Type", Blanket_Type);
                    cmd.Parameters.AddWithValue("@Blanket_Level", LVL_Posting);
                    cmd.Parameters.AddWithValue("@Usr", oCom.UserName);
                    cmd.Parameters.AddWithValue("@Approve", 0);
                    cmd.Parameters.AddWithValue("@Usr_Comment", this.EditText3.Value.Trim());
                    conn.Open();
                    int row_count = cmd.ExecuteNonQuery();
                    if (row_count == 0)
                        Application.SBO_Application.StatusBar.SetText("Approve Failed !");
                    else
                    {
                        //Cancel Blanket Agreement
                        if (LVL_Posting == 1 || LVL_Posting == 4 || LVL_Posting == 5)
                        {
                            SAPbobsCOM.CompanyService oCompSer = oCom.GetCompanyService();
                            SAPbobsCOM.BlanketAgreementsService oBAService = (SAPbobsCOM.BlanketAgreementsService)oCompSer.GetBusinessService(SAPbobsCOM.ServiceTypes.BlanketAgreementsService);
                            SAPbobsCOM.BlanketAgreementParams oParams = (SAPbobsCOM.BlanketAgreementParams)oBAService.GetDataInterface(SAPbobsCOM.BlanketAgreementsServiceDataInterfaces.basBlanketAgreementParams);
                            oParams.AgreementNo = int.Parse(AbsID);
                            oBAService.CancelBlanketAgreement(oParams);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.MessageBox("Can't approve - Error: " + ex.Message);
                }
                finally
                {
                    conn.Close();
                    cmd.Dispose();
                    Parent_Form.Select();
                    Application.SBO_Application.ActivateMenuItem("1304");
                    this.UIAPIRawForm.Close();
                }
            }
        }

        private void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //In Process
            if (Fr_Authorise)
            {
                SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oR_RecordSet.DoQuery("Select * from [@ADDONCFG]");
                string uid = oR_RecordSet.Fields.Item("Code").Value.ToString();
                string pwd = oR_RecordSet.Fields.Item("Name").Value.ToString();
                SqlConnection conn = new SqlConnection(string.Format("Data Source={0}; Initial Catalog={1}; User id={2}; Password={3};", oCom.Server, oCom.CompanyDB, uid, pwd));
                SqlCommand cmd = null;
                try
                {
                    cmd = new SqlCommand("Update_Blanket_Post_Level", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@AbsID", AbsID);
                    cmd.Parameters.AddWithValue("@Blanket_Type", Blanket_Type);
                    cmd.Parameters.AddWithValue("@Blanket_Level", LVL_Posting);
                    cmd.Parameters.AddWithValue("@Usr", oCom.UserName);
                    cmd.Parameters.AddWithValue("@Approve", 2);
                    cmd.Parameters.AddWithValue("@Usr_Comment", this.EditText3.Value.Trim());
                    conn.Open();
                    int row_count = cmd.ExecuteNonQuery();
                    if (row_count == 0)
                        Application.SBO_Application.StatusBar.SetText("Approve Failed !");
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.MessageBox("Can't approve - Error: " + ex.Message);
                }
                finally
                {
                    conn.Close();
                    cmd.Dispose();
                    Parent_Form.Select();
                    Application.SBO_Application.ActivateMenuItem("1304");
                    this.UIAPIRawForm.Close();
                }
            }
        }
    }
}