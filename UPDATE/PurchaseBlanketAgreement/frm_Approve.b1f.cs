using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data;
using System.Data.SqlClient;
using System.Net.Mail;

namespace PurchaseBlanketAgreement
{
    [FormAttribute("PurchaseBlanketAgreement.frm_Approve", "frm_Approve.b1f")]
    class frm_Approve : UserFormBase
    {
        string Blanket_Agreement_No = "";
        string Blanket_Type = "";
        string AbsID = "";
        string User_Created = "";
        string FProject = "";
        int LVL_Posting = 0;
        int Usr_LVL = 0;
        bool Fr_Authorise = false;
        SAPbouiCOM.Form Parent_Form = null;
        Dictionary<int, string> Post_Level;
        int deptcreate = -100;
        int position = -100;
        int dept = -100;
        int lvl1 = 0;
        string CGroup = "";
        string E_BpName = "";
        //Email Server Config
        string Email_From = "";
        string Email_From_Name = "";
        string Host_Address = "";
        int Host_Port = 25;
        bool EnableSSL = false;
        string Uid = "";
        string Pwd = "";
        //End Email Config
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCom = null;
        SqlConnection conn = null;
        public frm_Approve(string p_Blanket_No, string p_Blanket_Type, SAPbouiCOM.Form oForm_Parent)
        {            
            Blanket_Agreement_No = p_Blanket_No;
            Blanket_Type = p_Blanket_Type;
            Post_Level = new Dictionary<int, string>();
            //oCompany = ((SAPbobsCOM.Company)(Application.SBO_Application.Company.GetDICompany()));
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(string.Format("SELECT b.position,b.dept FROM OUSR a LEFT JOIN OHEM b ON a.USERID = b.userId WHERE a.USER_CODE ='{0}'", oCom.UserName));
            int.TryParse(rs.Fields.Item("position").Value.ToString(), out position);
            int.TryParse(rs.Fields.Item("dept").Value.ToString(), out dept);

            rs.DoQuery(string.Format("SELECT a.U_CGroup,b.dept,c.USER_CODE From OOAT a LEFT JOIN OHEM b ON a.UserSign = b.userId inner join OUSR c on b.userId=c.USERID WHERE a.Number = '{0}'", Blanket_Agreement_No));
            CGroup = rs.Fields.Item("U_CGroup").Value.ToString();
            User_Created = rs.Fields.Item("USER_CODE").Value.ToString();
            int.TryParse(rs.Fields.Item("dept").Value.ToString(), out deptcreate);

            lvl1 = Get_Level_1();
            Get_Posting_Level();
            Usr_LVL = Get_User_Posting_Level();
            LVL_Posting = Check_Current_Level();
            int USR_Posting = Get_User_Level();
            if (LVL_Posting == USR_Posting && USR_Posting != -9)
            {
                Fr_Authorise = true;
                Parent_Form = oForm_Parent;
                Load_Data();
                if (LVL_Posting != 4 || position != 1)
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
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_6").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_7").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_10").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter); 
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_9").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);             
            this.StaticText0.Item.FontSize = 16;
            this.StaticText0.Item.Height = 20;
            this.StaticText0.Item.ForeColor = 26879;
            this.StaticText0.Item.TextStyle = 1;    
            this.OnCustomInitialize();
        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.CloseBefore += new CloseBeforeHandler(this.Form_CloseBefore);

        }

        private void OnCustomInitialize()
        {
            this.oApp = (SAPbouiCOM.Application)SAPbouiCOM.Framework.Application.SBO_Application;
            this.oCom = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            //Create Connection SQL
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery("Select * from [@ADDONCFG]");
            if (oR_RecordSet.RecordCount > 0)
            {
                string uid = oR_RecordSet.Fields.Item("Code").Value.ToString();
                string pwd = oR_RecordSet.Fields.Item("Name").Value.ToString();
                conn = new SqlConnection(string.Format("Data Source={0}; Initial Catalog={1}; User id={2}; Password={3};", oCom.Server, oCom.CompanyDB, uid, pwd));
            }
            else
            {
                oApp.MessageBox("Can't connect DB !");
            }
            DataTable rs = Get_Email_Conf();
            if (rs.Rows.Count == 1)
            {
                Email_From = rs.Rows[0]["Email_From"].ToString();
                Email_From_Name = rs.Rows[0]["Email_From_Name"].ToString();
                Host_Address = rs.Rows[0]["Host_Address"].ToString();
                int.TryParse(rs.Rows[0]["Host_Port"].ToString(), out Host_Port);
                EnableSSL = false;
                if (rs.Rows[0]["EnableSSL"].ToString() == "0")
                    EnableSSL = false;
                else
                    EnableSSL = true;
                Uid = rs.Rows[0]["User"].ToString();
                Pwd = rs.Rows[0]["Pwd"].ToString();
            }
        }

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;

        void Get_Posting_Level()
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string str_query = string.Format("Select U_Apprv1, U_Apprv2, U_Apprv3, U_Apprv4, U_Apprv5, U_Apprv6, U_Apprv7, U_Apprv8, U_Apprv9, U_Apprv10, U_Apprv11 "
                                + " from OOAT where BpType = '{0}' and Status ='D' and Cancelled ='N' and Number ='{1}'", Blanket_Type, Blanket_Agreement_No);
            rs.DoQuery(str_query);
            if (rs.RecordCount > 0)
            {
                Post_Level.Add(1, rs.Fields.Item("U_Apprv1").Value.ToString());
                Post_Level.Add(2, rs.Fields.Item("U_Apprv2").Value.ToString());
                Post_Level.Add(3, rs.Fields.Item("U_Apprv3").Value.ToString());
                Post_Level.Add(4, rs.Fields.Item("U_Apprv4").Value.ToString());
                Post_Level.Add(5, rs.Fields.Item("U_Apprv5").Value.ToString());
                Post_Level.Add(6, rs.Fields.Item("U_Apprv6").Value.ToString());
                Post_Level.Add(7, rs.Fields.Item("U_Apprv7").Value.ToString());
                Post_Level.Add(8, rs.Fields.Item("U_Apprv8").Value.ToString());
                Post_Level.Add(9, rs.Fields.Item("U_Apprv9").Value.ToString());
                Post_Level.Add(10, rs.Fields.Item("U_Apprv10").Value.ToString());
                Post_Level.Add(11, rs.Fields.Item("U_Apprv11").Value.ToString());
            }
        }

        int Get_Level_1()
        {
            SqlCommand cmd = null;
            int result = -9;
            try
            {

                cmd = new SqlCommand("BL_Get_Level_1", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@BlanketNo", int.Parse(Blanket_Agreement_No));

                SqlParameter returnPara = cmd.Parameters.Add("@ReturnVal", SqlDbType.Int);
                returnPara.Direction = ParameterDirection.ReturnValue;

                conn.Open();
                cmd.ExecuteNonQuery();
                result = (int)returnPara.Value;
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
            return result;
        }

        int Get_User_Posting_Level()
        {
            SqlCommand cmd = null;
            int result = -9;
            try
            {

                cmd = new SqlCommand("BL_Get_User_Posting_Level", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCom.UserName);
                cmd.Parameters.AddWithValue("@BlanketNo", int.Parse(Blanket_Agreement_No));

                SqlParameter returnPara = cmd.Parameters.Add("@ReturnVal", SqlDbType.Int);
                returnPara.Direction = ParameterDirection.ReturnValue;

                conn.Open();
                cmd.ExecuteNonQuery();
                result = (int)returnPara.Value;
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

        int Get_User_Level()
        {
            SqlCommand cmd = null;
            int result = -9;
            try
            {

                cmd = new SqlCommand("BL_Check_User_Level", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCom.UserName);
                cmd.Parameters.AddWithValue("@BlanketNo", int.Parse(Blanket_Agreement_No));

                SqlParameter returnPara = cmd.Parameters.Add("@ReturnVal", SqlDbType.Int);
                returnPara.Direction = ParameterDirection.ReturnValue;

                conn.Open();
                cmd.ExecuteNonQuery();
                result = (int)returnPara.Value;
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
            return result;
        }

        int Check_Current_Level()
        {
            SqlCommand cmd = null;
            int result = 0;
            try
            {

                cmd = new SqlCommand("BL_Check_Current_Level", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.AddWithValue("@UserName", oCom.UserName);
                cmd.Parameters.AddWithValue("@BlanketNo", int.Parse(Blanket_Agreement_No));

                SqlParameter returnPara = cmd.Parameters.Add("@ReturnVal", SqlDbType.Int);
                returnPara.Direction = ParameterDirection.ReturnValue;
                if(conn.State != ConnectionState.Open)
                    conn.Open();
                cmd.ExecuteNonQuery();
                result = (int)returnPara.Value;

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
            return result;
        }

        void Load_Data()
        {
            try
            {
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(string.Format("Select AbsID from OOAT where Number ='{0}' and BpType = '{1}' and Status ='D' and Cancelled ='N'", Blanket_Agreement_No, Blanket_Type));
                AbsID = rs.Fields.Item("AbsID").Value.ToString();
                SAPbobsCOM.CompanyService oCompSer = oCom.GetCompanyService();
                SAPbobsCOM.BlanketAgreementsService oBAService = (SAPbobsCOM.BlanketAgreementsService)oCompSer.GetBusinessService(SAPbobsCOM.ServiceTypes.BlanketAgreementsService);
                SAPbobsCOM.BlanketAgreementParams oParams = (SAPbobsCOM.BlanketAgreementParams)oBAService.GetDataInterface(SAPbobsCOM.BlanketAgreementsServiceDataInterfaces.basBlanketAgreementParams);
                oParams.AgreementNo = int.Parse(AbsID);
                SAPbobsCOM.BlanketAgreement oBA = oBAService.GetBlanketAgreement(oParams);
                EditText0.Value = oBA.UserFields.Item("U_PRJ").Value.ToString();
                FProject = oBA.UserFields.Item("U_PRJ").Value.ToString();
                EditText1.Value = oBA.DocNum.ToString();
                E_BpName = oBA.BPName;
            }
            catch
            {
                return;
            }
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

        private DataTable Get_lst_User_Next_LV(int pLVL_Posting)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BL_Get_Lst_Usr_LV", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@BlanketNo", Blanket_Agreement_No);
                cmd.Parameters.AddWithValue("@LVL_Posting", pLVL_Posting);
                if (conn.State == ConnectionState.Closed) 
                    conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
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
            return result;
        }

        private DataTable Get_Email_Conf()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("GET_EMAIL_CONF", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
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
            return result;
        }

        private void Send_Email(string pToEMail, string pNameTo, string pBody, string pSubject)
        {
            MailMessage msg = new MailMessage();
            msg.To.Add(new MailAddress(pToEMail, pNameTo));
            msg.Priority = MailPriority.High;
            msg.From = new MailAddress(Email_From, Email_From_Name);
            msg.Subject = pSubject;
            msg.Body = pBody;
            msg.IsBodyHtml = true;

            SmtpClient client = new SmtpClient();
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential(Uid, Pwd);
            client.Port = Host_Port; // You can use Port 25 if 587 is blocked (mine is!)
            client.Host = Host_Address;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            //client.Timeout = 30;
            client.EnableSsl = EnableSSL;
            try
            {
                client.Send(msg);
            }
            catch (Exception ex)
            {
                oApp.SetStatusBarMessage("Send Email error: " + ex.Message);
            }
        }

        //Approve with Note
        private void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Fr_Authorise)
            {
                SqlCommand cmd = null;
                try
                {
                    cmd = new SqlCommand("BL_Update_Post_Level_HD_WithNote", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@AbsID", AbsID);
                    cmd.Parameters.AddWithValue("@Blanket_Type", Blanket_Type);
                    cmd.Parameters.AddWithValue("@Blanket_Level", Usr_LVL);
                    cmd.Parameters.AddWithValue("@Usr", oCom.UserName);
                    cmd.Parameters.AddWithValue("@Lvl1", lvl1);
                    cmd.Parameters.AddWithValue("@Approve", 2);
                    cmd.Parameters.AddWithValue("@Usr_Comment", this.EditText2.Value.Trim());
                    conn.Open();
                    int row_count = cmd.ExecuteNonQuery();
                    if (row_count == 0)
                        Application.SBO_Application.StatusBar.SetText("Approve Failed !");
                    else
                    {
                        //Send Notification
                        SAPbobsCOM.Messages msg = null;
                        msg = (SAPbobsCOM.Messages)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                        msg.MessageText = string.Format(@"Số HD {0} của Phòng ban/Dự án {1} nhận được yêu cầu điều chỉnh thông tin.{2}Nội dung: {3}.{2}Anh chị vui lòng vào xem.", Blanket_Agreement_No, FProject, Environment.NewLine, this.EditText2.Value.Trim());
                        msg.Subject = "Yêu cầu điều chỉnh Hợp đồng số " + Blanket_Agreement_No.ToString();
                        msg.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;
                        msg.Recipients.SetCurrentLine(0);
                        msg.Recipients.UserCode = User_Created;
                        msg.Recipients.NameTo = User_Created;
                        msg.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                        msg.Add();
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
                        if (Parent_Form.Title != "Danh sách Hợp đồng")
                        {
                            Parent_Form.Select();
                            Application.SBO_Application.ActivateMenuItem("1304");
                        }
                    }
                    this.UIAPIRawForm.Close();
                }
            }
        }

        //Approve
        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Fr_Authorise)
            {
                SqlCommand cmd = null;
                try
                {
                    cmd = new SqlCommand("BL_Update_Post_Level_HD", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@AbsID", AbsID);
                    cmd.Parameters.AddWithValue("@Blanket_Type", Blanket_Type);
                    cmd.Parameters.AddWithValue("@Blanket_Level", Usr_LVL);
                    cmd.Parameters.AddWithValue("@Usr", oCom.UserName);
                    cmd.Parameters.AddWithValue("@Approve", 1);
                    cmd.Parameters.AddWithValue("@Usr_Comment", this.EditText2.Value.Trim());
                    conn.Open();
                    int row_count = cmd.ExecuteNonQuery();
                    if (row_count == 0)
                        Application.SBO_Application.StatusBar.SetText("Approve Failed !");
                    else
                    {
                        int New_LV = Check_Current_Level();
                        if (LVL_Posting != New_LV)
                        {
                            //Send notification
                            DataTable lst = Get_lst_User_Next_LV(New_LV);
                            if (lst.Rows.Count > 0)
                            {
                                SAPbobsCOM.Messages msg = null;
                                msg = (SAPbobsCOM.Messages)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                                msg.MessageText = string.Format(@"Số HD {0} của phòng ban/dự án {1} đang chờ duyệt.{2}Anh chị vui lòng vào xem.", Blanket_Agreement_No,FProject,Environment.NewLine);
                                msg.Subject = "Yêu cầu phê duyệt Hợp đồng số " + Blanket_Agreement_No.ToString();
                                msg.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;
                                for (int i = 0; i < lst.Rows.Count; i++)
                                {
                                    msg.Recipients.SetCurrentLine(i);
                                    msg.Recipients.UserCode = lst.Rows[i]["USER_CODE"].ToString();
                                    msg.Recipients.NameTo = lst.Rows[i]["NAME"].ToString();
                                    msg.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                                    if (lst.Rows[i]["EMAIL"].ToString() != "" && !string.IsNullOrEmpty(lst.Rows[i]["EMAIL"].ToString()))
                                    {
                                        string Body_Content = string.Format(@"Dear Anh/Chị,<br/><br/>"
                                           + @"Có <b>phiếu trình ký hợp đồng</b> đang chờ bạn xử lý trên hệ thống SAP<br/><br/>"
                                           + @"<b>Thông tin phiếu trình ký hợp đồng:</b><br/><br/>"
                                           + @"P/B/BP/Dự án : <b>{0}</b><br/>"
                                           + @"Số hợp đồng : <b>{1}</b><br/>"
                                           + @"Đối tượng : <b>{2}</b><br/><br/>"
                                           + @"Đây là email được gửi tự động từ hệ thống SAP, vui lòng không trả lời lại email này. Xin cảm ơn.<br/><br/>"
                                           + @"--------------<br/>"
                                           + @"Trân trọng<br/>"
                                           + @"SAP Business One", FProject, Blanket_Agreement_No, E_BpName);
                                        Send_Email(lst.Rows[i]["EMAIL"].ToString(), lst.Rows[i]["NAME"].ToString(), Body_Content, msg.Subject);
                                    }
                                    if (i < lst.Rows.Count - 1)
                                    {
                                        msg.Recipients.Add();
                                    }
                                }
                                msg.Add();
                            }
                            
                        }
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
                    if (conn.State == ConnectionState.Open)
                        conn.Close();
                    cmd.Dispose();
                    if (!Check_Parent_Form_Closed(Parent_Form.UniqueID))
                    {
                        if (Parent_Form.Title != "Danh sách Hợp đồng")
                        {
                            Parent_Form.Select();
                            Application.SBO_Application.ActivateMenuItem("1304");
                        }
                    }
                    this.UIAPIRawForm.Close();
                }

            }
        }

        //Not Approve
        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        { 
            if (Fr_Authorise)
            {
                SqlCommand cmd = null;
                try
                {
                    cmd = new SqlCommand("BL_Update_Post_Level_HD", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@AbsID", AbsID);
                    cmd.Parameters.AddWithValue("@Blanket_Type", Blanket_Type);
                    cmd.Parameters.AddWithValue("@Blanket_Level", Usr_LVL);
                    cmd.Parameters.AddWithValue("@Usr", oCom.UserName);
                    cmd.Parameters.AddWithValue("@Approve", 0);
                    cmd.Parameters.AddWithValue("@Usr_Comment", this.EditText2.Value.Trim());
                    conn.Open();
                    int row_count = cmd.ExecuteNonQuery();
                    if (row_count == 0)
                        Application.SBO_Application.StatusBar.SetText("Approve Failed !");
                    else
                    {
                        //Cancel Blanket Agreement
                        if (LVL_Posting == 5)
                        {
                            SAPbobsCOM.CompanyService oCompSer = oCom.GetCompanyService();
                            SAPbobsCOM.BlanketAgreementsService oBAService = (SAPbobsCOM.BlanketAgreementsService)oCompSer.GetBusinessService(SAPbobsCOM.ServiceTypes.BlanketAgreementsService);
                            SAPbobsCOM.BlanketAgreementParams oParams = (SAPbobsCOM.BlanketAgreementParams)oBAService.GetDataInterface(SAPbobsCOM.BlanketAgreementsServiceDataInterfaces.basBlanketAgreementParams);
                            oParams.AgreementNo = int.Parse(AbsID);
                            oBAService.CancelBlanketAgreement(oParams);
                        }
                        else if (LVL_Posting == 4 && position == 1)
                        {
                            SAPbobsCOM.CompanyService oCompSer = oCom.GetCompanyService();
                            SAPbobsCOM.BlanketAgreementsService oBAService = (SAPbobsCOM.BlanketAgreementsService)oCompSer.GetBusinessService(SAPbobsCOM.ServiceTypes.BlanketAgreementsService);
                            SAPbobsCOM.BlanketAgreementParams oParams = (SAPbobsCOM.BlanketAgreementParams)oBAService.GetDataInterface(SAPbobsCOM.BlanketAgreementsServiceDataInterfaces.basBlanketAgreementParams);
                            oParams.AgreementNo = int.Parse(AbsID);
                            oBAService.CancelBlanketAgreement(oParams);
                        }
                        //Send Notification to infor creator
                        SAPbobsCOM.Messages msg = null;
                        msg = (SAPbobsCOM.Messages)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                        msg.MessageText = string.Format(@"Hợp đồng số {0} đã bị từ chối.{1}Nội dung Comment: {2}", Blanket_Agreement_No, Environment.NewLine, EditText2.Value);
                        msg.Subject = "Hợp đồng số" + Blanket_Agreement_No.ToString() + " bị từ chối";
                        msg.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;
                        msg.Recipients.SetCurrentLine(0);
                        msg.Recipients.UserCode = User_Created;
                        msg.Recipients.NameTo = User_Created;
                        msg.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                        msg.Add();
                    }
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.MessageBox("Can't reject - Error: " + ex.Message);
                }
                finally
                {
                    conn.Close();
                    cmd.Dispose();
                    if (!Check_Parent_Form_Closed(Parent_Form.UniqueID))
                    {
                        if (Parent_Form.Title != "Danh sách Hợp đồng")
                        {
                            Parent_Form.Select();
                            Application.SBO_Application.ActivateMenuItem("1304");
                        }
                    }
                    this.UIAPIRawForm.Close();
                }
            }
        }

        private void Form_CloseBefore(SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                oCom.Disconnect();
            }
            catch
            { }

        }

        //int Check_Current_Level()
        //{
        //    int result = -9;
        //    if (Post_Level.Count > 0)
        //    {
        //        if (CGroup == "CD")
        //        {
        //            if (Post_Level[10] != "2")
        //            {
        //                if (deptcreate == 1 || deptcreate == 2)
        //                {
        //                    result = 3;
        //                    if (Post_Level[4] != "" && Post_Level[6] != "" && Post_Level[8] != "" && result == 3)
        //                    {
        //                        result = 4;
        //                        if (Post_Level[10] == "1" && result == 4)
        //                        {
        //                            result = 5;
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    result = 1;
        //                    if (Post_Level[1] != "" && result == 1)
        //                    {
        //                        result = 3;
        //                        if (Post_Level[4] != "" && Post_Level[6] != "" && Post_Level[8] != "" && result == 3)
        //                        {
        //                            result = 4;
        //                            if (Post_Level[10] == "1" && result == 4)
        //                            {
        //                                result = 5;
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                if (deptcreate == 1 || deptcreate == 2)
        //                {
        //                    result = 4;
        //                    if (Post_Level[10] == "1" && result == 4)
        //                    {
        //                        result = 5;
        //                    }
        //                }
        //                else
        //                {
        //                    result = 1;
        //                    if (Post_Level[1] != "2" && result == 1)
        //                    {
        //                        result = 4;
        //                        if (Post_Level[10] == "1" && result == 4)
        //                        {
        //                            result = 5;
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        else if (CGroup == "XD")
        //        {
        //            if (Post_Level[10] != "2")
        //            {
        //                if (deptcreate == 1 || deptcreate == 2)
        //                {
        //                    result = 3;
        //                    if (Post_Level[4] != "" && Post_Level[8] != "" && result == 3)
        //                    {
        //                        result = 4;
        //                        if (Post_Level[10] == "1" && result == 4)
        //                        {
        //                            result = 5;
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    result = 2;
        //                    if (Post_Level[2] != "" && result == 2)
        //                    {
        //                        result = 3;
        //                        if (Post_Level[4] != "" && Post_Level[8] != "" && result == 3)
        //                        {
        //                            result = 4;
        //                            if (Post_Level[10] == "1" && result == 4)
        //                            {
        //                                result = 5;
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                if (deptcreate == 1 || deptcreate == 2)
        //                {
        //                    result = 4;
        //                    if (Post_Level[10] == "1" && result == 4)
        //                    {
        //                        result = 5;
        //                    }
        //                }
        //                else
        //                {
        //                    result = 2;
        //                    if (Post_Level[2] != "2" && result == 2)
        //                    {
        //                        result = 4;
        //                        if (Post_Level[10] == "1" && result == 4)
        //                        {
        //                            result = 5;
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        else if (CGroup == "CDXD")
        //        {
        //            if (Post_Level[10] != "2")
        //            {
        //                if (deptcreate == 1 || deptcreate == 2)
        //                {
        //                    result = 3;
        //                    if (Post_Level[4] != "" && Post_Level[6] != "" && Post_Level[8] != "" && result == 3)
        //                    {
        //                        result = 4;
        //                        if (Post_Level[10] == "1" && result == 4)
        //                        {
        //                            result = 5;
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    result = 1;
        //                    if (Post_Level[1] != "" && result == 1)
        //                    {
        //                        result = 2;
        //                        if (Post_Level[2] != "" && result == 2)
        //                        {
        //                            result = 3;
        //                            if (Post_Level[4] != "" && Post_Level[6] != "" && Post_Level[8] != "" && result == 3)
        //                            {
        //                                result = 4;
        //                                if (Post_Level[10] == "1" && result == 4)
        //                                {
        //                                    result = 5;
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                if (deptcreate == 1 || deptcreate == 2)
        //                {
        //                    result = 3;
        //                    if (Post_Level[4] != "" && Post_Level[6] != "" && Post_Level[8] != "" && result == 3)
        //                    {
        //                        result = 4;
        //                        if (Post_Level[10] == "1" && result == 4)
        //                        {
        //                            result = 5;
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    result = 1;
        //                    if (Post_Level[1] != "2" && result == 1)
        //                    {
        //                        result = 4;
        //                        if (Post_Level[10] == "1" && result == 4)
        //                        {
        //                            result = 5;
        //                        }
        //                    }
        //                }
        //            }

        //        }
        //        else if (CGroup == "TB")
        //        {
        //            if (Post_Level[10] != "2")
        //            {
        //                if (deptcreate == 1 )//|| deptcreate == 2)
        //                {
        //                    result = 3;
        //                    if (Post_Level[4] != "" && Post_Level[8] != "" && result == 3)
        //                    {
        //                        result = 4;
        //                        if (Post_Level[10] == "1" && result == 4)
        //                        {
        //                            result = 5;
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    result = 1;
        //                    if (Post_Level[1] != "" && result == 1)
        //                    {
        //                        result = 3;
        //                        if (Post_Level[4] != "" && Post_Level[8] != "" && result == 3)
        //                        {
        //                            result = 4;
        //                            if (Post_Level[10] == "1" && result == 4)
        //                            {
        //                                result = 5;
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                if (deptcreate == 1 || deptcreate == 2)
        //                {
        //                    result = 4;
        //                    if (Post_Level[10] == "1" && result == 4)
        //                    {
        //                        result = 5;
        //                    }
        //                }
        //                else
        //                {
        //                    result = 1;
        //                    if (Post_Level[1] != "2" && result == 1)
        //                    {
        //                        result = 4;
        //                        if (Post_Level[10] == "1" && result == 4)
        //                        {
        //                            result = 5;
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        else if (CGroup == "TBXD")
        //        {
        //            if (Post_Level[10] != "2")
        //            {
        //                if (deptcreate == 1 || deptcreate == 2)
        //                {
        //                    result = 3;
        //                    if (Post_Level[4] != "" && Post_Level[6] != "" && Post_Level[8] != "" && result == 3)
        //                    {
        //                        result = 4;
        //                        if (Post_Level[10] == "1" && result == 4)
        //                        {
        //                            result = 5;
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    result = 2;
        //                    if (Post_Level[2] != "" && result == 2)
        //                    {
        //                        result = 3;
        //                        if (Post_Level[4] != "" && Post_Level[6] != "" && Post_Level[8] != "" && result == 3)
        //                        {
        //                            result = 4;
        //                            if (Post_Level[10] == "1" && result == 4)
        //                            {
        //                                result = 5;
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                if (deptcreate == 1 || deptcreate == 2)
        //                {
        //                    result = 4;
        //                    if (Post_Level[10] == "1" && result == 4)
        //                    {
        //                        result = 5;
        //                    }
        //                }
        //                else
        //                {
        //                    result = 2;
        //                    if (Post_Level[2] != "2" && result == 2)
        //                    {
        //                        result = 4;
        //                        if (Post_Level[10] == "1" && result == 4)
        //                        {
        //                            result = 5;
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        else if (CGroup == "VP")
        //        {
        //            if (Post_Level[10] != "2" )
        //            {
        //                result = 1;
        //                if (Post_Level[1] != "" && result == 1)
        //                {
        //                    result = 3;
        //                    if (Post_Level[4] != "" && Post_Level[8] != "" && result == 3)
        //                    {
        //                        result = 4;
        //                        if (Post_Level[10] == "1" && result == 4)
        //                        {
        //                            result = 5;
        //                        }
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                result = 1;
        //                if (Post_Level[1] != "2" && result == 1)
        //                {
        //                    result = 4;
        //                    if (Post_Level[10] == "1" && result == 4)
        //                    {
        //                        result = 5;
        //                    }
        //                }
        //            }
        //        }
        //    }

        //    return result;
        //}

        //int Get_User_Level()
        //{
        //    int result = -9;;
        //    if (CGroup == "CD")
        //    {
        //        //chỉ huy trưởng ME		 phòng pháp chế	phòng ME	phòng kế toán	phòng CCM	 giám đốc dự án  
        //        if (position == 6)
        //            result = 1;
        //        else if (position == 2 && dept == 4)
        //            result = 3;
        //        else if (position == 1 && dept == 4)
        //            result = 3;
        //        else if (position == 2 && dept == 5)
        //            result = 3;
        //        else if (position == 1 && dept == 5)
        //            result = 3;
        //        else if (position == 2 && dept == -2)
        //            result = 3;
        //        else if (position == 1 && dept == -2)
        //            result = 3;
        //        else if (position == 2 && dept == 1)
        //            result = 4;
        //        else if (position == 1 && dept == 1)
        //            result = 4;
        //        else if (position == 3)
        //            result = 5;
        //    }
        //    else if (CGroup == "XD")
        //    {
        //        //Xây dựng		chỉ huy trưởng dự án	 phòng pháp chế		phòng kế toán	phòng CCM	 giám đốc dự án.
        //        if (position == 5)
        //            result = 2;
        //        else if (position == 2 && dept == 4)
        //            result = 3;
        //        else if (position == 1 && dept == 4)
        //            result = 3;
        //        else if (position == 2 && dept == -2)
        //            result = 3;
        //        else if (position == 1 && dept == -2)
        //            result = 3;
        //        else if (position == 2 && dept == 1)
        //            result = 4;
        //        else if (position == 1 && dept == 1)
        //            result = 4;
        //        else if (position == 3)
        //            result = 5;
        //    }
        //    else if (CGroup == "CDXD")
        //    {
        //        if (position == 6)
        //            result = 1;
        //        else if (position == 5)
        //            result = 2;
        //        else if (position == 2 && dept == 4)
        //            result = 3;
        //        else if (position == 1 && dept == 4)
        //            result = 3;
        //        else if (position == 2 && dept == 5)
        //            result = 3;
        //        else if (position == 1 && dept == 5)
        //            result = 3;
        //        else if (position == 2 && dept == -2)
        //            result = 3;
        //        else if (position == 1 && dept == -2)
        //            result = 3;
        //        else if (position == 2 && dept == 1)
        //            result = 4;
        //        else if (position == 1 && dept == 1)
        //            result = 4;
        //        else if (position == 3)
        //            result = 5;
        //    }
        //    else if (CGroup == "TB")
        //    {
        //        //Trưởng phòng Thiết bị		 phòng pháp chế		phòng kế toán	phòng CCM	 giám đốc dự án.
        //        if (position == 1 && dept == 2)
        //            result = 1;
        //        else if (position == 2 && dept == 4)
        //            result = 3;
        //        else if (position == 1 && dept == 4)
        //            result = 3;
        //        else if (position == 2 && dept == -2)
        //            result = 3;
        //        else if (position == 1 && dept == -2)
        //            result = 3;
        //        else if (position == 2 && dept == 1)
        //            result = 4;
        //        else if (position == 1 && dept == 1)
        //            result = 4;
        //        else if (position == 3)
        //            result = 5;
        //    }
        //    else if (CGroup == "TBXD")
        //    {
        //        if (position == 5)
        //            result = 2;
        //        else if (position == 2 && dept == 4)
        //            result = 3;
        //        else if (position == 1 && dept == 4)
        //            result = 3;
        //        else if (position == 2 && dept == 2)
        //            result = 3;
        //        else if (position == 1 && dept == 2)
        //            result = 3;
        //        else if (position == 2 && dept == -2)
        //            result = 3;
        //        else if (position == 1 && dept == -2)
        //            result = 3;
        //        else if (position == 2 && dept == 1)
        //            result = 4;
        //        else if (position == 1 && dept == 1)
        //            result = 4;
        //        else if (position == 3)
        //            result = 5;
        //    }
        //    else if (CGroup == "VP")
        //    {
        //        if (position == 1 && dept == deptcreate && (Post_Level[1] == "" || Post_Level[1] == "2"))
        //            result = 1;                 
        //        else if (position == 2 && dept == 4)
        //            result = 3;
        //        else if (position == 1 && dept == 4)
        //            result = 3;
        //        else if (position == 2 && dept == -2)
        //            result = 3;
        //        else if (position == 1 && dept == -2)
        //            result = 3;
        //        else if (position == 2 && dept == 1)
        //            result = 4;
        //        else if (position == 1 && dept == 1)
        //            result = 4;
        //        else if (position == 4)
        //            result = 5;
        //    }
        //    else
        //    {
        //        result = -9;
        //    }
        //    return result;
        //}

        //int Get_Level_1()
        //{
        //    int result = 0;
        //    if (CGroup == "CD")
        //    {
        //        result = 1;
        //    }
        //    else if (CGroup == "XD")
        //    {
        //        result = 2;
        //    }
        //    else if (CGroup == "CDXD")
        //    {
        //        result = 1;
        //    }
        //    else if (CGroup == "TB")
        //    {
        //        result = 1;
        //    }
        //    else if (CGroup == "TBXD")
        //    {
        //        result = 2;
        //    }
        //    else if (CGroup == "VP")
        //    {
        //        result = 1;
        //    }
        //    else
        //    {
        //        result = 0;
        //    }
        //    return result;
        //}

        //int Get_User_Posting_Level()
        //{
        //    int result = -9;
        //    if (CGroup == "CD")
        //    {
        //        //chỉ huy trưởng ME		 phòng pháp chế	phòng ME	phòng kế toán	phòng CCM	 giám đốc dự án  
        //        if (position == 6)
        //            result = 1;
        //        else if (position == 2 && dept == 4)
        //            result = 3;
        //        else if (position == 1 && dept == 4)
        //            result = 4;
        //        else if (position == 2 && dept == 5)
        //            result = 5;
        //        else if (position == 1 && dept == 5)
        //            result = 6;
        //        else if (position == 2 && dept == -2)
        //            result = 7;
        //        else if (position == 1 && dept == -2)
        //            result = 8;
        //        else if (position == 2 && dept == 1)
        //            result = 9;
        //        else if (position == 1 && dept == 1)
        //            result = 10;
        //        else if (position == 3)
        //            result = 11;
        //    }
        //    else if (CGroup == "XD")
        //    {
        //        //Xây dựng		chỉ huy trưởng dự án	 phòng pháp chế		phòng kế toán	phòng CCM	 giám đốc dự án.
        //        if (position == 5)
        //            result = 2;
        //        else if (position == 2 && dept == 4)
        //            result = 3;
        //        else if (position == 1 && dept == 4)
        //            result = 4;
        //        else if (position == 2 && dept == -2)
        //            result = 7;
        //        else if (position == 1 && dept == -2)
        //            result = 8;
        //        else if (position == 2 && dept == 1)
        //            result = 9;
        //        else if (position == 1 && dept == 1)
        //            result = 10;
        //        else if (position == 3)
        //            result = 11;
        //    }
        //    else if (CGroup == "CDXD")
        //    {
        //        if (position == 6)
        //            result = 1;
        //        else if (position == 5)
        //            result = 2;
        //        else if (position == 2 && dept == 4)
        //            result = 3;
        //        else if (position == 1 && dept == 4)
        //            result = 4;
        //        else if (position == 2 && dept == 5)
        //            result = 5;
        //        else if (position == 1 && dept == 5)
        //            result = 6;
        //        else if (position == 2 && dept == -2)
        //            result = 7;
        //        else if (position == 1 && dept == -2)
        //            result = 8;
        //        else if (position == 2 && dept == 1)
        //            result = 9;
        //        else if (position == 1 && dept == 1)
        //            result = 10;
        //        else if (position == 3)
        //            result = 11;
        //    }
        //    else if (CGroup == "TB")
        //    {
        //        //Trưởng phòng Thiết bị		 phòng pháp chế		phòng kế toán	phòng CCM	 giám đốc dự án.
        //        if (position == 1 && dept == 2)
        //            result = 1;
        //        else if (position == 2 && dept == 4)
        //            result = 3;
        //        else if (position == 1 && dept == 4)
        //            result = 4;
        //        else if (position == 2 && dept == -2)
        //            result = 7;
        //        else if (position == 1 && dept == -2)
        //            result = 8;
        //        else if (position == 2 && dept == 1)
        //            result = 9;
        //        else if (position == 1 && dept == 1)
        //            result = 10;
        //        else if (position == 3)
        //            result = 11;
        //    }
        //    else if (CGroup == "TBXD")
        //    {
        //        if (position == 5)
        //            result = 2;
        //        else if (position == 2 && dept == 4)
        //            result = 3;
        //        else if (position == 1 && dept == 4)
        //            result = 4;
        //        else if (position == 2 && dept == 2)
        //            result = 5;
        //        else if (position == 1 && dept == 2)
        //            result = 6;
        //        else if (position == 2 && dept == -2)
        //            result = 7;
        //        else if (position == 1 && dept == -2)
        //            result = 8;
        //        else if (position == 2 && dept == 1)
        //            result = 9;
        //        else if (position == 1 && dept == 1)
        //            result = 10;
        //        else if (position == 3)
        //            result = 11;
        //    }
        //    else if (CGroup == "VP")
        //    {
        //        if (position == 1 && dept == deptcreate && (Post_Level[1] == "" || Post_Level[1] == "2"))
        //            result = 1;             
        //        else if (position == 2 && dept == 4)
        //            result = 3;
        //        else if (position == 1 && dept == 4)
        //            result = 4;
        //        else if (position == 2 && dept == -2)
        //            result = 7;
        //        else if (position == 1 && dept == -2)
        //            result = 8;
        //        else if (position == 2 && dept == 1)
        //            result = 9;
        //        else if (position == 1 && dept == 1)
        //            result = 10;
        //        else if (position == 4)
        //            result = 11;
        //    }
        //    else
        //    {
        //        result = -9;
        //    }
        //    return result;
        //}
    }
}
