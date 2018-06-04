using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Net.Mail;

namespace U_ApproveJV
{
    [FormAttribute("U_ApproveJV.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
        int DocEntry = 0, BatchNum = 0,Period =0;
        string User_Create = "", Type = "", Dep_BpName = "", FProject = "";
        //Email Server Config
        string Email_From = "";
        string Email_From_Name = "";
        string Host_Address = "";
        int Host_Port = 25;
        bool EnableSSL = false;
        string Uid = "";
        string Pwd = "";
        //End Email Config
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Grid Grid1;
        private SAPbouiCOM.Grid Grid2;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.OptionBtn OptionBtn0;
        private SAPbouiCOM.OptionBtn OptionBtn1;
        private SAPbouiCOM.OptionBtn OptionBtn2;
        private SAPbouiCOM.OptionBtn OptionBtn3;

        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("grd_lst").Specific));
            this.Grid0.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid0_DoubleClickAfter);
            this.Grid0.PressedAfter += new SAPbouiCOM._IGridEvents_PressedAfterEventHandler(this.Grid0_PressedAfter);
            this.Grid1 = ((SAPbouiCOM.Grid)(this.GetItem("grd_pro").Specific));
            this.Grid2 = ((SAPbouiCOM.Grid)(this.GetItem("grd_info").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_appr").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_23").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("bt_rej").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("bt_appr2").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_comm").Specific));
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("bt_cover").Specific));
            this.Button3.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button3_PressedAfter);
            this.OptionBtn0 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_cur").Specific));
            this.OptionBtn0.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn0_PressedAfter);
            this.OptionBtn1 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_appr").Specific));
            this.OptionBtn1.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn1_PressedAfter);
            this.OptionBtn2 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_rej").Specific));
            this.OptionBtn2.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn2_PressedAfter);
            this.OptionBtn3 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_all").Specific));
            this.OptionBtn3.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn3_PressedAfter);
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
            //Load_Grid_Period();

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

            OptionBtn1.GroupWith("op_cur");
            OptionBtn2.GroupWith("op_cur");
            OptionBtn3.GroupWith("op_cur");
            OptionBtn0.Selected = true;
        }

        //Get Email Config
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

        //Send Email
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

        //Grid Approve
        private void Load_Grid_Period()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("JV_GetList_Approve_Current", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);
                
                Grid0.AutoResizeColumns();

                Grid0.Columns.Item("NumOfTrans").Visible = false;
                //Grid0.Columns.Item("DocEntry").Visible = false;
                Grid0.Columns.Item("POST_LVL").Visible = false;

            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
                conn.Close();
                cmd.Dispose();
            }
        }

        //Grid Approved
        private void Load_Grid_Approved_Period()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("JV_GetList_Approved_Current", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);
                
                Grid0.Columns.Item("NumOfTrans").Visible = false;
                Grid0.Columns.Item("POST_LVL").Visible = false;
                Grid0.Columns.Item("ProfitCode").Visible = false;

                Grid0.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
                conn.Close();
                cmd.Dispose();
            }
        }

        //Grid Rejected
        private void Load_Grid_Rejected_Period()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("JV_GetList_Rejected_Current", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);
 
                Grid0.Columns.Item("NumOfTrans").Visible = false;
                Grid0.Columns.Item("POST_LVL").Visible = false;
                Grid0.Columns.Item("ProfitCode").Visible = false;

                Grid0.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
                conn.Close();
                cmd.Dispose();
            }
        }

        //Grid All
        private void Load_Grid_Period_All()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("JV_GetList_Current", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();

                Grid0.Columns.Item("NumOfTrans").Visible = false;
                //Grid0.Columns.Item("DocEntry").Visible = false;
                Grid0.Columns.Item("POST_LVL").Visible = false;

            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
                conn.Close();
                cmd.Dispose();
            }
        }

        private void Load_Grid_Info()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                if (Type != "BILLVP")
                {
                    cmd = new SqlCommand("JV_Get_Total_Approve", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@BatchNum", BatchNum);
                }
                else
                {
                    cmd = new SqlCommand("VPBILL_Get_Total_Approve", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@DocEntry", DocEntry);
                }
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid2.DataTable = Convert_SAP_DataTable_Info(result);
                Grid2.AutoResizeColumns();
                this.UIAPIRawForm.Freeze(false);

            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
                conn.Close();
                cmd.Dispose();
            }
        }

        private SAPbouiCOM.DataTable Convert_SAP_DataTable(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_JVList"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_JVList");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_JVList");
            }

            //Add column to DataTable
            foreach (System.Data.DataColumn c in pDataTable.Columns)
            {
                try
                {
                    if (c.DataType.ToString() == "System.DateTime")
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Date);
                    else
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Text);
                }
                catch
                { }

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

            return oDT;
        }

        private SAPbouiCOM.DataTable Convert_SAP_DataTable_Info(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_JVInfo"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_JVInfo");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_JVInfo");
            }

            //Add column to DataTable
            foreach (System.Data.DataColumn c in pDataTable.Columns)
            {
                try
                {
                    if (c.DataType.ToString() == "System.DateTime")
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Date);
                    else
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Text);
                }
                catch
                { }

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

            return oDT;
        }

        private bool CheckExistUniqueID(SAPbouiCOM.Form pForm, string pItemID)
        {
            if (pForm.DataSources.DataTables.Count > 0)
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

        private void Grid0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            DocEntry = 0;
            if (Grid0.Rows.SelectedRows.Count == 1)
            {
                int.TryParse(Grid0.DataTable.GetValue("DocEntry", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out DocEntry);
                int.TryParse(Grid0.DataTable.GetValue("BatchNum", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out BatchNum);
                Period = 0;
                int.TryParse(Grid0.DataTable.GetValue("Period", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out Period);
                User_Create = Grid0.DataTable.GetValue("Username", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                Type = Grid0.DataTable.GetValue("Type", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                Dep_BpName = Grid0.DataTable.GetValue(@"Department/BPName", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                FProject = Grid0.DataTable.GetValue(@"Project", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                //JV
                if (Type != "BILLVP")
                {
                    //Load Info Posting Level
                    string lenh = @"Select U_Level as 'DeptCode',(Select [Name] from OUDP where Code=U_Level) as 'DeptName' , U_Position as 'PosCode', (Select [Name] from OHPS where posID=U_Position) as 'Position',(Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') as 'NAME' from OHEM"
                        + @" where userId = (Select UserId from OUSR where User_Code = U_Usr)) as 'Approved by' ,case U_Status when 1 then 'Approved' when 2 then 'Rejected' when 3 then 'By Pass' when 4 then 'Approved with Comment' end as 'Status'"
                        + @",U_Time as 'Approved on', U_Comment as 'Comment' From [@JV_APROVE_D] where DocEntry =" + DocEntry;
                    Grid1.DataTable.ExecuteQuery(lenh);
                }
                //BILL VP
                else
                {
                    //Load Info Posting Level
                    string lenh = @"Select U_Level as 'DeptCode',(Select [Name] from OUDP where Code=U_Level) as 'DeptName' , U_Position as 'PosCode', (Select [Name] from OHPS where posID=U_Position) as 'Position',(Select ISNULL(lastName,'') +' ' + ISNULL(middleName,'') +' ' +ISNULL(firstName,'') as 'NAME' from OHEM"
                        + @" where userId = (Select UserId from OUSR where User_Code = U_Usr)) as 'Approved by' ,case U_Status when 1 then 'Approved' when 2 then 'Rejected' when 3 then 'By Pass' when 4 then 'Approved with Comment' end as 'Status'"
                        + @",U_Time as 'Approved on', U_Comment as 'Comment' From [@BILLVP2] where DocEntry =" + DocEntry;
                    Grid1.DataTable.ExecuteQuery(lenh);
                }
                Grid1.Columns.Item(0).Visible = false;
                Grid1.Columns.Item(2).Visible = false;
                Grid1.AutoResizeColumns();
                //this.Button0.Item.Enabled = true;
                //this.Button1.Item.Enabled = true;
                //if (Check_CCM(oCompany.UserName))
                //    this.Button2.Item.Enabled = true;
                //else
                //    this.Button2.Item.Enabled = false;
                Load_Grid_Info();
            }
        }

        private bool Check_GD_DA(string pUsrName)
        {
            string sql = string.Format("Select position from OHEM  where userID = (Select t.USERID from OUSR t where t.User_Code='{0}')", pUsrName);
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(sql);
            if (oR_RecordSet.RecordCount > 0)
            {
                string position = oR_RecordSet.Fields.Item("position").Value.ToString();
                if (position == "3" || position == "4") return true;
                else return false;
            }
            return false;
        }

        private bool Check_Manager(string pUsrName)
        {
            string sql = string.Format("Select position from OHEM  where userID = (Select t.USERID from OUSR t where t.User_Code='{0}')", pUsrName);
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(sql);
            if (oR_RecordSet.RecordCount > 0)
            {
                string position = oR_RecordSet.Fields.Item("position").Value.ToString();
                if (position == "1") return true;
                else return false;
            }
            return false;
        }

        private bool Check_All_Approve()
        {
            string sql= "";
            if (Type != "BILLVP")
                sql = string.Format("Select case when Count(U_Status) = Count(*) then 1 else 0 end as Result from [@JV_APROVE_D] where DocEntry = {0}", DocEntry);
            else
                sql = string.Format("Select case when Count(U_Status) = Count(*) then 1 else 0 end as Result from [@BILLVP2] where DocEntry = {0}", DocEntry);
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(sql);
            if (oR_RecordSet.RecordCount > 0)
            {
                string Result = oR_RecordSet.Fields.Item("Result").Value.ToString();
                if (Result == "1") return true;
                else return false;
            }
            return false;
        }

        private bool Check_CCM(string pUsrName)
        {
            string sql = string.Format("Select dept from OHEM  where userID = (Select t.USERID from OUSR t where t.User_Code='{0}')", pUsrName);
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(sql);
            if (oR_RecordSet.RecordCount > 0)
            {
                string dept = oR_RecordSet.Fields.Item("dept").Value.ToString();
                if (dept == "1") return true;
                else return false;
            }
            return false;
        }
        
        private void Send_Alert()
        {
            DataTable lst = Get_lst_User_Next_LV();
            if (lst.Rows.Count > 0)
            {
                SAPbobsCOM.Messages msg = null;
                msg = (SAPbobsCOM.Messages)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                string E_Body = "";
                if (Type != "BILLVP")
                {
                    msg.MessageText = string.Format("Journal Voucher số {0} đang chờ phê duyệt", BatchNum);
                    msg.Subject = "Yêu cầu phê duyệt Journal Voucher số " + BatchNum.ToString();
                    E_Body = string.Format(@"Dear Anh/Chị,<br/><br/>"
                       + @"Có <b>đề nghị thanh toán</b> đang chờ bạn xử lý trên hệ thống SAP<br/><br/>"
                       + @"<b>Thông tin đề nghị thanh toán (Journal Vouchers):</b><br/><br/>"
                       + @"P/B/BP/Dự án : <b>{0}</b><br/>"
                       + @"Kỳ thanh toán : <b>{1}</b><br/>"
                       + @"Số JV : <b>{2}</b><br/>"
                       + @"Đối tượng : <b>{3}</b><br/><br/>"
                        //+ "Trạng thái           : <b>{4}</b> đã duyệt<br/>"
                       + @"Đây là email được gửi tự động từ hệ thống SAP, vui lòng không trả lời lại email này. Xin cảm ơn.<br/><br/>"
                       + @"--------------<br/>"
                       + @"Trân trọng<br/>"
                       + @"SAP Business One", FProject, Period, BatchNum, Dep_BpName);
                }
                else
                {
                    msg.MessageText = string.Format("Bill văn phòng số {0} đang chờ phê duyệt", DocEntry);
                    msg.Subject = "Yêu cầu phê duyệt Bill văn phòng số " + DocEntry.ToString();
                    E_Body = string.Format(@"Dear Anh/Chị,<br/><br/>"
                      + @"Có <b>đề nghị thanh toán</b> đang chờ bạn xử lý trên hệ thống SAP<br/><br/>"
                      + @"<b>Thông tin đề nghị thanh toán (Bill Văn phòng):</b><br/><br/>"
                      + @"P/B/BP/Dự án : <b>{0}</b><br/>"
                      + @"Kỳ thanh toán : <b>{1}</b><br/>"
                      + @"Số bill : <b>{2}</b><br/>"
                      + @"Đối tượng : <b>{3}</b><br/><br/>"
                        //+ "Trạng thái           : <b>{4}</b> đã duyệt<br/>"
                      + @"Đây là email được gửi tự động từ hệ thống SAP, vui lòng không trả lời lại email này. Xin cảm ơn.<br/><br/>"
                      + @"--------------<br/>"
                      + @"Trân trọng<br/>"
                      + @"SAP Business One", FProject, Period, DocEntry, Dep_BpName);
                }
                msg.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;

                for (int i = 0; i < lst.Rows.Count; i++)
                {
                    msg.Recipients.SetCurrentLine(i);
                    msg.Recipients.UserCode = lst.Rows[i]["USER_CODE"].ToString();
                    msg.Recipients.NameTo = lst.Rows[i]["NAME"].ToString();
                    msg.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                    msg.Recipients.UserType = SAPbobsCOM.BoMsgRcpTypes.rt_InternalUser;
                    if (lst.Rows[i]["EMAIL"].ToString() != "" && !string.IsNullOrEmpty(lst.Rows[i]["EMAIL"].ToString()))
                    {
                        Send_Email(lst.Rows[i]["EMAIL"].ToString(), lst.Rows[i]["NAME"].ToString(), E_Body, msg.Subject);
                    }
                    if (i < lst.Rows.Count - 1)
                    {
                        msg.Recipients.Add();
                    }
                }
                msg.Add();
            }
        }

        private DataTable Get_lst_User_Next_LV()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                if (Type != "BILLVP")
                {
                    cmd = new SqlCommand("JV_Get_Lst_Usr_LV", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@DocEntry", DocEntry);
                }
                else
                {
                    cmd = new SqlCommand("VPBILL_Get_Lst_Usr_LV", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@DocEntry", DocEntry);
                }
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            return result;
        }

        private string Get_User_Accountant()
        {
            string result = "";
            string lenh = @"Select U_Usr From [@JV_APROVE_D] where U_Level = '-2' and U_Position = '2' and DocEntry =" + DocEntry;
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(lenh);
            result = oR_RecordSet.Fields.Item("U_Usr").Value.ToString();
            return result;
        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            int new_heigt = this.UIAPIRawForm.ClientHeight;
            int new_width = this.UIAPIRawForm.ClientWidth;

            //Application.SBO_Application.MessageBox(this.UIAPIRawForm.Width.ToString() + "x" + this.UIAPIRawForm.Height.ToString());
            //Resize List JV
            Grid0.Item.Width = 400;
            Grid0.Item.Height = new_heigt - 40;
            //Resize List Apprv
            Grid1.Item.Left = 422;
            Grid1.Item.Height = 130;

            //Resize List Infor
            Grid2.Item.Left = 422;

            //Resize Button
            Button0.Item.Top = new_heigt - 40;
            Button0.Item.Left = 422;
            Button1.Item.Top = new_heigt - 40;
            Button1.Item.Left = 497;
            Button2.Item.Top = new_heigt - 40;
            Button2.Item.Left = 572;
            Button3.Item.Top = new_heigt - 40;
            Button3.Item.Left = 676;

            //Resize option button
            OptionBtn1.Item.Left = OptionBtn0.Item.Left + OptionBtn0.Item.Width + 10;
            OptionBtn2.Item.Left = OptionBtn1.Item.Left + OptionBtn1.Item.Width + 10;
            OptionBtn3.Item.Left = OptionBtn2.Item.Left + OptionBtn2.Item.Width + 10;
            //this.UIAPIRawForm.Refresh();
        }

        System.Data.DataTable Get_Data_HD_VP(string pFProject, string pCGroup, string pPUType, DateTime pToDate)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {

                cmd = new SqlCommand("JV_GET_HDINFO", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFProject);
                cmd.Parameters.AddWithValue("@BatchNum", BatchNum);
                cmd.Parameters.AddWithValue("@CGroup", pCGroup);
                cmd.Parameters.AddWithValue("@PUType", pPUType);
                cmd.Parameters.AddWithValue("@ToDate", pToDate);
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
        
        System.Data.DataTable Get_Data_Cover_BCH()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {

                cmd = new SqlCommand("JV_Get_Data_BCH_Cover", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@BatchNum", BatchNum);
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

        //Get MenuUID Crystal Report
        System.Data.DataTable Get_MenuUID(string pReportName)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("GET_MENUUID", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@ReportName", pReportName);
                cmd.CommandTimeout = 0;
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

        public String convert(decimal total)
        {
            try
            {
                string rs = "";
                total = Math.Abs(Math.Round(total, 0));
                string[] ch = { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };
                string[] rch = { "lẻ", "mốt", "", "", "", "lăm" };
                string[] u = { "", "mươi", "trăm", "ngàn", "", "", "triệu", "", "", "tỷ", "", "", "ngàn", "", "", "triệu" };
                string nstr = total.ToString();
                int[] n = new int[nstr.Length];
                int len = n.Length;
                for (int i = 0; i < len; i++)
                {
                    n[len - 1 - i] = Convert.ToInt32(nstr.Substring(i, 1));
                }
                for (int i = len - 1; i >= 0; i--)
                {
                    if (i % 3 == 2)// số 0 ở hàng trăm
                    {
                        if (n[i] == 0 && n[i - 1] == 0 && n[i - 2] == 0) continue;//nếu cả 3 số là 0 thì bỏ qua không đọc
                    }
                    else if (i % 3 == 1) // số ở hàng chục
                    {
                        if (n[i] == 0)
                        {
                            if (n[i - 1] == 0) { continue; }// nếu hàng chục và hàng đơn vị đều là 0 thì bỏ qua.
                            else
                            {
                                rs += " " + rch[n[i]]; continue;// hàng chục là 0 thì bỏ qua, đọc số hàng đơn vị
                            }
                        }
                        if (n[i] == 1)//nếu số hàng chục là 1 thì đọc là mười
                        {
                            rs += " mười"; continue;
                        }
                    }
                    else if (i != len - 1)// số ở hàng đơn vị (không phải là số đầu tiên)
                    {
                        if (n[i] == 0)// số hàng đơn vị là 0 thì chỉ đọc đơn vị
                        {
                            if (i + 2 <= len - 1 && n[i + 2] == 0 && n[i + 1] == 0) continue;
                            rs += " " + (i % 3 == 0 ? u[i] : u[i % 3]);
                            continue;
                        }
                        if (n[i] == 1)// nếu là 1 thì tùy vào số hàng chục mà đọc: 0,1: một / còn lại: mốt
                        {
                            rs += " " + ((n[i + 1] == 1 || n[i + 1] == 0) ? ch[n[i]] : rch[n[i]]);
                            rs += " " + (i % 3 == 0 ? u[i] : u[i % 3]);
                            continue;
                        }
                        if (n[i] == 5) // cách đọc số 5
                        {
                            if (n[i + 1] != 0) //nếu số hàng chục khác 0 thì đọc số 5 là lăm
                            {
                                rs += " " + rch[n[i]];// đọc số 
                                rs += " " + (i % 3 == 0 ? u[i] : u[i % 3]);// đọc đơn vị
                                continue;
                            }
                        }
                    }
                    rs += (rs == "" ? " " : ", ") + ch[n[i]];// đọc số
                    rs += " " + (i % 3 == 0 ? u[i] : u[i % 3]);// đọc đơn vị
                }
                if (rs[rs.Length - 1] != ' ')
                    rs += " đồng.";
                else
                    rs += "đồng.";
                if (rs.Length > 2)
                {
                    string rs1 = rs.Substring(0, 2);
                    rs1 = rs1.ToUpper();
                    rs = rs.Substring(2);
                    rs = rs1 + rs;
                }
                return rs.Trim().Replace("lẻ,", "lẻ").Replace("mươi,", "mươi").Replace("trăm,", "trăm").Replace("mười,", "mười");
            }
            catch
            {
                return "Error";
            }
        }

        private void Grid0_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            try
            {
                if (Grid0.Rows.SelectedRows.Count == 1)
                {
                    if (Type != "BILLVP")
                    {
                        int BatchNum = 0;
                        int.TryParse(Grid0.DataTable.GetValue("BatchNum", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out BatchNum);
                        if (BatchNum > 0)
                        {
                            oApp.ActivateMenuItem("1541");
                            SAPbouiCOM.Form oActiveFrm = oApp.Forms.ActiveForm;
                            if (oActiveFrm.Type == 229)
                            {
                                SAPbouiCOM.Matrix mtx1 = (SAPbouiCOM.Matrix)oActiveFrm.Items.Item("8").Specific;
                                for (int i = 0; i < mtx1.RowCount; i++)
                                {
                                    string tmp = ((SAPbouiCOM.EditText)mtx1.Columns.Item("1").Cells.Item(i + 1).Specific).Value;
                                    if (tmp == BatchNum.ToString())
                                    {
                                        mtx1.Columns.Item("1").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                        SAPbouiCOM.Matrix mtx2 = (SAPbouiCOM.Matrix)oActiveFrm.Items.Item("7").Specific;
                                        if (mtx2.RowCount >= 1)
                                        {
                                            mtx2.Columns.Item("1").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                oApp.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }

        }

        //Approve Button
        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SqlCommand cmd = null;
            try
            {
                if (Type != "BILLVP")
                {
                    cmd = new SqlCommand("JV_Approve_LV", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                    cmd.Parameters.AddWithValue("@DocEntry", DocEntry);
                    cmd.Parameters.AddWithValue("@Status", "1");
                    cmd.Parameters.AddWithValue("@Comment", EditText0.Value.Trim());
                }
                else
                {
                    cmd = new SqlCommand("VPBILL_Approve_LV", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                    cmd.Parameters.AddWithValue("@DocEntry", DocEntry);
                    cmd.Parameters.AddWithValue("@Status", "1");
                    cmd.Parameters.AddWithValue("@Comment", EditText0.Value.Trim());
                }

                conn.Open();
                int rowupdate = cmd.ExecuteNonQuery();
                if (rowupdate >= 1)
                {
                    oApp.MessageBox("Phê duyệt thành công");
                    //Gui tin nhan
                    Send_Alert();
                    //Update Status neu da qua day du cac buoc duyet
                    if (Check_All_Approve())
                    {
                        SAPbobsCOM.GeneralService oGeneralService = null;
                        SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                        SAPbobsCOM.CompanyService sCmp = null;
                        sCmp = oCompany.GetCompanyService();
                        if (Type != "BILLVP")
                        {
                            oGeneralService = sCmp.GetGeneralService("JVAPPROVE");
                            oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                            oGeneralParams.SetProperty("DocEntry", DocEntry);
                            oGeneralService.Close(oGeneralParams);
                            //Send SMS to infor Ke toan
                            SAPbobsCOM.Messages msg = null;
                            msg = (SAPbobsCOM.Messages)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                            msg.MessageText = string.Format("Journal Voucher số {0} đã được duyệt.{1}Nội dung Comment: {2}.{3}Anh/Chị vui lòng quay lại màn hình Financial -> Journal Voucher để post giao dịch lên hệ thống", BatchNum, Environment.NewLine, EditText0.Value, Environment.NewLine);
                            msg.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;
                            msg.Subject = "Journal Voucher số " + BatchNum.ToString() + " đã được duyệt";
                            msg.Recipients.SetCurrentLine(0);
                            msg.Recipients.UserCode = Get_User_Accountant();
                            msg.Recipients.NameTo = Get_User_Accountant();
                            msg.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                            msg.Add();
                        }
                        else
                        {
                            oGeneralService = sCmp.GetGeneralService("BILLVP");
                            oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                            oGeneralParams.SetProperty("DocEntry", DocEntry);
                            oGeneralService.Close(oGeneralParams);
                            //Send SMS to infor Ke toan
                            SAPbobsCOM.Messages msg = null;
                            msg = (SAPbobsCOM.Messages)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                            msg.MessageText = string.Format("Bill văn phòng số {0} đã được duyệt.{1}Nội dung Comment: {2}.{3}", DocEntry, Environment.NewLine, EditText0.Value, Environment.NewLine);
                            msg.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;
                            msg.Subject = "Bill văn phòng số " + DocEntry.ToString() + " đã được duyệt";
                            msg.Recipients.SetCurrentLine(0);
                            msg.Recipients.UserCode = Get_User_Accountant();
                            msg.Recipients.NameTo = Get_User_Accountant();
                            msg.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                            msg.Add();
                        }
                    }
                }
                else
                {
                    oApp.MessageBox("Phê duyệt không thành công !");
                }
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
        }

        //Reject Button
        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SqlCommand cmd = null;
            try
            {
                if (Type != "BILLVP")
                {
                    cmd = new SqlCommand("JV_Approve_LV", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                    cmd.Parameters.AddWithValue("@DocEntry", DocEntry);
                    cmd.Parameters.AddWithValue("@Status", "2");
                    cmd.Parameters.AddWithValue("@Comment", EditText0.Value.Trim());
                }
                else
                {
                    cmd = new SqlCommand("VPBILL_Approve_LV", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                    cmd.Parameters.AddWithValue("@DocEntry", DocEntry);
                    cmd.Parameters.AddWithValue("@Status", "2");
                    cmd.Parameters.AddWithValue("@Comment", EditText0.Value.Trim());
                }
                conn.Open();
                int rowupdate = cmd.ExecuteNonQuery();
                if (rowupdate >= 1)
                {
                    oApp.MessageBox("Reject thành công");
                    if (Check_GD_DA(oCompany.UserName) || Check_Manager(oCompany.UserName))
                    {
                        SAPbobsCOM.GeneralService oGeneralService = null;
                        SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                        SAPbobsCOM.CompanyService sCmp = null;
                        sCmp = oCompany.GetCompanyService();
                        if (Type != "BILLVP")
                        {
                            oGeneralService = sCmp.GetGeneralService("JVAPPROVE");
                            oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                            oGeneralParams.SetProperty("DocEntry", DocEntry);
                            oGeneralService.Cancel(oGeneralParams);
                        }
                        else
                        {
                            oGeneralService = sCmp.GetGeneralService("BILLVP");
                            oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                            oGeneralParams.SetProperty("DocEntry", DocEntry);
                            oGeneralService.Cancel(oGeneralParams);
                        }
                    }
                    //Send alert to infor creator
                    SAPbobsCOM.Messages msg = null;
                    msg = (SAPbobsCOM.Messages)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                    if (Type != "BILLVP")
                    {
                        msg.MessageText = string.Format(@"Journal Voucher số {0} đã bị từ chối.{1}Nội dung Comment: {2}", BatchNum, Environment.NewLine, EditText0.Value);
                        msg.Subject = "Journal Voucher số " + BatchNum.ToString() + " bị từ chối";
                    }
                    else
                    {
                        msg.MessageText = string.Format(@"Bill văn phòng số {0} đã bị từ chối.{1}Nội dung Comment: {2}", DocEntry, Environment.NewLine, EditText0.Value);
                        msg.Subject = "Bill văn phòng số " + DocEntry.ToString() + " bị từ chối";
                    }
                    msg.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;
                    msg.Recipients.SetCurrentLine(0);
                    msg.Recipients.UserCode = User_Create;
                    msg.Recipients.NameTo = User_Create;
                    msg.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                    msg.Add();
                }
                else
                {
                    oApp.MessageBox("Reject không thành công!");
                }
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
        }

        //Approve with Note for CCM Button
        private void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Type != "BILLVP")
            {
                SqlCommand cmd = null;
                try
                {
                    cmd = new SqlCommand("JV_Approve_LV", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                    cmd.Parameters.AddWithValue("@DocEntry", DocEntry);
                    cmd.Parameters.AddWithValue("@Status", "4");
                    cmd.Parameters.AddWithValue("@Comment", EditText0.Value.Trim());

                    conn.Open();
                    int rowupdate = cmd.ExecuteNonQuery();
                    if (rowupdate >= 1)
                    {
                        oApp.MessageBox("Phê duyệt thành công");
                        //Gui tin nhan cho nguoi tao JV
                        SAPbobsCOM.Messages msg = null;
                        msg = (SAPbobsCOM.Messages)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                        msg.MessageText = string.Format("Journal Voucher số {0} có yêu cầu điều chỉnh thông tin từ phòng CCM.{1}Nội dung Comment: {2}", BatchNum, Environment.NewLine, EditText0.Value);
                        msg.Subject = "Yêu cầu điều chỉnh Journal Voucher số " + BatchNum.ToString();
                        msg.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;
                        msg.Recipients.SetCurrentLine(0);
                        msg.Recipients.UserCode = User_Create;
                        msg.Recipients.NameTo = User_Create;
                        msg.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                        msg.Add();
                        //Reset cấp duyệt đầu tiên
                        if (Type == "BCH")
                        {
                            cmd = new SqlCommand("JV_Reset_Approve_LV1", conn);
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@DocEntry", DocEntry);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        oApp.MessageBox("Phê duyệt không thành công !");
                    }

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
            }
            else
            {
                oApp.MessageBox("Không hỗ trợ cho Bill văn phòng");
            }
        }

        //Print Cover Button
        private void Button3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Type != "BILLVP")
            {
                Microsoft.Office.Interop.Excel.Application oXL;
                Microsoft.Office.Interop.Excel._Workbook oWB;
                Microsoft.Office.Interop.Excel._Worksheet oSheet;
                Microsoft.Office.Interop.Excel.Range oRng_Source;
                Microsoft.Office.Interop.Excel.Range oRng_Dest;
                object misvalue = System.Reflection.Missing.Value;
                //Get DATA
                //int BatchNum = 0;
                //int.TryParse(Grid0.DataTable.GetValue("BatchNum", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out BatchNum);
                string FProject = Grid0.DataTable.GetValue("Project", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                //Type = Grid0.DataTable.GetValue("Type", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                string Period = Grid0.DataTable.GetValue("Period", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                string SToDate = Grid0.DataTable.GetValue("Create Date", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                DateTime Todate = DateTime.Parse(SToDate);
                try
                {
                    //Print Cover BCH
                    if (Type == "BCH")
                    {
                        #region EXCEL
                        //DataTable tmp = Get_Data_Cover_BCH();
                        ////Start Excel and get Application object.
                        //oXL = new Microsoft.Office.Interop.Excel.Application();
                        //oXL.Visible = true;
                        ////Open Template
                        //oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_BCH.xlsx");
                        //oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                        //oSheet.Cells[1, 4] = "P/B/BP/DA: " + FProject;
                        //oSheet.Cells[2, 4] = "Ngày: " + DateTime.Now.ToString("dd/MM/yyyy");
                        //oSheet.Cells[3, 4] = "Số: " + Period;
                        //oSheet.Cells[5, 1] = "Kỳ: " + Period;
                        //oSheet.Cells[6, 1] = "Yêu cầu số: " + BatchNum;

                        //if (tmp.Rows.Count == 1)
                        //{
                        //    oSheet.Cells[8, 6] = tmp.Rows[0]["TONGGT"].ToString();
                        //    oSheet.Cells[9, 6] = tmp.Rows[0]["KYTRUOC"].ToString();
                        //    oSheet.Cells[10, 6] = tmp.Rows[0]["KYNAY"].ToString();
                        //    decimal KN = 0;
                        //    decimal.TryParse(tmp.Rows[0]["KYNAY"].ToString(), out KN);
                        //    oSheet.Cells[11, 1] = "Bằng chữ: " + convert(KN);

                        //    string Comment = "";
                        //    DateTime t_date_appr = DateTime.Today;
                        //    for (int t = 0; t < Grid1.Rows.Count; t++)
                        //    {
                        //        string t_date = Grid1.DataTable.GetValue("Approved on", t).ToString();
                        //        if (!string.IsNullOrEmpty(t_date))
                        //            t_date_appr = DateTime.ParseExact(t_date, "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture);
                        //        string t_name = Grid1.DataTable.GetValue("Approved by", t).ToString();
                        //        string t_status = Grid1.DataTable.GetValue("Status", t).ToString();
                        //        string t_comm = Grid1.DataTable.GetValue("Comment", t).ToString();
                        //        string t_dept = Grid1.DataTable.GetValue("DeptName", t).ToString();
                        //        string t_position = Grid1.DataTable.GetValue("Position", t).ToString();
                        //        string t_deptcode = Grid1.DataTable.GetValue("DeptCode", t).ToString();
                        //        string t_poscode = Grid1.DataTable.GetValue("PosCode", t).ToString();
                        //        if (t_deptcode == "3") // t_dept == "Dự Án")
                        //        {
                        //            oSheet.Cells[18, 1] = t_status;
                        //            if (t_status == "Rejected")
                        //                oSheet.Cells[18, 1].Font.Color = System.Drawing.Color.FromArgb(255, 0, 0);
                        //            oSheet.Cells[19, 1] = t_name;
                        //            if (!string.IsNullOrEmpty(t_date))
                        //                oSheet.Cells[20, 1] = "'" + t_date_appr.ToString("dd/MM/yyyy HH:mm:ss");
                        //            if (!string.IsNullOrEmpty(t_comm))
                        //            {
                        //                Comment += string.Format("{0}: {1}{2}", t_name, t_comm, Environment.NewLine);
                        //            }
                        //        }
                        //        else if (t_deptcode == "3" && t_poscode == "1")//(t_dept == "CCM" && t_position == "Trưởng phòng")
                        //        {
                        //            oSheet.Cells[18, 3] = t_status;
                        //            if (t_status == "Rejected")
                        //                oSheet.Cells[18, 3].Font.Color = System.Drawing.Color.FromArgb(255, 0, 0);
                        //            oSheet.Cells[19, 3] = t_name;
                        //            if (!string.IsNullOrEmpty(t_date))
                        //                oSheet.Cells[20, 3] = "'" + t_date_appr.ToString("dd/MM/yyyy HH:mm:ss");
                        //            if (!string.IsNullOrEmpty(t_comm))
                        //            {

                        //                Comment += string.Format("{0}: {1}{2}", t_name, t_comm, Environment.NewLine);
                        //            }
                        //        }
                        //        else if (t_deptcode == "-2" && t_poscode == "1")//(t_dept == "Kế toán" && t_position == "Trưởng phòng")
                        //        {
                        //            oSheet.Cells[18, 5] = t_status;
                        //            if (t_status == "Rejected")
                        //                oSheet.Cells[18, 5].Font.Color = System.Drawing.Color.FromArgb(255, 0, 0);
                        //            oSheet.Cells[19, 5] = t_name;
                        //            if (!string.IsNullOrEmpty(t_date))
                        //                oSheet.Cells[20, 5] = "'" + t_date_appr.ToString("dd/MM/yyyy HH:mm:ss");
                        //            if (!string.IsNullOrEmpty(t_comm))
                        //            {
                        //                Comment += string.Format("{0}: {1}{2}", t_name, t_comm, Environment.NewLine);
                        //            }
                        //        }
                        //        else if (t_poscode == "3")//(t_position == "Giám đốc dự án")
                        //        {
                        //            oSheet.Cells[18, 7] = t_status;
                        //            if (t_status == "Rejected")
                        //                oSheet.Cells[18, 7].Font.Color = System.Drawing.Color.FromArgb(255, 0, 0);
                        //            oSheet.Cells[19, 7] = t_name;
                        //            if (!string.IsNullOrEmpty(t_date))
                        //                oSheet.Cells[20, 7] = "'" + t_date_appr.ToString("dd/MM/yyyy HH:mm:ss");
                        //            if (!string.IsNullOrEmpty(t_comm))
                        //            {
                        //                Comment += string.Format("{0}: {1}{2}", t_name, t_comm, Environment.NewLine);
                        //            }
                        //        }
                        //    }
                        //    oSheet.Cells[12, 2] = Comment.Substring(0, Comment.Length - 2);              
                        //}
                    #endregion

                        #region Crystal Report
                        DataTable rs = Get_MenuUID("Cover_CPVP_BCH");
                        if (rs.Rows.Count > 0)
                        {
                            oApp.ActivateMenuItem(rs.Rows[0]["MenuUID"].ToString());
                            SAPbouiCOM.Form act_frm = oApp.Forms.ActiveForm;
                            ((SAPbouiCOM.EditText)act_frm.Items.Item("1000003").Specific).Value = BatchNum.ToString();
                            act_frm.Items.Item("1").Click();
                        }
                        #endregion
                    }
                    else if (Type == "VP")
                    {
                        #region Excel
                //        DataTable tmp = Get_Data_Cover_BCH();
                //        if (tmp.Rows.Count == 1)
                //        {
                //            string VP_TYPE = tmp.Rows[0]["VP_TYPE"].ToString();
                //            if (VP_TYPE == "TTBT")
                //            {
                //                //Start Excel and get Application object.
                //                oXL = new Microsoft.Office.Interop.Excel.Application();
                //                oXL.Visible = true;
                //                //Open Template
                //                oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_VP2.xlsx");
                //                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //                //oSheet.Cells[1, 4] = "P/B/BP/DA: " + FProject;
                //                oSheet.Cells[1, 4] = "P/B/BP/DA: " + tmp.Rows[0]["PBTT"].ToString();
                //                oSheet.Cells[2, 4] = "Ngày: " + DateTime.Now.ToString("dd/MM/yyyy");
                //                oSheet.Cells[3, 4] = "Số: " + Period;
                //                oSheet.Cells[5, 1] = "Kỳ: " + Period;
                //                oSheet.Cells[6, 1] = "Yêu cầu số: " + BatchNum;

                //                oSheet.Cells[8, 6] = tmp.Rows[0]["TONGGT"].ToString();
                //                oSheet.Cells[9, 6] = tmp.Rows[0]["KYTRUOC"].ToString();
                //                oSheet.Cells[10, 6] = tmp.Rows[0]["KYNAY"].ToString();
                //                decimal KN = 0;
                //                decimal.TryParse(tmp.Rows[0]["KYNAY"].ToString(), out KN);
                //                oSheet.Cells[11, 1] = "Bằng chữ: " + convert(KN);

                //                string Comment = "";
                //                DateTime t_date_appr = DateTime.Today;
                //                for (int t = 0; t < Grid1.Rows.Count; t++)
                //                {
                //                    string t_date = Grid1.DataTable.GetValue("Approved on", t).ToString();
                //                    if (!string.IsNullOrEmpty(t_date))
                //                        t_date_appr = DateTime.ParseExact(t_date, "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture);
                //                    string t_name = Grid1.DataTable.GetValue("Approved by", t).ToString();
                //                    string t_status = Grid1.DataTable.GetValue("Status", t).ToString();
                //                    string t_comm = Grid1.DataTable.GetValue("Comment", t).ToString();
                //                    string t_dept = Grid1.DataTable.GetValue("DeptName", t).ToString();
                //                    string t_position = Grid1.DataTable.GetValue("Position", t).ToString();
                //                    string t_deptcode = Grid1.DataTable.GetValue("DeptCode", t).ToString();
                //                    string t_poscode = Grid1.DataTable.GetValue("PosCode", t).ToString();
                //                    if (t_poscode == "1" && t == 1)
                //                    {
                //                        oSheet.Cells[18, 1] = t_status;
                //                        if (t_status == "Rejected")
                //                            oSheet.Cells[18, 1].Font.Color = System.Drawing.Color.FromArgb(255, 0, 0);
                //                        oSheet.Cells[19, 1] = t_name;
                //                        if (!string.IsNullOrEmpty(t_date))
                //                            oSheet.Cells[20, 1] = "'" + t_date_appr.ToString("dd/MM/yyyy HH:mm:ss");
                //                        if (!string.IsNullOrEmpty(t_comm))
                //                        {
                //                            Comment += string.Format("{0}: {1}{2}", t_name, t_comm, Environment.NewLine);
                //                        }
                //                    }
                //                    else if (t_deptcode == "1" && t_poscode == "1" && t == 3)
                //                    {
                //                        oSheet.Cells[18, 3] = t_status;
                //                        if (t_status == "Rejected")
                //                            oSheet.Cells[18, 3].Font.Color = System.Drawing.Color.FromArgb(255, 0, 0);
                //                        oSheet.Cells[19, 3] = t_name;
                //                        if (!string.IsNullOrEmpty(t_date))
                //                            oSheet.Cells[20, 3] = "'" + t_date_appr.ToString("dd/MM/yyyy HH:mm:ss");
                //                        if (!string.IsNullOrEmpty(t_comm))
                //                        {

                //                            Comment += string.Format("{0}: {1}{2}", t_name, t_comm, Environment.NewLine);
                //                        }
                //                    }
                //                    else if (t_deptcode == "-2" && t_poscode == "1")
                //                    {
                //                        oSheet.Cells[18, 5] = t_status;
                //                        if (t_status == "Rejected")
                //                            oSheet.Cells[18, 5].Font.Color = System.Drawing.Color.FromArgb(255, 0, 0);
                //                        oSheet.Cells[19, 5] = t_name;
                //                        if (!string.IsNullOrEmpty(t_date))
                //                            oSheet.Cells[20, 5] = "'" + t_date_appr.ToString("dd/MM/yyyy HH:mm:ss");
                //                        if (!string.IsNullOrEmpty(t_comm))
                //                        {
                //                            Comment += string.Format("{0}: {1}{2}", t_name, t_comm, Environment.NewLine);
                //                        }
                //                    }
                //                    else if (t_poscode == "4")
                //                    {
                //                        oSheet.Cells[18, 7] = t_status;
                //                        if (t_status == "Rejected")
                //                            oSheet.Cells[18, 7].Font.Color = System.Drawing.Color.FromArgb(255, 0, 0);
                //                        oSheet.Cells[19, 7] = t_name;
                //                        if (!string.IsNullOrEmpty(t_date))
                //                            oSheet.Cells[20, 7] = "'" + t_date_appr.ToString("dd/MM/yyyy HH:mm:ss");
                //                        if (!string.IsNullOrEmpty(t_comm))
                //                        {
                //                            Comment += string.Format("{0}: {1}{2}", t_name, t_comm, Environment.NewLine);
                //                        }
                //                    }
                //                }
                //                oSheet.Cells[12, 2] = Comment.Substring(0, Comment.Length - 2);
                //            }
                //            else if (VP_TYPE == "TTBN")
                //            {
                //                decimal GTHD = 0, PLT = 0, PLTT = 0;
                //                string SoHD = "";
                //                string NgayHD = "";
                //                string BpCode = "";
                //                string BpName = "";
                //                //Get Data HD
                //                DataTable tmp_hd = Get_Data_HD_VP(FProject, "", "", Todate);
                //                if (tmp_hd.Rows.Count >= 1)
                //                {
                //                    BpCode = tmp_hd.Rows[0]["BPCode"].ToString();
                //                    BpName = tmp_hd.Rows[0]["BPName"].ToString();
                //                    foreach (DataRow r in tmp_hd.Rows)
                //                    {
                //                        decimal tmp1 = 0;
                //                        decimal.TryParse(r["GTHD"].ToString(), out tmp1);
                //                        if (r["Type"].ToString() == "HD")
                //                        {
                //                            GTHD += tmp1;
                //                            SoHD = r["Number"].ToString();
                //                            NgayHD = r["StartDate"].ToString();
                //                            //HD_Descript = r["Descript"].ToString();
                //                        }
                //                        else if (r["Type"].ToString() == "PLT")
                //                            PLT += tmp1;
                //                        else if (r["Type"].ToString() == "PLTT")
                //                        {
                //                            PLTT += tmp1;
                //                            SoHD = r["U_SHD"].ToString();
                //                            NgayHD = r["StartDate"].ToString();
                //                            //HD_Descript = r["Descript"].ToString();
                //                        }
                //                    }
                //                }
                //                //Start Excel and get Application object.
                //                oXL = new Microsoft.Office.Interop.Excel.Application();
                //                oXL.Visible = true;
                //                //Open Template
                //                oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_VP1.xlsx");
                //                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //                oSheet.Cells[1, 4] = "P/B/BP/DA: " + FProject;
                //                oSheet.Cells[2, 4] = "Ngày: " + DateTime.Now.ToString("dd/MM/yyyy");
                //                oSheet.Cells[3, 4] = "Số: " + Period;
                //                oSheet.Cells[5, 1] = "Nhà cung cấp: " + BpCode + " - " + BpName;
                //                oSheet.Cells[6, 1] = "Số hợp đồng: " + SoHD;
                //                oSheet.Cells[6, 6] = "Ngày: " + NgayHD;

                //                oSheet.Cells[8, 6] = (GTHD + PLTT).ToString();
                //                oSheet.Cells[9, 6] = PLT.ToString();
                //                oSheet.Cells[10, 6] = (GTHD + PLTT + PLT).ToString();


                //                oSheet.Cells[12, 6] = tmp.Rows[0]["TONGGT"].ToString();
                //                oSheet.Cells[13, 6] = tmp.Rows[0]["KYTRUOC"].ToString();
                //                oSheet.Cells[14, 6] = tmp.Rows[0]["KYNAY"].ToString();
                //                decimal KN = 0;
                //                decimal.TryParse(tmp.Rows[0]["KYNAY"].ToString(), out KN);
                //                oSheet.Cells[15, 2] = "Bằng chữ: " + convert(KN);

                //            }
                //        }
                #endregion

                        #region Crystal Report
                        DataTable rs = Get_MenuUID("Cover_CPVP_BCH");
                        if (rs.Rows.Count > 0)
                        {
                            oApp.ActivateMenuItem(rs.Rows[0]["MenuUID"].ToString());
                            SAPbouiCOM.Form act_frm = oApp.Forms.ActiveForm;
                            ((SAPbouiCOM.EditText)act_frm.Items.Item("1000003").Specific).Value = BatchNum.ToString();
                            act_frm.Items.Item("1").Click();
                        }
                        #endregion
                    }
                }
                catch
                {

                }
                finally
                {

                }
                        
            }
            else
            {
                DataTable rs = Get_MenuUID("VPBILL_Cover");
                if (rs.Rows.Count > 0)
                {
                    oApp.ActivateMenuItem(rs.Rows[0]["MenuUID"].ToString());
                    SAPbouiCOM.Form act_frm = oApp.Forms.ActiveForm;
                    ((SAPbouiCOM.EditText)act_frm.Items.Item("1000003").Specific).Value = DocEntry.ToString();                    
                    act_frm.Items.Item("1").Click();
                }
            }

        }

        private void Form_CloseBefore(SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                oCompany.Disconnect();
            }
            catch
            { }
        }

        //Show Current
        private void OptionBtn0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn0.Selected)
            {
                Load_Grid_Period();
                Button0.Item.Enabled = true;
                Button1.Item.Enabled = true;
                Button2.Item.Enabled = true;
            }
        }

        //Show Approved
        private void OptionBtn1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn1.Selected)
            {
                Load_Grid_Approved_Period();
                Button0.Item.Enabled = false;
                Button1.Item.Enabled = false;
                Button2.Item.Enabled = false;
            }
        }

        //Show Rejected
        private void OptionBtn2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn2.Selected)
            {
                Load_Grid_Rejected_Period();
                Button0.Item.Enabled = false;
                Button1.Item.Enabled = false;
                Button2.Item.Enabled = false;
            }
        }

        //Show All
        private void OptionBtn3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn3.Selected)
            {
                Load_Grid_Period_All();
                Button0.Item.Enabled = false;
                Button1.Item.Enabled = false;
                Button2.Item.Enabled = false;
            }
        }


    }
}