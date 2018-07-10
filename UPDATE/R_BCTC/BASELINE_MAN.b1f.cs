using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;
using System.Runtime.InteropServices;
using System.Net.Mail;

namespace R_BCTC
{
    [FormAttribute("R_BCTC.BASELINE_MAN", "BASELINE_MAN.b1f")]
    class BASELINE_MAN : UserFormBase
    {
        public BASELINE_MAN()
        {
        }
        int DocEntry_BaseLine = 0;
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
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
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("grd_bl").Specific));
            this.Grid0.PressedAfter += new SAPbouiCOM._IGridEvents_PressedAfterEventHandler(this.Grid0_PressedAfter);
            this.Grid1 = ((SAPbouiCOM.Grid)(this.GetItem("grd_app").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_appr").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("bt_rej").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("bt_ceo").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("Item_11").Specific));
            this.Grid3 = ((SAPbouiCOM.Grid)(this.GetItem("grd_subp").Specific));
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("bt_ce").Specific));
            this.Button3.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button3_PressedAfter);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("bt_fi").Specific));
            this.Button4.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button4_PressedAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_15").Specific));
            this.OptionBtn0 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_sc").Specific));
            this.OptionBtn0.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn0_PressedAfter);
            this.OptionBtn1 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_sap").Specific));
            this.OptionBtn1.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn1_PressedAfter);
            this.OptionBtn2 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_sr").Specific));
            this.OptionBtn2.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn2_PressedAfter);
            this.OptionBtn3 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_sha").Specific));
            this.OptionBtn3.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn3_PressedAfter);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_note").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
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
            OptionBtn1.GroupWith("op_sc");
            OptionBtn2.GroupWith("op_sc");
            OptionBtn3.GroupWith("op_sc");
            OptionBtn0.Selected = true;

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
        
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Grid Grid1;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.Grid Grid3;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.Button Button4;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.OptionBtn OptionBtn0;
        private SAPbouiCOM.OptionBtn OptionBtn1;
        private SAPbouiCOM.OptionBtn OptionBtn2;
        private SAPbouiCOM.OptionBtn OptionBtn3;
        private SAPbouiCOM.EditText EditText0;

        private void Load_Grid_Period()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_GetList_Current", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Usr", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();
                //Disable all grid column
                foreach (SAPbouiCOM.GridColumn c in Grid0.Columns)
                {
                    c.Editable = false;
                }
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

        private void Load_Grid_Approved_Period()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_GetList_Approved_Current", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Usr", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();
                //Disable all grid column
                foreach (SAPbouiCOM.GridColumn c in Grid0.Columns)
                {
                    c.Editable = false;
                }
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

        private void Load_Grid_Rejected_Period()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_GetList_Rejected_Current", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Usr", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();
                //Disable all grid column
                foreach (SAPbouiCOM.GridColumn c in Grid0.Columns)
                {
                    c.Editable = false;
                }
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

        private void Load_Grid_All()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_GetList", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Usr", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();
                //Disable all grid column
                foreach (SAPbouiCOM.GridColumn c in Grid0.Columns)
                {
                    c.Editable = false;
                }
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

        private SAPbouiCOM.DataTable Convert_SAP_DataTable(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_BASELINEList"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_BASELINEList");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_BASELINEList");
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
                    if (c.DataType.ToString() == "System.DateTime")
                        oDT.SetValue(c.ColumnName, oDT.Rows.Count - 1, r[c.ColumnName]);
                    else
                        oDT.SetValue(c.ColumnName, oDT.Rows.Count - 1, r[c.ColumnName].ToString());
                }
            }

            return oDT;
        }

        //Show Current
        private void OptionBtn0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn0.Selected)
            {
                Button0.Item.Enabled = true;
                Button1.Item.Enabled = true;
                Button2.Item.Enabled = true;
                Load_Grid_Period();
            }
        }
        
        //Show Approved
        private void OptionBtn1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn1.Selected)
            {
                Button0.Item.Enabled = false;
                Button1.Item.Enabled = false;
                Button2.Item.Enabled = false;
                Load_Grid_Approved_Period();
            }
        }
        
        //Show Rejected
        private void OptionBtn2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn2.Selected)
            {
                Button0.Item.Enabled = false;
                Button1.Item.Enabled = false;
                Button2.Item.Enabled = false;
                Load_Grid_Rejected_Period();
            }
        }

        //Show All
        private void OptionBtn3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn3.Selected)
            {
                Button0.Item.Enabled = false;
                Button1.Item.Enabled = false;
                Button2.Item.Enabled = false;
                Load_Grid_All();
            }
        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                int new_heigt = this.UIAPIRawForm.ClientHeight;
                int new_width = this.UIAPIRawForm.ClientWidth;

                this.Grid3.Item.Left = this.StaticText1.Item.Left;
                this.Grid3.Item.Width = this.StaticText1.Item.Width;
                this.Grid1.Item.Top = this.Grid0.Item.Top + this.Grid0.Item.Height + 15;
                this.Grid1.Item.Height = new_heigt - this.Grid1.Item.Top - 10;
            }
            catch
            { }

        }

        private void Grid0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            DocEntry_BaseLine = 0;
            if (Grid0.Rows.SelectedRows.Count == 1)
            {
                try
                {
                    int.TryParse(Grid0.DataTable.GetValue("DocEntry", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out DocEntry_BaseLine);
                    Grid1.DataTable.ExecuteQuery("Select U_DeptName as 'Department', U_PosName as 'Position', U_Status as 'Status', U_Time as 'Approved on', U_Usr as 'Approved by', U_Comment as 'Comment' from [@BASELINE_APPR] where DocEntry=" + DocEntry_BaseLine);
                    Grid1.AutoResizeColumns();

                    Grid3.DataTable.ExecuteQuery("Select 'N' as 'Checked',AbsEntry,[NAME] as 'SubProject Name' from BASELINE_OPMG Where DocEntry_BaseLine=" + DocEntry_BaseLine);
                    Grid3.Columns.Item(1).Editable = false;
                    Grid3.Columns.Item(2).Editable = false;
                    Grid3.Columns.Item("Checked").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    Grid3.AutoResizeColumns();
                }
                catch
                {
 
                }
            }

        }

        //System.Data.DataTable Get_Data_BCDTA(string pFinancialProject)
        //{
        //    DataTable result = new DataTable();
        //    SqlCommand cmd = null;
        //    try
        //    {
        //        cmd = new SqlCommand("BASELINE_GET_DATA_BCDT_A", conn);
        //        cmd.CommandType = CommandType.StoredProcedure;
        //        cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
        //        conn.Open();
        //        SqlDataReader rd = cmd.ExecuteReader();
        //        result.Load(rd);
        //    }
        //    catch (Exception ex)
        //    {
        //        oApp.MessageBox(ex.Message);
        //    }
        //    finally
        //    {
        //        conn.Close();
        //        cmd.Dispose();
        //    }
        //    return result;
        //}

        System.Data.DataTable Get_Data_BCDTA(string pGoithauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_MM_CE_GET_DATA_A", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry_BaseLine", DocEntry_BaseLine);
                cmd.Parameters.AddWithValue("@Goithau_Key", pGoithauKey);
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

        System.Data.DataTable Get_Data_DUTRU( string pGoiThauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_MM_CE_GETDATA_DETAILS", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry_BaseLine", DocEntry_BaseLine);
                cmd.Parameters.AddWithValue("@GoiThauKey", pGoiThauKey);
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

        System.Data.DataTable Get_Data_DUTRU_SUM(string pGoiThauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_MM_CE_GETDATA_SUM", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry_BaseLine", DocEntry_BaseLine);
                cmd.Parameters.AddWithValue("@GoiThauKey", pGoiThauKey);
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

        System.Data.DataTable Get_Data_BCH_CE(string pGoiThauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_MM_CE_GET_DATA_BCH", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry_BaseLine", DocEntry_BaseLine);
                cmd.Parameters.AddWithValue("@GoiThauKey", pGoiThauKey);
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

        System.Data.DataTable Get_Data_BCH_FI(string pGoiThauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_MM_FI_GET_DATA_BCH", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry_BaseLine", DocEntry_BaseLine);
                cmd.Parameters.AddWithValue("@GoiThauKey", pGoiThauKey);
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

        System.Data.DataTable Get_Data_FI_VII( string pGoiThauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_MM_FI_GET_DATA_VII", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry_BaseLine", DocEntry_BaseLine);
                cmd.Parameters.AddWithValue("@GoiThauKey", pGoiThauKey);
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

        System.Data.DataTable Get_Prj_Info(string pFinancialProject)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand(string.Format("Select * from BASELINE_OPMG a where a.DocEntry_BaseLine='{0}' and a.STATUS <> 'T'", DocEntry_BaseLine), conn);
                cmd.CommandType = CommandType.Text;
                //cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
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

        //View CE Report Button
        private void Button3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string GoiThau_Key = "";
            string GoiThau_Name = "";
            for (int i = 0; i < Grid3.Rows.Count; i++)
            {
                string IsSelected = Grid3.DataTable.GetValue("Checked", i).ToString();
                if (IsSelected == "Y")
                {
                    GoiThau_Key += Grid3.DataTable.GetValue("AbsEntry", i).ToString() + ",";
                    GoiThau_Name = Grid3.DataTable.GetValue("SubProject Name", i).ToString();
                }
            }
            if (GoiThau_Key.Length > 0)
                GoiThau_Key = GoiThau_Key.Substring(0, GoiThau_Key.Length - 1);

            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            //Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;

            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            //Open Template
            oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_BCDT.xlsx");

            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            try
            {
                //Fill Header
                string PrjName = Grid0.DataTable.GetValue("Project Name", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                string FProject = Grid0.DataTable.GetValue("Financial Project", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                DateTime BaseLine_Date = DateTime.Parse(Grid0.DataTable.GetValue("BaseLine Date", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString());
                //Project Name
                oSheet.Cells[2, 3] = "Dự án: " + PrjName;
                //Subproject Name
                if (!string.IsNullOrEmpty(GoiThau_Key))
                {
                    if (GoiThau_Key.Split(',').Count() == 1)
                        oSheet.Cells[3, 3] = "Gói thầu: " + GoiThau_Name;
                }
                //oSheet.Cells[3, 3] = "Gói thầu: " + this.ComboBox1.Selected.Description;
                //Thang
                oSheet.Cells[4, 3] = "Tháng: " + BaseLine_Date.ToString("MM-yyyy");//this.ComboBox2.Selected.Value;

                DataTable A = Get_Data_BCDTA(GoiThau_Key);
                List<int> Group_No_RowNum = new List<int>();
                List<int> Section_RowNum = new List<int>();
                decimal sum_tb = 0, sum_dp2 = 0;
                //A- Doanh thu (truoc VAT)
                //Gia tri hop dong
                oSheet.Cells[7, 1] = "1";
                oSheet.Cells[7, 2] = "Giá trị hợp đồng";
                oSheet.Cells[7, 4].Value2 = A.Rows[0]["GTHD"];

                //Gia tri hop dong 1A ( truong hop CDT gui chi phi)
                oSheet.Cells[8, 1] = "1A";
                oSheet.Cells[8, 2] = "Giá trị hợp đồng 1A";
                oSheet.Cells[8, 4].Value2 = A.Rows[0]["KHAC"];

                //Phụ lục HĐ
                oSheet.Cells[9, 1] = "2";
                oSheet.Cells[9, 2] = "Phụ lục HĐ";
                oSheet.Cells[9, 4].Value2 = A.Rows[0]["PLHD"];

                //Giảm giá thương mại
                oSheet.Cells[10, 1] = "3";
                oSheet.Cells[10, 2] = "Giảm giá thương mại";
                oSheet.Cells[10, 4].Value2 = A.Rows[0]["GGTM"];

                //Giảm giá thương mại
                oSheet.Cells[11, 1] = "4";
                oSheet.Cells[11, 2] = "Phương án đề xuất tiết kiệm chi phí";
                oSheet.Cells[11, 4].Value2 = A.Rows[0]["PA"];

                //Phí quản lý
                oSheet.Cells[12, 1] = "5";
                oSheet.Cells[12, 2] = "Phí Quản lý";
                oSheet.Cells[12, 4].Value2 = A.Rows[0]["PhiQL"];
                //Total
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[6, 4]).Formula = string.Format("=SUM({0}:{1})", "D7", "D12");
                int current_rownum = 13;
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 1] = "B";
                oSheet.Cells[current_rownum, 2] = "CHI PHÍ (Trước VAT)";
                current_rownum++;
                //I - CÔNG TÁC THI CÔNG TRỰC TIẾP PHẦN XÂY DỰNG
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 1] = "I";
                oSheet.Cells[current_rownum, 2] = "CÔNG TÁC THI CÔNG TRỰC TIẾP PHẦN XÂY DỰNG";
                Section_RowNum.Add(current_rownum);
                current_rownum++;
                //LOAD DU TRU
                DataTable B = null;
                DataTable C = null;
                try
                {
                    //int.TryParse(this.ComboBox1.Selected.Value.ToString(), out GoithauKey);
                }
                catch
                {

                }
                if (GoiThau_Key == "")
                {
                    B = Get_Data_DUTRU_SUM();
                    C = Get_Data_DUTRU();
                }
                else
                {
                    B = Get_Data_DUTRU_SUM(GoiThau_Key);
                    C = Get_Data_DUTRU(GoiThau_Key);
                }
                int STT_GROUP = 1;

                foreach (DataRow r in B.Rows)
                {
                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                    oSheet.Cells[current_rownum, 1] = STT_GROUP;
                    oSheet.Cells[current_rownum, 2] = r["U_SubProjectDesc"].ToString();
                    oSheet.Cells[current_rownum, 4] = r["TTHD"];
                    Group_No_RowNum.Add(current_rownum);
                    current_rownum++;
                    #region Detail CT
                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Italic = true;
                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    oSheet.Cells[current_rownum, 2] = "Nhà cung cấp";
                    oSheet.Cells[current_rownum, 5].Value2 = r["U_CP_NCC"];
                    int detail_rownum = 0;
                    detail_rownum = current_rownum + 1;
                    current_rownum++;
                    foreach (DataRow rd in C.Select("U_DTT_LineID=" + r["LineID"]))
                    {
                        if (decimal.Parse(rd["U_CP_NCC"].ToString()) != 0)
                        {
                            oSheet.Cells[current_rownum, 2] = rd["U_BPNAME"].ToString();
                            oSheet.Cells[current_rownum, 6].Value2 = rd["U_CP_NCC"];
                            oSheet.Cells[current_rownum, 7].Value2 = rd["KL_TT_DD"];
                            oSheet.Cells[current_rownum, 9].Value2 = rd["CTQL"];
                            current_rownum++;
                        }
                    }
                    if (current_rownum - detail_rownum >= 1)
                    {
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 6]).Formula = string.Format("=SUBTOTAL(9,{0})", "F" + detail_rownum + ":F" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + detail_rownum + ":G" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{1}", detail_rownum, detail_rownum);
                    }

                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Italic = true;
                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    oSheet.Cells[current_rownum, 2] = "Nhà thầu phụ";
                    oSheet.Cells[current_rownum, 5].Value2 = r["U_CP_NTP"];
                    detail_rownum = current_rownum + 1;
                    current_rownum++;
                    foreach (DataRow rd in C.Select(string.Format("U_DTT_LineID={0} " ,r["LineID"])))
                    {
                        if (decimal.Parse(rd["U_CP_NTP"].ToString()) != 0)
                        {
                            oSheet.Cells[current_rownum, 2] = rd["U_BPNAME"].ToString();
                            oSheet.Cells[current_rownum, 6].Value2 = rd["U_CP_NTP"];
                            oSheet.Cells[current_rownum, 7].Value2 = rd["KL_TT_DD"];
                            oSheet.Cells[current_rownum, 9].Value2 = rd["CTQL"];
                            current_rownum++;
                        }
                    }
                    if (current_rownum - detail_rownum >= 1)
                    {
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 6]).Formula = string.Format("=SUBTOTAL(9,{0})", "F" + detail_rownum + ":F" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + detail_rownum + ":G" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{1}", detail_rownum, detail_rownum);
                    }

                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Italic = true;
                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    oSheet.Cells[current_rownum, 2] = "Đội thi công";
                    oSheet.Cells[current_rownum, 5].Value2 = r["U_CP_DTC"];
                    detail_rownum = current_rownum + 1;
                    current_rownum++;
                    foreach (DataRow rd in C.Select("U_DTT_LineID=" + r["LineID"]))
                    {
                        if (decimal.Parse(rd["U_CP_DTC"].ToString()) != 0)
                        {
                            oSheet.Cells[current_rownum, 2] = rd["U_BPNAME"].ToString();
                            oSheet.Cells[current_rownum, 6].Value2 = rd["U_CP_DTC"];
                            oSheet.Cells[current_rownum, 7].Value2 = rd["KL_TT_DD"];
                            oSheet.Cells[current_rownum, 9].Value2 = rd["CTQL"];
                            current_rownum++;
                        }
                    }
                    if (current_rownum - detail_rownum >= 1)
                    {
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 6]).Formula = string.Format("=SUBTOTAL(9,{0})", "F" + detail_rownum + ":F" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + detail_rownum + ":G" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{1}", detail_rownum, detail_rownum);
                    }

                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Italic = true;
                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    oSheet.Cells[current_rownum, 2] = "Vật tư phụ";
                    oSheet.Cells[current_rownum, 5].Value2 = r["U_CP_VTP"];
                    detail_rownum = current_rownum + 1;
                    current_rownum++;
                    foreach (DataRow rd in C.Select("U_DTT_LineID=" + r["LineID"]))
                    {
                        if (decimal.Parse(rd["U_CP_VTP"].ToString()) != 0)
                        {
                            oSheet.Cells[current_rownum, 2] = rd["U_BPNAME"].ToString();
                            oSheet.Cells[current_rownum, 6].Value2 = rd["U_CP_VTP"];
                            oSheet.Cells[current_rownum, 7].Value2 = rd["KL_TT_DD"];
                            oSheet.Cells[current_rownum, 9].Value2 = rd["CTQL"];
                            current_rownum++;
                        }
                    }
                    if (current_rownum - detail_rownum >= 1)
                    {
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 6]).Formula = string.Format("=SUBTOTAL(9,{0})", "F" + detail_rownum + ":F" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + detail_rownum + ":G" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{1}", detail_rownum, detail_rownum);
                    }

                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Italic = true;
                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    oSheet.Cells[current_rownum, 2] = "Vận chuyển";
                    oSheet.Cells[current_rownum, 5].Value2 = r["U_CP_VC"];
                    detail_rownum = current_rownum + 1;
                    current_rownum++;
                    foreach (DataRow rd in C.Select("U_DTT_LineID=" + r["LineID"]))
                    {
                        if (decimal.Parse(rd["U_CP_VC"].ToString()) != 0)
                        {
                            oSheet.Cells[current_rownum, 2] = rd["U_BPNAME"].ToString();
                            oSheet.Cells[current_rownum, 6].Value2 = rd["U_CP_VC"];
                            oSheet.Cells[current_rownum, 7].Value2 = rd["KL_TT_DD"];
                            oSheet.Cells[current_rownum, 9].Value2 = rd["CTQL"];
                            current_rownum++;
                        }
                    }
                    if (current_rownum - detail_rownum >= 1)
                    {
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 6]).Formula = string.Format("=SUBTOTAL(9,{0})", "F" + detail_rownum + ":F" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + detail_rownum + ":G" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{1}", detail_rownum, detail_rownum);
                    }

                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Italic = true;
                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    oSheet.Cells[current_rownum, 2] = "Công nhật";
                    oSheet.Cells[current_rownum, 5].Value2 = r["U_CP_CN"];
                    detail_rownum = current_rownum + 1;
                    current_rownum++;
                    foreach (DataRow rd in C.Select("U_DTT_LineID=" + r["LineID"]))
                    {
                        if (decimal.Parse(rd["U_CP_CN"].ToString()) != 0)
                        {
                            oSheet.Cells[current_rownum, 2] = rd["U_BPNAME"].ToString();
                            oSheet.Cells[current_rownum, 6].Value2 = rd["U_CP_CN"];
                            oSheet.Cells[current_rownum, 7].Value2 = rd["KL_TT_DD"];
                            oSheet.Cells[current_rownum, 9].Value2 = rd["CTQL"];
                            current_rownum++;
                        }
                    }
                    if (current_rownum - detail_rownum >= 1)
                    {
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 6]).Formula = string.Format("=SUBTOTAL(9,{0})", "F" + detail_rownum + ":F" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + detail_rownum + ":G" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{1}", detail_rownum, detail_rownum);
                    }

                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Italic = true;
                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    oSheet.Cells[current_rownum, 2] = "Dự phòng";
                    oSheet.Cells[current_rownum, 5].Value2 = r["U_CP_DP"];
                    detail_rownum = current_rownum + 1;
                    current_rownum++;
                    foreach (DataRow rd in C.Select("U_DTT_LineID=" + r["LineID"]))
                    {
                        if (decimal.Parse(rd["U_CP_DP"].ToString()) != 0)
                        {
                            oSheet.Cells[current_rownum, 2] = rd["U_BPNAME"].ToString();
                            oSheet.Cells[current_rownum, 6].Value2 = rd["U_CP_DP"];
                            oSheet.Cells[current_rownum, 7].Value2 = rd["KL_TT_DD"];
                            oSheet.Cells[current_rownum, 9].Value2 = rd["CTQL"];
                            current_rownum++;
                        }
                    }
                    if (current_rownum - detail_rownum >= 1)
                    {
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 6]).Formula = string.Format("=SUBTOTAL(9,{0})", "F" + detail_rownum + ":F" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + detail_rownum + ":G" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{1}", detail_rownum, detail_rownum);
                    }

                    //SUM DU PHONG 2
                    decimal tmp_dp2 = 0;
                    decimal.TryParse(r["U_CP_TB"].ToString(), out tmp_dp2);
                    sum_dp2 += tmp_dp2;
                    //oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Italic = true;
                    //oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    //oSheet.Cells[current_rownum, 2] = "Dự phòng 2";
                    //oSheet.Cells[current_rownum, 5] = r["U_CP_DP2"].ToString();
                    //current_rownum++;
                    //foreach (DataRow rd in C.Select("U_DTT_LineID=" + r["LineID"]))
                    //{
                    //    if (decimal.Parse(rd["U_CP_DP2"].ToString()) > 0)
                    //    {
                    //        oSheet.Cells[current_rownum, 2] = rd["U_BPNAME"].ToString();
                    //        oSheet.Cells[current_rownum, 6] = rd["U_CP_DP2"].ToString();
                    //        current_rownum++;
                    //    }
                    //}

                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Italic = true;
                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    oSheet.Cells[current_rownum, 2] = "PRELIM";
                    oSheet.Cells[current_rownum, 5].Value2 = r["U_CP_PRELIMs"];
                    detail_rownum = current_rownum + 1;
                    current_rownum++;
                    foreach (DataRow rd in C.Select("U_DTT_LineID=" + r["LineID"]))
                    {
                        if (decimal.Parse(rd["U_CP_PRELIMs"].ToString()) != 0)
                        {
                            oSheet.Cells[current_rownum, 2] = rd["U_BPNAME"].ToString();
                            oSheet.Cells[current_rownum, 6].Value2 = rd["U_CP_PRELIMs"];
                            oSheet.Cells[current_rownum, 7].Value2 = rd["KL_TT_DD"];
                            oSheet.Cells[current_rownum, 9].Value2 = rd["CTQL"];
                            current_rownum++;
                        }
                    }
                    if (current_rownum - detail_rownum >= 1)
                    {
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 6]).Formula = string.Format("=SUBTOTAL(9,{0})", "F" + detail_rownum + ":F" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + detail_rownum + ":G" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{1}", detail_rownum, detail_rownum);
                    }

                    //SUM THIET BI CONG TAC
                    decimal tmp_tb = 0;
                    decimal.TryParse(r["U_CP_TB"].ToString(), out tmp_tb);
                    sum_tb += tmp_tb;
                    //oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Italic = true;
                    //oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    //oSheet.Cells[current_rownum, 2] = "Thiết bị";
                    //oSheet.Cells[current_rownum, 5] = r["U_CP_TB"].ToString();
                    //current_rownum++;
                    //foreach (DataRow rd in C.Select("U_DTT_LineID=" + r["LineID"]))
                    //{
                    //    if (decimal.Parse(rd["U_CP_TB"].ToString()) > 0)
                    //    {
                    //        oSheet.Cells[current_rownum, 2] = rd["U_BPNAME"].ToString();
                    //        oSheet.Cells[current_rownum, 6] = rd["U_CP_TB"].ToString();
                    //        current_rownum++;
                    //    }
                    //}

                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Italic = true;
                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    oSheet.Cells[current_rownum, 2] = "Khác";
                    oSheet.Cells[current_rownum, 5].Value2 = r["U_CP_K"];
                    detail_rownum = current_rownum + 1;
                    current_rownum++;
                    foreach (DataRow rd in C.Select("U_DTT_LineID=" + r["LineID"]))
                    {
                        if (decimal.Parse(rd["U_CP_K"].ToString()) != 0)
                        {
                            oSheet.Cells[current_rownum, 2] = rd["U_BPNAME"].ToString();
                            oSheet.Cells[current_rownum, 6].Value2 = rd["U_CP_K"];
                            oSheet.Cells[current_rownum, 7].Value2 = rd["KL_TT_DD"];
                            current_rownum++;
                        }
                    }
                    if (current_rownum - detail_rownum >= 1)
                    {
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 6]).Formula = string.Format("=SUBTOTAL(9,{0})", "F" + detail_rownum + ":F" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + detail_rownum + ":G" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{1}", detail_rownum, detail_rownum);
                    }
                    current_rownum++;
                    STT_GROUP++;
                    ////Total Cong tac
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Group_No_RowNum[(Group_No_RowNum.Count - 1)] + 1, current_rownum - 2);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 6]).Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", Group_No_RowNum[(Group_No_RowNum.Count - 1)] + 1, current_rownum - 2);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Group_No_RowNum[(Group_No_RowNum.Count - 1)] + 1, current_rownum - 2);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 8]).Formula = string.Format("=G{0}/F{1}", Group_No_RowNum[(Group_No_RowNum.Count - 1)], Group_No_RowNum[(Group_No_RowNum.Count - 1)]);
                    #endregion
                }
                //THIET BI
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 1] = STT_GROUP;
                oSheet.Cells[current_rownum, 2] = "CHI PHÍ THIẾT BỊ";
                oSheet.Cells[current_rownum, 5].Value2 = sum_tb;
                Group_No_RowNum.Add(current_rownum);
                STT_GROUP++;
                current_rownum++;
                current_rownum++;

                //CP NCC - NTP Khac
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 1] = STT_GROUP;
                oSheet.Cells[current_rownum, 2] = "CHI PHÍ NCC/NTP KHÁC";
                oSheet.Cells[current_rownum, 5].Value2 = sum_dp2;
                Group_No_RowNum.Add(current_rownum);
                current_rownum++;
                current_rownum++;

                //Total I
                if (Group_No_RowNum.Count > 0)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 4]).Formula = string.Format("=SUBTOTAL(9,{0})", "D" + (Section_RowNum[(Section_RowNum.Count - 1)] + 1) + ":D" + (current_rownum - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,{0})", "E" + (Section_RowNum[(Section_RowNum.Count - 1)] + 1) + ":E" + (current_rownum - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 6]).Formula = string.Format("=SUBTOTAL(9,{0})", "F" + (Section_RowNum[(Section_RowNum.Count - 1)] + 1) + ":F" + (current_rownum - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (Section_RowNum[(Section_RowNum.Count - 1)] + 1) + ":G" + (current_rownum - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 8]).Formula = string.Format("=G{0}/F{1}", Section_RowNum[(Section_RowNum.Count - 1)], Section_RowNum[(Section_RowNum.Count - 1)]);
                }


                //II - CHI PHÍ QUẢN LÝ BCH TRỰC TIẾP
                DataTable D = Get_Data_BCH_CE(GoiThau_Key);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 1] = "II";
                oSheet.Cells[current_rownum, 2] = "CHI PHÍ QUẢN LÝ BCH TRỰC TIẾP";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", (current_rownum + 1), (current_rownum + 37));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", (current_rownum + 1), (current_rownum + 37));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", (current_rownum + 1), (current_rownum + 37));
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = string.Format("=SUM({0})", "G" + (current_rownum + 1) + ",G" + (current_rownum + 8) + ",G" + (current_rownum + 13) + ",G" + (current_rownum + 20));
                Section_RowNum.Add(current_rownum);
                current_rownum++;
                #region Details
                //1
                //oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                oSheet.Cells[current_rownum, 1] = "1";
                oSheet.Cells[current_rownum, 2] = "Chi phí lương, bảo hiểm, phụ cấp, công trường ...";
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=SUM({0}:{1})", "F"+(current_rownum +1),"F"+ (current_rownum + 7));
                oSheet.Cells[current_rownum, 5].Value2 = D.Select("U_TKKT='CPQL0000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='CPQL0000'")[0]["U_GTDP"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Phải trả công nhân viên";
                oSheet.Cells[current_rownum, 3] = "3341";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='33410000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33410000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='33410000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33410000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Phải trả người lao động khác (đội thi công)";
                oSheet.Cells[current_rownum, 3] = "33481";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='33481000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33481000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='33481000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33481000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phí lương kỹ thuật viên";
                oSheet.Cells[current_rownum, 3] = "33482";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='33482000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33482000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='33482000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33482000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phí vệ sinh, giữ xe,.. công trường (BCH)";
                oSheet.Cells[current_rownum, 3] = "33483";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='33483000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33483000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='33483000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33483000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phí lương an toàn viên";
                oSheet.Cells[current_rownum, 3] = "33484";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='33484000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33484000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='33484000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33484000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "BHXH,BHYT,KPCĐ,BHTN";
                oSheet.Cells[current_rownum, 3] = "62712";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62712000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62712000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62712000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62712000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Cells[current_rownum, 1] = "2";
                oSheet.Cells[current_rownum, 2] = "Chi phí vật tư lẻ";
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=SUM({0}:{1})", "F" + (current_rownum + 1), "F" + (current_rownum + 5));
                oSheet.Cells[current_rownum, 5].Value2 = D.Select("U_TKKT='CPVTL000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='CPVTL000'")[0]["U_GTDP"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phí nguyên vật liệu trực tiếp";
                oSheet.Cells[current_rownum, 3] = "621";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62100000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62100000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62100000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62100000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Nhiên liệu";
                oSheet.Cells[current_rownum, 3] = "62781";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62781000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62781000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62781000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62781000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phí bằng tiền khác";
                oSheet.Cells[current_rownum, 3] = "62788";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62788000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62788000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62788000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62788000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Bảo hộ lao động";
                oSheet.Cells[current_rownum, 3] = "62733";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62733000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62733000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62733000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62733000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Cells[current_rownum, 1] = "3";
                oSheet.Cells[current_rownum, 2] = "Chi phí máy móc thiết bị";
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=SUM({0}:{1})", "F" + (current_rownum + 1), "F" + (current_rownum + 7));
                oSheet.Cells[current_rownum, 5].Value2 = D.Select("U_TKKT='MMTB0000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='MMTB0000'")[0]["U_GTDP"] : "";
                //oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='MMTB0000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='MMTB0000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Công cụ, dụng cụ, thiết bị Ban chỉ huy CT";
                oSheet.Cells[current_rownum, 3] = "62731";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62731000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62731000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62731000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62731000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "VPP, photocopy";
                oSheet.Cells[current_rownum, 3] = "62732";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62732000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62732000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62732000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62732000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phí vận chuyển";
                oSheet.Cells[current_rownum, 3] = "62734";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62734000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62734000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62734000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62734000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Điện, nước thi công";
                oSheet.Cells[current_rownum, 3] = "62774";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62774000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62774000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62774000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62774000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Điện thoại cố định";
                oSheet.Cells[current_rownum, 3] = "62775";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62775000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62775000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62775000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62775000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Thuê TSCĐ, thiết bị thi công";
                oSheet.Cells[current_rownum, 3] = "62776";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62776000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62776000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62776000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62776000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Cells[current_rownum, 1] = "4";
                oSheet.Cells[current_rownum, 2] = "Chi phí ban chỉ huy văn phòng";
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=SUM({0}:{1})", "F" + (current_rownum + 1), "F" + (current_rownum + 18));
                oSheet.Cells[current_rownum, 5].Value2 = D.Select("U_TKKT='BCHVP000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='BCHVP000'")[0]["U_GTDP"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Ăn trưa";
                oSheet.Cells[current_rownum, 3] = "62713";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62713000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62713000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62713000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62713000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Điện thoại di động";
                oSheet.Cells[current_rownum, 3] = "62714";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62714000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62714000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62714000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62714000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phí thuê nhà";
                oSheet.Cells[current_rownum, 3] = "62716";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62716000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62716000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62716000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62716000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Thuế xuất nhập khẩu";
                oSheet.Cells[current_rownum, 3] = "62723";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62723000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62723000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62723000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62723000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Bao chí, bưu phí, tài liệu";
                oSheet.Cells[current_rownum, 3] = "62735";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62735000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62735000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62735000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62735000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Phí, lệ phí";
                oSheet.Cells[current_rownum, 3] = "62770";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62770000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62770000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Tiếp khách";
                oSheet.Cells[current_rownum, 3] = "62771";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62771000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62771000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62771000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62771000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Phí kiểm định, thí nghiệm";
                oSheet.Cells[current_rownum, 3] = "62773";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62773000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62773000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62773000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62773000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Phí ngân hàng";
                oSheet.Cells[current_rownum, 3] = "62777";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62777000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62777000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62777000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62777000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Quảng cáo, đào tạo";
                oSheet.Cells[current_rownum, 3] = "62778";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62778000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62778000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62778000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62778000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phí nhà thầu phụ";
                oSheet.Cells[current_rownum, 3] = "62779";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62779000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62779000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62779000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62779000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phí giao nhận hàng hóa nhập khẩu";
                oSheet.Cells[current_rownum, 3] = "62782";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62782000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62782000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62782000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62782000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Công tác phí";
                oSheet.Cells[current_rownum, 3] = "62783";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62783000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62783000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62783000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62783000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phí bị loại trừ";
                oSheet.Cells[current_rownum, 3] = "62784";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62784000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62784000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62784000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62784000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Thuốc, y tế, đồ dùng lặt vặt";
                oSheet.Cells[current_rownum, 3] = "62785";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62785000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62785000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62785000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62785000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Hồ sơ thầu";
                oSheet.Cells[current_rownum, 3] = "62786";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62786000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62786000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62786000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62786000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Phí bảo hiểm";
                oSheet.Cells[current_rownum, 3] = "62787";
                oSheet.Cells[current_rownum, 6].Value2 = D.Select("U_TKKT='62787000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62787000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 7].Value2 = D.Select("U_TKKT='62787000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62787000'")[0]["TOTAL_BCH"] : "";
                current_rownum++;
                #endregion

                //III - CHI PHÍ HỔ TRỢ
                DataTable VII = Get_Data_FI_VII(GoiThau_Key);
                string f_ht1 = "", f_ht2 = "", f_ng = "", f_dpcp = "", f_dpbh = "", f_cpqlct = "";
                if (VII.Rows.Count > 0)
                {
                    foreach (DataRow r in VII.Rows)
                    {
                        f_ht1 += string.Format(@"{0}*{1}/100 + ", r["Total"], r["HT1"]);
                        f_ht2 += string.Format(@"{0}*{1}/100 + ", r["Total"], r["HT2"]);
                        f_dpcp += string.Format(@"{0}*{1}/100 + ", r["Total"], r["DPCP"]);
                        f_dpbh += string.Format(@"{0}*{1}/100 + ", r["Total"], r["DPBH"]);
                        f_cpqlct += string.Format(@"{0}*{1}/100 + ", r["Total"], r["CPQLCT"]);
                        f_ng += string.Format(@"{0} + ", r["CPNG"]);
                    }
                }
                //DataTable E = Get_Prj_Info(FProject);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 1] = "III";
                oSheet.Cells[current_rownum, 2] = "CHI PHÍ HỔ TRỢ";
                Section_RowNum.Add(current_rownum);
                current_rownum++;
                //Chi phi ho tro 1
                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phi hỗ trợ 1";
                //oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                if (f_ht1 != "")
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = "=" + f_ht1.Substring(0, f_ht1.Length - 3);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = "=" + f_ht1.Substring(0, f_ht1.Length - 3);
                }
                current_rownum++;
                //Chi phi ho tro 2
                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phi hỗ trợ 2";
                if (f_ht2 != "")
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = "=" + f_ht2.Substring(0, f_ht2.Length - 3);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = "=" + f_ht2.Substring(0, f_ht2.Length - 3);
                }
                current_rownum++;
                //Chi phi quan ly cong ty
                //oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                //oSheet.Cells[current_rownum, 2] = "Chi phi quản lý công ty";
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("={0}*{1}/100", "D6", E.Rows[0]["U_CPQLCT"].ToString());
                //current_rownum++;
                //Chi phi NG
                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phi NG";
                if (f_ng != "")
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = "=" + f_ng.Substring(0, f_ng.Length - 3);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = "=" + f_ng.Substring(0, f_ng.Length - 3);
                }
                current_rownum++;

                //Total III
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Section_RowNum[(Section_RowNum.Count - 1)] + 1, current_rownum - 1);
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 6]).Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", Section_RowNum[(Section_RowNum.Count - 1)] + 1, current_rownum - 1);
                //IV - CHI PHÍ NCC/NTP KHÁC
                //DataTable D = Get_Data_BCH(this.ComboBox0.Selected.Description);
                //oSheet.Range["A" + current_rownum, "H" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                //oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                //oSheet.Cells[current_rownum, 1] = "IV";
                //oSheet.Cells[current_rownum, 2] = "CHI PHÍ NCC/NTP KHÁC";
                //current_rownum++;

                //oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                ////oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 4]).Formula = string.Format("={0}*{1}/100", "D6", E.Rows[0]["U_DPCP"].ToString());
                //current_rownum++;

                //IV - DỰ PHÒNG PHÍ
                //DataTable D = Get_Data_BCH(this.ComboBox0.Selected.Description);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 1] = "IV";
                oSheet.Cells[current_rownum, 2] = "DỰ PHÒNG PHÍ";
                Section_RowNum.Add(current_rownum);
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                //oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 2] = "Dự phòng chi phí cho ĐTC/ NTP/ NCC (0.5% giá trị doanh thu)";
                if (f_dpcp != "")
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = "=" + f_dpcp.Substring(0, f_dpcp.Length - 3);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = "=" + f_dpcp.Substring(0, f_dpcp.Length - 3);
                }
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                //oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 2] = "Dự phòng chi phí bảo hành (0.5% giá trị doanh thu)";
                if (f_dpbh != "")
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = "=" + f_dpbh.Substring(0, f_dpbh.Length - 3);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = "=" + f_dpbh.Substring(0, f_dpbh.Length - 3);
                }
                current_rownum++;

                //Total IV
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUM(E{0}:E{1})", Section_RowNum[(Section_RowNum.Count - 1)] + 1, current_rownum - 1);
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Section_RowNum[(Section_RowNum.Count - 1)] + 1, current_rownum - 1);
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 6]).Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", Section_RowNum[(Section_RowNum.Count - 1)] + 1, current_rownum - 1);
                //Total B
                if (Section_RowNum.Count > 0)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[13, 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Section_RowNum[0], Section_RowNum[(Section_RowNum.Count - 1)] + 2);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[13, 4]).Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", Section_RowNum[0], Section_RowNum[(Section_RowNum.Count - 1)] + 2);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[13, 6]).Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", Section_RowNum[0], Section_RowNum[(Section_RowNum.Count - 1)] + 2);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[13, 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Section_RowNum[0], Section_RowNum[(Section_RowNum.Count - 1)] + 2);
                }

                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[6, 4]).Formula = string.Format("=SUM({0}:{1})", "D7", "D12");

                //C
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 1] = "C";
                oSheet.Cells[current_rownum, 2] = "LỢI NHUẬN GỘP CỦA CÔNG TRƯỜNG A";
                //Section_RowNum.Add(current_rownum);
                current_rownum++;

                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 2] = "DOANH THU";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 4]).Formula = "=D6";
                current_rownum++;

                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 2] = "CHI PHÍ";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = "=E13";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = "=F13";

                current_rownum++;

                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 2] = "LỢI NHUẬN GỘP A";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("={0}-{1}", "D" + (current_rownum - 2), "E" + (current_rownum - 1));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = string.Format("={0}-{1}", "D" + (current_rownum - 2), "F" + (current_rownum - 1));
                current_rownum++;

                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 2] = "TỶ SUẤT LỢI NHUẬN GỘP A/ DOANH THU";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("={0}/{1}", "E" + (current_rownum - 1), "D" + (current_rownum - 3));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = string.Format("={0}/{1}", "F" + (current_rownum - 1), "D" + (current_rownum - 3));
                oSheet.Range["E" + current_rownum].NumberFormat = "0.00%";
                oSheet.Range["F" + current_rownum].NumberFormat = "0.00%";

                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = string.Format("=((D6-D8)-(F13-{0}))/(D6-D8)", "F" + Group_No_RowNum[(Group_No_RowNum.Count - 1)]);
                oSheet.Range["G" + current_rownum].NumberFormat = "0.00%";
                current_rownum++;

                //D
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[6, 4]).Formula = string.Format("=SUM({0}:{1})", "D7", "D12");
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 1] = "D";
                oSheet.Cells[current_rownum, 2] = "LỢI NHUẬN TUYỆT ĐỐI C (Bao gồm phí quản lý Công ty)";
                //Section_RowNum.Add(current_rownum);
                current_rownum++;

                //Chi phi quan ly cong ty
                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 2] = "CHI PHÍ QUẢN LÝ CÔNG TY";
                if (f_cpqlct != "")
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = "=" + f_cpqlct.Substring(0, f_cpqlct.Length - 3);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = "=" + f_cpqlct.Substring(0, f_cpqlct.Length - 3);
                }
                current_rownum++;

                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 2] = "DOANH THU";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 4]).Formula = "=D6";
                current_rownum++;

                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 2] = "CHI PHÍ";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=E{1}+E{0}", current_rownum - 2, current_rownum - 6);
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = string.Format("=F{1}+F{0}", current_rownum - 2, current_rownum - 6);
                current_rownum++;

                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 2] = "LỢI NHUẬN TUYỆT ĐỐI C";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("={0}-{1}", "D" + (current_rownum - 2), "E" + (current_rownum - 1));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = string.Format("={0}-{1}", "D" + (current_rownum - 2), "F" + (current_rownum - 1));
                current_rownum++;

                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 2] = "TỶ SUẤT LỢI TUYỆT ĐỐI C/ DOANH THU";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("={0}/{1}", "E" + (current_rownum - 1), "D" + (current_rownum - 3));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = string.Format("={0}/{1}", "F" + (current_rownum - 1), "D" + (current_rownum - 3));
                oSheet.Range["E" + current_rownum].NumberFormat = "0.00%";
                oSheet.Range["F" + current_rownum].NumberFormat = "0.00%";
                current_rownum++;

                //Border
                ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A7", "H" + (current_rownum - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                oXL.ActiveWindow.Activate();

            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
            }
        }

        //View FI Report Button
        private void Button4_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string PrjName = Grid0.DataTable.GetValue("Project Name", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
            string FProject = Grid0.DataTable.GetValue("Financial Project", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
            DateTime BaseLine_Date = DateTime.Parse(Grid0.DataTable.GetValue("BaseLine Date", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString());
            string GoiThau_Key = "";
            string GoiThau_Name = "";
            for (int i = 0; i < Grid3.Rows.Count; i++)
            {
                string IsSelected = Grid3.DataTable.GetValue("Checked", i).ToString();
                if (IsSelected == "Y")
                {
                    GoiThau_Key += Grid3.DataTable.GetValue("AbsEntry", i).ToString() + ",";
                    GoiThau_Name = Grid3.DataTable.GetValue("SubProject Name", i).ToString();
                }
            }
            if (GoiThau_Key.Length > 0)
                GoiThau_Key = GoiThau_Key.Substring(0, GoiThau_Key.Length - 1);

            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            //Open Template
            oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_BCTC.xlsx");
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            int current_row = 7;
            try
            {
                DataTable A = Get_Data_BCTCA(GoiThau_Key);
                List<int> Group_No_RowNum = new List<int>();
                List<int> Section_RowNum = new List<int>();
                //bool ME_Project = false;
                //Fill Header
                //Project Name
                oSheet.Cells[2, 4] = "Dự án: " + PrjName;
                //Subproject Name
                if (!string.IsNullOrEmpty(GoiThau_Key))
                {
                    if (GoiThau_Key.Split(',').Count() == 1)
                        oSheet.Cells[3, 4] = "Gói thầu: " + GoiThau_Name;
                }
                //oSheet.Cells[3, 4] = "Gói thầu: " + this.ComboBox1.Selected.Description;
                //Thang
                oSheet.Cells[4, 4] = "Tháng: " + BaseLine_Date.ToString("MM-yyyy");// this.ComboBox2.Selected.Value;

                //A- Doanh thu (truoc VAT)
                //Gia tri hop dong
                oSheet.Cells[current_row, 1] = "1";
                oSheet.Cells[current_row, 3] = "Giá trị hợp đồng";
                oSheet.Cells[current_row, 6].Value2 = A.Rows[0]["GTHD"];
                current_row++;

                //Gia tri hop dong 1A
                oSheet.Cells[current_row, 1] = "1A";
                oSheet.Cells[current_row, 3] = "Giá trị hợp đồng 1A";
                oSheet.Cells[current_row, 6].Value2 = A.Rows[0]["KHAC"];
                current_row++;

                //Phụ lục HĐ
                oSheet.Cells[current_row, 1] = "2";
                oSheet.Cells[current_row, 3] = "Phụ lục HĐ";
                oSheet.Cells[current_row, 6].Value2 = A.Rows[0]["PLHD"];
                current_row++;

                //Giảm giá thương mại
                oSheet.Cells[current_row, 1] = "3";
                oSheet.Cells[current_row, 3] = "Giảm giá thương mại";
                oSheet.Cells[current_row, 6].Value2 = A.Rows[0]["GGTM"];
                current_row++;

                //Giảm giá thương mại
                oSheet.Cells[current_row, 1] = "4";
                oSheet.Cells[current_row, 3] = "Phương án đề xuất tiết kiệm chi phí";
                oSheet.Cells[current_row, 6].Value2 = A.Rows[0]["PA"];
                current_row++;

                //Phí quản lý
                oSheet.Cells[current_row, 1] = "5";
                oSheet.Cells[current_row, 3] = "Phí Quản lý";
                oSheet.Cells[current_row, 6].Value2 = A.Rows[0]["PhiQL"];
                current_row++;

                //Doanh thu cung cấp dịch vụ (có hóa đơn)
                oSheet.Cells[current_row, 1] = "6";
                oSheet.Cells[current_row, 3] = "Doanh thu cung cấp dịch vụ (có hóa đơn)";
                oSheet.Cells[current_row, 6].Value2 = 0;
                current_row++;
                //Total
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[6, 6]).Formula = string.Format("=SUM({0}:{1})", "F7", "F13");
                //B-CHI PHI (Trước VAT)
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(244, 176, 132);
                oSheet.Cells[current_row, 1] = "B";
                oSheet.Cells[current_row, 3] = "CHI PHÍ (Trước VAT)";
                current_row++;
                //DataTable B = null;
                DataTable C = null;
                if (GoiThau_Key == "")
                {
                    C = Get_Data_BCTC_DUTRU();
                }
                else
                {
                    C = Get_Data_BCTC_DUTRU(GoiThau_Key);
                }

                //DOI THI CONG
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Cells[current_row, 1] = "I";
                oSheet.Cells[current_row, 3] = "ĐỘI THI CÔNG";
                //oSheet.Cells[current_row, 6].Value2 = B.Compute("SUM(U_CP_DTC)", "");
                Group_No_RowNum.Add(current_row);
                int detail_row_num = 0;
                detail_row_num = current_row;
                current_row++;
                foreach (DataRow d in C.Select("U_TYPE = 'XD'"))
                {
                    decimal tmp_cp = 0, tmp_klhd = 0, tmp_kltt = 0, tmp_klttdd, tmp_total_apinvoice = 0;
                    decimal.TryParse(d["U_CP_DTC"].ToString(), out tmp_cp);
                    decimal.TryParse(d["KL_HD"].ToString(), out tmp_klhd);
                    decimal.TryParse(d["KL_TT"].ToString(), out tmp_kltt);
                    decimal.TryParse(d["KL_TT_DD"].ToString(), out tmp_klttdd);
                    decimal.TryParse(d["TOTAL_AP_INVOICE"].ToString(), out tmp_total_apinvoice);
                    string CTQL = d["U_BPCode2"].ToString();
                    if ((tmp_cp != 0 && d["U_PUType"].ToString() == "PUT09") || (tmp_cp != 0 && string.IsNullOrEmpty(d["U_PUType"].ToString()))) // || tmp_klhd != 0 || tmp_kltt != 0 || tmp_klttdd != 0)
                    {
                        oSheet.Cells[current_row, 2] = d["U_BPCode"];
                        oSheet.Cells[current_row, 3] = d["U_BPName"];
                        oSheet.Cells[current_row, 6] = d["GTHD"];
                        oSheet.Cells[current_row, 7] = tmp_cp;
                        oSheet.Cells[current_row, 8] = tmp_klhd;
                        oSheet.Cells[current_row, 9] = tmp_klttdd;
                        oSheet.Cells[current_row, 10] = tmp_total_apinvoice;
                        oSheet.Cells[current_row, 11] = CTQL;
                        current_row++;
                    }
                }
                //Total
                if (current_row - detail_row_num > 1)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (detail_row_num + 1) + ":G" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H" + (detail_row_num + 1) + ":H" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (detail_row_num + 1) + ":I" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (detail_row_num + 1) + ":J" + (current_row - 1));
                }

                //NCC,NTP XAY DUNG
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Cells[current_row, 1] = "II";
                oSheet.Cells[current_row, 3] = "NCC, NTP XÂY DỰNG";
                int NCC_NTP_Row = 0;
                NCC_NTP_Row = current_row;
                current_row++;

                //NHA CUNG CAP XAY DUNG
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(155, 194, 230);
                oSheet.Cells[current_row, 1] = "II.1";
                oSheet.Cells[current_row, 3] = "NHÀ CUNG CẤP XÂY DỰNG";
                //oSheet.Cells[current_row, 6].Value2 = B.Compute("SUM(U_CP_NCC)", "");
                detail_row_num = current_row;
                current_row++;
                //Details
                foreach (DataRow d in C.Select("U_TYPE = 'XD'"))
                {
                    decimal tmp_ncc = 0, tmp_klhd = 0, tmp_kltt = 0, tmp_klttdd, tmp_total_apinvoice = 0;
                    decimal.TryParse(d["U_CP_NCC"].ToString(), out tmp_ncc);
                    decimal.TryParse(d["KL_HD"].ToString(), out tmp_klhd);
                    decimal.TryParse(d["KL_TT"].ToString(), out tmp_kltt);
                    decimal.TryParse(d["KL_TT_DD"].ToString(), out tmp_klttdd);
                    decimal.TryParse(d["TOTAL_AP_INVOICE"].ToString(), out tmp_total_apinvoice);
                    if ((tmp_ncc != 0 && (d["U_PUType"].ToString() == "PUT01" || d["U_PUType"].ToString() == "PUT08")) || (tmp_ncc != 0 && string.IsNullOrEmpty(d["U_PUType"].ToString()))) // || tmp_klhd != 0 || tmp_kltt != 0 || tmp_klttdd != 0)
                    {
                        oSheet.Cells[current_row, 2] = d["U_BPCode"];
                        oSheet.Cells[current_row, 3] = d["U_BPName"];
                        oSheet.Cells[current_row, 6] = d["GTHD"];
                        oSheet.Cells[current_row, 7] = tmp_ncc;
                        oSheet.Cells[current_row, 8] = tmp_klhd;
                        oSheet.Cells[current_row, 9] = tmp_klttdd;
                        oSheet.Cells[current_row, 10] = tmp_total_apinvoice;
                        current_row++;
                    }
                }
                //Total
                if (current_row - detail_row_num > 1)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (detail_row_num + 1) + ":G" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H" + (detail_row_num + 1) + ":H" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (detail_row_num + 1) + ":I" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (detail_row_num + 1) + ":J" + (current_row - 1));
                }

                //NHA THAU PHU XAY DUNG
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(155, 194, 230);
                oSheet.Cells[current_row, 1] = "II.2";
                oSheet.Cells[current_row, 3] = "NHÀ THẦU PHỤ XÂY DỰNG";
                //oSheet.Cells[current_row, 6].Value2 = B.Compute("SUM(U_CP_NTP)", "");
                detail_row_num = current_row;
                current_row++;
                //Details
                foreach (DataRow d in C.Select("U_TYPE = 'XD'"))
                {
                    decimal tmp_ntp = 0, tmp_klhd = 0, tmp_kltt = 0, tmp_klttdd = 0, tmp_total_apinvoice = 0;
                    decimal.TryParse(d["U_CP_NTP"].ToString(), out tmp_ntp);
                    decimal.TryParse(d["KL_HD"].ToString(), out tmp_klhd);
                    decimal.TryParse(d["KL_TT"].ToString(), out tmp_kltt);
                    decimal.TryParse(d["KL_TT_DD"].ToString(), out tmp_klttdd);
                    decimal.TryParse(d["TOTAL_AP_INVOICE"].ToString(), out tmp_total_apinvoice);
                    if ((tmp_ntp != 0 && d["U_PUType"].ToString() == "PUT02") || (tmp_ntp != 0 && string.IsNullOrEmpty(d["U_PUType"].ToString()))) //|| tmp_klhd != 0 || tmp_kltt != 0 || tmp_klttdd != 0)
                    {
                        oSheet.Cells[current_row, 2] = d["U_BPCode"];
                        oSheet.Cells[current_row, 3] = d["U_BPName"];
                        oSheet.Cells[current_row, 6] = d["GTHD"];
                        oSheet.Cells[current_row, 7] = tmp_ntp;
                        oSheet.Cells[current_row, 8] = tmp_klhd;
                        oSheet.Cells[current_row, 9] = tmp_klttdd;
                        oSheet.Cells[current_row, 10] = tmp_total_apinvoice;
                        current_row++;
                    }
                }
                //Total
                if (current_row - detail_row_num > 1)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (detail_row_num + 1) + ":G" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H" + (detail_row_num + 1) + ":H" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (detail_row_num + 1) + ":I" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (detail_row_num + 1) + ":J" + (current_row - 1));
                }

                //Total II
                if (current_row - detail_row_num >= 1)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_Row, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (NCC_NTP_Row + 1) + ":G" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_Row, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H" + (NCC_NTP_Row + 1) + ":H" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_Row, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (NCC_NTP_Row + 1) + ":I" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_Row, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (NCC_NTP_Row + 1) + ":J" + (current_row - 1));
                }
                //NHA CUNG CAP, NHA THAU PHU M&E
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Cells[current_row, 1] = "III";
                oSheet.Cells[current_row, 3] = "NHÀ CUNG CẤP, NHÀ THẦU PHỤ M&E";
                NCC_NTP_Row = current_row;
                current_row++;
                //NHA CUNG CAP M&E
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(155, 194, 230);
                oSheet.Cells[current_row, 1] = "III.1";
                oSheet.Cells[current_row, 3] = "NHÀ CUNG CẤP M&E";
                detail_row_num = current_row;
                current_row++;

                //Details
                foreach (DataRow d in C.Select("U_TYPE = 'CDXD'"))
                {
                    decimal tmp_ncc = 0, tmp_klhd = 0, tmp_kltt = 0, tmp_klttdd = 0, tmp_total_apinvoice = 0;
                    decimal.TryParse(d["U_CP_NCC"].ToString(), out tmp_ncc);
                    decimal.TryParse(d["KL_HD"].ToString(), out tmp_klhd);
                    decimal.TryParse(d["KL_TT"].ToString(), out tmp_kltt);
                    decimal.TryParse(d["KL_TT_DD"].ToString(), out tmp_klttdd);
                    decimal.TryParse(d["TOTAL_AP_INVOICE"].ToString(), out tmp_total_apinvoice);
                    if ((tmp_ncc != 0 && d["U_PUType"].ToString() == "PUT01") || (tmp_ncc != 0 && string.IsNullOrEmpty(d["U_PUType"].ToString()))) // || tmp_klhd != 0 || tmp_kltt != 0 || tmp_klttdd != 0)
                    {
                        oSheet.Cells[current_row, 2] = d["U_BPCode"];
                        oSheet.Cells[current_row, 3] = d["U_BPName"];
                        oSheet.Cells[current_row, 6] = d["GTHD"];
                        oSheet.Cells[current_row, 7] = tmp_ncc;
                        oSheet.Cells[current_row, 8] = tmp_klhd;
                        oSheet.Cells[current_row, 9] = tmp_klttdd;
                        oSheet.Cells[current_row, 10] = tmp_total_apinvoice;
                        current_row++;
                    }
                }
                //Total
                if (current_row - detail_row_num > 1)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (detail_row_num + 1) + ":G" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H" + (detail_row_num + 1) + ":H" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (detail_row_num + 1) + ":I" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (detail_row_num + 1) + ":J" + (current_row - 1));
                }

                //NHA THAU PHU M&E
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(155, 194, 230);
                oSheet.Cells[current_row, 1] = "III.2";
                oSheet.Cells[current_row, 3] = "NHÀ THẦU PHỤ M&E";
                detail_row_num = current_row;
                current_row++;
                //Details
                foreach (DataRow d in C.Select("U_TYPE = 'CDXD'"))
                {
                    decimal tmp_ntp = 0, tmp_klhd = 0, tmp_kltt = 0, tmp_klttdd = 0, tmp_total_apinvoice = 0;
                    decimal.TryParse(d["U_CP_NTP"].ToString(), out tmp_ntp);
                    decimal.TryParse(d["KL_HD"].ToString(), out tmp_klhd);
                    decimal.TryParse(d["KL_TT"].ToString(), out tmp_kltt);
                    decimal.TryParse(d["KL_TT_DD"].ToString(), out tmp_klttdd);
                    decimal.TryParse(d["TOTAL_AP_INVOICE"].ToString(), out tmp_total_apinvoice);
                    if ((tmp_ntp != 0 && d["U_PUType"].ToString() == "PUT02") || (tmp_ntp != 0 && string.IsNullOrEmpty(d["U_PUType"].ToString()))) // || tmp_klhd != 0 || tmp_kltt != 0 || tmp_klttdd != 0)
                    {
                        oSheet.Cells[current_row, 2] = d["U_BPCode"];
                        oSheet.Cells[current_row, 3] = d["U_BPName"];
                        oSheet.Cells[current_row, 6] = d["GTHD"];
                        oSheet.Cells[current_row, 7] = tmp_ntp;
                        oSheet.Cells[current_row, 8] = tmp_klhd;
                        oSheet.Cells[current_row, 9] = tmp_klttdd;
                        oSheet.Cells[current_row, 10] = tmp_total_apinvoice;
                        current_row++;
                    }
                }
                //Total
                if (current_row - detail_row_num > 1)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (detail_row_num + 1) + ":G" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H" + (detail_row_num + 1) + ":H" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (detail_row_num + 1) + ":I" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (detail_row_num + 1) + ":J" + (current_row - 1));
                }
                //Total III
                if (current_row - detail_row_num > 1)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_Row, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (NCC_NTP_Row + 1) + ":G" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_Row, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H" + (NCC_NTP_Row + 1) + ":H" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_Row, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (NCC_NTP_Row + 1) + ":I" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_Row, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (NCC_NTP_Row + 1) + ":J" + (current_row - 1));
                }
                //CHI PHI THIET BI
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Cells[current_row, 1] = "IV";
                oSheet.Cells[current_row, 3] = "CHI PHÍ THIẾT BỊ";
                NCC_NTP_Row = current_row;
                current_row++;
                //NHA CUNG CAP THIET BI
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(155, 194, 230);
                oSheet.Cells[current_row, 1] = "IV.1";
                oSheet.Cells[current_row, 3] = "NHÀ CUNG THIẾT BỊ";
                detail_row_num = current_row;
                current_row++;

                //Details
                foreach (DataRow d in C.Select("U_TYPE = 'TBXD'"))
                {
                    decimal tmp_ncc = 0, tmp_klhd = 0, tmp_kltt = 0, tmp_klttdd = 0, tmp_total_apinvoice = 0;
                    decimal.TryParse(d["U_CP_NCC"].ToString(), out tmp_ncc);
                    decimal.TryParse(d["KL_HD"].ToString(), out tmp_klhd);
                    decimal.TryParse(d["KL_TT"].ToString(), out tmp_kltt);
                    decimal.TryParse(d["KL_TT_DD"].ToString(), out tmp_klttdd);
                    decimal.TryParse(d["TOTAL_AP_INVOICE"].ToString(), out tmp_total_apinvoice);
                    if ((tmp_ncc != 0 && d["U_PUType"].ToString() == "PUT01") || (tmp_ncc != 0 && string.IsNullOrEmpty(d["U_PUType"].ToString()))) // || tmp_klhd != 0 || tmp_kltt != 0 || tmp_klttdd != 0)
                    {
                        oSheet.Cells[current_row, 2] = d["U_BPCode"];
                        oSheet.Cells[current_row, 3] = d["U_BPName"];
                        oSheet.Cells[current_row, 6] = d["GTHD"];
                        oSheet.Cells[current_row, 7] = tmp_ncc;
                        oSheet.Cells[current_row, 8] = tmp_klhd;
                        oSheet.Cells[current_row, 9] = tmp_klttdd;
                        oSheet.Cells[current_row, 10] = tmp_total_apinvoice;
                        current_row++;
                    }
                }
                //Total
                if (current_row - detail_row_num > 1)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (detail_row_num + 1) + ":G" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H" + (detail_row_num + 1) + ":H" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (detail_row_num + 1) + ":I" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (detail_row_num + 1) + ":J" + (current_row - 1));
                }

                //NHA THAU PHU THIET BI
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(155, 194, 230);
                oSheet.Cells[current_row, 1] = "IV.2";
                oSheet.Cells[current_row, 3] = "NHÀ THẦU PHỤ THIẾT BỊ";
                detail_row_num = current_row;
                current_row++;

                //Details
                foreach (DataRow d in C.Select("U_TYPE = 'TBXD'"))
                {
                    decimal tmp_ntp = 0, tmp_klhd = 0, tmp_kltt = 0, tmp_klttdd = 0, tmp_total_apinvoice = 0;
                    decimal.TryParse(d["U_CP_NTP"].ToString(), out tmp_ntp);
                    decimal.TryParse(d["KL_HD"].ToString(), out tmp_klhd);
                    decimal.TryParse(d["KL_TT"].ToString(), out tmp_kltt);
                    decimal.TryParse(d["KL_TT_DD"].ToString(), out tmp_klttdd);
                    decimal.TryParse(d["TOTAL_AP_INVOICE"].ToString(), out tmp_total_apinvoice);
                    if ((tmp_ntp != 0 && d["U_PUType"].ToString() == "PUT02") || (tmp_ntp != 0 && string.IsNullOrEmpty(d["U_PUType"].ToString()))) // || tmp_klhd != 0 || tmp_kltt != 0 || tmp_klttdd != 0)
                    {
                        oSheet.Cells[current_row, 2] = d["U_BPCode"];
                        oSheet.Cells[current_row, 3] = d["U_BPName"];
                        oSheet.Cells[current_row, 6] = d["GTHD"];
                        oSheet.Cells[current_row, 7] = tmp_ntp;
                        oSheet.Cells[current_row, 8] = tmp_klhd;
                        oSheet.Cells[current_row, 9] = tmp_klttdd;
                        oSheet.Cells[current_row, 10] = tmp_total_apinvoice;
                        current_row++;
                    }
                }
                //Total
                if (current_row - detail_row_num > 1)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (detail_row_num + 1) + ":G" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H" + (detail_row_num + 1) + ":H" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (detail_row_num + 1) + ":I" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (detail_row_num + 1) + ":J" + (current_row - 1));
                }
                //Total IV
                if (current_row - detail_row_num >= 1)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_Row, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (NCC_NTP_Row + 1) + ":G" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_Row, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H" + (NCC_NTP_Row + 1) + ":H" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_Row, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (NCC_NTP_Row + 1) + ":I" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_Row, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (NCC_NTP_Row + 1) + ":J" + (current_row - 1));
                }

                //CHI PHI BAN CHI HUY
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Cells[current_row, 1] = "V";
                oSheet.Cells[current_row, 3] = "CHI PHÍ BAN CHỈ HUY";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (current_row + 1) + ":G" + (current_row + 37));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (current_row + 1) + ":I" + (current_row + 37));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (current_row + 1) + ":J" + (current_row + 37));
                current_row++;
                DataTable D = Get_Data_BCH_FI(GoiThau_Key);
                //1
                //oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                oSheet.Cells[current_row, 1] = "1";
                oSheet.Cells[current_row, 3] = "Chi phí lương, bảo hiểm, phụ cấp, công trường ...";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (current_row + 1) + ":I" + (current_row + 6));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (current_row + 1) + ":J" + (current_row + 6));
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Phải trả công nhân viên";
                oSheet.Cells[current_row, 2] = "3341";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='33410000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33410000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='33410000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33410000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='33410000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33410000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Phải trả người lao động khác (đội thi công)";
                oSheet.Cells[current_row, 2] = "33481";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='33481000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33481000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='33481000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33481000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='33481000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33481000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí lương kỹ thuật viên";
                oSheet.Cells[current_row, 2] = "33482";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='33482000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33482000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='33482000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33482000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='33482000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33482000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí vệ sinh, giữ xe,.. công trường (BCH)";
                oSheet.Cells[current_row, 2] = "33483";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='33483000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33483000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='33483000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33483000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='33483000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33483000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí lương an toàn viên";
                oSheet.Cells[current_row, 2] = "33484";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='33484000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33484000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='33484000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33484000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='33484000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33484000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "BHXH,BHYT,KPCĐ,BHTN";
                oSheet.Cells[current_row, 2] = "62712";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62712000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62712000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62712000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62712000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62712000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62712000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Cells[current_row, 1] = "2";
                oSheet.Cells[current_row, 3] = "Chi phí vật tư lẻ";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (current_row + 1) + ":I" + (current_row + 4));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (current_row + 1) + ":J" + (current_row + 4));
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí nguyên vật liệu trực tiếp";
                oSheet.Cells[current_row, 2] = "621";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62100000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62100000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62100000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62100000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62100000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62100000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Nhiên liệu";
                oSheet.Cells[current_row, 2] = "62781";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62781000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62781000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62781000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62781000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62781000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62781000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí bằng tiền khác";
                oSheet.Cells[current_row, 2] = "62788";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62788000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62788000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62788000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62788000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62788000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62788000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Bảo hộ lao động";
                oSheet.Cells[current_row, 2] = "62733";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62733000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62733000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62733000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62733000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62733000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62733000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Cells[current_row, 1] = "3";
                oSheet.Cells[current_row, 3] = "Chi phí máy móc thiết bị";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (current_row + 1) + ":I" + (current_row + 6));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (current_row + 1) + ":J" + (current_row + 6));
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Công cụ, dụng cụ, thiết bị Ban chỉ huy CT";
                oSheet.Cells[current_row, 2] = "62731";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62731000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62731000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62731000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62731000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62731000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62731000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "VPP, photocopy";
                oSheet.Cells[current_row, 2] = "62732";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62732000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62732000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62732000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62732000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62732000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62732000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí vận chuyển";
                oSheet.Cells[current_row, 2] = "62734";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62734000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62734000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62734000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62734000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62734000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62734000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Điện, nước thi công";
                oSheet.Cells[current_row, 2] = "62774";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62774000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62774000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62774000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62774000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62774000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62774000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Điện thoại cố định";
                oSheet.Cells[current_row, 2] = "62775";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62775000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62775000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62775000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62775000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62775000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62775000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Thuê TSCĐ, thiết bị thi công";
                oSheet.Cells[current_row, 2] = "62776";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62776000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62776000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62776000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62776000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62776000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62776000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Cells[current_row, 1] = "4";
                oSheet.Cells[current_row, 3] = "Chi phí ban chỉ huy văn phòng";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (current_row + 1) + ":I" + (current_row + 17));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (current_row + 1) + ":J" + (current_row + 17));
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Ăn trưa";
                oSheet.Cells[current_row, 2] = "62713";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62713000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62713000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62713000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62713000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62713000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62713000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Điện thoại di động";
                oSheet.Cells[current_row, 2] = "62714";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62714000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62714000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62714000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62714000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62714000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62714000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí thuê nhà";
                oSheet.Cells[current_row, 2] = "62716";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62716000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62716000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62716000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62716000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62716000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62716000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Thuế xuất nhập khẩu";
                oSheet.Cells[current_row, 2] = "62723";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62723000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62723000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62723000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62723000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62723000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62723000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Bao chí, bưu phí, tài liệu";
                oSheet.Cells[current_row, 2] = "62735";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62735000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62735000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62735000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62735000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62735000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62735000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Phí, lệ phí";
                oSheet.Cells[current_row, 2] = "62770";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62770000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62770000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62770000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Tiếp khách";
                oSheet.Cells[current_row, 2] = "62771";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62771000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62771000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62771000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62771000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62771000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62771000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Phí kiểm định, thí nghiệm";
                oSheet.Cells[current_row, 2] = "62773";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62773000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62773000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62773000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62773000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62773000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62773000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Phí ngân hàng";
                oSheet.Cells[current_row, 2] = "62777";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62777000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62777000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62777000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62777000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62777000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62777000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Quảng cáo, đào tạo";
                oSheet.Cells[current_row, 2] = "62778";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62778000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62778000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62778000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62778000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62778000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62778000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí nhà thầu phụ";
                oSheet.Cells[current_row, 2] = "62779";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62779000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62779000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62779000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62779000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62779000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62779000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí giao nhận hàng hóa nhập khẩu";
                oSheet.Cells[current_row, 2] = "62782";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62782000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62782000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62782000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62782000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62782000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62782000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Công tác phí";
                oSheet.Cells[current_row, 2] = "62783";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62783000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62783000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62783000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62783000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62783000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62783000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí bị loại trừ";
                oSheet.Cells[current_row, 2] = "62784";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62784000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62784000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62784000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62784000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62784000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62784000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Thuốc, y tế, đồ dùng lặt vặt";
                oSheet.Cells[current_row, 2] = "62785";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62785000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62785000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62785000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62785000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62785000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62785000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Hồ sơ thầu";
                oSheet.Cells[current_row, 2] = "62786";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62786000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62786000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62786000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62786000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62786000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62786000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Phí bảo hiểm";
                oSheet.Cells[current_row, 2] = "62787";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62787000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62787000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62787000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62787000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62787000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62787000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                //DU PHONG PHI
                DataTable VII = Get_Data_FI_VII( GoiThau_Key);
                string f_ht1 = "", f_ht2 = "", f_ng = "", f_dpcp = "", f_dpbh = "", f_cpqlct = "";
                if (VII.Rows.Count > 0)
                {
                    foreach (DataRow r in VII.Rows)
                    {
                        f_ht1 += string.Format(@"{0}*{1}/100 + ", r["Total"], r["HT1"]);
                        f_ht2 += string.Format(@"{0}*{1}/100 + ", r["Total"], r["HT2"]);
                        f_dpcp += string.Format(@"{0}*{1}/100 + ", r["Total"], r["DPCP"]);
                        f_dpbh += string.Format(@"{0}*{1}/100 + ", r["Total"], r["DPBH"]);
                        f_cpqlct += string.Format(@"{0}*{1}/100 + ", r["Total"], r["CPQLCT"]);
                        f_ng += string.Format(@"{0} + ", r["CPNG"]);
                    }
                }

                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Cells[current_row, 1] = "VI";
                oSheet.Cells[current_row, 3] = "DỰ PHÒNG PHÍ";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                //oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 3] = "Dự phòng chi phí cho ĐTC/ NTP/ NCC (0.5% giá trị doanh thu)";
                if (f_dpcp != "")
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = "=" + f_dpcp.Substring(0, f_dpcp.Length - 3);
                    //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("={0}*{1}/100", "F6", E.Rows[0]["U_DPCP"].ToString());
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                //oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 3] = "Dự phòng chi phí bảo hành (0.5% giá trị doanh thu)";
                if (f_dpbh != "")
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = "=" + f_dpbh.Substring(0, f_dpbh.Length - 3);
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("={0}*{1}/100", "F6", E.Rows[0]["U_DPBH"].ToString());
                current_row++;
                //Total Du phong phi
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row - 3, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (current_row - 2) + ":G" + (current_row - 1));

                //HO TRO
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Cells[current_row, 1] = "VII";
                oSheet.Cells[current_row, 3] = "HỖ TRỢ";
                current_row++;

                //Chi phi ho tro 1
                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phi hỗ trợ 1";
                if (f_ht1 != "")
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = "=" + f_ht1.Substring(0, f_ht1.Length - 3);
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("={0}*{1}/100", "F6", E.Rows[0]["U_CPHT1"].ToString());
                current_row++;
                //Chi phi ho tro 2
                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phi hỗ trợ 2";
                if (f_ht2 != "")
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = "=" + f_ht2.Substring(0, f_ht2.Length - 3);
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("={0}*{1}/100", "F6", E.Rows[0]["U_CPHT2"].ToString());
                current_row++;
                //Chi phi NG
                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phi NG";
                if (f_ng != "")
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = "=" + f_ng.Substring(0, f_ng.Length - 3);
                //oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = E.Rows[0]["U_CPNG"].ToString(); //string.Format("={0}*{1}/100", "D6",
                current_row++;

                //Total Ho tro
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row - 4, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (current_row - 3) + ":G" + (current_row - 1));

                //NHA CUNG CAP/ NHA THAU PHU KHAC
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Cells[current_row, 1] = "VIII";
                oSheet.Cells[current_row, 3] = @"NHÀ CUNG CẤP / NHÀ THẦU PHỤ KHÁC";
                int NCC_NTP_KHAC_row = current_row;
                current_row++;
                //NHA CUNG CAP KHAC
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(155, 194, 230);
                oSheet.Cells[current_row, 1] = "VIII.1";
                oSheet.Cells[current_row, 3] = "NHÀ CUNG CẤP KHÁC";
                //oSheet.Cells[current_row, 6].Value2 = B.Compute("SUM(U_CP_NCC)", "");
                detail_row_num = current_row;
                current_row++;
                //Details
                foreach (DataRow d in C.Select("U_TYPE = 'KH'"))
                {
                    decimal tmp_ncc = 0, tmp_klhd = 0, tmp_kltt = 0, tmp_klttdd;
                    decimal.TryParse(d["U_CP_NCC"].ToString(), out tmp_ncc);
                    decimal.TryParse(d["KL_HD"].ToString(), out tmp_klhd);
                    decimal.TryParse(d["KL_TT"].ToString(), out tmp_kltt);
                    decimal.TryParse(d["KL_TT_DD"].ToString(), out tmp_klttdd);
                    if (tmp_ncc != 0) // || tmp_klhd != 0 || tmp_kltt != 0 || tmp_klttdd != 0)
                    {
                        oSheet.Cells[current_row, 2] = d["U_BPCode"];
                        oSheet.Cells[current_row, 3] = d["U_BPName"];
                        oSheet.Cells[current_row, 7] = tmp_ncc;
                        oSheet.Cells[current_row, 8] = tmp_klhd;
                        oSheet.Cells[current_row, 9] = tmp_klttdd;
                        current_row++;
                    }
                }
                //Total
                if (current_row - detail_row_num > 1)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (detail_row_num + 1) + ":G" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H" + (detail_row_num + 1) + ":H" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (detail_row_num + 1) + ":I" + (current_row - 1));
                }

                //NHA THAU PHU XAY DUNG
                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(155, 194, 230);
                oSheet.Cells[current_row, 1] = "VIII.2";
                oSheet.Cells[current_row, 3] = "NHÀ THẦU PHỤ KHÁC";
                //oSheet.Cells[current_row, 6].Value2 = B.Compute("SUM(U_CP_NTP)", "");
                detail_row_num = current_row;
                current_row++;
                //Details
                foreach (DataRow d in C.Select("U_TYPE = 'KH'"))
                {
                    decimal tmp_ntp = 0, tmp_klhd = 0, tmp_kltt = 0, tmp_klttdd;
                    decimal.TryParse(d["U_CP_NTP"].ToString(), out tmp_ntp);
                    decimal.TryParse(d["KL_HD"].ToString(), out tmp_klhd);
                    decimal.TryParse(d["KL_TT"].ToString(), out tmp_kltt);
                    decimal.TryParse(d["KL_TT_DD"].ToString(), out tmp_klttdd);
                    if (tmp_ntp != 0) //|| tmp_klhd != 0 || tmp_kltt != 0 || tmp_klttdd != 0)
                    {
                        oSheet.Cells[current_row, 2] = d["U_BPCode"];
                        oSheet.Cells[current_row, 3] = d["U_BPName"];
                        oSheet.Cells[current_row, 6] = d["GTHD"];
                        oSheet.Cells[current_row, 7] = tmp_ntp;
                        oSheet.Cells[current_row, 8] = tmp_klhd;
                        oSheet.Cells[current_row, 9] = tmp_klttdd;
                        current_row++;
                    }
                }
                //Total
                if (current_row - detail_row_num > 1)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (detail_row_num + 1) + ":G" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H" + (detail_row_num + 1) + ":H" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_row_num, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (detail_row_num + 1) + ":I" + (current_row - 1));
                }

                //Total VIII
                if (current_row - detail_row_num >= 1)
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_KHAC_row, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (NCC_NTP_KHAC_row + 1) + ":G" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_KHAC_row, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H" + (NCC_NTP_KHAC_row + 1) + ":H" + (current_row - 1));
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[NCC_NTP_KHAC_row, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (NCC_NTP_KHAC_row + 1) + ":I" + (current_row - 1));
                }
                //Total B
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[14, 6]).Formula = string.Format("=SUBTOTAL(9,{0})", "F15" + ":F" + (current_row - 1));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[14, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G15" + ":G" + (current_row - 1));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[14, 8]).Formula = string.Format("=SUBTOTAL(9,{0})", "H15" + ":H" + (current_row - 1));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[14, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I15" + ":I" + (current_row - 1));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[14, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J15" + ":J" + (current_row - 1));

                //C
                //oSheet.Range["A" + current_row, "H" + current_row].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(244, 176, 132);
                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 1] = "C";
                oSheet.Cells[current_row, 3] = "LỢI NHUẬN GỘP CỦA CÔNG TRƯỜNG A";
                //Section_RowNum.Add(current_rownum);
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "DOANH THU";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 6]).Formula = "=F6";
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "CHI PHÍ";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = "=G14";
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "LỢI NHUẬN GỘP A";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("={0}-{1}", "F" + (current_row - 2), "G" + (current_row - 1));
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "TỶ SUẤT LỢI NHUẬN GỘP A/ DOANH THU";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("={0}/{1}", "G" + (current_row - 1), "F" + (current_row - 3));
                oSheet.Range["G" + current_row].NumberFormat = "0.00%";

                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 8]).Formula = string.Format("=((F6-F8)-(G14-{0}))/(F6-F8)", "E" + NCC_NTP_KHAC_row);
                oSheet.Range["H" + current_row].NumberFormat = "0.00%";
                current_row++;

                //D
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(244, 176, 132);
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[6, 4]).Formula = string.Format("=SUM({0}:{1})", "D7", "D12");
                //oSheet.Range["A" + current_row, "H" + current_row].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 1] = "D";
                oSheet.Cells[current_row, 3] = "LỢI NHUẬN TUYỆT ĐỐI C (Bao gồm phí quản lý Công ty)";
                //Section_RowNum.Add(current_rownum);
                current_row++;

                //Chi phi quan ly cong ty
                oSheet.Range["B" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "CHI PHÍ QUẢN LÝ CÔNG TY";
                if (f_cpqlct != "")
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = "=" + f_cpqlct.Substring(0, f_cpqlct.Length - 3);
                oSheet.Range["E" + current_row].NumberFormat = "_(* #,##0_);_(* (#,##0);_(* \" - \"??_);_(@_)";
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "DOANH THU";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 6]).Formula = "=F6";
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "CHI PHÍ";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("=G14+E{0}", current_row - 2);
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "LỢI NHUẬN TUYỆT ĐỐI C";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("={0}-{1}", "F" + (current_row - 2), "G" + (current_row - 1));
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "TỶ SUẤT LỢI TUYỆT ĐỐI C/ DOANH THU";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("={0}/{1}", "G" + (current_row - 1), "F" + (current_row - 3));
                oSheet.Range["G" + current_row].NumberFormat = "0.00%";
                current_row++;
                //Hide Column
                oSheet.Range["D:D", Type.Missing].EntireColumn.Hidden = true;
                //Border
                ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A7", "K" + (current_row - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                oXL.ActiveWindow.Activate();
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            { }

        }

        System.Data.DataTable Get_Data_BCTCA(string pGoithauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_MM_FI_GET_DATA_A", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry_BaseLine", DocEntry_BaseLine);
                cmd.Parameters.AddWithValue("@Goithau_Key", pGoithauKey);
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

        System.Data.DataTable Get_Data_BCTC_DUTRU(string pGoithauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_MM_FI_GET_DATA_B", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry_BaseLine", DocEntry_BaseLine);
                cmd.Parameters.AddWithValue("@Goithau_Key", pGoithauKey);
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

        System.Data.DataTable Get_Data_BCTC_BCH(string pGoiThauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_MM_FI_GET_DATA_BCH", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry_BaseLine", DocEntry_BaseLine);
                cmd.Parameters.AddWithValue("@Goithau_Key", pGoiThauKey);
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

        private DataTable Get_lst_User_Next_LV()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_Get_Lst_Usr_LV", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry", DocEntry_BaseLine);
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
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

        private void Send_Alert()
        {
            DataTable lst = Get_lst_User_Next_LV();
            if (lst.Rows.Count > 0)
            {
                string FProject = Grid0.DataTable.GetValue("Financial Project", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                DateTime BaseLine_Date = DateTime.Parse(Grid0.DataTable.GetValue("BaseLine Date", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString());
                string Note = Grid0.DataTable.GetValue("Note", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                SAPbobsCOM.Messages msg = null;
                msg = (SAPbobsCOM.Messages)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                msg.MessageText = string.Format("BASELINE số {0} của dự án {1} đang chờ duyệt.{2}Anh chị vui lòng vào xem.", DocEntry_BaseLine, FProject, Environment.NewLine);
                msg.Subject = "Yêu cầu phê duyệt BASELINE số " + DocEntry_BaseLine.ToString();
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
                        + @"Có <b>BASELINE</b> đang chờ bạn xử lý trên hệ thống SAP<br/><br/>"
                        + @"<b>Thông tin chi tiết:</b><br/><br/>"
                        + @"P/B/BP/Dự án : <b>{0}</b><br/>"
                        //+ @"Kỳ thanh toán : <b>{1}</b><br/>"
                        + @"Số bill : <b>{1}</b><br/>"
                        + @"Ngày tạo : <b>{2}</b><br/>"
                        + @"Ghi chú : <b>{3}</b><br/><br/>"
                        + @"Đây là email được gửi tự động từ hệ thống SAP, vui lòng không trả lời lại email này. Xin cảm ơn.<br/><br/>"
                        + @"--------------<br/>"
                        + @"Trân trọng<br/>"
                        + @"SAP Business One", FProject, DocEntry_BaseLine, BaseLine_Date.ToString("dd/MM/yyyy"), Note);
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

        private void Send_Alert_Rejected()
        {
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(string.Format("Select USER_CODE, ISNULL(a.LastName,'') +' '+ ISNULL(a.MiddleName,'')+ ' '+ ISNULL(a.FirstName,'') as 'NAME',a.email--,a.empID,c.teamID,d.name "
                            + "from OHEM a inner join OUSR b on a.USERID = b.UserID "
                            + "where b.USERID in (Select UserSign from [@BASELINE] where DocEntry={0})", DocEntry_BaseLine));
            if (oR_RecordSet.RecordCount > 0)
            {
                SAPbobsCOM.Messages msg = null;
                msg = (SAPbobsCOM.Messages)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                msg.MessageText = string.Format("BASELINE số {0} đã bị từ chối", DocEntry_BaseLine);
                msg.Subject = "BASELINE số " + DocEntry_BaseLine.ToString() + " bị từ chối";
                msg.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;

                msg.Recipients.SetCurrentLine(0);
                msg.Recipients.UserCode = oR_RecordSet.Fields.Item("USER_CODE").Value.ToString();
                msg.Recipients.NameTo = oR_RecordSet.Fields.Item("NAME").Value.ToString();
                msg.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                if (oR_RecordSet.Fields.Item("email").Value.ToString() != "" && !string.IsNullOrEmpty(oR_RecordSet.Fields.Item("email").Value.ToString()))
                    Send_Email(oR_RecordSet.Fields.Item("email").Value.ToString(), msg.Recipients.NameTo, msg.MessageText, msg.Subject);
                int kq = msg.Add();
                if (kq != 0)
                    oApp.SetStatusBarMessage("Send Email Result: " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private bool Check_Ketoan_truong(string pUsrName)
        {
            string sql = string.Format("Select position,dept from OHEM  where userID = (Select t.USERID from OUSR t where t.User_Code='{0}')", pUsrName);
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(sql);
            if (oR_RecordSet.RecordCount > 0)
            {
                string position = oR_RecordSet.Fields.Item("position").Value.ToString();
                string dept = oR_RecordSet.Fields.Item("dept").Value.ToString();
                if ((position == "1" && dept == "-2") || dept == "16") return true;
                else return false;
            }
            return false;
        }

        private bool Check_CCM(string pUsrName)
        {
            string sql = string.Format("Select position,dept from OHEM  where userID = (Select t.USERID from OUSR t where t.User_Code='{0}')", pUsrName);
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(sql);
            if (oR_RecordSet.RecordCount > 0)
            {
                string position = oR_RecordSet.Fields.Item("position").Value.ToString();
                string dept = oR_RecordSet.Fields.Item("dept").Value.ToString();
                if (dept == "-") return true;
                else return false;
            }
            return false;
        }

        //Approve Button
        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_Approve_LV", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                cmd.Parameters.AddWithValue("@DocEntry", DocEntry_BaseLine);
                cmd.Parameters.AddWithValue("@Status", "1");
                cmd.Parameters.AddWithValue("@Comment", EditText0.Value.Trim());

                conn.Open();
                int rowupdate = cmd.ExecuteNonQuery();
                if (rowupdate >= 1)
                {
                    SAPbobsCOM.GeneralService oGeneralService = null;
                    SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                    SAPbobsCOM.CompanyService sCmp = null;
                    SAPbobsCOM.GeneralData oGeneralData = null;
                    SAPbobsCOM.GeneralData oChild = null;
                    SAPbobsCOM.GeneralDataCollection oChildren = null;

                    oApp.MessageBox("Phê duyệt thành công");
                    //Add More Level Approve
                    sCmp = oCompany.GetCompanyService();
                    oGeneralService = sCmp.GetGeneralService("BaseLine");
                    oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oGeneralParams.SetProperty("DocEntry", DocEntry_BaseLine);
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                    oChildren = oGeneralData.Child("BASELINE_APPR");
                    if (oChildren.Count == 4)
                    {
                        oChild = oChildren.Add();
                        oChild.SetProperty("U_Level", "-2");
                        oChild.SetProperty("U_Posistion", "2");
                        oChild.SetProperty("U_DeptName", "Kế toán");
                        oChild.SetProperty("U_PosName", "Nhân viên");

                        oChild = oChildren.Add();
                        oChild.SetProperty("U_Level", "-2");
                        oChild.SetProperty("U_Posistion", "1");
                        oChild.SetProperty("U_DeptName", "Kế toán");
                        oChild.SetProperty("U_PosName", "Trưởng phòng");

                        oGeneralService.Update(oGeneralData);
                    }
                    if (Check_Ketoan_truong(oCompany.UserName))
                    {
                        oGeneralService.Close(oGeneralParams);
                    }
                    //Gui tin nhan
                    Send_Alert();
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
                cmd = new SqlCommand("BASELINE_Approve_LV", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                cmd.Parameters.AddWithValue("@DocEntry", DocEntry_BaseLine);
                cmd.Parameters.AddWithValue("@Status", "2");
                cmd.Parameters.AddWithValue("@Comment", EditText0.Value.Trim());

                conn.Open();
                int rowupdate = cmd.ExecuteNonQuery();
                if (rowupdate >= 1)
                {
                    oApp.MessageBox("Reject thành công");
                    //SAPbobsCOM.GeneralService oGeneralService = null;
                    //SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                    //SAPbobsCOM.CompanyService sCmp = null;
                    //sCmp = oCompany.GetCompanyService();
                    //oGeneralService = sCmp.GetGeneralService("KLTT");
                    //oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    //oGeneralParams.SetProperty("DocEntry", DocEntry);
                    //oGeneralService.Cancel(oGeneralParams);
                    ////Send Alert
                    Send_Alert_Rejected();
                }
                else
                {
                    oApp.MessageBox("Phê duyệt không thành công!");
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

        //Approve and request to CEO
        private void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_Approve_LV", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                cmd.Parameters.AddWithValue("@DocEntry", DocEntry_BaseLine);
                cmd.Parameters.AddWithValue("@Status", "1");
                cmd.Parameters.AddWithValue("@Comment", EditText0.Value.Trim());

                conn.Open();
                int rowupdate = cmd.ExecuteNonQuery();
                if (rowupdate >= 1)
                {
                    SAPbobsCOM.GeneralService oGeneralService = null;
                    SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                    SAPbobsCOM.CompanyService sCmp = null;
                    SAPbobsCOM.GeneralData oGeneralData = null;
                    SAPbobsCOM.GeneralData oChild = null;
                    SAPbobsCOM.GeneralDataCollection oChildren = null;

                    oApp.MessageBox("Phê duyệt thành công");
                    //Add More Level Approve
                    sCmp = oCompany.GetCompanyService();
                    oGeneralService = sCmp.GetGeneralService("BaseLine");
                    oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oGeneralParams.SetProperty("DocEntry", DocEntry_BaseLine);
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                    oChildren = oGeneralData.Child("BASELINE_APPR");
                    if (oChildren.Count == 4 && Check_CCM(oCompany.UserName))
                    {
                        oChild = oChildren.Add();
                        oChild.SetProperty("U_Level", "6");
                        oChild.SetProperty("U_Posistion", "10");
                        oChild.SetProperty("U_DeptName", "Ban giám đốc");
                        oChild.SetProperty("U_PosName", "Tổng giám đốc");

                        //oChild = oChildren.Add();
                        //oChild.SetProperty("U_Level", "-2");
                        //oChild.SetProperty("U_Posistion", "2");
                        //oChild.SetProperty("U_DeptName", "Kế toán");
                        //oChild.SetProperty("U_PosName", "Nhân viên");

                        //oChild = oChildren.Add();
                        //oChild.SetProperty("U_Level", "-2");
                        //oChild.SetProperty("U_Posistion", "1");
                        //oChild.SetProperty("U_DeptName", "Kế toán");
                        //oChild.SetProperty("U_PosName", "Trưởng phòng");

                        oGeneralService.Update(oGeneralData);
                    }
                    if (Check_Ketoan_truong(oCompany.UserName))
                    {
                        oGeneralService.Close(oGeneralParams);
                    }
                    //Gui tin nhan
                    Send_Alert();
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


    }
}
