using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;

namespace R_BCTC
{
    [FormAttribute("R_BCTC.MM_ITEM_CON", "MM_ITEM_CON.b1f")]
    class MM_ITEM_CON : UserFormBase
    {
        public MM_ITEM_CON()
        {
        }
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("txt_fpro").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_8").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("grd_sub").Specific));
            this.Grid0.PressedAfter += new SAPbouiCOM._IGridEvents_PressedAfterEventHandler(this.Grid0_PressedAfter);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_date").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
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
            Load_Financial_Project();
        }

        void Load_Financial_Project()
        {
            System.Data.DataTable tb_fprj = Get_List_FProject();
            if (tb_fprj.Rows.Count > 0)
            {
                foreach (DataRow r in tb_fprj.Rows)
                {
                    ComboBox0.ValidValues.Add(r["PrjCode"].ToString(), r["PrjName"].ToString());
                }
            }
        }

        private SAPbouiCOM.DataTable Convert_SAP_DataTable(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_SubList"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_SubList");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_SubList");
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

        System.Data.DataTable Get_List_FProject()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("MM_FI_GET_FPROJECT", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Username", oCompany.UserName);
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

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.Button Button1;

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            Grid0.DataTable.ExecuteQuery(string.Format("SELECT 'N' as 'Checked',AbsEntry,NAME FROM OPMG T0 WHERE T0.[FIPROJECT] = '{0}'and T0.[STATUS] <> 'T' ORDER BY AbsEntry", ComboBox0.Selected.Value));
            Grid0.Columns.Item(1).Editable = false;
            Grid0.Columns.Item(2).Editable = false;
            Grid0.Columns.Item("Checked").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
            Grid0.AutoResizeColumns();
        }

        System.Data.DataTable Get_Data(string pFinancialProject, DateTime pToDate, string pGoiThauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("MM_QC_ITEM", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                cmd.Parameters.AddWithValue("@GoiThauKey", pGoiThauKey);
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

        System.Data.DataTable Get_Data_HM(string pFinancialProject, DateTime pToDate, string pGoiThauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("MM_QC_ITEM_HM", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                cmd.Parameters.AddWithValue("@GoiThauKey", pGoiThauKey);
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

        private void Grid0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            if (pVal.ColUID == "Checked")
            {
                if (pVal.Row >= 0)
                {
                    if (Grid0.DataTable.GetValue("Checked", pVal.Row).ToString() == "Y")
                    {
                        Grid0.Rows.SelectedRows.Add(pVal.Row);
                    }
                    else
                    {
                        Grid0.Rows.SelectedRows.Remove(pVal.Row);
                    }
                }
            }
        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

            //throw new System.NotImplementedException();
            string GoiThau_Key = "";
            string GoiThau_Name = "";
            for (int i = 0; i < Grid0.Rows.Count; i++)
            {
                string IsSelected = Grid0.DataTable.GetValue("Checked", i).ToString();
                if (IsSelected == "Y")
                {
                    GoiThau_Key += Grid0.DataTable.GetValue("AbsEntry", i).ToString() + ",";
                    GoiThau_Name = Grid0.DataTable.GetValue("NAME", i).ToString();
                    //oApp.MessageBox(Grid0.DataTable.GetValue("AbsEntry", i).ToString());
                }
            }
            if (GoiThau_Key.Length > 0)
                GoiThau_Key = GoiThau_Key.Substring(0, GoiThau_Key.Length - 1);
            string FProject = ComboBox0.Selected.Value;
            string tmp_d = EditText0.Value;
            if (string.IsNullOrEmpty(tmp_d))
            {
                oApp.MessageBox("Please Enter To Date !");
                return;
            }
            DateTime ToDate = DateTime.ParseExact(tmp_d, "yyyyMMdd", CultureInfo.InvariantCulture);
            DataTable rs = Get_Data(FProject, ToDate, GoiThau_Key);
            if (rs.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application oXL;
                Microsoft.Office.Interop.Excel._Workbook oWB;
                Microsoft.Office.Interop.Excel._Worksheet oSheet;
                Microsoft.Office.Interop.Excel.Range oRng;

                object misvalue = System.Reflection.Missing.Value;
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;
                //Open Template
                oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_QC_ITEM.xlsx");
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Project Name
                oSheet.Cells[2, 4] = "Dự án: " + this.ComboBox0.Selected.Description;
                //SubProject
                if (!string.IsNullOrEmpty(GoiThau_Key))
                {
                    if (GoiThau_Key.Split(',').Count() == 1)
                        oSheet.Cells[3, 4] = "Gói thầu: " + GoiThau_Name;
                }
                //Todate
                oSheet.Cells[4, 4] = "Đến ngày: " + ToDate.ToString("dd/MM/yyyy");
                int current_row = 6;
                for (int i = 0; i < rs.Rows.Count; i++)
                {
                    //STT
                    oSheet.Cells[current_row, 1] = (i + 1);
                    //Ma vat tu
                    oSheet.Cells[current_row, 2] = rs.Rows[i]["U_ITEMNO"].ToString();
                    //Ten vat tu
                    oSheet.Cells[current_row, 3] = rs.Rows[i]["U_ITEMNAME"].ToString();
                    //DVT
                    oSheet.Cells[current_row, 4] = rs.Rows[i]["U_DVT"].ToString();
                    //KL BoQ
                    oSheet.Cells[current_row, 5].Value2 = rs.Rows[i]["KL_BOQ"];
                    //KL Ban ve
                    oSheet.Cells[current_row, 6].Value2 = rs.Rows[i]["KL_BV"];
                    //KL NCC
                    oSheet.Cells[current_row, 7].Value2 = rs.Rows[i]["KL_DN"];
                    //DVT NCC
                    oSheet.Cells[current_row, 8] = rs.Rows[i]["DVT_NCC"].ToString();
                    //KL Con Lai
                    oSheet.Cells[current_row, 9].Formula = string.Format("=F{0}-G{0}", current_row);
                    //KL hao hut
                    oSheet.Cells[current_row, 10].Formula = string.Format("=IF(ISERROR((G{0}-F{0})/G{0}),0,(G{0}-F{0})/G{0})", current_row);
                    current_row++;
                }
                //Border
                ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A6", "K" + (current_row - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                oXL.ActiveWindow.Activate();
            }
        }

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string GoiThau_Key = "";
            string GoiThau_Name = "";
            for (int i = 0; i < Grid0.Rows.Count; i++)
            {
                string IsSelected = Grid0.DataTable.GetValue("Checked", i).ToString();
                if (IsSelected == "Y")
                {
                    GoiThau_Key += Grid0.DataTable.GetValue("AbsEntry", i).ToString() + ",";
                    GoiThau_Name = Grid0.DataTable.GetValue("NAME", i).ToString();
                    //oApp.MessageBox(Grid0.DataTable.GetValue("AbsEntry", i).ToString());
                }
            }
            if (GoiThau_Key.Length > 0)
                GoiThau_Key = GoiThau_Key.Substring(0, GoiThau_Key.Length - 1);
            string FProject = ComboBox0.Selected.Value;
            string tmp_d = EditText0.Value;
            if (string.IsNullOrEmpty(tmp_d))
            {
                oApp.MessageBox("Please Enter To Date !");
                return;
            }
            DateTime ToDate = DateTime.ParseExact(tmp_d, "yyyyMMdd", CultureInfo.InvariantCulture);
            DataTable rs = Get_Data_HM(FProject, ToDate, GoiThau_Key);
            if (rs.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application oXL;
                Microsoft.Office.Interop.Excel._Workbook oWB;
                Microsoft.Office.Interop.Excel._Worksheet oSheet;
                Microsoft.Office.Interop.Excel.Range oRng;

                object misvalue = System.Reflection.Missing.Value;
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;
                //Open Template
                oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_QC_ITEM_HM.xlsx");
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Project Name
                oSheet.Cells[2, 4] = "Dự án: " + this.ComboBox0.Selected.Description;
                //SubProject
                if (!string.IsNullOrEmpty(GoiThau_Key))
                {
                    if (GoiThau_Key.Split(',').Count() == 1)
                        oSheet.Cells[3, 4] = "Gói thầu: " + GoiThau_Name;
                }
                //Todate
                oSheet.Cells[4, 4] = "Đến ngày: " + ToDate.ToString("dd/MM/yyyy");
                int current_row = 6;
                string HM_Key = "";
                int STT = 1;
                for (int i = 0; i < rs.Rows.Count; i++)
                {
                    if (rs.Rows[i]["HM_Key"].ToString() != HM_Key)
                    {
                        //Ma HM
                        oSheet.Cells[current_row, 2] = rs.Rows[i]["HM_Code"].ToString();
                        //Ten HM
                        oSheet.Cells[current_row, 3] = rs.Rows[i]["HM_Name"].ToString();
                        //Format style
                        //oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                        oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                        current_row++;

                        //Cong tac
                        string CT_Key = "";
                        foreach (DataRow r in rs.Select("HM_Key=" + rs.Rows[i]["HM_Key"].ToString()))
                        {
                            if (CT_Key != r["CT_Key"].ToString())
                            {
                                STT = 1;
                                //Ma CT
                                oSheet.Cells[current_row, 2] = r["CT_Code"].ToString();
                                //Ten CT
                                oSheet.Cells[current_row, 3] = r["CT_Name"].ToString();
                                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(248, 203, 173); //(248, 203, 173)
                                current_row++;
                                //CV
                                foreach (DataRow r1 in rs.Select("HM_Key=" + rs.Rows[i]["HM_Key"].ToString() + " and CT_Key=" + r["CT_Key"].ToString()))
                                {
                                    //STT
                                    oSheet.Cells[current_row, 1].Value2 = STT;
                                    //Ma CV
                                    oSheet.Cells[current_row, 2] = r1["CV_Code"].ToString();
                                    //Ten CV
                                    oSheet.Cells[current_row, 3] = r1["CV_Name"].ToString();
                                    //DVT
                                    oSheet.Cells[current_row, 4] = r1["CV_DVT"].ToString();
                                    //KLDT
                                    oSheet.Cells[current_row, 5].Value2 = r1["CV_KLDT"];
                                    //KLBV
                                    oSheet.Cells[current_row, 6].Value2 = r1["CV_KLBV"];
                                    //KL Nha thau phu
                                    oSheet.Cells[current_row, 7].Value2 = r1["NTP"];
                                    //Don vi NTP
                                    oSheet.Cells[current_row, 8].Value2 = r1["DV_NTP"].ToString();
                                    //Doi thi cong
                                    oSheet.Cells[current_row, 9].Value2 = r1["DTC"];
                                    //Don vi DTC
                                    oSheet.Cells[current_row, 10].Value2 = r1["DV_DTC"].ToString();
                                    STT++;
                                    current_row++;
                                }
                                CT_Key = r["CT_Key"].ToString();
                            }
 
                        }
                        HM_Key = rs.Rows[i]["HM_Key"].ToString();
                    }
                    ////STT
                    //oSheet.Cells[current_row, 1] = (i + 1);
                    ////Ten vat tu
                    //oSheet.Cells[current_row, 2] = rs.Rows[i]["U_ITEMNAME"].ToString();
                    ////DVT
                    //oSheet.Cells[current_row, 3] = rs.Rows[i]["U_DVT"].ToString();
                    ////KL BoQ
                    //oSheet.Cells[current_row, 4] = rs.Rows[i]["KL_BOQ"].ToString();
                    ////KL Ban ve
                    //oSheet.Cells[current_row, 5] = rs.Rows[i]["KL_BV"].ToString();
                    ////KL Du an nhap
                    //oSheet.Cells[current_row, 6] = rs.Rows[i]["KL_DN"].ToString();
                    ////KL Con Lai
                    //oSheet.Cells[current_row, 7].Formula = string.Format("=E{0}-F{0}", current_row);
                    ////KL hao hut
                    //oSheet.Cells[current_row, 8].Formula = string.Format("=IF(ISERROR((F{0}-E{0})/F{0}),0,(F{0}-E{0})/F{0})", current_row);
                    //current_row++;
                }
                //Border
                ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A6", "K" + (current_row - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                oXL.ActiveWindow.Activate();
            }
        }
    }
}
