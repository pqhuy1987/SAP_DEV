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
    [FormAttribute("R_BCTC.CCM_SANLUONG", "CCM_SANLUONG.b1f")]
    class CCM_SANLUONG : UserFormBase
    {
        public CCM_SANLUONG()
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
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.OptionBtn0 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_dt").Specific));
            this.OptionBtn0.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn0_PressedAfter);
            this.OptionBtn1 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_ct").Specific));
            this.OptionBtn1.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn1_PressedAfter);
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_ct").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_dt").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_frd").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txt_tod").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_8").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {
            OptionBtn1.GroupWith("op_dt");
            OptionBtn0.Selected = true;
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
            Load_DT();
            Load_CT();
        }

        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.OptionBtn OptionBtn0;
        private SAPbouiCOM.OptionBtn OptionBtn1;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.ComboBox ComboBox1;

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            

        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.Button Button0;

        private void OptionBtn0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn0.Selected)
            {
                ComboBox1.Item.Enabled = true;
                ComboBox1.Item.Click();
                ComboBox0.Item.Enabled = false;
                
            }
        }

        private void OptionBtn1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn1.Selected)
            {
                ComboBox0.Item.Enabled = true;
                ComboBox0.Item.Click();
                ComboBox1.Item.Enabled = false;
            }

        }

        void Load_DT()
        {
            System.Data.DataTable tb_dt = Get_List_DT();
            if (tb_dt.Rows.Count > 0)
            {
                foreach (DataRow r in tb_dt.Rows)
                {
                    ComboBox1.ValidValues.Add(r["CardCode"].ToString(), r["CardName"].ToString());
                }
                ComboBox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            }
        }
        void Load_CT()
        {
            System.Data.DataTable tb_ct = Get_List_CT();
            if (tb_ct.Rows.Count > 0)
            {
                foreach (DataRow r in tb_ct.Rows)
                {
                    ComboBox0.ValidValues.Add(r["U_001"].ToString(), r["NAME"].ToString());
                }
                ComboBox0.ExpandType = SAPbouiCOM.BoExpandType.et_ValueDescription;
            }
        }
        System.Data.DataTable Get_List_DT()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_GET_LST_DT", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.AddWithValue("@Username", oCompany.UserName);
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

        System.Data.DataTable Get_List_CT()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_GET_LST_CT", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.AddWithValue("@Username", oCompany.UserName);
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

        System.Data.DataTable Get_Data_CT(DateTime pFrDate, DateTime pToDate, string pCT)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_SANLUONG_DATA_CT", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FrDate", pFrDate);
                cmd.Parameters.AddWithValue("@ToDate", pToDate);
                cmd.Parameters.AddWithValue("@CT", pCT);
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

        System.Data.DataTable Get_Data_DT(DateTime pFrDate, DateTime pToDate, string pBpCode )
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_SANLUONG_DATA_DT", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FrDate", pFrDate);
                cmd.Parameters.AddWithValue("@ToDate", pToDate);
                cmd.Parameters.AddWithValue("@BpCode", pBpCode);
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

        private decimal Get_GTHD(DateTime pFrDate, DateTime pToDate, string pBpCode, string pFProject)
        {
            decimal GTHD = 0;
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_SANLUONG_GTHD", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FrDate", pFrDate);
                cmd.Parameters.AddWithValue("@ToDate", pToDate);
                cmd.Parameters.AddWithValue("@BpCode", pBpCode);
                cmd.Parameters.AddWithValue("@FProject", pFProject);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                if (result.Rows.Count > 0)
                {
                    decimal.TryParse(result.Rows[0]["GTHD"].ToString(),out GTHD);
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
            return GTHD;
        }

        private decimal Get_GTHD_CT(DateTime pFrDate, DateTime pToDate, string pBpCode, string pFProject, string pPUType)
        {
            decimal GTHD = 0;
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_SANLUONG_GTHD_CT", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FrDate", pFrDate);
                cmd.Parameters.AddWithValue("@ToDate", pToDate);
                cmd.Parameters.AddWithValue("@BpCode", pBpCode);
                cmd.Parameters.AddWithValue("@FProject", pFProject);
                cmd.Parameters.AddWithValue("@PuType", pPUType);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                if (result.Rows.Count > 0)
                {
                    decimal.TryParse(result.Rows[0]["GTHD"].ToString(), out GTHD);
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
            return GTHD;
        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn0.Selected)
            {
                //Bao cao theo doi tuong
                if (ComboBox1.Selected.Value == "")
                {
                    oApp.MessageBox("Please select Business Partner !");
                    return;
                }
                else
                {
                    DateTime FrDate = DateTime.ParseExact(EditText0.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                    DateTime ToDate = DateTime.ParseExact(EditText1.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                    DataTable tb_data = Get_Data_DT(FrDate, ToDate, ComboBox1.Selected.Value);
                    if (tb_data.Rows.Count > 0)
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
                        oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_SANLUONG_DT.xlsx");
                        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                        int STT = 1;
                        int current_row = 5;
                        int group_row = 5;
                        string Fproject = "";
                        decimal total_GTHD = 0;
                        string BpCode = "";
                        oSheet.Cells[3, 1] = string.Format("Thời gian: Từ ngày {0} đến ngày {1}", FrDate.ToString("dd/MM/yyyy"), ToDate.ToString("dd/MM/yyyy"));

                        foreach (DataRow r in tb_data.Rows)
                        {
                            if (Fproject != r["Project"].ToString())
                            {
                                
                                STT = 1;
                                //Format
                                oSheet.Range["A" + current_row, "J" + current_row].Font.Bold = true;
                                oSheet.Range["A" + current_row, "J" + current_row].Interior.Color = System.Drawing.Color.FromArgb(248, 203, 173);
                                //Total
                                if (Fproject != "")
                                {
                                    oSheet.Cells[group_row, 7] = total_GTHD;
                                    oSheet.Cells[group_row, 9].Formula = string.Format("=G{0} - SUM(H{1}:H{2})", group_row, group_row + 1, current_row - 1);
                                }
                                group_row = current_row;
                                Fproject = r["Project"].ToString();
                                BpCode = "";
                                total_GTHD = 0;
                                current_row++;

                            }
                            //Details
                            oSheet.Cells[current_row, 1] = STT;
                            oSheet.Cells[current_row, 2] = r["Project"];
                            oSheet.Cells[current_row, 3] = r["MA_CT"];
                            oSheet.Cells[current_row, 4] = r["TEN_CT"];
                            oSheet.Cells[current_row, 5] = r["CardName"];
                            oSheet.Cells[current_row, 6] = r["NDT"];
                            if (BpCode != r["CardCode"].ToString())
                            {
                                BpCode=r["CardCode"].ToString();
                                decimal tmp = Get_GTHD(FrDate, ToDate, BpCode, Fproject);
                                total_GTHD += tmp;
                                //oSheet.Cells[current_row, 7] = tmp;
                            }
                            oSheet.Cells[current_row, 8] = r["Total"];
                            STT++;
                            current_row++;
                        }
                        //Total
                        if (Fproject != "")
                        {
                            oSheet.Cells[group_row, 7] = total_GTHD;
                            oSheet.Cells[group_row, 9].Formula = string.Format("=G{0} - SUM(H{1}:H{2})", group_row, group_row + 1, current_row - 1);
                        }
                        //Total
                        oSheet.Range["A" + current_row, "F" + current_row].Merge(false);
                        oSheet.Range["A" + current_row, "F" + current_row].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        oSheet.Range["A" + current_row, "J" + current_row].Font.Bold = true;
                        oSheet.Range["A" + current_row, "F" + current_row].Value2 = "TỔNG CỘNG";
                        oSheet.Cells[current_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1}",5,current_row-1);
                        oSheet.Cells[current_row, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1}", 5, current_row - 1);
                        oSheet.Cells[current_row, 9].Formula = string.Format("=SUBTOTAL(9,I{0}:I{1}", 5, current_row - 1);
                        current_row++;

                        //Border
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A5", "J" + (current_row - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        oXL.ActiveWindow.Activate();

                    }
                    else
                    {
                        oApp.MessageBox("No records !");
                    }
 
                }

            }
            else if (OptionBtn1.Selected)
            {
                //Bao cao theo cong tac
                if (ComboBox0.Selected.Value == "")
                {
                    oApp.MessageBox("Please select value !");
                    return;
                }
                else
                {
                    DateTime FrDate = DateTime.ParseExact(EditText0.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                    DateTime ToDate = DateTime.ParseExact(EditText1.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                    DataTable tb_data = Get_Data_CT(FrDate, ToDate, ComboBox0.Selected.Value);
                    if (tb_data.Rows.Count > 0)
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
                        oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_SANLUONG_CT.xlsx");
                        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                        int STT = 1;
                        int current_row = 5;
                        decimal total_GTHD = 0;
                        string BpCode = "";
                        oSheet.Cells[3, 1] = string.Format("Thời gian: Từ ngày {0} đến ngày {1}", FrDate.ToString("dd/MM/yyyy"), ToDate.ToString("dd/MM/yyyy"));

                        foreach (DataRow r in tb_data.Rows)
                        {
                            //Details
                            oSheet.Cells[current_row, 1] = STT;
                            oSheet.Cells[current_row, 2] = r["Project"];
                            oSheet.Cells[current_row, 3] = r["MA_CT"];
                            oSheet.Cells[current_row, 4] = r["TEN_CT"];
                            oSheet.Cells[current_row, 5] = r["CardName"];
                            string NhomDoitruong = "";
                            if (r["NDT"].ToString() == "PUT01") NhomDoitruong = "NCC";
                            else if (r["NDT"].ToString() == "PUT02") NhomDoitruong = "NTP";
                            else if (r["NDT"].ToString() == "PUT09") NhomDoitruong = "DTC";
                            oSheet.Cells[current_row, 6] = NhomDoitruong;
                            BpCode = r["CardCode"].ToString();
                            decimal tmp = Get_GTHD_CT(FrDate, ToDate, BpCode, r["Project"].ToString(), r["NDT"].ToString());
                            total_GTHD += tmp;
                            oSheet.Cells[current_row, 7] = tmp;
                            oSheet.Cells[current_row, 8] = r["Total"];
                            oSheet.Cells[current_row, 9].Formula = string.Format("=G{0} - H{1}", current_row, current_row);
                            STT++;
                            current_row++;
                        }
                        //Total
                        //if (Fproject != "")
                        //{
                        //    oSheet.Cells[group_row, 7] = total_GTHD;
                        //    oSheet.Cells[group_row, 9].Formula = string.Format("=G{0} - SUM(H{1}:H{2})", group_row, group_row + 1, current_row - 1);
                        //}
                        //Total
                        oSheet.Range["A" + current_row, "F" + current_row].Merge(false);
                        oSheet.Range["A" + current_row, "F" + current_row].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        oSheet.Range["A" + current_row, "J" + current_row].Font.Bold = true;
                        oSheet.Range["A" + current_row, "F" + current_row].Value2 = "TỔNG CỘNG";
                        oSheet.Cells[current_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1}", 5, current_row - 1);
                        oSheet.Cells[current_row, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1}", 5, current_row - 1);
                        oSheet.Cells[current_row, 9].Formula = string.Format("=SUBTOTAL(9,I{0}:I{1}", 5, current_row - 1);
                        current_row++;

                        //Border
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A5", "J" + (current_row - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        oXL.ActiveWindow.Activate();

                    }
                    else
                    {
                        oApp.MessageBox("No records !");
                    }
 
                }
            }

        }
    }
}
