using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Globalization;
using System.Data;

namespace R_BCTC
{
    [FormAttribute("R_BCTC.CCM_DT_LN_LIST", "CCM_DT_LN_LIST.b1f")]
    class CCM_DT_LN_LIST : UserFormBase
    {
        public CCM_DT_LN_LIST()
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
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_fr").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txt_to").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_view").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.StaticText StaticText0;

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

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.Button Button0;

        System.Data.DataTable Get_Data(DateTime pFrDate, DateTime pToDate)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_DT_LN_Project_List", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FrDate", pFrDate);
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

        System.Data.DataTable Get_Data_CURRENT(string pFProject, int pProjectID)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_CURRENT_A_INDEX", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FProject", pFProject);
                cmd.Parameters.AddWithValue("@ProjectID", pProjectID);
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

        //System.Data.DataTable Get_Data_BASELINE(int pBASELINE_DocEntry, int pProjectID)
        //{
        //    DataTable result = new DataTable();
        //    SqlCommand cmd = null;
        //    try
        //    {
        //        cmd = new SqlCommand("CCM_BASELINE_A_INDEX", conn);
        //        cmd.CommandType = CommandType.StoredProcedure;
        //        cmd.Parameters.AddWithValue("@BASELINE_DocEntry", pBASELINE_DocEntry);
        //        cmd.Parameters.AddWithValue("@ProjectID", pProjectID);
        //        //cmd.Parameters.AddWithValue("@ProjectID", pProjectID);
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

        System.Data.DataTable Get_Data_BASELINE(int pBASELINE_DocEntry, string pFProject, int pProjectID)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_BASELINE_A_INDEX", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@BASELINE_DocEntry", pBASELINE_DocEntry);
                cmd.Parameters.AddWithValue("@FinancialProject", pFProject);
                cmd.Parameters.AddWithValue("@ProjectID", pProjectID);
                //cmd.Parameters.AddWithValue("@ProjectID", pProjectID);
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

        System.Data.DataTable Get_Data_BASELINE_Date(int pBASELINE_DocEntry, string pFProject, int pProjectID, DateTime pToDate)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_BASELINE_DATE_A_INDEX", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@BASELINE_DocEntry", pBASELINE_DocEntry);
                cmd.Parameters.AddWithValue("@FinancialProject", pFProject);
                cmd.Parameters.AddWithValue("@ProjectID", pProjectID);
                cmd.Parameters.AddWithValue("@ToDate", pToDate);
                //cmd.Parameters.AddWithValue("@ProjectID", pProjectID);
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

        System.Data.DataTable Get_A_Index_Fproject(int pBASELINE_DocEntry, string pFProject)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_BASELINE_FPROJECT_A_INDEX", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@BASELINE_DocEntry", pBASELINE_DocEntry);
                cmd.Parameters.AddWithValue("@FinancialProject", pFProject);
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

        System.Data.DataTable Get_A_Index_Fproject_Date(int pBASELINE_DocEntry, string pFProject, DateTime pToDate)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_BASELINE_FPROJECT_DATE_A_INDEX", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@BASELINE_DocEntry", pBASELINE_DocEntry);
                cmd.Parameters.AddWithValue("@FinancialProject", pFProject);
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

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (string.IsNullOrEmpty(EditText0.Value) || string.IsNullOrEmpty(EditText1.Value))
            {
                oApp.MessageBox("Please enter value !");
                return;
            }
            try
            {
                DateTime frdate = DateTime.ParseExact(EditText0.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                DateTime todate = DateTime.ParseExact(EditText1.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                DataTable rs_detail = Get_Data(frdate,todate);

                if (rs_detail.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Application oXL;
                    Microsoft.Office.Interop.Excel._Workbook oWB;
                    Microsoft.Office.Interop.Excel._Worksheet oSheet;

                    object misvalue = System.Reflection.Missing.Value;
                    //Start Excel and get Application object.
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oXL.Visible = true;
                    //Open Template
                    oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_DTLN_THONGKE.xlsx");
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                    int current_row = 4;
                    int current_group = 4;
                    int STT = 1;
                    string FProject = "";
                    oSheet.Cells[2, 3] = string.Format("Từ ngày {0} đến ngày {1}",frdate.ToString("dd/MM/yyyy"),todate.ToString("dd/MM/yyyy"));

                    foreach (DataRow r in rs_detail.Rows)
                    {
                        if (FProject != r["PrjCode"].ToString())
                        {
                            //Print subtotal PreGroup
                            if (FProject != "")
                                oSheet.Cells[current_group, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1}", current_group+1, current_row - 1);
                            //Print FProject Group
                            current_group = current_row;
                            oSheet.Cells[current_row, 1] = STT;
                            oSheet.Cells[current_row, 2] = r["PrjName"];
                            oSheet.Range["A" + current_row, "J" + current_row].Interior.Color = System.Drawing.Color.FromArgb(252, 228, 214);
                            //Get A-Index Baseline
                            DataTable rs_baseline_f = Get_A_Index_Fproject_Date(-1, r["PrjCode"].ToString(), frdate);
                            if (rs_baseline_f.Rows.Count >= 1)
                                oSheet.Cells[current_row, 8] = rs_baseline_f.Rows[0]["A-INDEX"];
                            //Get A-Index Current
                            DataTable rs_current_f = Get_A_Index_Fproject_Date(-1, r["PrjCode"].ToString(), todate);
                            if (rs_current_f.Rows.Count >= 1)
                                oSheet.Cells[current_row, 9] = rs_current_f.Rows[0]["A-INDEX"];
                            current_row++;
                            STT++;
                            FProject = r["PrjCode"].ToString();
                        }
                        //oSheet.Cells[current_row, 1] = STT;
                        //oSheet.Cells[current_row, 2] = r["PrjName"];
                        oSheet.Cells[current_row, 3] = r["NAME"];
                        oSheet.Cells[current_row, 4] = r["CARDNAME"];
                        oSheet.Cells[current_row, 5] = r["PRJGROUP"];
                        oSheet.Cells[current_row, 6] = r["PRJTYPE"];
                        oSheet.Cells[current_row, 7] = r["GTHD"];
                        //Get A-Index Baseline
                        DataTable rs_baseline = Get_Data_BASELINE_Date(-1, r["PrjCode"].ToString(),int.Parse(r["AbsEntry"].ToString()),frdate);
                        if (rs_baseline.Rows.Count >= 1)
                            oSheet.Cells[current_row, 8] = rs_baseline.Rows[0]["A-INDEX"];
                        //Get A-Index Current
                        DataTable rs_current = Get_Data_BASELINE_Date(-1, r["PrjCode"].ToString(), int.Parse(r["AbsEntry"].ToString()), todate);
                        if (rs_current.Rows.Count >= 1)
                            oSheet.Cells[current_row, 9] = rs_current.Rows[0]["A-INDEX"];
                        current_row++;
                    }
                    oSheet.Cells[current_group, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1}", current_group + 1, current_row - 1);
                    //Border
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A4", "J" + (current_row - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    oXL.ActiveWindow.Activate();
                }
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }

        }
    }
}
