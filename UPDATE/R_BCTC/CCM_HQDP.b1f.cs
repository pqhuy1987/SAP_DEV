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
    [FormAttribute("R_BCTC.CCM_HQDP", "CCM_HQDP.b1f")]
    class CCM_HQDP : UserFormBase
    {
        public CCM_HQDP()
        {
        }
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.Button Button0;

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_frd").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txt_tod").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
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
            SAPbobsCOM.Recordset or_RecoderSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            or_RecoderSet.DoQuery("Select * from [@ADDONCFG]");
            if (or_RecoderSet.RecordCount > 0)
            {
                string uid = or_RecoderSet.Fields.Item("Code").Value.ToString();
                string pwd = or_RecoderSet.Fields.Item("Name").Value.ToString();
                conn = new SqlConnection(string.Format("Data Source={0}; Initial Catalog={1}; User id={2}; Password={3};", oCompany.Server, oCompany.CompanyDB, uid, pwd));
            }
            else
            {
                oApp.MessageBox("Can't connect to DB !");
            }
        }

        System.Data.DataTable Get_Data( DateTime pFrDate, DateTime pToDate)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_HQDP", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FromDate", pFrDate);
                cmd.Parameters.AddWithValue("@ToDate", pToDate);
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

        private decimal Get_GTHD_CT(DateTime pFrDate, DateTime pToDate, string pBpCode, string pFProject, string pPUType)
        {
            decimal GTHD = 0;
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_HQDP_GTHD", conn);
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
            if (string.IsNullOrEmpty(EditText0.Value) || string.IsNullOrEmpty(EditText1.Value))
            {
                oApp.MessageBox("Please enter date");
                return;
            }
            else
            {
                DateTime FrDate = DateTime.ParseExact(EditText0.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                DateTime ToDate = DateTime.ParseExact(EditText1.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                DataTable tb_data = Get_Data(FrDate, ToDate);
                if (tb_data.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Application oXL;
                    Microsoft.Office.Interop.Excel._Workbook oWB;
                    Microsoft.Office.Interop.Excel._Worksheet oSheet;

                    object misvalue = System.Reflection.Missing.Value;
                    //Start Excel and get Application object.
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oXL.Visible = true;
                    //Open Template
                    oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_HQDP.xlsx");
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                    int STT = 1;
                    int current_row = 5, g_row = 5, g1_row = 6;

                    //Parameter
                    oSheet.Cells[3, 1] = string.Format("Thời gian từ ngày {0} đến ngày {1}", FrDate.ToString("dd/MM/yyyy"), ToDate.ToString("dd/MM/yyyy"));
                    #region XD
                    //Group NCC-NTP Xay dung
                    oSheet.Cells[g_row, 1] = "I";
                    oSheet.Cells[g_row, 2] = "NCC, NTP XÂY DỰNG";
                    oSheet.Range["A" + g_row, "M" + g_row].Interior.Color = System.Drawing.Color.FromArgb(191, 191, 191);
                    oSheet.Range["A" + g_row, "M" + g_row].Font.Bold = true;
                    current_row++;

                    //Group NCC XD
                    oSheet.Cells[g1_row, 1] = "I.1";
                    oSheet.Cells[g1_row, 2] = "NHÀ CUNG CẤP XÂY DỰNG";
                    oSheet.Range["A" + g1_row, "M" + g1_row].Interior.Color = System.Drawing.Color.FromArgb(244, 176, 132);
                    oSheet.Range["A" + g1_row, "M" + g1_row].Font.Bold = true;
                    current_row++;

                    //Print Details
                    foreach (DataRow r in tb_data.Select("U_RECTYPE='XD' and NDT='PUT01'"))
                    {
                        oSheet.Cells[current_row, 1] = STT;
                        oSheet.Cells[current_row, 2] = r["CardName"];
                        oSheet.Cells[current_row, 3] = r["Project"];
                        oSheet.Cells[current_row, 4] = r["HM_NAME"];
                        //Gia tri ky hop dong
                        oSheet.Cells[current_row, 5] = Get_GTHD_CT(FrDate,ToDate,r["CardCode"].ToString(),r["Project"].ToString(),r["NDT"].ToString());
                        oSheet.Cells[current_row, 6] = r["GT_BOQ"];
                        oSheet.Cells[current_row, 7] = r["GG_DT"];
                        oSheet.Cells[current_row, 8] = r["Total_GRPO"];
                        //Hiệu quả tăng thêm so với giá gốc đấu thầu
                        oSheet.Cells[current_row, 9].Formula = string.Format("=1-H{0}/G{0}", current_row);
                        oSheet.Cells[current_row, 10].Formula = string.Format("=IF(H{0}<>0,G{0}-H{0},0)", current_row);
                        //Hiệu quả tăng thêm so giá trị BOQ 
                        oSheet.Cells[current_row, 11].Formula = string.Format("=1-H{0}/F{0}", current_row);
                        oSheet.Cells[current_row, 12].Formula = string.Format("=IF(H{0}<>0,F{0}-H{0},0)", current_row);
                        current_row++;
                        STT++;
                    }
                    //Total g1_row
                    if (g1_row < current_row - 1)
                    {
                        oSheet.Cells[g1_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 10].Formula = string.Format("=SUBTOTAL(9,J{0}:J{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1})", g1_row + 1, current_row - 1);
                    }
                    g1_row = current_row;
                    STT = 1;
                    //Group NTP XD
                    oSheet.Cells[g1_row, 1] = "I.2";
                    oSheet.Cells[g1_row, 2] = "NHÀ THẦU PHỤ XÂY DỰNG";
                    oSheet.Range["A" + g1_row, "M" + g1_row].Interior.Color = System.Drawing.Color.FromArgb(244, 176, 132);
                    oSheet.Range["A" + g1_row, "M" + g1_row].Font.Bold = true;
                    current_row++;

                    //Print Details
                    foreach (DataRow r in tb_data.Select("U_RECTYPE='XD' and NDT='PUT02'"))
                    {
                        oSheet.Cells[current_row, 1] = STT;
                        oSheet.Cells[current_row, 2] = r["CardName"];
                        oSheet.Cells[current_row, 3] = r["Project"];
                        oSheet.Cells[current_row, 4] = r["HM_NAME"];
                        //Gia tri ky hop dong
                        oSheet.Cells[current_row, 5] = Get_GTHD_CT(FrDate, ToDate, r["CardCode"].ToString(), r["Project"].ToString(), r["NDT"].ToString());
                        oSheet.Cells[current_row, 6] = r["GT_BOQ"];
                        oSheet.Cells[current_row, 7] = r["GG_DT"];
                        oSheet.Cells[current_row, 8] = r["Total_GRPO"];
                        //Hiệu quả tăng thêm so với giá gốc đấu thầu
                        oSheet.Cells[current_row, 9].Formula = string.Format("=1-H{0}/G{0}", current_row);
                        oSheet.Cells[current_row, 10].Formula = string.Format("=IF(H{0}<>0,G{0}-H{0},0)", current_row);
                        //Hiệu quả tăng thêm so giá trị BOQ 
                        oSheet.Cells[current_row, 11].Formula = string.Format("=1-H{0}/F{0}", current_row);
                        oSheet.Cells[current_row, 12].Formula = string.Format("=IF(H{0}<>0,F{0}-H{0},0)", current_row);
                        current_row++;
                        STT++;
                    }
                    //Total g1_row
                    if (g1_row < current_row - 1)
                    {
                        oSheet.Cells[g1_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 10].Formula = string.Format("=SUBTOTAL(9,J{0}:J{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1})", g1_row + 1, current_row - 1);
                    }
                    //Total g_row
                    if (g_row < current_row - 1)
                    {
                        oSheet.Cells[g_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 10].Formula = string.Format("=SUBTOTAL(9,J{0}:J{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1})", g_row + 1, current_row - 1);
                    }
                    #endregion

                    g_row = current_row;
                    g1_row = g_row + 1;
                    STT = 1;
                    #region ME
                    //Group NCC-NTP ME
                    oSheet.Cells[g_row, 1] = "II";
                    oSheet.Cells[g_row, 2] = "NHÀ CUNG CẤP, NHÀ THẦU PHỤ M&E";
                    oSheet.Range["A" + g_row, "M" + g_row].Interior.Color = System.Drawing.Color.FromArgb(191, 191, 191);
                    oSheet.Range["A" + g_row, "M" + g_row].Font.Bold = true;
                    current_row++;

                    //Group NCC XD
                    oSheet.Cells[g1_row, 1] = "II.1";
                    oSheet.Cells[g1_row, 2] = "NHÀ CUNG CẤP M&E";
                    oSheet.Range["A" + g1_row, "M" + g1_row].Interior.Color = System.Drawing.Color.FromArgb(244, 176, 132);
                    oSheet.Range["A" + g1_row, "M" + g1_row].Font.Bold = true;
                    current_row++;

                    //Print Details
                    foreach (DataRow r in tb_data.Select("U_RECTYPE='CDXD' and NDT='PUT01'"))
                    {
                        oSheet.Cells[current_row, 1] = STT;
                        oSheet.Cells[current_row, 2] = r["CardName"];
                        oSheet.Cells[current_row, 3] = r["Project"];
                        oSheet.Cells[current_row, 4] = r["HM_NAME"];
                        //Gia tri ky hop dong
                        oSheet.Cells[current_row, 5] = Get_GTHD_CT(FrDate, ToDate, r["CardCode"].ToString(), r["Project"].ToString(), r["NDT"].ToString());
                        oSheet.Cells[current_row, 6] = r["GT_BOQ"];
                        oSheet.Cells[current_row, 7] = r["GG_DT"];
                        oSheet.Cells[current_row, 8] = r["Total_GRPO"];
                        //Hiệu quả tăng thêm so với giá gốc đấu thầu
                        oSheet.Cells[current_row, 9].Formula = string.Format("=1-H{0}/G{0}", current_row);
                        oSheet.Cells[current_row, 10].Formula = string.Format("=IF(H{0}<>0,G{0}-H{0},0)", current_row);
                        //Hiệu quả tăng thêm so giá trị BOQ 
                        oSheet.Cells[current_row, 11].Formula = string.Format("=1-H{0}/F{0}", current_row);
                        oSheet.Cells[current_row, 12].Formula = string.Format("=IF(H{0}<>0,F{0}-H{0},0)", current_row);
                        current_row++;
                        STT++;
                    }
                    //Total g1_row
                    if (g1_row < current_row - 1)
                    {
                        oSheet.Cells[g1_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 10].Formula = string.Format("=SUBTOTAL(9,J{0}:J{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1})", g1_row + 1, current_row - 1);
                    }
                    g1_row = current_row;
                    STT = 1;
                    //Group NTP ME
                    oSheet.Cells[g1_row, 1] = "II.2";
                    oSheet.Cells[g1_row, 2] = "NHÀ THẦU PHỤ M&E";
                    oSheet.Range["A" + g1_row, "M" + g1_row].Interior.Color = System.Drawing.Color.FromArgb(244, 176, 132);
                    oSheet.Range["A" + g1_row, "M" + g1_row].Font.Bold = true;
                    current_row++;

                    //Print Details
                    foreach (DataRow r in tb_data.Select("U_RECTYPE='CDXD' and NDT='PUT02'"))
                    {
                        oSheet.Cells[current_row, 1] = STT;
                        oSheet.Cells[current_row, 2] = r["CardName"];
                        oSheet.Cells[current_row, 3] = r["Project"];
                        oSheet.Cells[current_row, 4] = r["HM_NAME"];
                        //Gia tri ky hop dong
                        oSheet.Cells[current_row, 5] = Get_GTHD_CT(FrDate, ToDate, r["CardCode"].ToString(), r["Project"].ToString(), r["NDT"].ToString());
                        oSheet.Cells[current_row, 6] = r["GT_BOQ"];
                        oSheet.Cells[current_row, 7] = r["GG_DT"];
                        oSheet.Cells[current_row, 8] = r["Total_GRPO"];
                        //Hiệu quả tăng thêm so với giá gốc đấu thầu
                        oSheet.Cells[current_row, 9].Formula = string.Format("=1-H{0}/G{0}", current_row);
                        oSheet.Cells[current_row, 10].Formula = string.Format("=IF(H{0}<>0,G{0}-H{0},0)", current_row);
                        //Hiệu quả tăng thêm so giá trị BOQ 
                        oSheet.Cells[current_row, 11].Formula = string.Format("=1-H{0}/F{0}", current_row);
                        oSheet.Cells[current_row, 12].Formula = string.Format("=IF(H{0}<>0,F{0}-H{0},0)", current_row);
                        current_row++;
                        STT++;
                    }
                    //Total g1_row
                    if (g1_row < current_row - 1)
                    {
                        oSheet.Cells[g1_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 10].Formula = string.Format("=SUBTOTAL(9,J{0}:J{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1})", g1_row + 1, current_row - 1);
                    }
                    //Total g_row
                    if (g_row < current_row - 1)
                    {
                        oSheet.Cells[g_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 10].Formula = string.Format("=SUBTOTAL(9,J{0}:J{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1})", g_row + 1, current_row - 1);
                    }
                    #endregion

                    g_row = current_row;
                    g1_row = g_row + 1;
                    STT = 1;
                    #region TB
                    //Group NCC-NTP TB
                    oSheet.Cells[g_row, 1] = "III";
                    oSheet.Cells[g_row, 2] = "NHÀ CUNG CẤP, NHÀ THẦU PHỤ THIẾT BỊ";
                    oSheet.Range["A" + g_row, "M" + g_row].Interior.Color = System.Drawing.Color.FromArgb(191, 191, 191);
                    oSheet.Range["A" + g_row, "M" + g_row].Font.Bold = true;
                    current_row++;

                    //Group NCC TB
                    oSheet.Cells[g1_row, 1] = "III.1";
                    oSheet.Cells[g1_row, 2] = "NHÀ CUNG CẤP THIẾT BỊ";
                    oSheet.Range["A" + g1_row, "M" + g1_row].Interior.Color = System.Drawing.Color.FromArgb(244, 176, 132);
                    oSheet.Range["A" + g1_row, "M" + g1_row].Font.Bold = true;
                    current_row++;

                    //Print Details
                    foreach (DataRow r in tb_data.Select("U_RECTYPE='TBXD' and NDT='PUT01'"))
                    {
                        oSheet.Cells[current_row, 1] = STT;
                        oSheet.Cells[current_row, 2] = r["CardName"];
                        oSheet.Cells[current_row, 3] = r["Project"];
                        oSheet.Cells[current_row, 4] = r["HM_NAME"];
                        //Gia tri ky hop dong
                        oSheet.Cells[current_row, 5] = Get_GTHD_CT(FrDate, ToDate, r["CardCode"].ToString(), r["Project"].ToString(), r["NDT"].ToString());
                        oSheet.Cells[current_row, 6] = r["GT_BOQ"];
                        oSheet.Cells[current_row, 7] = r["GG_DT"];
                        oSheet.Cells[current_row, 8] = r["Total_GRPO"];
                        //Hiệu quả tăng thêm so với giá gốc đấu thầu
                        oSheet.Cells[current_row, 9].Formula = string.Format("=1-H{0}/G{0}", current_row);
                        oSheet.Cells[current_row, 10].Formula = string.Format("=IF(H{0}<>0,G{0}-H{0},0)", current_row);
                        //Hiệu quả tăng thêm so giá trị BOQ 
                        oSheet.Cells[current_row, 11].Formula = string.Format("=1-H{0}/F{0}", current_row);
                        oSheet.Cells[current_row, 12].Formula = string.Format("=IF(H{0}<>0,F{0}-H{0},0)", current_row);
                        current_row++;
                        STT++;
                    }
                    //Total g1_row
                    if (g1_row < current_row - 1)
                    {
                        oSheet.Cells[g1_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 10].Formula = string.Format("=SUBTOTAL(9,J{0}:J{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1})", g1_row + 1, current_row - 1);
                    }
                    g1_row = current_row;
                    STT = 1;
                    //Group NTP TB
                    oSheet.Cells[g1_row, 1] = "III.2";
                    oSheet.Cells[g1_row, 2] = "NHÀ THẦU PHỤ THIẾT BỊ";
                    oSheet.Range["A" + g1_row, "M" + g1_row].Interior.Color = System.Drawing.Color.FromArgb(244, 176, 132);
                    oSheet.Range["A" + g1_row, "M" + g1_row].Font.Bold = true;
                    current_row++;

                    //Print Details
                    foreach (DataRow r in tb_data.Select("U_RECTYPE='TBXD' and NDT='PUT02'"))
                    {
                        oSheet.Cells[current_row, 1] = STT;
                        oSheet.Cells[current_row, 2] = r["CardName"];
                        oSheet.Cells[current_row, 3] = r["Project"];
                        oSheet.Cells[current_row, 4] = r["HM_NAME"];
                        //Gia tri ky hop dong
                        oSheet.Cells[current_row, 5] = Get_GTHD_CT(FrDate, ToDate, r["CardCode"].ToString(), r["Project"].ToString(), r["NDT"].ToString());
                        oSheet.Cells[current_row, 6] = r["GT_BOQ"];
                        oSheet.Cells[current_row, 7] = r["GG_DT"];
                        oSheet.Cells[current_row, 8] = r["Total_GRPO"];
                        //Hiệu quả tăng thêm so với giá gốc đấu thầu
                        oSheet.Cells[current_row, 9].Formula = string.Format("=1-H{0}/G{0}", current_row);
                        oSheet.Cells[current_row, 10].Formula = string.Format("=IF(H{0}<>0,G{0}-H{0},0)", current_row);
                        //Hiệu quả tăng thêm so giá trị BOQ 
                        oSheet.Cells[current_row, 11].Formula = string.Format("=1-H{0}/F{0}", current_row);
                        oSheet.Cells[current_row, 12].Formula = string.Format("=IF(H{0}<>0,F{0}-H{0},0)", current_row);
                        current_row++;
                        STT++;
                    }
                    //Total g1_row
                    if (g1_row < current_row - 1)
                    {
                        oSheet.Cells[g1_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 10].Formula = string.Format("=SUBTOTAL(9,J{0}:J{1})", g1_row + 1, current_row - 1);
                        oSheet.Cells[g1_row, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1})", g1_row + 1, current_row - 1);
                    }
                    //Total g_row
                    if (g_row < current_row - 1)
                    {
                        oSheet.Cells[g_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 10].Formula = string.Format("=SUBTOTAL(9,J{0}:J{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1})", g_row + 1, current_row - 1);
                    }
                    #endregion

                    //Total all
                    oSheet.Range["A" + current_row, "M" + current_row].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                    oSheet.Range["A" + current_row, "M" + current_row].Font.Bold = true;
                    oSheet.Range["A" + current_row, "B" + current_row].Merge(true);
                    oSheet.Cells[current_row, 1] = "TỔNG CỘNG";
                    oSheet.Cells[current_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})",  5, current_row - 1);
                    oSheet.Cells[current_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})",  5, current_row - 1);
                    oSheet.Cells[current_row, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})",  5, current_row - 1);
                    oSheet.Cells[current_row, 10].Formula = string.Format("=SUBTOTAL(9,J{0}:J{1})", 5, current_row - 1);
                    oSheet.Cells[current_row, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1})", 5, current_row - 1);
                    //Border
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A5", "M" + (current_row))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    oXL.ActiveWindow.Activate();
                }
            }

        }

    }
}
