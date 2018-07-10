using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;

namespace R_BCTC
{
    [FormAttribute("R_BCTC.CCM_APPROVE_BILL", "CCM_APPROVE_BILL.b1f")]
    class CCM_APPROVE_BILL : UserFormBase
    {
        public CCM_APPROVE_BILL()
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
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_view").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_fprj").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_fp").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txt_tp").Specific));
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

        System.Data.DataTable Get_List_FProject()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_SUMMARY_GET_FPROJECT", conn);
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
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText1;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (string.IsNullOrEmpty(ComboBox0.Selected.Value) || string.IsNullOrEmpty(EditText0.Value) || string.IsNullOrEmpty(EditText1.Value))
                {
                    oApp.MessageBox("Please select Project and Period !");
                    return;
                }
                int fr_period = 0;
                int to_period = 0;
                string fProject = ComboBox0.Selected.Value;
                int.TryParse(EditText0.Value, out fr_period);
                int.TryParse(EditText1.Value, out to_period);

                DataTable rs = Get_Data(fProject, fr_period, to_period);
                if (rs.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Application oXL;
                    Microsoft.Office.Interop.Excel._Workbook oWB;
                    Microsoft.Office.Interop.Excel._Worksheet oSheet;

                    object misvalue = System.Reflection.Missing.Value;
                    //Start Excel and get Application object.
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oXL.Visible = true;
                    //Open Template
                    oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_DUYET_BILL.xlsx");
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                    //Project Name
                    oSheet.Cells[3, 3] = "Dự án: " + this.ComboBox0.Selected.Description;
                    //Period
                    oSheet.Cells[4, 3] = string.Format("Từ kỳ {0} đến kỳ {1}", fr_period, to_period);

                    int current_row = 6, STT = 1;

                    foreach (DataRow r in rs.Rows)
                    {
                        //STT
                        oSheet.Cells[current_row, 1] = STT;
                        //Project Name
                        oSheet.Cells[current_row, 2] = r["U_FIPROJECT"].ToString();
                        //Period
                        oSheet.Cells[current_row, 3] = r["U_Period"];
                        //BP Name
                        oSheet.Cells[current_row, 4] = r["U_BPName"].ToString();
                        #region BILL CACULATOR
                        decimal Kynay = 0, Kytruoc = 0;
                        string Phamvi = "";
                        Bill_Caculator(ComboBox0.Selected.Value, int.Parse(r["U_Period"].ToString()), r["U_BPCode"].ToString(), r["U_BGroup"].ToString(), r["U_PUType"].ToString(), int.Parse(r["U_BType"].ToString()), int.Parse(r["DocEntry"].ToString()), DateTime.Parse(r["U_DATETO"].ToString()), out Kytruoc, out Kynay, out Phamvi);
                        oSheet.Cells[current_row, 6].Value2 = Kynay;
                        //Pham vi cong viec
                        oSheet.Cells[current_row, 5].Value2 = Phamvi;
                        #endregion
                        //CHT
                        oSheet.Cells[current_row, 7] = r["CHT"];
                        //Thiet bi
                        oSheet.Cells[current_row, 8] = r["TB"];
                        //Co dien
                        oSheet.Cells[current_row, 9] = r["CD"];
                        //CCM
                        oSheet.Cells[current_row, 10] = r["CCM"];
                        //GDDA
                        oSheet.Cells[current_row, 11] = r["GDDA"];
                        //Ke toan
                        oSheet.Cells[current_row, 12] = r["KT"];
                        //Total
                        oSheet.Cells[current_row, 13].Formula = string.Format("=IF(DAYS(L{0},G{0})<0,0,DAYS(L{0},G{0}))", current_row);
                        //Ghi chu
                        oSheet.Cells[current_row, 14] = r["Rejected"];
                        STT++;
                        current_row++;
                    }

                    //Border
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A6", "N" + (current_row - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    oXL.ActiveWindow.Activate();
                }
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
 
            }
        }

        System.Data.DataTable Get_Data(string pFproject, int pFrPeriod, int pToPeriod)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_DUYET_BILL_GET_DATA", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FProject", pFproject);
                cmd.Parameters.AddWithValue("@Fr_Period", pFrPeriod);
                cmd.Parameters.AddWithValue("@To_Period", pToPeriod);
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

        System.Data.DataTable Load_Data_KLTT(string pFinancialProject, string pType, int pPeriod, string pBPCode, string pCGroup, string pPUType, int pDocEntry, DateTime pToDate)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                if (pType == "AI")
                {
                    cmd = new SqlCommand("KLTT_APPROVE_GET_ADDITIONALINFO", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                    cmd.Parameters.AddWithValue("@Period", pPeriod);
                    cmd.Parameters.AddWithValue("@BP_Code", pBPCode);
                    cmd.Parameters.AddWithValue("@CGroup", pCGroup);
                    cmd.Parameters.AddWithValue("@PUType", pPUType);
                    cmd.Parameters.AddWithValue("@ToDate", pToDate);
                }
                else if (pType == "SPP")
                {
                    cmd = new SqlCommand("KLTT_APPROVE_TOTAL", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                    cmd.Parameters.AddWithValue("@Period", pPeriod);
                    cmd.Parameters.AddWithValue("@BP_Code", pBPCode);
                    cmd.Parameters.AddWithValue("@BGroup", pCGroup);
                    cmd.Parameters.AddWithValue("@PUType", pPUType);
                }
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

        private void Bill_Caculator(string FProject, int Period, string BPCode, string BGroup, string PUType, int BType, int DocEntry, DateTime Todate, out decimal KLTT_Kytruoc, out decimal KLTT_Kynay, out string Phamvi)
        {
            KLTT_Kytruoc = 0;
            KLTT_Kynay = 0;
            Phamvi = "";
            DataTable SPP = Load_Data_KLTT(FProject, "SPP", Period - 1, BPCode, BGroup, PUType, DocEntry, Todate);
            DataTable SPP_CURRENT = Load_Data_KLTT(FProject, "SPP", Period, BPCode, BGroup, PUType, DocEntry, Todate);
            DataTable AI = Load_Data_KLTT(FProject, "AI", Period, BPCode, BGroup, PUType, DocEntry, Todate);

            #region Thong tin chi tiet HD
            decimal GTHD = 0, PLTT = 0, PLT = 0;
            float PTBH = 0, PTTU = 0, PTHU = 0, PTGL = 0;
            string TTTU = "", HTBH = "" ;
            if (AI.Rows.Count >= 1)
            {
                float.TryParse(AI.Rows[0]["PTBH"].ToString(), out PTBH);
                float.TryParse(AI.Rows[0]["PTGL"].ToString(), out PTGL);
                float.TryParse(AI.Rows[0]["PTTU"].ToString(), out PTTU);
                float.TryParse(AI.Rows[0]["PTHU"].ToString(), out PTHU);
                TTTU = AI.Rows[0]["TTTU"].ToString();
                HTBH = AI.Rows[0]["HTBH"].ToString();
                foreach (DataRow r in AI.Rows)
                {
                    decimal tmp = 0;
                    decimal.TryParse(r["GTHD"].ToString(), out tmp);
                    if (r["Type"].ToString() == "HD")
                    {
                        GTHD += tmp;
                        Phamvi = r["Descript"].ToString();
                    }
                    else if (r["Type"].ToString() == "PLT")
                        PLT += tmp;
                    else if (r["Type"].ToString() == "PLTT")
                    {
                        Phamvi = r["Descript"].ToString();
                        PLTT += tmp;
                    }
                }
            }
            #endregion

            #region Thong tin bill
            decimal GTTU = 0, GTHU = 0;
            decimal pp_pl = 0, pp_ca = 0, pp_ca_no_VAT = 0, pp_tu_lastbill = 0, pp_hu_lastbill = 0;
            decimal sum_ca_novat = 0;
            if (SPP_CURRENT.Rows.Count == 1)
            {
                decimal.TryParse(SPP_CURRENT.Rows[0]["SUM_CA_NOVAT"].ToString(), out sum_ca_novat);
                decimal.TryParse(SPP_CURRENT.Rows[0]["TOTAL_TU"].ToString(), out GTTU);
                decimal.TryParse(SPP_CURRENT.Rows[0]["TOTAL_HU"].ToString(), out GTHU);
            }


            if (SPP.Rows.Count == 1)
            {
                decimal.TryParse(SPP.Rows[0]["SUM_PL"].ToString(), out pp_pl);
                decimal.TryParse(SPP.Rows[0]["SUM_CA"].ToString(), out pp_ca);
                decimal.TryParse(SPP.Rows[0]["SUM_CA_NOVAT"].ToString(), out pp_ca_no_VAT);
                //decimal.TryParse(SPP.Rows[0]["TOTAL_TU"].ToString(), out GTTU);
                decimal.TryParse(SPP.Rows[0]["TOTAL_HU"].ToString(), out pp_hu_lastbill);
                decimal.TryParse(SPP.Rows[0]["TOTAL_TU_LASTBILL"].ToString(), out pp_tu_lastbill);
            }
            //Gia tri thuc hien den ky nay
            decimal GTKN = 0;
            if (BType > 1)
                decimal.TryParse(SPP_CURRENT.Rows[0]["SUM_CA"].ToString(), out GTKN);
            //Tong gia tri thi cong
            //EditText4.Value = BType == "1" ? "0" : String.Format("{0:n}", SPP_CURRENT.Rows[0]["SUM_PL_VAT"]);

            //Tam ung
            decimal TU = 0;
            if (BType == 1)
            {
                string sql_cmd = string.Format("Select a.U_GTTU from [@KLTT] a where a.U_FIProject='{0}' and a.DocEntry = {1};", FProject, DocEntry);
                try
                {
                    SqlCommand cmd = new SqlCommand(sql_cmd, conn);
                    conn.Open();
                    decimal.TryParse(cmd.ExecuteScalar().ToString(), out TU);
                }
                catch
                {

                }
                finally
                {
                    conn.Close();
                }
                //EditText12.Value = String.Format("{0:n}", U_GTTU);
            }
            else
            {
                TU = GTTU;
                //EditText12.Value = String.Format("{0:n}", GTTU);
            }
            //Hoan tra TU
            decimal HU = 0;
            if (BType == 3)
                HU = GTTU;
            else
                HU = GTHU;
            //GT thanh toan den ky nay
            decimal GTTT_KN = 0;
            if (BType == 3)
                GTTT_KN = GTKN;
            else
                GTTT_KN = Math.Round((1 - (decimal)PTGL) * GTKN, 0);

            //GT thanh toan giu lai
            //if (BType == "3")
            //{
            //    if (HTBH == "TM")
            //    {
            //        StaticText10.Caption = "4. GT giữ lại bảo hành";
            //        EditText10.Value = String.Format("{0:n}", Math.Round((decimal)PTBH * decimal.Parse(EditText6.Value), 0));
            //    }
            //    else
            //    {
            //        StaticText10.Caption = "4. GT giữ lại bảo hành (Chứng thư)";
            //        EditText10.Value = "0";
            //    }
            //}
            //else
            //{
            //    StaticText10.Caption = "4. GT thanh toán giữ lại";
            //    EditText10.Value = String.Format("{0:n}", Math.Round((decimal)PTGL * decimal.Parse(EditText6.Value), 0));
            //}
            //Tong GT duoc thanh toan den ky nay
            decimal TongGT = 0;
            TongGT = GTTT_KN + TU - HU;// String.Format("{0:n}", decimal.Parse(EditText8.Value) + decimal.Parse(EditText12.Value) - decimal.Parse(EditText14.Value));
            
            //Tong GT thanh toan den ky truoc
            if (BType == 1)
            {
                KLTT_Kytruoc = 0;
                //EditText18.Value = "0";
            }
            else if (BType == 2)
            {
                if (pp_tu_lastbill > 0)
                    KLTT_Kytruoc = Math.Round((pp_ca * (1 - (decimal)PTGL)) + pp_tu_lastbill - pp_hu_lastbill, 0);
                else
                    KLTT_Kytruoc = Math.Round((pp_ca * (1 - (decimal)PTGL)), 0);
            }
            else if (BType == 3)
            {
                if (pp_tu_lastbill > 0)
                    KLTT_Kytruoc = Math.Round((pp_ca * (1 - (decimal)PTGL)) + pp_tu_lastbill - pp_hu_lastbill);
                else
                    KLTT_Kytruoc = Math.Round((pp_ca * (1 - (decimal)PTGL)));
            }

            //GT thanh toan ky nay
            if (BType != 1)
            {
                KLTT_Kynay = Math.Round(GTTT_KN - KLTT_Kytruoc, 0);
            }
            else
            {
                KLTT_Kynay = TU;
            }

            #endregion
        }
    }
}
