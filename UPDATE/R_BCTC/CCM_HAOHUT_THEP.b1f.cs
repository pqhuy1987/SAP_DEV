using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;
using System.Globalization;

namespace R_BCTC
{
    [FormAttribute("R_BCTC.CCM_HAOHUT_THEP", "CCM_HAOHUT_THEP.b1f")]
    class CCM_HAOHUT_THEP : UserFormBase
    {
        public CCM_HAOHUT_THEP()
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
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_tdate").Specific));
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

        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.EditText EditText0;

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

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (ComboBox0.Selected.Value == "" || EditText0.Value == "")
            {
                oApp.MessageBox("Please select Project and To Date !");
                return;
            }
            DateTime ToDate = DateTime.ParseExact(EditText0.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
            string FProject = ComboBox0.Selected.Value;
            string FProjectName = ComboBox0.Selected.Description;
            DataTable rs = Get_Data(ComboBox0.Selected.Value,ToDate);
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
                oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_HAOHUT_THEP.xlsx");
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                
                //Project Code
                oSheet.Cells[1, 3] = "Mã dự án: " + FProject;
                //Project Name
                oSheet.Cells[2, 3] = "Tên dự án: " + FProjectName;
                //From Date - To Date
                oSheet.Cells[3, 3] = string.Format("Ngày: {0}", ToDate.ToString("dd/MM/yyyy"));

                int STT = 1;
                int STT_G = 1;
                int current_row = 8,g_row= 8;
                string CT_CODE = "";
                foreach (DataRow r in rs.Rows)
                {
                    //Print Cong tac
                    if (r["HM_CODE"].ToString() != CT_CODE)
                    {
                        CT_CODE = r["HM_CODE"].ToString();
                        //STT Group
                        oSheet.Cells[current_row, 1].Formula = string.Format("=ROMAN({0})", STT_G);
                        //Ma cong tac
                        oSheet.Cells[current_row, 2] = r["HM_CODE"].ToString();
                        //Ten cong tac
                        oSheet.Cells[current_row, 3] = r["HM_NAME"].ToString();
                        oSheet.Cells[current_row, 3].Font.Color = System.Drawing.Color.Red;

                        //Format style
                        oSheet.Range["A" + current_row, "M" + current_row].Font.Bold = true;
                        oSheet.Range["A" + current_row, "M" + current_row].Interior.Color = System.Drawing.Color.FromArgb(219, 219, 219);

                        //Total Group
                        if (g_row < current_row)
                        {
                            oSheet.Cells[g_row, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1})", g_row+1,current_row-1);
                            oSheet.Cells[g_row, 13].Formula = string.Format("=SUBTOTAL(9,M{0}:M{1})", g_row + 1, current_row - 1);
                        }
                        STT_G++;
                        g_row = current_row;
                        current_row++;
                    }
                    //STT
                    oSheet.Cells[current_row, 1] = STT;
                    //Ma CT
                    oSheet.Cells[current_row, 2] = r["CV_CODE"].ToString();
                    //Ten CT
                    oSheet.Cells[current_row, 3] = r["CV_NAME"].ToString();
                    //DVT
                    oSheet.Cells[current_row, 4] = r["CV_DVT"].ToString();
                    //KL Dau thau
                    oSheet.Cells[current_row, 5] = r["CV_KLDT"];
                    //KL Cong truong
                    oSheet.Cells[current_row, 6] = r["CV_KLBV"];
                    //KL Nhap ve
                    oSheet.Cells[current_row, 7] = r["KLNV"];
                    //% KL den hien tai
                    oSheet.Cells[current_row, 8] = r["PT_HOANTHANH"];
                    //KL Den hien tai
                    oSheet.Cells[current_row, 9].Formula = string.Format("=H{0}*F{0}", current_row);
                    //KL Nguyen con ton
                    oSheet.Cells[current_row, 10] = r["KL_NGUYEN"];
                    //KL vun va thanh ly
                    oSheet.Cells[current_row, 11] = r["KL_VUN"];
                    //Hao hut khong ro nguyen nhan
                    oSheet.Cells[current_row, 12].Formula = string.Format("=G{0}-I{0}-J{0}-K{0}", current_row);
                    //KL Hao hut
                    oSheet.Cells[current_row, 13].Formula = string.Format("=(70/100)*K{0}+L{0}", current_row);

                    STT++;
                    current_row++;


                }
                //Total Last Group
                if (g_row < current_row)
                {
                    oSheet.Cells[g_row, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1})", g_row + 1, current_row - 1);
                    oSheet.Cells[g_row, 13].Formula = string.Format("=SUBTOTAL(9,M{0}:M{1})", g_row + 1, current_row - 1);
                }

                //TOTAL
                oSheet.Range["A" + current_row, "C" + current_row].Merge();
                oSheet.Range["A" + current_row, "M" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "M" + current_row].Interior.Color = System.Drawing.Color.FromArgb(247, 202, 172);
                
                oSheet.Cells[current_row, 1] = "TỔNG CỘNG (16)";
                oSheet.Cells[current_row, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", 8 , current_row - 1);
                oSheet.Cells[current_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", 8, current_row - 1);
                oSheet.Cells[current_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", 8, current_row - 1);

                oSheet.Cells[current_row, 9].Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", 8, current_row - 1);
                oSheet.Cells[current_row, 10].Formula = string.Format("=SUBTOTAL(9,J{0}:J{1})", 8, current_row - 1);
                oSheet.Cells[current_row, 11].Formula = string.Format("=SUBTOTAL(9,K{0}:K{1})", 8, current_row - 1);
                oSheet.Cells[current_row, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1})", 8, current_row - 1);
                oSheet.Cells[current_row, 13].Formula = string.Format("=SUBTOTAL(9,M{0}:M{1})", 8, current_row - 1);
                current_row++;

                oSheet.Range["A" + current_row, "C" + current_row].Merge();
                oSheet.Range["A" + current_row, "C" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "M" + current_row].Interior.Color = System.Drawing.Color.FromArgb(247, 202, 172);
                oSheet.Range["A" + current_row, "M" + current_row].RowHeight = 60;
                oSheet.Range["A" + current_row, "M" + current_row].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                oSheet.Cells[current_row, 1] = "Phần trăm hao hụt khối lượng nhập là:";
                oSheet.Cells[current_row, 13].Formula = string.Format("=M{0}/(G{0}-J{0})", current_row - 1);
                oSheet.Cells[current_row, 13].NumberFormat = "0.00%";
                current_row++;

                //Border
                ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A8", "M" + (current_row - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                oXL.ActiveWindow.Activate();
            }

        }

        System.Data.DataTable Get_Data(string pFproject, DateTime pToDate)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_HAOHUT_THEP_GET_DATA", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FProject", pFproject);
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
    }
}
