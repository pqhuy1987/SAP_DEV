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
    [FormAttribute("R_BCTC.CCM_TONGHOP_VP", "CCM_TONGHOP_VP.b1f")]
    class CCM_TONGHOP_VP : UserFormBase
    {
        public CCM_TONGHOP_VP()
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
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));
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
        }

        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.EditText EditText0;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (string.IsNullOrEmpty(EditText0.Value))
            {
                oApp.MessageBox("Please enter value !");
                return;
            }
            try
            {
                DateTime todate = DateTime.ParseExact(EditText0.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                DataTable rs_detail = Get_Data(todate);

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
                    oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_TH_CP_VP.xlsx");
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                    //Project Code
                    //oSheet.Cells[1, 3] = "Dự án: " + ComboBox0.Selected.Description;
                    //Project Name
                    //oSheet.Cells[2, 3] = "Hạng mục: Chi phí quản lý văn phòng" ;
                    //Period
                    //oSheet.Cells[3, 3] = string.Format("Kỳ: {0}", period);
                    //Date view report
                    oSheet.Cells[3, 1] = string.Format("Từ ngày {0} đến ngày {1}",new DateTime(todate.Year,1,1).ToString("dd/MM/yyyy") ,todate.ToString("dd/MM/yyyy"));
                    int STT = 1;
                    int STT_G = 1;
                    int current_row = 5, g_row = 5;
                    string NCP = "";
                    string CP = "";
                    foreach (DataRow r in rs_detail.Rows)
                    {
                        //Print Nhom chi phi
                        if (r["MA_NHOM_CP"].ToString() != NCP)
                        {
                            string NCP_old = NCP;
                            NCP = r["MA_NHOM_CP"].ToString();
                            //STT Group
                            oSheet.Cells[current_row, 1].Formula = string.Format("=ROMAN({0})", STT_G);
                            //Ten nhom CP
                            oSheet.Cells[current_row, 2] = r["TEN_NHOM_CP"].ToString();
                            //Gia tri cho Khau tru TSCD
                            if (NCP == "C8")
                            {
                                //KT No VAT
                                oSheet.Cells[current_row, 4] = r["KT_NO_VAT"];
                                //CCM no VAT
                                oSheet.Cells[current_row, 5] = r["GT_NO_VAT"];
                            }

                            //Format style
                            oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                            oSheet.Range["A" + current_row, "H" + current_row].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);

                            //Total Group
                            if (g_row < current_row)
                            {
                                oSheet.Cells[g_row, 3].Formula = string.Format("=SUBTOTAL(9,C{0}:C{1})", g_row + 1, current_row - 1);
                                if (NCP_old != "C8")
                                {
                                    oSheet.Cells[g_row, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", g_row + 1, current_row - 1);
                                    oSheet.Cells[g_row, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", g_row + 1, current_row - 1);
                                    oSheet.Cells[g_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", g_row + 1, current_row - 1);
                                }
                                else
                                {
                                    oSheet.Cells[g_row, 6].Formula = string.Format("=C{0}-E{0}", g_row );
                                }
                                
                            }
                            STT_G++;
                            STT = 1;
                            g_row = current_row;
                            current_row++;
                        }

                        //STT
                        oSheet.Cells[current_row, 1] = STT;
                        //Ten CP
                        oSheet.Cells[current_row, 2] = r["TEN_CP"].ToString();
                        //Du Tru
                        oSheet.Cells[current_row, 3] = r["DuTru"];
                        if (NCP != "C8")
                        {
                            //KT No VAT
                            oSheet.Cells[current_row, 4] = r["KT_NO_VAT"];
                            //CCM no VAT
                            oSheet.Cells[current_row, 5] = r["GT_NO_VAT"];
                            //CP Con lai
                            oSheet.Cells[current_row, 6].Formula = string.Format("=C{0}-E{0}", current_row);
                        }
                        else
                        {
                            //KT No VAT
                            oSheet.Cells[current_row, 4] = 0;
                            //CCM no VAT
                            oSheet.Cells[current_row, 5] = 0;
                            //CP Con lai
                            oSheet.Cells[current_row, 6] = 0;
                        }
                        
                        //Phan tram hoan thanh
                        oSheet.Cells[current_row, 7].Formula = string.Format("=IF(ISERROR(E{0}/C{0}),0,E{0}/C{0})", current_row);

                        STT++;
                        current_row++;
                    }
                    //Total Last CP
                    //if (g2_row < current_row)
                    //{
                    //        //oSheet.Cells[g2_row, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", g2_row + 1, current_row - 1);
                    //        oSheet.Cells[g2_row, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", g2_row + 1, current_row - 1);
                    //        oSheet.Cells[g2_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", g2_row + 1, current_row - 1);
                    //        oSheet.Cells[g2_row, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", g2_row + 1, current_row - 1);
                    //        oSheet.Cells[g2_row, 7].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", g2_row + 1, current_row - 1);
                    //}
                    //Total Last Group
                    if (g_row < current_row)
                    {
                        oSheet.Cells[g_row, 3].Formula = string.Format("=SUBTOTAL(9,C{0}:C{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", g_row + 1, current_row - 1);
                        oSheet.Cells[g_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", g_row + 1, current_row - 1);
                    }

                    //TOTAL
                    oSheet.Range["A" + current_row, "B" + current_row].Merge();
                    oSheet.Range["A" + current_row, "B" + current_row].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                    oSheet.Range["A" + current_row, "H" + current_row].Interior.Color = System.Drawing.Color.FromArgb(255, 217, 102);

                    oSheet.Cells[current_row, 1] = "TỔNG CỘNG";
                    oSheet.Cells[current_row, 3].Formula = string.Format("=SUBTOTAL(9,C{0}:C{1})", 5, current_row - 1);
                    oSheet.Cells[current_row, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", 5, current_row - 1);
                    oSheet.Cells[current_row, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", 5, current_row - 1);
                    oSheet.Cells[current_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", 5, current_row - 1);

                    current_row++;

                    //Border
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A5", "H" + (current_row - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    oXL.ActiveWindow.Activate();

                }
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
        }

        System.Data.DataTable Get_Data(DateTime pToDate)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_TONGHOP_VP_DETAILS_DATA", conn);
                cmd.CommandType = CommandType.StoredProcedure;
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
