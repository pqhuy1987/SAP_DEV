using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
namespace R_BCTC
{
    [FormAttribute("R_BCTC.MM_FI", "MM_FI.b1f")]
    class MM_FI : UserFormBase
    {
        public MM_FI()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;

        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.ComboBox ComboBox2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.ComboBox ComboBox3;
        private SAPbouiCOM.Button Button0;

        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_prj").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_sprj").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_m").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.ComboBox3 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_y").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_rep").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("bt_bl").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
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
            Load_Month();
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

        void Load_SubProject(string pFProject)
        {
            if (this.ComboBox1.ValidValues.Count > 1)
            {
                //Remove Valid Value
                this.ComboBox1.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                int itm_count = ComboBox1.ValidValues.Count;
                for (int i = 0; i < itm_count-1; i++)
                {
                    this.ComboBox1.ValidValues.Remove(1, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            }
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //oR_RecordSet.DoQuery("Select AbsEntry, NAME as 'Description' from OPHA where TYP =1 and ProjectID = " + pProject_AbsEntry.ToString());
            oR_RecordSet.DoQuery(string.Format("SELECT AbsEntry,NAME FROM OPMG T0 WHERE T0.[FIPROJECT] = '{0}'and T0.[STATUS] <> 'T' ORDER BY AbsEntry", pFProject));
            try
            {
                this.ComboBox1.ValidValues.Add("", "");
            }
            catch
            { }
            if (oR_RecordSet.RecordCount > 0)
            {
                while (!oR_RecordSet.EoF)
                {
                    ComboBox1.ValidValues.Add(oR_RecordSet.Fields.Item("AbsEntry").Value.ToString(), oR_RecordSet.Fields.Item("NAME").Value.ToString());
                    oR_RecordSet.MoveNext();
                }
            }
        }

        void Load_Month()
        {
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery("Select Code from OFPR order by Code desc");
            if (oR_RecordSet.RecordCount > 0)
            {
                while (!oR_RecordSet.EoF)
                {
                    ComboBox2.ValidValues.Add(oR_RecordSet.Fields.Item("Code").Value.ToString(), oR_RecordSet.Fields.Item("Code").Value.ToString());
                    oR_RecordSet.MoveNext();
                }
            }
 
        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
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
            oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_BCTC.xlsx");
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            int current_row = 7;
            try
            {
                DataTable A = Get_Data_BCDTA(this.ComboBox0.Selected.Value);
                List<int> Group_No_RowNum = new List<int>();
                List<int> Section_RowNum = new List<int>();
                //bool ME_Project = false;
                //Fill Header
                //Project Name
                oSheet.Cells[2, 4] = "Dự án: " + this.ComboBox0.Selected.Description;
                //Subproject Name
                oSheet.Cells[3, 4] = "Gói thầu: " + this.ComboBox1.Selected.Description;
                //Thang
                oSheet.Cells[4, 4] = "Tháng: " + DateTime.Today.ToString("MM-yyyy");// this.ComboBox2.Selected.Value;

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
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[6,6]).Formula = string.Format("=SUM({0}:{1})", "F7", "F13");
                //B-CHI PHI (Trước VAT)
                oSheet.Range["A" + current_row, "K"+current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(244, 176, 132);
                oSheet.Cells[current_row, 1] = "B";
                oSheet.Cells[current_row, 3] = "CHI PHÍ (Trước VAT)";
                current_row++;
                //DataTable B = null;
                DataTable C = null;
                int GoiThauKey = -1;
                int.TryParse(this.ComboBox1.Selected.Value, out GoiThauKey);
                if (GoiThauKey == 0)
                {
                    //B = Get_Data_DUTRU_SUM(this.ComboBox0.Selected.Description);
                    C = Get_Data_DUTRU(this.ComboBox0.Selected.Value);
                }
                else
                {
                    //B = Get_Data_DUTRU_SUM(this.ComboBox0.Selected.Description,-1, SubProjectKey);
                    C = Get_Data_DUTRU(this.ComboBox0.Selected.Value, -1, GoiThauKey);
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
                    decimal tmp_cp = 0,tmp_klhd =0,tmp_kltt = 0, tmp_klttdd , tmp_total_apinvoice = 0;
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
                foreach (DataRow d in C.Select("U_TYPE = 'ME'"))
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
                foreach (DataRow d in C.Select("U_TYPE = 'ME'"))
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
                foreach (DataRow d in C.Select("U_TYPE = 'TB'"))
                {
                    decimal tmp_ncc = 0, tmp_klhd = 0, tmp_kltt = 0, tmp_klttdd = 0, tmp_total_apinvoice=0;
                    decimal.TryParse(d["U_CP_NCC"].ToString(), out tmp_ncc);
                    decimal.TryParse(d["KL_HD"].ToString(), out tmp_klhd);
                    decimal.TryParse(d["KL_TT"].ToString(), out tmp_kltt);
                    decimal.TryParse(d["KL_TT_DD"].ToString(), out tmp_klttdd);
                    decimal.TryParse(d["TOTAL_AP_INVOICE"].ToString(), out tmp_total_apinvoice);
                    if ((tmp_ncc != 0 && d["U_PUType"].ToString() == "PUT03") || (tmp_ncc != 0 && string.IsNullOrEmpty(d["U_PUType"].ToString()))) // || tmp_klhd != 0 || tmp_kltt != 0 || tmp_klttdd != 0)
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
                foreach (DataRow d in C.Select("U_TYPE = 'TB'"))
                {
                    decimal tmp_ntp = 0, tmp_klhd = 0, tmp_kltt = 0, tmp_klttdd = 0, tmp_total_apinvoice = 0;
                    decimal.TryParse(d["U_CP_NTP"].ToString(), out tmp_ntp);
                    decimal.TryParse(d["KL_HD"].ToString(), out tmp_klhd);
                    decimal.TryParse(d["KL_TT"].ToString(), out tmp_kltt);
                    decimal.TryParse(d["KL_TT_DD"].ToString(), out tmp_klttdd);
                    decimal.TryParse(d["TOTAL_AP_INVOICE"].ToString(), out tmp_total_apinvoice);
                    if ((tmp_ntp != 0 && d["U_PUType"].ToString() == "PUT04") || (tmp_ntp != 0 && string.IsNullOrEmpty(d["U_PUType"].ToString()))) // || tmp_klhd != 0 || tmp_kltt != 0 || tmp_klttdd != 0)
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
                DataTable D = Get_Data_BCH(this.ComboBox0.Selected.Value);
                //1
                //oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                oSheet.Cells[current_row, 1] = "1";
                oSheet.Cells[current_row, 3] = "Chi phí lương, bảo hiểm, phụ cấp, công trường ...";
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=SUM({0}:{1})", "F"+(current_rownum +1),"F"+ (current_rownum + 7));
                //oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='CPQL0000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='CPQL0000'")[0]["U_GTDP"] : "";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (current_row + 1) + ":I" + (current_row + 6));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (current_row + 1) + ":J" + (current_row + 6));
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Phải trả công nhân viên";
                oSheet.Cells[current_row, 4] = "3341";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='33410000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33410000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='33410000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33410000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='33410000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33410000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Phải trả người lao động khác (đội thi công)";
                oSheet.Cells[current_row, 4] = "33481";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='33481000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33481000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='33481000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33481000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='33481000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33481000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí lương kỹ thuật viên";
                oSheet.Cells[current_row, 4] = "33482";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='33482000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33482000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='33482000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33482000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='33482000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33482000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí vệ sinh, giữ xe,.. công trường (BCH)";
                oSheet.Cells[current_row, 4] = "33483";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='33483000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33483000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='33483000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33483000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='33483000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33483000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí lương an toàn viên";
                oSheet.Cells[current_row, 4] = "33484";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='33484000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33484000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='33484000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33484000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='33484000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='33484000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "BHXH,BHYT,KPCĐ,BHTN";
                oSheet.Cells[current_row, 4] = "62712";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62712000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62712000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62712000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62712000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62712000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62712000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Cells[current_row, 1] = "2";
                oSheet.Cells[current_row, 3] = "Chi phí vật tư lẻ";
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=SUM({0}:{1})", "F" + (current_rownum + 1), "F" + (current_rownum + 5));
                //oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='CPVTL000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='CPVTL000'")[0]["U_GTDP"] : "";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (current_row + 1) + ":I" + (current_row + 4));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (current_row + 1) + ":J" + (current_row + 4));
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí nguyên vật liệu trực tiếp";
                oSheet.Cells[current_row, 4] = "621";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62100000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62100000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62100000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62100000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62100000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62100000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Nhiên liệu";
                oSheet.Cells[current_row, 4] = "62781";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62781000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62781000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62781000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62781000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62781000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62781000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí bằng tiền khác";
                oSheet.Cells[current_row, 4] = "62788";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62788000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62788000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62788000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62788000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62788000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62788000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Bảo hộ lao động";
                oSheet.Cells[current_row, 4] = "62733";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62733000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62733000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62733000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62733000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62733000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62733000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Cells[current_row, 1] = "3";
                oSheet.Cells[current_row, 3] = "Chi phí máy móc thiết bị";
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=SUM({0}:{1})", "F" + (current_rownum + 1), "F" + (current_rownum + 7));
                //oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='MMTB0000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='MMTB0000'")[0]["U_GTDP"] : "";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (current_row + 1) + ":I" + (current_row + 6));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (current_row + 1) + ":J" + (current_row + 6));
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Công cụ, dụng cụ, thiết bị Ban chỉ huy CT";
                oSheet.Cells[current_row, 4] = "62731";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62731000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62731000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62731000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62731000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62731000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62731000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "VPP, photocopy";
                oSheet.Cells[current_row, 4] = "62732";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62732000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62732000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62732000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62732000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62732000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62732000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí vận chuyển";
                oSheet.Cells[current_row, 4] = "62734";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62734000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62734000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62734000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62734000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62734000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62734000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Điện, nước thi công";
                oSheet.Cells[current_row, 4] = "62774";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62774000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62774000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62774000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62774000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62774000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62774000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Điện thoại cố định";
                oSheet.Cells[current_row, 4] = "62775";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62775000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62775000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62775000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62775000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62775000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62775000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Thuê TSCĐ, thiết bị thi công";
                oSheet.Cells[current_row, 4] = "62776";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62776000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62776000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62776000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62776000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62776000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62776000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Cells[current_row, 1] = "4";
                oSheet.Cells[current_row, 3] = "Chi phí ban chỉ huy văn phòng";
                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=SUM({0}:{1})", "F" + (current_rownum + 1), "F" + (current_rownum + 18));
                //oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='BCHVP000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='BCHVP000'")[0]["U_GTDP"] : "";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 9]).Formula = string.Format("=SUBTOTAL(9,{0})", "I" + (current_row + 1) + ":I" + (current_row + 17));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 10]).Formula = string.Format("=SUBTOTAL(9,{0})", "J" + (current_row + 1) + ":J" + (current_row + 17));
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Ăn trưa";
                oSheet.Cells[current_row, 4] = "62713";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62713000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62713000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62713000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62713000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62713000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62713000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Điện thoại di động";
                oSheet.Cells[current_row, 4] = "62714";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62714000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62714000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62714000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62714000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62714000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62714000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí thuê nhà";
                oSheet.Cells[current_row, 4] = "62716";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62716000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62716000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62716000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62716000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62716000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62716000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Thuế xuất nhập khẩu";
                oSheet.Cells[current_row, 4] = "62723";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62723000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62723000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62723000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62723000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62723000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62723000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Bao chí, bưu phí, tài liệu";
                oSheet.Cells[current_row, 4] = "62735";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62735000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62735000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62735000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62735000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62735000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62735000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Phí, lệ phí";
                oSheet.Cells[current_row, 4] = "62770";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62770000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62770000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62770000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Tiếp khách";
                oSheet.Cells[current_row, 4] = "62771";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62771000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62771000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62771000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62771000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62771000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62771000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Phí kiểm định, thí nghiệm";
                oSheet.Cells[current_row, 4] = "62773";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62773000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62773000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62773000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62773000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62773000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62773000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Phí ngân hàng";
                oSheet.Cells[current_row, 4] = "62777";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62777000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62777000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62777000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62777000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62777000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62777000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Quảng cáo, đào tạo";
                oSheet.Cells[current_row, 4] = "62778";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62778000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62778000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62778000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62778000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62778000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62778000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí nhà thầu phụ";
                oSheet.Cells[current_row, 4] = "62779";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62779000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62779000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62779000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62779000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62779000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62779000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí giao nhận hàng hóa nhập khẩu";
                oSheet.Cells[current_row, 4] = "62782";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62782000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62782000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62782000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62782000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62782000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62782000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Công tác phí";
                oSheet.Cells[current_row, 4] = "62783";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62783000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62783000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62783000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62783000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62783000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62783000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phí bị loại trừ";
                oSheet.Cells[current_row, 4] = "62784";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62784000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62784000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62784000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62784000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62784000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62784000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Thuốc, y tế, đồ dùng lặt vặt";
                oSheet.Cells[current_row, 4] = "62785";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62785000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62785000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62785000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62785000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62785000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62785000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Hồ sơ thầu";
                oSheet.Cells[current_row, 4] = "62786";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62786000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62786000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62786000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62786000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62786000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62786000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Phí bảo hiểm";
                oSheet.Cells[current_row, 4] = "62787";
                oSheet.Cells[current_row, 7].Value2 = D.Select("U_TKKT='62787000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62787000'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 9].Value2 = D.Select("U_TKKT='62787000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62787000'")[0]["TOTAL_BCH"] : "";
                oSheet.Cells[current_row, 10].Value2 = D.Select("U_TKKT='62787000'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62787000'")[0]["TOTAL_BCH"] : "";
                current_row++;

                //DU PHONG PHI
                DataTable E = Get_Prj_Info(this.ComboBox0.Selected.Value);

                oSheet.Range["A" + current_row, "K" + current_row].Font.Bold = true;
                oSheet.Range["A" + current_row, "K" + current_row].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Cells[current_row, 1] = "VI";
                oSheet.Cells[current_row, 3] = "DỰ PHÒNG PHÍ";
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                //oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 3] = "Dự phòng chi phí cho ĐTC/ NTP/ NCC (0.5% giá trị doanh thu)";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("={0}*{1}/100", "F6", E.Rows[0]["U_DPCP"].ToString());
                current_row++;

                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                //oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_row, 3] = "Dự phòng chi phí bảo hành (0.5% giá trị doanh thu)";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("={0}*{1}/100", "F6", E.Rows[0]["U_DPBH"].ToString());
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
                //oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("={0}*{1}/100", "F6", E.Rows[0]["U_CPHT1"].ToString());
                current_row++;
                //Chi phi ho tro 2
                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phi hỗ trợ 2";
                //oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = string.Format("={0}*{1}/100", "F6", E.Rows[0]["U_CPHT2"].ToString());
                current_row++;
                //Chi phi NG
                oSheet.Range["B" + current_row, "H" + current_row].Font.Italic = true;
                oSheet.Cells[current_row, 3] = "Chi phi NG";
                //oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 7]).Formula = E.Rows[0]["U_CPNG"].ToString(); //string.Format("={0}*{1}/100", "D6",
                current_row++;

                //Total Ho tro
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row -4, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + (current_row - 3) + ":G" + (current_row - 1));

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
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 4]).Formula = "=F6";
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "CHI PHÍ";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 5]).Formula = "=G14";
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "LỢI NHUẬN GỘP A";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 5]).Formula = string.Format("={0}-{1}", "D" + (current_row - 2), "E" + (current_row - 1));
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "TỶ SUẤT LỢI NHUẬN GỘP A/ DOANH THU";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 5]).Formula = string.Format("={0}/{1}", "E" + (current_row - 1), "D" + (current_row - 3));
                oSheet.Range["E" + current_row].NumberFormat = "0.00%";

                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 6]).Formula = string.Format("=((F6-F8)-(G14-{0}))/(F6-F8)", "E" + NCC_NTP_KHAC_row);
                oSheet.Range["F" + current_row].NumberFormat = "0.00%";
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
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 5]).Formula = string.Format("={0}*{1}/100", "F6", E.Rows[0]["U_CPQLCT"].ToString());
                oSheet.Range["E" + current_row].NumberFormat = "_(* #,##0_);_(* (#,##0);_(* \" - \"??_);_(@_)";
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "DOANH THU";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 4]).Formula = "=F6";
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "CHI PHÍ";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 5]).Formula = string.Format("=G14+E{0}", current_row - 2);
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "LỢI NHUẬN TUYỆT ĐỐI C";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 5]).Formula = string.Format("={0}-{1}", "D" + (current_row - 2), "E" + (current_row - 1));
                current_row++;

                oSheet.Range["A" + current_row, "H" + current_row].Font.Bold = true;
                oSheet.Cells[current_row, 3] = "TỶ SUẤT LỢI TUYỆT ĐỐI C/ DOANH THU";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_row, 5]).Formula = string.Format("={0}/{1}", "E" + (current_row - 1), "D" + (current_row - 3));
                oSheet.Range["E" + current_row].NumberFormat = "0.00%";
                current_row++;

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

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            Load_SubProject(this.ComboBox0.Selected.Value);
            this.ComboBox1.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
        }

        System.Data.DataTable Get_Data_BCDTA(string pFinancialProject)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("MM_FI_GET_DATA_A", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
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

        System.Data.DataTable Get_Data_DUTRU(string pFinancialProject, int pCTG_DocEntry = -1, int pGoithauKey = -1)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("MM_FI_GET_DATA_B", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                cmd.Parameters.AddWithValue("@Goithau_Key", pGoithauKey);
                //cmd.Parameters.AddWithValue("@GoithauKey", pGoithauKey);
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

        //System.Data.DataTable Get_Data_DUTRU_SUM(string pFinancialProject, int pCTG_DocEntry = -1, int pGoithauKey = -1)
        //{
        //    DataTable result = new DataTable();
        //    SqlCommand cmd = null;
        //    try
        //    {
        //        cmd = new SqlCommand("GET_DATA_DUTRU_SUM", conn);
        //        cmd.CommandType = CommandType.StoredProcedure;
        //        cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
        //        cmd.Parameters.AddWithValue("@CTG_Entry", pCTG_DocEntry);
        //        cmd.Parameters.AddWithValue("@GoithauKey", pGoithauKey);
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

        System.Data.DataTable Get_Data_BCH(string pFinancialProject, int pCTG_DocEntry = -1, int pGoiThauKey = -1)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("MM_FI_GET_DATA_BCH", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                //cmd.Parameters.AddWithValue("@CTG_Entry", pCTG_DocEntry);
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
                cmd = new SqlCommand(string.Format("Select * from OPMG a where a.FIPROJECT='{0}' and a.STATUS <> 'T'", pFinancialProject), conn);
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

        private SAPbouiCOM.Button Button1;

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                BASELINE frm = new BASELINE(this.ComboBox0.Selected.Value);
                frm.Show();
            }
            catch
            { }
        }
    }
}