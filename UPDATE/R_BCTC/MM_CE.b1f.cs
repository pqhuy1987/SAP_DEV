using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;
using System.Runtime.InteropServices;

namespace R_BCTC
{
    [FormAttribute("R_BCTC.MM_CE", "MM_CE.b1f")]
    class MM_CE : UserFormBase
    {
        public MM_CE()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_prj").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_v").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("grd_sp").Specific));
            this.Grid0.PressedAfter += new SAPbouiCOM._IGridEvents_PressedAfterEventHandler(this.Grid0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("bt_bl").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.CloseBefore += new CloseBeforeHandler(this.Form_CloseBefore);

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
            //Load_Month();
        }

        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button1;

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Grid0.DataTable.ExecuteQuery(string.Format("SELECT 'N' as 'Checked',AbsEntry,NAME FROM OPMG T0 WHERE T0.[FIPROJECT] = '{0}'and T0.[STATUS] <> 'T' ORDER BY AbsEntry", this.ComboBox0.Selected.Value));
                Grid0.Columns.Item(1).Editable = false;
                Grid0.Columns.Item(2).Editable = false;
                Grid0.Columns.Item("Checked").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                Grid0.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
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

        //void Load_SubProject(string pFProject)
        //{
        //    if (this.ComboBox1.ValidValues.Count > 1)
        //    {
        //        //Remove Valid Value
        //        this.ComboBox1.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //        int itm_count = ComboBox1.ValidValues.Count;
        //        for (int i = 0; i < itm_count - 1; i++)
        //        {
        //            this.ComboBox1.ValidValues.Remove(1, SAPbouiCOM.BoSearchKey.psk_Index);
        //        }
        //    }
        //    SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    //oR_RecordSet.DoQuery("Select AbsEntry, NAME as 'Description' from OPHA where TYP =1 and ProjectID = " + pProject_AbsEntry.ToString());
        //    oR_RecordSet.DoQuery(string.Format("SELECT AbsEntry,NAME FROM OPMG T0 WHERE T0.[FIPROJECT] = '{0}'and T0.[STATUS] <> 'T' ORDER BY AbsEntry", pFProject));
        //    try
        //    {
        //        this.ComboBox1.ValidValues.Add("", "");
        //    }
        //    catch
        //    { }
        //    if (oR_RecordSet.RecordCount > 0)
        //    {
        //        while (!oR_RecordSet.EoF)
        //        {
        //            ComboBox1.ValidValues.Add(oR_RecordSet.Fields.Item("AbsEntry").Value.ToString(), oR_RecordSet.Fields.Item("NAME").Value.ToString());
        //            oR_RecordSet.MoveNext();
        //        }
        //    }
        //}

        //void Load_Month()
        //{
        //    SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    oR_RecordSet.DoQuery("Select Code from OFPR order by Code desc");
        //    if (oR_RecordSet.RecordCount > 0)
        //    {
        //        while (!oR_RecordSet.EoF)
        //        {
        //            ComboBox2.ValidValues.Add(oR_RecordSet.Fields.Item("Code").Value.ToString(), oR_RecordSet.Fields.Item("Code").Value.ToString());
        //            oR_RecordSet.MoveNext();
        //        }
        //    }

        //}

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
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

                //Project Name
                oSheet.Cells[2, 3] = "Dự án: " + this.ComboBox0.Selected.Description;
                //Subproject Name
                if (!string.IsNullOrEmpty(GoiThau_Key))
                {
                    if (GoiThau_Key.Split(',').Count() == 1)
                        oSheet.Cells[3, 3] = "Gói thầu: " + GoiThau_Name;
                }
                //oSheet.Cells[3, 3] = "Gói thầu: " + this.ComboBox1.Selected.Description;
                //Thang
                oSheet.Cells[4, 3] = "Tháng: " + DateTime.Today.ToString("MM-yyyy");//this.ComboBox2.Selected.Value;
                DataTable A = null;
                if (!string.IsNullOrEmpty(GoiThau_Key))
                    A = Get_Data_BCDTA(this.ComboBox0.Selected.Value, GoiThau_Key);
                else
                    A = Get_Data_BCDTA(this.ComboBox0.Selected.Value);
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
                    B = Get_Data_DUTRU_SUM(this.ComboBox0.Selected.Value);
                    C = Get_Data_DUTRU(this.ComboBox0.Selected.Value);
                }
                else
                {
                    B = Get_Data_DUTRU_SUM(this.ComboBox0.Selected.Value, GoiThau_Key);
                    C = Get_Data_DUTRU(this.ComboBox0.Selected.Value, GoiThau_Key);
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
                            oSheet.Cells[current_rownum, 3] = rd["U_BPCode"].ToString();
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
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{0}", detail_rownum -1 );
                    }

                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Italic = true;
                    oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    oSheet.Cells[current_rownum, 2] = "Nhà thầu phụ";
                    oSheet.Cells[current_rownum, 5].Value2 = r["U_CP_NTP"];
                    detail_rownum = current_rownum + 1;
                    current_rownum++;
                    foreach (DataRow rd in C.Select("U_DTT_LineID=" + r["LineID"]))
                    {
                        if (decimal.Parse(rd["U_CP_NTP"].ToString()) != 0)
                        {
                            oSheet.Cells[current_rownum, 2] = rd["U_BPNAME"].ToString();
                            oSheet.Cells[current_rownum, 3] = rd["U_BPCode"].ToString();
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
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{0}", detail_rownum -1);
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
                            oSheet.Cells[current_rownum, 3] = rd["U_BPCode"].ToString();
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
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{0}", detail_rownum -1);
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
                            oSheet.Cells[current_rownum, 3] = rd["U_BPCode"].ToString();
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
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{0}", detail_rownum -1);
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
                            oSheet.Cells[current_rownum, 3] = rd["U_BPCode"].ToString();
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
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{0}", detail_rownum -1);
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
                            oSheet.Cells[current_rownum, 3] = rd["U_BPCode"].ToString();
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
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{0}", detail_rownum -1);
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
                            oSheet.Cells[current_rownum, 3] = rd["U_BPCode"].ToString();
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
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{0}", detail_rownum -1);
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
                            oSheet.Cells[current_rownum, 3] = rd["U_BPCode"].ToString();
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
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{0}", detail_rownum -1);
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
                            oSheet.Cells[current_rownum, 3] = rd["U_BPCode"].ToString();
                            oSheet.Cells[current_rownum, 6].Value2 = rd["U_CP_K"];
                            oSheet.Cells[current_rownum, 7].Value2 = rd["KL_TT_DD"];
                            current_rownum++;
                        }
                    }
                    if (current_rownum - detail_rownum >= 1)
                    {
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 6]).Formula = string.Format("=SUBTOTAL(9,{0})", "F" + detail_rownum + ":F" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 7]).Formula = string.Format("=SUBTOTAL(9,{0})", "G" + detail_rownum + ":G" + (current_rownum - 1));
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[detail_rownum - 1, 8]).Formula = string.Format("=G{0}/F{0}", detail_rownum -1);
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
                DataTable D = Get_Data_BCH(this.ComboBox0.Selected.Value, GoiThau_Key);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 1] = "II";
                oSheet.Cells[current_rownum, 2] = "CHI PHÍ QUẢN LÝ BCH TRỰC TIẾP";
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", (current_rownum + 1) , (current_rownum + 37));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", (current_rownum + 1) , (current_rownum + 37));
                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", (current_rownum + 1) , (current_rownum + 37));
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
                //DataTable E = Get_Prj_Info(this.ComboBox0.Selected.Value);
                DataTable VII = Get_Data_VII(this.ComboBox0.Selected.Value, GoiThau_Key);
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
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 1] = "III";
                oSheet.Cells[current_rownum, 2] = "CHI PHÍ HỔ TRỢ";
                Section_RowNum.Add(current_rownum);
                current_rownum++;
                //Chi phi ho tro 1
                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phi hỗ trợ 1";
                //decimal tmp_cpht = 0;
                //decimal.TryParse(E.Rows[0]["U_CPHT1"].ToString(), out tmp_cpht);
                if (f_ht1 != "")
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = "=" + f_ht1.Substring(0, f_ht1.Length - 3);
                    //string.Format("={0}*{1}/100", "D6", tmp_cpht);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = "=" + f_ht1.Substring(0, f_ht1.Length - 3);
                    //string.Format("={0}*{1}/100", "D6", tmp_cpht);
                }
                current_rownum++;
                //Chi phi ho tro 2
                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                oSheet.Cells[current_rownum, 2] = "Chi phi hỗ trợ 2";
                //decimal tmp_cpht2 = 0;
                //decimal.TryParse(E.Rows[0]["U_CPHT2"].ToString(), out tmp_cpht2);
                if (f_ht2 != "")
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = "=" + f_ht2.Substring(0, f_ht2.Length - 3);
                        //string.Format("={0}*{1}/100", "D6", tmp_cpht2);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = "=" + f_ht2.Substring(0, f_ht2.Length - 3);
                        //string.Format("={0}*{1}/100", "D6", tmp_cpht2);
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
                //decimal tmp_cpng = 0;
                //decimal.TryParse(E.Rows[0]["U_CPNG"].ToString(), out tmp_cpng);
                if (f_ng != "")
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = "=" + f_ng.Substring(0, f_ng.Length - 3);
                    //.Value2 = tmp_cpng; //string.Format("={0}*{1}/100", "D6",
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = "=" + f_ng.Substring(0, f_ng.Length - 3);
                    //.Value2 = tmp_cpng;
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
                
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(201, 201, 201);
                oSheet.Range["A" + current_rownum, "H" + current_rownum].Font.Bold = true;
                oSheet.Cells[current_rownum, 1] = "IV";
                oSheet.Cells[current_rownum, 2] = "DỰ PHÒNG PHÍ";
                Section_RowNum.Add(current_rownum);
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                //oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 2] = "Dự phòng chi phí cho ĐTC/ NTP/ NCC (0.5% giá trị doanh thu)";
                //decimal tmp_dp1 = 0;
                //decimal.TryParse(E.Rows[0]["U_DPCP"].ToString(), out tmp_dp1);
                if (f_dpcp != "")
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = "=" + f_dpcp.Substring(0, f_dpcp.Length - 3);
                        //string.Format("={0}*{1}/100", "D6", tmp_dp1);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = "=" + f_dpcp.Substring(0, f_dpcp.Length - 3);
                        //string.Format("={0}*{1}/100", "D6", tmp_dp1);
                }
                current_rownum++;

                oSheet.Range["B" + current_rownum, "H" + current_rownum].Font.Italic = true;
                //oSheet.Cells[current_rownum, 4].Value2 = D.Select("U_TKKT='62770'").Count<DataRow>() > 0 ? D.Select("U_TKKT='62770'")[0]["U_GTDP"] : "";
                oSheet.Cells[current_rownum, 2] = "Dự phòng chi phí bảo hành (0.5% giá trị doanh thu)";
                //decimal tmp_dpbh = 0;
                //decimal.TryParse(E.Rows[0]["U_DPBH"].ToString(), out tmp_dpbh);
                if (f_dpbh != "")
                {
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = "=" + f_dpbh.Substring(0, f_dpbh.Length - 3);
                    //string.Format("={0}*{1}/100", "D6", tmp_dpbh);
                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 6]).Formula = "=" + f_dpbh.Substring(0, f_dpbh.Length - 3);
                    //string.Format("={0}*{1}/100", "D6", tmp_dpbh);
                }
                current_rownum++;

                //Total IV
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

        System.Data.DataTable Get_Data_BCDTA(string pFinancialProject, string pGoiThauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("GET_DATA_BCDT_A", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
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

        System.Data.DataTable Get_Data_DUTRU(string pFinancialProject, string pGoiThauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("MM_CE_GETDATA_DETAILS_NEW", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
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

        System.Data.DataTable Get_Data_DUTRU_SUM(string pFinancialProject, string pGoiThauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("MM_CE_GETDATA_SUM_NEW", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
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

        System.Data.DataTable Get_Data_BCH(string pFinancialProject, string pGoiThauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("MM_CE_GET_DATA_BCH_NEW", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
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

        System.Data.DataTable Get_Data_VII(string pFinancialProject, string pGoiThauKey = "")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("MM_FI_GET_DATA_VII", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
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

        System.Data.DataTable Get_List_FProject()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("MM_CE_GET_FPROJECT", conn);
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

        private void Grid0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
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

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                BASELINE frm = new BASELINE(this.ComboBox0.Selected.Value);
                frm.Show();
            }
            catch
            {
 
            }
        }

    }
}
