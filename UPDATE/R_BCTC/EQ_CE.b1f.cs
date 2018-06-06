using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;

namespace R_BCTC
{
    [FormAttribute("R_BCTC.EQ_CE", "EQ_CE.b1f")]
    class EQ_CE : UserFormBase
    {
        public EQ_CE()
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
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_fpro").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_subp").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_m").Specific));
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

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.ComboBox ComboBox2;
        private SAPbouiCOM.Button Button0;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            int GoiThauKey = -1;
            if (ComboBox1.Selected.Value == "")
            {
                oApp.MessageBox("Please select Subproject !");
                return;
            }
            int.TryParse(ComboBox1.Selected.Value, out GoiThauKey);
            //throw new System.NotImplementedException();
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            //Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;

            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            //Open Template
            oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_EQ_BCDT.xlsx");

            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            try
            {
                //Fill Header
                //Project Name
                oSheet.Cells[2, 4] = "Dự án: " + this.ComboBox0.Selected.Description;
                //Subproject Name
                oSheet.Cells[3, 4] = "Gói thầu: " + this.ComboBox1.Selected.Description;
                //Thang
                //oSheet.Cells[4, 3] = "Tháng: " + this.ComboBox2.Selected.Value;

                //Get DATA
                DataTable A = Get_Data_DUTRU(this.ComboBox0.Selected.Value, -1, GoiThauKey);
                int row_num = 8;
                int STT = 1;
                int Group_row_num = 0;
                oSheet.Range["A" + row_num, "T" + row_num].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + row_num, "T" + row_num].Font.Bold = true;
                oSheet.Cells[row_num, 1] = "I";
                oSheet.Cells[row_num, 2] = "THIẾT BỊ NÂNG HẠ";
                Group_row_num = row_num;
                row_num++;

                foreach (DataRow r in A.Select("U_001 = 'TBNH'"))
                {
                    oSheet.Cells[row_num, 1] = STT;
                    oSheet.Cells[row_num, 2] = /*r["U_MATHIETBI"].ToString() + " - " +*/ r["U_TENTHIETBI"];
                    oSheet.Cells[row_num, 3] = r["U_DVTTB"];
                    oSheet.Cells[row_num, 4] = r["U_TLTB"];
                    oSheet.Cells[row_num, 5] = r["U_SLDUTRU"];
                    oSheet.Cells[row_num, 6] = r["U_SLTHUE"];
                    oSheet.Cells[row_num, 7] = r["U_SLVANCHUYEN"];
                    oSheet.Cells[row_num, 8] = r["U_SLVANHANH"];
                    oSheet.Cells[row_num, 9] = r["U_NGAYCAP"];
                    oSheet.Cells[row_num, 10] = r["U_NGAYTRA"];
                    oSheet.Cells[row_num, 11].Formula = string.Format("=J{0}-I{0}", row_num);
                    oSheet.Cells[row_num, 12] = r["U_DGMUABAN"];
                    oSheet.Cells[row_num, 13] = r["U_DGTHUE"];
                    oSheet.Cells[row_num, 14] = r["U_DGVCTB"];
                    oSheet.Cells[row_num, 15] = r["U_DGVH"];
                    oSheet.Cells[row_num, 16] = r["U_GTMB"];
                    oSheet.Cells[row_num, 17] = r["U_GTTHUE"];
                    oSheet.Cells[row_num, 18] = r["U_GTVANCHUYEN"];
                    oSheet.Cells[row_num, 19] = r["U_GTVANHANH"];
                    STT++;
                    row_num++;
                }
                //GROUP TOTAL
                if (row_num - Group_row_num > 1)
                {
                    oSheet.Cells[Group_row_num, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 16].Formula = string.Format("=SUBTOTAL(9,P{0}:P{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 17].Formula = string.Format("=SUBTOTAL(9,Q{0}:Q{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 18].Formula = string.Format("=SUBTOTAL(9,R{0}:R{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 19].Formula = string.Format("=SUBTOTAL(9,S{0}:S{1})", Group_row_num + 1, row_num - 1);
                }


                oSheet.Range["A" + row_num, "T" + row_num].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + row_num, "T" + row_num].Font.Bold = true;
                oSheet.Cells[row_num, 1] = "II";
                oSheet.Cells[row_num, 2] = "THIẾT BỊ BAO CHE";
                STT = 1;
                Group_row_num = row_num;
                row_num++;

                foreach (DataRow r in A.Select("U_001 = 'TBBC'"))
                {
                    oSheet.Cells[row_num, 1] = STT;
                    oSheet.Cells[row_num, 2] = /*r["U_MATHIETBI"].ToString() + " - " +*/ r["U_TENTHIETBI"];
                    oSheet.Cells[row_num, 3] = r["U_DVTTB"];
                    oSheet.Cells[row_num, 4] = r["U_TLTB"];
                    oSheet.Cells[row_num, 5] = r["U_SLDUTRU"];
                    oSheet.Cells[row_num, 6] = r["U_SLTHUE"];
                    oSheet.Cells[row_num, 7] = r["U_SLVANCHUYEN"];
                    oSheet.Cells[row_num, 8] = r["U_SLVANHANH"];
                    oSheet.Cells[row_num, 9] = r["U_NGAYCAP"];
                    oSheet.Cells[row_num, 10] = r["U_NGAYTRA"];
                    oSheet.Cells[row_num, 11].Formula = string.Format("=J{0}-I{0}", row_num);
                    oSheet.Cells[row_num, 12] = r["U_DGMUABAN"];
                    oSheet.Cells[row_num, 13] = r["U_DGTHUE"];
                    oSheet.Cells[row_num, 14] = r["U_DGVCTB"];
                    oSheet.Cells[row_num, 15] = r["U_DGVH"];
                    oSheet.Cells[row_num, 16] = r["U_GTMB"];
                    oSheet.Cells[row_num, 17] = r["U_GTTHUE"];
                    oSheet.Cells[row_num, 18] = r["U_GTVANCHUYEN"];
                    oSheet.Cells[row_num, 19] = r["U_GTVANHANH"];
                    STT++;
                    row_num++;
                }
                //GROUP TOTAL
                if (row_num - Group_row_num > 1)
                {
                    oSheet.Cells[Group_row_num, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 16].Formula = string.Format("=SUBTOTAL(9,P{0}:P{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 17].Formula = string.Format("=SUBTOTAL(9,Q{0}:Q{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 18].Formula = string.Format("=SUBTOTAL(9,R{0}:R{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 19].Formula = string.Format("=SUBTOTAL(9,S{0}:S{1})", Group_row_num + 1, row_num - 1);
                }
                oSheet.Range["A" + row_num, "T" + row_num].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + row_num, "T" + row_num].Font.Bold = true;
                oSheet.Cells[row_num, 1] = "III";
                oSheet.Cells[row_num, 2] = "THIẾT BỊ CHỐNG ĐỠ";
                STT = 1;
                Group_row_num = row_num;
                row_num++;

                foreach (DataRow r in A.Select("U_001 = 'TBCD'"))
                {
                    oSheet.Cells[row_num, 1] = STT;
                    oSheet.Cells[row_num, 2] = /*r["U_MATHIETBI"].ToString() + " - " +*/ r["U_TENTHIETBI"];
                    oSheet.Cells[row_num, 3] = r["U_DVTTB"];
                    oSheet.Cells[row_num, 4] = r["U_TLTB"];
                    oSheet.Cells[row_num, 5] = r["U_SLDUTRU"];
                    oSheet.Cells[row_num, 6] = r["U_SLTHUE"];
                    oSheet.Cells[row_num, 7] = r["U_SLVANCHUYEN"];
                    oSheet.Cells[row_num, 8] = r["U_SLVANHANH"];
                    oSheet.Cells[row_num, 9] = r["U_NGAYCAP"];
                    oSheet.Cells[row_num, 10] = r["U_NGAYTRA"];
                    oSheet.Cells[row_num, 11].Formula = string.Format("=J{0}-I{0}", row_num);
                    oSheet.Cells[row_num, 12] = r["U_DGMUABAN"];
                    oSheet.Cells[row_num, 13] = r["U_DGTHUE"];
                    oSheet.Cells[row_num, 14] = r["U_DGVCTB"];
                    oSheet.Cells[row_num, 15] = r["U_DGVH"];
                    oSheet.Cells[row_num, 16] = r["U_GTMB"];
                    oSheet.Cells[row_num, 17] = r["U_GTTHUE"];
                    oSheet.Cells[row_num, 18] = r["U_GTVANCHUYEN"];
                    oSheet.Cells[row_num, 19] = r["U_GTVANHANH"];
                    STT++;
                    row_num++;
                }
                //GROUP TOTAL 
                if (row_num - Group_row_num > 1)
                {
                    oSheet.Cells[Group_row_num, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 16].Formula = string.Format("=SUBTOTAL(9,P{0}:P{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 17].Formula = string.Format("=SUBTOTAL(9,Q{0}:Q{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 18].Formula = string.Format("=SUBTOTAL(9,R{0}:R{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 19].Formula = string.Format("=SUBTOTAL(9,S{0}:S{1})", Group_row_num + 1, row_num - 1);
                }

                oSheet.Range["A" + row_num, "T" + row_num].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + row_num, "T" + row_num].Font.Bold = true;
                oSheet.Cells[row_num, 1] = "IV";
                oSheet.Cells[row_num, 2] = "COPPHA NHÔM";
                STT = 1;
                Group_row_num = row_num;
                row_num++;

                foreach (DataRow r in A.Select("U_001 = 'CFN'"))
                {
                    oSheet.Cells[row_num, 1] = STT;
                    oSheet.Cells[row_num, 2] = /*r["U_MATHIETBI"].ToString() + " - " +*/ r["U_TENTHIETBI"];
                    oSheet.Cells[row_num, 3] = r["U_DVTTB"];
                    oSheet.Cells[row_num, 4] = r["U_TLTB"];
                    oSheet.Cells[row_num, 5] = r["U_SLDUTRU"];
                    oSheet.Cells[row_num, 6] = r["U_SLTHUE"];
                    oSheet.Cells[row_num, 7] = r["U_SLVANCHUYEN"];
                    oSheet.Cells[row_num, 8] = r["U_SLVANHANH"];
                    oSheet.Cells[row_num, 9] = r["U_NGAYCAP"];
                    oSheet.Cells[row_num, 10] = r["U_NGAYTRA"];
                    oSheet.Cells[row_num, 11].Formula = string.Format("=J{0}-I{0}", row_num);
                    oSheet.Cells[row_num, 12] = r["U_DGMUABAN"];
                    oSheet.Cells[row_num, 13] = r["U_DGTHUE"];
                    oSheet.Cells[row_num, 14] = r["U_DGVCTB"];
                    oSheet.Cells[row_num, 15] = r["U_DGVH"];
                    oSheet.Cells[row_num, 16] = r["U_GTMB"];
                    oSheet.Cells[row_num, 17] = r["U_GTTHUE"];
                    oSheet.Cells[row_num, 18] = r["U_GTVANCHUYEN"];
                    oSheet.Cells[row_num, 19] = r["U_GTVANHANH"];
                    STT++;
                    row_num++;
                }
                //GROUP TOTAL 
                if (row_num - Group_row_num > 1)
                {
                    oSheet.Cells[Group_row_num, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 16].Formula = string.Format("=SUBTOTAL(9,P{0}:P{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 17].Formula = string.Format("=SUBTOTAL(9,Q{0}:Q{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 18].Formula = string.Format("=SUBTOTAL(9,R{0}:R{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 19].Formula = string.Format("=SUBTOTAL(9,S{0}:S{1})", Group_row_num + 1, row_num - 1);
                }

                oSheet.Range["A" + row_num, "T" + row_num].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + row_num, "T" + row_num].Font.Bold = true;
                oSheet.Cells[row_num, 1] = "V";
                oSheet.Cells[row_num, 2] = "THIẾT BỊ AN TOÀN";
                STT = 1;
                Group_row_num = row_num;
                row_num++;

                foreach (DataRow r in A.Select("U_001 = 'TBAT'"))
                {
                    oSheet.Cells[row_num, 1] = STT;
                    oSheet.Cells[row_num, 2] = /*r["U_MATHIETBI"].ToString() + " - " +*/ r["U_TENTHIETBI"];
                    oSheet.Cells[row_num, 3] = r["U_DVTTB"];
                    oSheet.Cells[row_num, 4] = r["U_TLTB"];
                    oSheet.Cells[row_num, 5] = r["U_SLDUTRU"];
                    oSheet.Cells[row_num, 6] = r["U_SLTHUE"];
                    oSheet.Cells[row_num, 7] = r["U_SLVANCHUYEN"];
                    oSheet.Cells[row_num, 8] = r["U_SLVANHANH"];
                    oSheet.Cells[row_num, 9] = r["U_NGAYCAP"];
                    oSheet.Cells[row_num, 10] = r["U_NGAYTRA"];
                    oSheet.Cells[row_num, 11].Formula = string.Format("=J{0}-I{0}", row_num);
                    oSheet.Cells[row_num, 12] = r["U_DGMUABAN"];
                    oSheet.Cells[row_num, 13] = r["U_DGTHUE"];
                    oSheet.Cells[row_num, 14] = r["U_DGVCTB"];
                    oSheet.Cells[row_num, 15] = r["U_DGVH"];
                    oSheet.Cells[row_num, 16] = r["U_GTMB"];
                    oSheet.Cells[row_num, 17] = r["U_GTTHUE"];
                    oSheet.Cells[row_num, 18] = r["U_GTVANCHUYEN"];
                    oSheet.Cells[row_num, 19] = r["U_GTVANHANH"];
                    STT++;
                    row_num++;
                }
                //GROUP TOTAL 
                if (row_num - Group_row_num > 1)
                {
                    oSheet.Cells[Group_row_num, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 16].Formula = string.Format("=SUBTOTAL(9,P{0}:P{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 17].Formula = string.Format("=SUBTOTAL(9,Q{0}:Q{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 18].Formula = string.Format("=SUBTOTAL(9,R{0}:R{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 19].Formula = string.Format("=SUBTOTAL(9,S{0}:S{1})", Group_row_num + 1, row_num - 1);
                }

                oSheet.Range["A" + row_num, "T" + row_num].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + row_num, "T" + row_num].Font.Bold = true;
                oSheet.Cells[row_num, 1] = "VI";
                oSheet.Cells[row_num, 2] = "MÁY MÓC";
                STT = 1;
                Group_row_num = row_num;
                row_num++;

                foreach (DataRow r in A.Select("U_001 = 'MM'"))
                {
                    oSheet.Cells[row_num, 1] = STT;
                    oSheet.Cells[row_num, 2] = /*r["U_MATHIETBI"].ToString() + " - " +*/ r["U_TENTHIETBI"];
                    oSheet.Cells[row_num, 3] = r["U_DVTTB"];
                    oSheet.Cells[row_num, 4] = r["U_TLTB"];
                    oSheet.Cells[row_num, 5] = r["U_SLDUTRU"];
                    oSheet.Cells[row_num, 6] = r["U_SLTHUE"];
                    oSheet.Cells[row_num, 7] = r["U_SLVANCHUYEN"];
                    oSheet.Cells[row_num, 8] = r["U_SLVANHANH"];
                    oSheet.Cells[row_num, 9] = r["U_NGAYCAP"];
                    oSheet.Cells[row_num, 10] = r["U_NGAYTRA"];
                    oSheet.Cells[row_num, 11].Formula = string.Format("=J{0}-I{0}", row_num);
                    oSheet.Cells[row_num, 12] = r["U_DGMUABAN"];
                    oSheet.Cells[row_num, 13] = r["U_DGTHUE"];
                    oSheet.Cells[row_num, 14] = r["U_DGVCTB"];
                    oSheet.Cells[row_num, 15] = r["U_DGVH"];
                    oSheet.Cells[row_num, 16] = r["U_GTMB"];
                    oSheet.Cells[row_num, 17] = r["U_GTTHUE"];
                    oSheet.Cells[row_num, 18] = r["U_GTVANCHUYEN"];
                    oSheet.Cells[row_num, 19] = r["U_GTVANHANH"];
                    STT++;
                    row_num++;
                }
                //GROUP TOTAL 
                if (row_num - Group_row_num > 1)
                {
                    oSheet.Cells[Group_row_num, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 16].Formula = string.Format("=SUBTOTAL(9,P{0}:P{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 17].Formula = string.Format("=SUBTOTAL(9,Q{0}:Q{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 18].Formula = string.Format("=SUBTOTAL(9,R{0}:R{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 19].Formula = string.Format("=SUBTOTAL(9,S{0}:S{1})", Group_row_num + 1, row_num - 1);
                }

                oSheet.Range["A" + row_num, "T" + row_num].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + row_num, "T" + row_num].Font.Bold = true;
                oSheet.Cells[row_num, 1] = "VI";
                oSheet.Cells[row_num, 2] = "THIẾT BỊ KHÁC";
                STT = 1;
                Group_row_num = row_num;
                row_num++;

                foreach (DataRow r in A.Select("U_001 = 'TBK'"))
                {
                    oSheet.Cells[row_num, 1] = STT;
                    oSheet.Cells[row_num, 2] = /*r["U_MATHIETBI"].ToString() + " - " +*/ r["U_TENTHIETBI"];
                    oSheet.Cells[row_num, 3] = r["U_DVTTB"];
                    oSheet.Cells[row_num, 4] = r["U_TLTB"];
                    oSheet.Cells[row_num, 5] = r["U_SLDUTRU"];
                    oSheet.Cells[row_num, 6] = r["U_SLTHUE"];
                    oSheet.Cells[row_num, 7] = r["U_SLVANCHUYEN"];
                    oSheet.Cells[row_num, 8] = r["U_SLVANHANH"];
                    oSheet.Cells[row_num, 9] = r["U_NGAYCAP"];
                    oSheet.Cells[row_num, 10] = r["U_NGAYTRA"];
                    oSheet.Cells[row_num, 11].Formula = string.Format("=J{0}-I{0}", row_num);
                    oSheet.Cells[row_num, 12] = r["U_DGMUABAN"];
                    oSheet.Cells[row_num, 13] = r["U_DGTHUE"];
                    oSheet.Cells[row_num, 14] = r["U_DGVCTB"];
                    oSheet.Cells[row_num, 15] = r["U_DGVH"];
                    oSheet.Cells[row_num, 16] = r["U_GTMB"];
                    oSheet.Cells[row_num, 17] = r["U_GTTHUE"];
                    oSheet.Cells[row_num, 18] = r["U_GTVANCHUYEN"];
                    oSheet.Cells[row_num, 19] = r["U_GTVANHANH"];
                    STT++;
                    row_num++;
                }
                //GROUP TOTAL 
                if (row_num - Group_row_num > 1)
                {
                    oSheet.Cells[Group_row_num, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 16].Formula = string.Format("=SUBTOTAL(9,P{0}:P{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 17].Formula = string.Format("=SUBTOTAL(9,Q{0}:Q{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 18].Formula = string.Format("=SUBTOTAL(9,R{0}:R{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 19].Formula = string.Format("=SUBTOTAL(9,S{0}:S{1})", Group_row_num + 1, row_num - 1);
                }
                //TOTAL DUTRU
                oSheet.Cells[7, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", 8, row_num - 1);
                oSheet.Cells[7, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", 8, row_num - 1);
                oSheet.Cells[7, 7].Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", 8, row_num - 1);
                oSheet.Cells[7, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1})", 8, row_num - 1);
                oSheet.Cells[7, 16].Formula = string.Format("=SUBTOTAL(9,P{0}:P{1})", 8, row_num - 1);
                oSheet.Cells[7, 17].Formula = string.Format("=SUBTOTAL(9,Q{0}:Q{1})", 8, row_num - 1);
                oSheet.Cells[7, 18].Formula = string.Format("=SUBTOTAL(9,R{0}:R{1})", 8, row_num - 1);
                oSheet.Cells[7, 19].Formula = string.Format("=SUBTOTAL(9,S{0}:S{1})", 8, row_num - 1);

                //Border
                ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A8", "T" + (row_num - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
 
            }
        }
        void Load_Financial_Project()
        {
            //SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //oR_RecordSet.DoQuery("SELECT T0.[PrjCode], T0.[PrjName] FROM OPRJ T0 WHERE T0.[ValidFrom] >= '01-01-2018' and T0.[Active] = 'Y'");
            //if (oR_RecordSet.RecordCount > 0)
            //{
            //    while (!oR_RecordSet.EoF)
            //    {
            //        ComboBox0.ValidValues.Add(oR_RecordSet.Fields.Item("PrjCode").Value.ToString(), oR_RecordSet.Fields.Item("PrjName").Value.ToString());
            //        oR_RecordSet.MoveNext();
            //    }
            //}
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
                for (int i = 0; i < itm_count - 1; i++)
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

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            Load_SubProject(this.ComboBox0.Selected.Value);
            this.ComboBox1.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
        }
        
        System.Data.DataTable Get_Data_DUTRU(string pFinancialProject, int pCTG_DocEntry = -1, int pGoiThauKey = -1)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("EQ_CE_GET_DATA_DUTRU", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                cmd.Parameters.AddWithValue("@CTG_Entry", pCTG_DocEntry);
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
                cmd = new SqlCommand("EQ_CE_GET_FPROJECT", conn);
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
    }
}
