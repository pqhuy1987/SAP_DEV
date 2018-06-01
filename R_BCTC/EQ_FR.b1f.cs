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
    [FormAttribute("R_BCTC.EQ_FR", "EQ_FR.b1f")]
    class EQ_FR : UserFormBase
    {
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
        public EQ_FR()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_1").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_3").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_6").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_5").Specific));
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
                cmd = new SqlCommand("EQ_GET_FPROJECT", conn);
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
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.Button Button0;

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Load_SubProject(this.ComboBox0.Selected.Value);
            this.ComboBox1.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
        }
        System.Data.DataTable Get_Data_DUTRU(string pFinancialProject, string pType,DateTime pToDate, int pCTG_DocEntry = -1, int pGoiThauKey = -1)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("EQ_FR_GET_DATA", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                cmd.Parameters.AddWithValue("@CTG_Entry", pCTG_DocEntry);
                cmd.Parameters.AddWithValue("@GoiThauKey", pGoiThauKey);
                cmd.Parameters.AddWithValue("@Type", pType);
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
            //throw new System.NotImplementedException();
            int GoiThauKey = -1;
            DateTime tDate = DateTime.Today;
            if (ComboBox1.Selected.Value == "")
            {
                oApp.MessageBox("Please select Subproject !");
                return;
            }
            if (string.IsNullOrEmpty(EditText0.Value))
            {
                oApp.MessageBox("Please enter To date !");
                return;
            }
            else
            {
                tDate = DateTime.ParseExact(EditText0.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
            }
            int.TryParse(ComboBox1.Selected.Value, out GoiThauKey);
            //throw new System.NotImplementedException();
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;

            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            //Open Template
            oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_EQ_BCTC.xlsx");

            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            try
            {
                //Fill Header
                //Project Name
                oSheet.Cells[2, 3] = "Dự án: " + this.ComboBox0.Selected.Description;
                //Subproject Name
                oSheet.Cells[3, 3] = "Gói thầu: " + this.ComboBox1.Selected.Description;
                //Thang
                oSheet.Cells[4, 3] = "Tháng: " + DateTime.Today.ToString("MM/yyyy");

                //Get DATA
                //DataTable A = Get_Data_DUTRU(this.ComboBox0.Selected.Value, "S", -1, GoiThauKey);
                DataTable B = Get_Data_DUTRU(this.ComboBox0.Selected.Value, "D", tDate, - 1, GoiThauKey);
                int row_num = 8;
                int STT = 1;
                int Group_row_num = 0;
                foreach (DataRow r in B.Select("TYPE='NCC'"))
                {
                    oSheet.Cells[row_num, 1] = STT;
                    oSheet.Cells[row_num, 2] = r["U_BPName"].ToString();
                    oSheet.Cells[row_num, 3] = r["U_BPCode"].ToString();
                    oSheet.Cells[row_num, 4] = r["CP"];
                    oSheet.Cells[row_num, 5].Formula = string.Format("=F{0}/D{0}", row_num);
                    oSheet.Cells[row_num, 6] = r["GT"];
                    row_num++;
                    STT++;
                }
                //TONG NCC
                oSheet.Cells[7, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", 8, row_num - 1);
                oSheet.Cells[7, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", 8, row_num - 1);

                //NTP
                oSheet.Range["A" + row_num, "G" + row_num].Interior.Color = System.Drawing.Color.FromArgb(184, 208, 224);
                //oSheet.Range["B" + row_num, "G" + row_num].Merge(true);
                oSheet.Cells[row_num, 1] = "II";
                oSheet.Cells[row_num, 2] = "Nhà thầu phụ";

                Group_row_num = row_num;
                row_num++;
                STT = 1;
                //Details NTP
                foreach (DataRow r in B.Select("TYPE='NTP'"))
                {
                    oSheet.Cells[row_num, 1] = STT;
                    oSheet.Cells[row_num, 2] = r["U_BPName"].ToString();
                    oSheet.Cells[row_num, 3] = r["U_BPCode"].ToString();
                    oSheet.Cells[row_num, 4] = r["CP"];
                    oSheet.Cells[row_num, 5].Formula = string.Format("=F{0}/D{0}", row_num);
                    oSheet.Cells[row_num, 6] = r["GT"];
                    row_num++;
                    STT++;
                }

                //TONG NTP
                if (row_num - Group_row_num > 1)
                {
                    oSheet.Cells[Group_row_num, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", Group_row_num + 1, row_num - 1);
                    oSheet.Cells[Group_row_num, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", Group_row_num + 1, row_num - 1);
                }

                //TOTAL
                oSheet.Cells[6, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", 7, row_num - 1);
                oSheet.Cells[6, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", 7, row_num - 1);
                //Border
                ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A8", "G" + (row_num - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {

            }
        }

        private SAPbouiCOM.EditText EditText0;
    }
}
