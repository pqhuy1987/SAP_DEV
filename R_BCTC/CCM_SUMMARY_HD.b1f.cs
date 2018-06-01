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
    [FormAttribute("R_BCTC.CCM_SUMMARY_HD", "CCM_SUMMARY_HD.b1f")]
    class CCM_SUMMARY_HD : UserFormBase
    {
        public CCM_SUMMARY_HD()
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
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_fpro").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_fdate").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txt_tdate").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            
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
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText1;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            if (ComboBox0.Selected.Value == "" )
            {
                oApp.MessageBox("Please select Project !");
                return;
            }
            if (string.IsNullOrEmpty(EditText0.Value) || string.IsNullOrEmpty(EditText1.Value))
            {
                oApp.MessageBox("Please enter Date !");
                return;
            }
            string FProject = ComboBox0.Selected.Value;
            DateTime FrDate = DateTime.ParseExact(EditText0.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
            DateTime ToDate = DateTime.ParseExact(EditText1.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
            DataTable rs = Get_Data(ComboBox0.Selected.Value, FrDate,ToDate);
            if (rs.Rows.Count > 0)
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
                oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_DUYET_HD.xlsx");
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Project Name
                oSheet.Cells[3, 3] = "Dự án: " + this.ComboBox0.Selected.Description;
                //From Date - To Date
                oSheet.Cells[4, 3] = string.Format("Từ ngày: {0} đến ngày: {1}", FrDate.ToString("dd/MM/yyyy"), ToDate.ToString("dd/MM/yyyy"));

                int STT = 1;
                int current_row = 6;
                foreach (DataRow r in rs.Rows)
                {
                    //STT
                    oSheet.Cells[current_row, 1] = STT;
                    //Project Name
                    oSheet.Cells[current_row, 2] = r["Project"].ToString();
                    //BP Name
                    oSheet.Cells[current_row, 3] = r["BpName"].ToString();
                    //Pham vi cong viec
                    oSheet.Cells[current_row, 4] = r["Descript"].ToString();
                    //GTHD
                    oSheet.Cells[current_row, 5] = r["GTHD"];
                    //CHT
                    oSheet.Cells[current_row, 6] = r["CHT"];
                    //Phap che
                    oSheet.Cells[current_row, 7] = r["PC"];
                    //Thiet bi
                    oSheet.Cells[current_row, 8] = r["TB"];
                    //Co dien
                    oSheet.Cells[current_row, 9] = r["ME"];
                    //Ke toan
                    oSheet.Cells[current_row, 10] = r["KT"];
                    //CCM
                    oSheet.Cells[current_row, 11] = r["CCM"];
                    //PGD
                    oSheet.Cells[current_row, 12] = r["PGD"];
                    //Total
                    oSheet.Cells[current_row, 13].Formula = string.Format("=IF(DAYS(L{0},F{0})<0,0,DAYS(L{0},F{0}))", current_row);

                    STT++;
                    current_row++;
                }

                //Border
                ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A6", "N" + (current_row - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                oXL.ActiveWindow.Activate();
            }
        }

        System.Data.DataTable Get_Data(string pFproject, DateTime pFrDate, DateTime pToDate)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_SUMMARY_HD_GET_DATA", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FProject", pFproject);
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
    }
}
