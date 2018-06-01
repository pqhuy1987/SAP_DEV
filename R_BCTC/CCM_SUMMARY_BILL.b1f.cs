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
    [FormAttribute("R_BCTC.CCM_SUMMARY_BILL", "CCM_SUMMARY_BILL.b1f")]
    class CCM_SUMMARY_BILL : UserFormBase
    {
        public CCM_SUMMARY_BILL()
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
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_fp").Specific));
            //this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_per").Specific));
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

        System.Data.DataTable Get_List_Period(string pFproject)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_SUMMARY_BILL_GET_PERIOD", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FProject", pFproject);
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

        System.Data.DataTable Get_Data(string pFproject, int pPeriod)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_SUMMARY_BILL_GET_DATA", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FProject", pFproject);
                cmd.Parameters.AddWithValue("@Period", pPeriod);
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

        System.Data.DataTable Get_Data_BCH(string pFproject, int pPeriod)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_SUMMARY_BILL_GET_DATA_BCH", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FProject", pFproject);
                cmd.Parameters.AddWithValue("@Period", pPeriod);
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

        //private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
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
        //    DataTable rs = Get_List_Period(ComboBox0.Selected.Value);
        //    try
        //    {
        //        this.ComboBox1.ValidValues.Add("", "");
        //    }
        //    catch
        //    { }
        //    if (rs.Rows.Count > 0)
        //    {
        //        foreach (DataRow r in rs.Rows)
        //        {
        //            ComboBox1.ValidValues.Add(r["U_Period"].ToString(), "Kỳ " + r["U_Period"].ToString());
        //        }
        //    }

        //}

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (ComboBox0.Selected.Value == "" || EditText0.Value == "")
            {
                oApp.MessageBox("Please select Project and Period !");
                return;
            }
            int period = 0;
            int.TryParse(EditText0.Value, out period);
            DataTable rs = Get_Data(ComboBox0.Selected.Value, period);
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
                oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_TONG_HOP_THANH_TOAN_KY.xlsx");
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Project Name
                oSheet.Cells[1, 3] = "Dự án: " + this.ComboBox0.Selected.Description;
                //Period
                oSheet.Cells[3, 3] = "Kỳ: " + this.EditText0.Value;
                int current_row = 7, group_row = 7, STT = 1;
                //Doi thi cong
                oSheet.Cells[current_row, 1] = "I";
                oSheet.Cells[current_row, 2] = "CHI PHÍ ĐỘI THI CÔNG";
                oSheet.Range["A" + current_row, "N" + current_row].Interior.Color = System.Drawing.Color.FromArgb(191, 191, 191);
                oSheet.Range["A" + current_row, "N" + current_row].Font.Bold = true;
                current_row++;

                foreach (DataRow r in rs.Select("U_PUTYPE='PUT09'"))
                {
                    //STT
                    oSheet.Cells[current_row, 1] = STT;
                    //BP Name
                    oSheet.Cells[current_row, 2] = r["U_BPName"].ToString();
                    //BP Code
                    oSheet.Cells[current_row, 3] = r["U_BPCode"].ToString();
                    //GT Lũy kế đến kỳ này
                    oSheet.Cells[current_row, 4].Formula = string.Format("=E{0}+F{0}",current_row);
                    #region BILL CACULATOR
                    decimal Kynay = 0, Kytruoc = 0;
                    Bill_Caculator(ComboBox0.Selected.Value, int.Parse(r["U_Period"].ToString()), r["U_BPCode"].ToString(), r["U_BGroup"].ToString(), r["U_PUType_Origin"].ToString(), int.Parse(r["U_BType"].ToString()), int.Parse(r["DocEntry"].ToString()), DateTime.Parse(r["U_DATETO"].ToString()), out Kytruoc, out Kynay);
                    if (int.Parse(r["U_Period"].ToString()) == period)
                    {
                        //GT thanh toan ky truoc
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc;
                        //GT thanh toan ky nay
                        oSheet.Cells[current_row, 6].Value2 = Kynay;
                    }
                    else
                    {
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc + Kynay;
                    }
                    #endregion
                    //Last Approved by
                    oSheet.Cells[current_row, 7] = r["Last Approved by"].ToString();
                    //Last Approved on
                    if (!string.IsNullOrEmpty(r["Last Approved on"].ToString()))
                        oSheet.Cells[current_row, 8] = DateTime.ParseExact(r["Last Approved on"].ToString(), "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture);
                    //CCM duyet
                    if (!string.IsNullOrEmpty(r["CCM Approve"].ToString()))
                    {
                        if (r["CCM Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 10] = "x";
                        else if (r["CCM Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 9] = "x";
                    }
                    //Ke toan duyet
                    if (!string.IsNullOrEmpty(r["KT Approve"].ToString()))
                    {
                        if (r["KT Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 12] = "x";
                        else if (r["KT Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 11] = "x";
                    }
                    //Ngay quyet toan
                    if (!string.IsNullOrEmpty(r["U_BType"].ToString()))
                    {
                        if (r["U_BType"].ToString() == "3")
                        {
                            DateTime dt_qt = DateTime.Today;
                            DateTime.TryParse(r["U_DATETO"].ToString(), out dt_qt);
                            oSheet.Cells[current_row, 13] = dt_qt;
                        }
                    }
                    //Ghi chu
                    if (r["Canceled"].ToString() == "Y")
                    {
                        oSheet.Cells[current_row, 14] = "Rejected";
                    }
                    current_row++;
                    STT++;
                }

                //Total I
                if (current_row - group_row > 1)
                {
                    oSheet.Cells[group_row, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", group_row + 1, current_row - 1);
                    oSheet.Cells[group_row, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", group_row + 1, current_row - 1);
                    oSheet.Cells[group_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", group_row + 1, current_row - 1);
                }

                #region XD
                //XD
                oSheet.Cells[current_row, 1] = "II";
                oSheet.Cells[current_row, 2] = "NCC, NTP XÂY DỰNG";
                oSheet.Range["A" + current_row, "N" + current_row].Interior.Color = System.Drawing.Color.FromArgb(191, 191, 191);
                oSheet.Range["A" + current_row, "N" + current_row].Font.Bold = true;
                group_row = current_row;
                current_row++;
                //NCC XD
                oSheet.Cells[current_row, 1] = "II.1";
                oSheet.Cells[current_row, 2] = "NHÀ CUNG CẤP XÂY DỰNG";
                oSheet.Range["A" + current_row, "N" + current_row].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + current_row, "N" + current_row].Font.Bold = true;
                int group_row2 = current_row;
                current_row++;
                STT = 1;
                foreach (DataRow r in rs.Select("U_PUTYPE='PUT01' and U_BGroup='XD'"))
                {
                    //STT
                    oSheet.Cells[current_row, 1] = STT;
                    //BP Name
                    oSheet.Cells[current_row, 2] = r["U_BPName"].ToString();
                    //BP Code
                    oSheet.Cells[current_row, 3] = r["U_BPCode"].ToString();
                    //GT Lũy kế đến kỳ này
                    oSheet.Cells[current_row, 4].Formula = string.Format("=E{0}+F{0}", current_row);
                    #region BILL CACULATOR
                    decimal Kynay = 0, Kytruoc = 0;
                    Bill_Caculator(ComboBox0.Selected.Value, int.Parse(r["U_Period"].ToString()), r["U_BPCode"].ToString(), r["U_BGroup"].ToString(), r["U_PUType_Origin"].ToString(), int.Parse(r["U_BType"].ToString()), int.Parse(r["DocEntry"].ToString()), DateTime.Parse(r["U_DATETO"].ToString()), out Kytruoc, out Kynay);
                    if (int.Parse(r["U_Period"].ToString()) == period)
                    {
                        //GT thanh toan ky truoc
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc;
                        //GT thanh toan ky nay
                        oSheet.Cells[current_row, 6].Value2 = Kynay;
                    }
                    else
                    {
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc + Kynay;
                    }
                    #endregion
                    //Last Approved by
                    oSheet.Cells[current_row, 7] = r["Last Approved by"].ToString();
                    //Last Approved on
                    if (!string.IsNullOrEmpty(r["Last Approved on"].ToString()))
                        oSheet.Cells[current_row, 8] = DateTime.ParseExact(r["Last Approved on"].ToString(), "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture);
                    //CCM duyet
                    if (!string.IsNullOrEmpty(r["CCM Approve"].ToString()))
                    {
                        if (r["CCM Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 10] = "x";
                        else if (r["CCM Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 9] = "x";
                    }
                    //Ke toan duyet
                    if (!string.IsNullOrEmpty(r["KT Approve"].ToString()))
                    {
                        if (r["KT Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 12] = "x";
                        else if (r["KT Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 11] = "x";
                    }
                    //Ngay quyet toan
                    if (!string.IsNullOrEmpty(r["U_BType"].ToString()))
                    {
                        if (r["U_BType"].ToString() == "3")
                        {
                            DateTime dt_qt = DateTime.Today;
                            DateTime.TryParse(r["U_DATETO"].ToString(), out dt_qt);
                            oSheet.Cells[current_row, 13] = dt_qt;
                        }
                    }
                    //Ghi chu
                    if (r["Canceled"].ToString() == "Y")
                    {
                        oSheet.Cells[current_row, 14] = "Rejected";
                    }
                    current_row++;
                    STT++;
                }
                //Total II.1
                if (current_row - group_row2 > 1)
                {
                    oSheet.Cells[group_row2, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", group_row2 + 1, current_row - 1);
                    oSheet.Cells[group_row2, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", group_row2 + 1, current_row - 1);
                    oSheet.Cells[group_row2, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", group_row2 + 1, current_row - 1);
                }

                //NTP XD
                oSheet.Cells[current_row, 1] = "II.2";
                oSheet.Cells[current_row, 2] = "NHÀ THẦU PHỤ XÂY DỰNG";
                oSheet.Range["A" + current_row, "N" + current_row].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + current_row, "N" + current_row].Font.Bold = true;
                group_row2 = current_row;
                current_row++;
                STT = 1;
                foreach (DataRow r in rs.Select("U_PUTYPE='PUT02' and U_BGroup='XD'"))
                {
                    //STT
                    oSheet.Cells[current_row, 1] = STT;
                    //BP Name
                    oSheet.Cells[current_row, 2] = r["U_BPName"].ToString();
                    //BP Code
                    oSheet.Cells[current_row, 3] = r["U_BPCode"].ToString();
                    //GT Lũy kế đến kỳ này
                    oSheet.Cells[current_row, 4].Formula = string.Format("=E{0}+F{0}", current_row);
                    #region BILL CACULATOR
                    decimal Kynay = 0, Kytruoc = 0;
                    Bill_Caculator(ComboBox0.Selected.Value, int.Parse(r["U_Period"].ToString()), r["U_BPCode"].ToString(), r["U_BGroup"].ToString(), r["U_PUType_Origin"].ToString(), int.Parse(r["U_BType"].ToString()), int.Parse(r["DocEntry"].ToString()), DateTime.Parse(r["U_DATETO"].ToString()), out Kytruoc, out Kynay);
                    if (int.Parse(r["U_Period"].ToString()) == period)
                    {
                        //GT thanh toan ky truoc
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc;
                        //GT thanh toan ky nay
                        oSheet.Cells[current_row, 6].Value2 = Kynay;
                    }
                    else
                    {
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc + Kynay;
                    }
                    #endregion
                    //Last Approved by
                    oSheet.Cells[current_row, 7] = r["Last Approved by"].ToString();
                    //Last Approved on
                    if (!string.IsNullOrEmpty(r["Last Approved on"].ToString()))
                        oSheet.Cells[current_row, 8] = DateTime.ParseExact(r["Last Approved on"].ToString(), "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture);
                    //CCM duyet
                    if (!string.IsNullOrEmpty(r["CCM Approve"].ToString()))
                    {
                        if (r["CCM Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 10] = "x";
                        else if (r["CCM Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 9] = "x";
                    }
                    //Ke toan duyet
                    if (!string.IsNullOrEmpty(r["KT Approve"].ToString()))
                    {
                        if (r["KT Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 12] = "x";
                        else if (r["KT Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 11] = "x";
                    }
                    //Ngay quyet toan
                    if (!string.IsNullOrEmpty(r["U_BType"].ToString()))
                    {
                        if (r["U_BType"].ToString() == "3")
                        {
                            DateTime dt_qt = DateTime.Today;
                            DateTime.TryParse(r["U_DATETO"].ToString(), out dt_qt);
                            oSheet.Cells[current_row, 13] = dt_qt;
                        }
                    }
                    //Ghi chu
                    if (r["Canceled"].ToString() == "Y")
                    {
                        oSheet.Cells[current_row, 14] = "Rejected";
                    }
                    current_row++;
                    STT++;
                }
                //Total II.2
                if (current_row - group_row2 > 1)
                {
                    oSheet.Cells[group_row2, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", group_row2 + 1, current_row - 1);
                    oSheet.Cells[group_row2, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", group_row2 + 1, current_row - 1);
                    oSheet.Cells[group_row2, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", group_row2 + 1, current_row - 1);
                }

                //Total II
                if (current_row - group_row > 1)
                {
                    oSheet.Cells[group_row, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", group_row + 1, current_row - 1);
                    oSheet.Cells[group_row, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", group_row + 1, current_row - 1);
                    oSheet.Cells[group_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", group_row + 1, current_row - 1);
                }
                #endregion

                #region ME
                //ME
                oSheet.Cells[current_row, 1] = "III";
                oSheet.Cells[current_row, 2] = "NHÀ CUNG CẤP, NHÀ THẦU PHỤ M&E";
                oSheet.Range["A" + current_row, "N" + current_row].Interior.Color = System.Drawing.Color.FromArgb(191, 191, 191);
                oSheet.Range["A" + current_row, "N" + current_row].Font.Bold = true;
                group_row = current_row;
                current_row++;
                //NCC ME
                oSheet.Cells[current_row, 1] = "III.1";
                oSheet.Cells[current_row, 2] = "NHÀ CUNG CẤP M&E";
                oSheet.Range["A" + current_row, "N" + current_row].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + current_row, "N" + current_row].Font.Bold = true;
                group_row2 = current_row;
                current_row++;
                STT = 1;
                foreach (DataRow r in rs.Select("U_PUTYPE='PUT01' and U_BGroup='CD'"))
                {
                    //STT
                    oSheet.Cells[current_row, 1] = STT;
                    //BP Name
                    oSheet.Cells[current_row, 2] = r["U_BPName"].ToString();
                    //BP Code
                    oSheet.Cells[current_row, 3] = r["U_BPCode"].ToString();
                    //GT Lũy kế đến kỳ này
                    oSheet.Cells[current_row, 4].Formula = string.Format("=E{0}+F{0}", current_row);
                    #region BILL CACULATOR
                    decimal Kynay = 0, Kytruoc = 0;
                    Bill_Caculator(ComboBox0.Selected.Value, int.Parse(r["U_Period"].ToString()), r["U_BPCode"].ToString(), r["U_BGroup"].ToString(), r["U_PUType_Origin"].ToString(), int.Parse(r["U_BType"].ToString()), int.Parse(r["DocEntry"].ToString()), DateTime.Parse(r["U_DATETO"].ToString()), out Kytruoc, out Kynay);
                    if (int.Parse(r["U_Period"].ToString()) == period)
                    {
                        //GT thanh toan ky truoc
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc;
                        //GT thanh toan ky nay
                        oSheet.Cells[current_row, 6].Value2 = Kynay;
                    }
                    else
                    {
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc + Kynay;
                    }
                    #endregion
                    //Last Approved by
                    oSheet.Cells[current_row, 7] = r["Last Approved by"].ToString();
                    //Last Approved on
                    if (!string.IsNullOrEmpty(r["Last Approved on"].ToString()))
                        oSheet.Cells[current_row, 8] = DateTime.ParseExact(r["Last Approved on"].ToString(), "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture);
                    //CCM duyet
                    if (!string.IsNullOrEmpty(r["CCM Approve"].ToString()))
                    {
                        if (r["CCM Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 10] = "x";
                        else if (r["CCM Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 9] = "x";
                    }
                    //Ke toan duyet
                    if (!string.IsNullOrEmpty(r["KT Approve"].ToString()))
                    {
                        if (r["KT Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 12] = "x";
                        else if (r["KT Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 11] = "x";
                    }
                    //Ngay quyet toan
                    if (!string.IsNullOrEmpty(r["U_BType"].ToString()))
                    {
                        if (r["U_BType"].ToString() == "3")
                        {
                            DateTime dt_qt = DateTime.Today;
                            DateTime.TryParse(r["U_DATETO"].ToString(), out dt_qt);
                            oSheet.Cells[current_row, 13] = dt_qt;
                        }
                    }
                    //Ghi chu
                    if (r["Canceled"].ToString() == "Y")
                    {
                        oSheet.Cells[current_row, 14] = "Rejected";
                    }
                    current_row++;
                    STT++;
                }
                //Total III.1
                if (current_row - group_row2 > 1)
                {
                    oSheet.Cells[group_row2, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", group_row2 + 1, current_row - 1);
                    oSheet.Cells[group_row2, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", group_row2 + 1, current_row - 1);
                    oSheet.Cells[group_row2, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", group_row2 + 1, current_row - 1);
                }

                //NTP ME
                oSheet.Cells[current_row, 1] = "III.2";
                oSheet.Cells[current_row, 2] = "NHÀ THẦU PHỤ M&E";
                oSheet.Range["A" + current_row, "N" + current_row].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + current_row, "N" + current_row].Font.Bold = true;
                group_row2 = current_row;
                current_row++;
                STT = 1;
                foreach (DataRow r in rs.Select("U_PUTYPE='PUT02' and U_BGroup='CD'"))
                {
                    //STT
                    oSheet.Cells[current_row, 1] = STT;
                    //BP Name
                    oSheet.Cells[current_row, 2] = r["U_BPName"].ToString();
                    //BP Code
                    oSheet.Cells[current_row, 3] = r["U_BPCode"].ToString();
                    //GT Lũy kế đến kỳ này
                    oSheet.Cells[current_row, 4].Formula = string.Format("=E{0}+F{0}", current_row);
                    #region BILL CACULATOR
                    decimal Kynay = 0, Kytruoc = 0;
                    Bill_Caculator(ComboBox0.Selected.Value, int.Parse(r["U_Period"].ToString()), r["U_BPCode"].ToString(), r["U_BGroup"].ToString(), r["U_PUType_Origin"].ToString(), int.Parse(r["U_BType"].ToString()), int.Parse(r["DocEntry"].ToString()), DateTime.Parse(r["U_DATETO"].ToString()), out Kytruoc, out Kynay);
                    if (int.Parse(r["U_Period"].ToString()) == period)
                    {
                        //GT thanh toan ky truoc
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc;
                        //GT thanh toan ky nay
                        oSheet.Cells[current_row, 6].Value2 = Kynay;
                    }
                    else
                    {
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc + Kynay;
                    }
                    #endregion
                    //Last Approved by
                    oSheet.Cells[current_row, 7] = r["Last Approved by"].ToString();
                    //Last Approved on
                    if (!string.IsNullOrEmpty(r["Last Approved on"].ToString()))
                        oSheet.Cells[current_row, 8] = DateTime.ParseExact(r["Last Approved on"].ToString(), "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture);
                    //CCM duyet
                    if (!string.IsNullOrEmpty(r["CCM Approve"].ToString()))
                    {
                        if (r["CCM Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 10] = "x";
                        else if (r["CCM Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 9] = "x";
                    }
                    //Ke toan duyet
                    if (!string.IsNullOrEmpty(r["KT Approve"].ToString()))
                    {
                        if (r["KT Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 12] = "x";
                        else if (r["KT Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 11] = "x";
                    }
                    //Ngay quyet toan
                    if (!string.IsNullOrEmpty(r["U_BType"].ToString()))
                    {
                        if (r["U_BType"].ToString() == "3")
                        {
                            DateTime dt_qt = DateTime.Today;
                            DateTime.TryParse(r["U_DATETO"].ToString(), out dt_qt);
                            oSheet.Cells[current_row, 13] = dt_qt;
                        }
                    }
                    //Ghi chu
                    if (r["Canceled"].ToString() == "Y")
                    {
                        oSheet.Cells[current_row, 14] = "Rejected";
                    }
                    current_row++;
                    STT++;
                }
                //Total III.2
                if (current_row - group_row2 > 1)
                {
                    oSheet.Cells[group_row2, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", group_row2 + 1, current_row - 1);
                    oSheet.Cells[group_row2, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", group_row2 + 1, current_row - 1);
                    oSheet.Cells[group_row2, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", group_row2 + 1, current_row - 1);
                }

                //Total III
                if (current_row - group_row > 1)
                {
                    oSheet.Cells[group_row, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", group_row + 1, current_row - 1);
                    oSheet.Cells[group_row, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", group_row + 1, current_row - 1);
                    oSheet.Cells[group_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", group_row + 1, current_row - 1);
                }
                #endregion

                #region TB
                //TB
                oSheet.Cells[current_row, 1] = "IV";
                oSheet.Cells[current_row, 2] = "NHÀ CUNG CẤP, NHÀ THẦU PHỤ THIẾT BỊ";
                oSheet.Range["A" + current_row, "N" + current_row].Interior.Color = System.Drawing.Color.FromArgb(191, 191, 191);
                oSheet.Range["A" + current_row, "N" + current_row].Font.Bold = true;
                group_row = current_row;
                current_row++;
                //NCC TB
                oSheet.Cells[current_row, 1] = "IV.1";
                oSheet.Cells[current_row, 2] = "NHÀ CUNG CẤP THIẾT BỊ";
                oSheet.Range["A" + current_row, "N" + current_row].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + current_row, "N" + current_row].Font.Bold = true;
                group_row2 = current_row;
                current_row++;
                STT = 1;
                foreach (DataRow r in rs.Select("U_PUTYPE='PUT01' and U_BGroup='TBXD'"))
                {
                    //STT
                    oSheet.Cells[current_row, 1] = STT;
                    //BP Name
                    oSheet.Cells[current_row, 2] = r["U_BPName"].ToString();
                    //BP Code
                    oSheet.Cells[current_row, 3] = r["U_BPCode"].ToString();
                    //GT Lũy kế đến kỳ này
                    oSheet.Cells[current_row, 4].Formula = string.Format("=E{0}+F{0}", current_row);
                    #region BILL CACULATOR
                    decimal Kynay = 0, Kytruoc = 0;
                    Bill_Caculator(ComboBox0.Selected.Value, int.Parse(r["U_Period"].ToString()), r["U_BPCode"].ToString(), r["U_BGroup"].ToString(), r["U_PUType_Origin"].ToString(), int.Parse(r["U_BType"].ToString()), int.Parse(r["DocEntry"].ToString()), DateTime.Parse(r["U_DATETO"].ToString()), out Kytruoc, out Kynay);
                    if (int.Parse(r["U_Period"].ToString()) == period)
                    {
                        //GT thanh toan ky truoc
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc;
                        //GT thanh toan ky nay
                        oSheet.Cells[current_row, 6].Value2 = Kynay;
                    }
                    else
                    {
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc + Kynay;
                    }
                    #endregion
                    //Last Approved by
                    oSheet.Cells[current_row, 7] = r["Last Approved by"].ToString();
                    //Last Approved on
                    if (!string.IsNullOrEmpty(r["Last Approved on"].ToString()))
                        oSheet.Cells[current_row, 8] = DateTime.ParseExact(r["Last Approved on"].ToString(), "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture);
                    //CCM duyet
                    if (!string.IsNullOrEmpty(r["CCM Approve"].ToString()))
                    {
                        if (r["CCM Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 10] = "x";
                        else if (r["CCM Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 9] = "x";
                    }
                    //Ke toan duyet
                    if (!string.IsNullOrEmpty(r["KT Approve"].ToString()))
                    {
                        if (r["KT Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 12] = "x";
                        else if (r["KT Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 11] = "x";
                    }
                    //Ngay quyet toan
                    if (!string.IsNullOrEmpty(r["U_BType"].ToString()))
                    {
                        if (r["U_BType"].ToString() == "3")
                        {
                            DateTime dt_qt = DateTime.Today;
                            DateTime.TryParse(r["U_DATETO"].ToString(), out dt_qt);
                            oSheet.Cells[current_row, 13] = dt_qt;
                        }
                    }
                    //Ghi chu
                    if (r["Canceled"].ToString() == "Y")
                    {
                        oSheet.Cells[current_row, 14] = "Rejected";
                    }
                    current_row++;
                    STT++;
                }
                //Total IV.1
                if (current_row - group_row2 > 1)
                {
                    oSheet.Cells[group_row2, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", group_row2 + 1, current_row - 1);
                    oSheet.Cells[group_row2, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", group_row2 + 1, current_row - 1);
                    oSheet.Cells[group_row2, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", group_row2 + 1, current_row - 1);
                }

                //NTP TB
                oSheet.Cells[current_row, 1] = "IV.2";
                oSheet.Cells[current_row, 2] = "NHÀ THẦU PHỤ THIẾT BỊ";
                oSheet.Range["A" + current_row, "N" + current_row].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + current_row, "N" + current_row].Font.Bold = true;
                group_row2 = current_row;
                current_row++;
                STT = 1;
                foreach (DataRow r in rs.Select("U_PUTYPE='PUT02' and U_BGroup='TBXD'"))
                {
                    //STT
                    oSheet.Cells[current_row, 1] = STT;
                    //BP Name
                    oSheet.Cells[current_row, 2] = r["U_BPName"].ToString();
                    //BP Code
                    oSheet.Cells[current_row, 3] = r["U_BPCode"].ToString();
                    //GT Lũy kế đến kỳ này
                    oSheet.Cells[current_row, 4].Formula = string.Format("=E{0}+F{0}", current_row);
                    #region BILL CACULATOR
                    decimal Kynay = 0, Kytruoc = 0;
                    Bill_Caculator(ComboBox0.Selected.Value, int.Parse(r["U_Period"].ToString()), r["U_BPCode"].ToString(), r["U_BGroup"].ToString(), r["U_PUType_Origin"].ToString(), int.Parse(r["U_BType"].ToString()), int.Parse(r["DocEntry"].ToString()), DateTime.Parse(r["U_DATETO"].ToString()), out Kytruoc, out Kynay);
                    if (int.Parse(r["U_Period"].ToString()) == period)
                    {
                        //GT thanh toan ky truoc
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc;
                        //GT thanh toan ky nay
                        oSheet.Cells[current_row, 6].Value2 = Kynay;
                    }
                    else
                    {
                        oSheet.Cells[current_row, 5].Value2 = Kytruoc + Kynay;
                    }
                    #endregion
                    //Last Approved by
                    oSheet.Cells[current_row, 7] = r["Last Approved by"].ToString();
                    //Last Approved on
                    if (!string.IsNullOrEmpty(r["Last Approved on"].ToString()))
                        oSheet.Cells[current_row, 8] = DateTime.ParseExact(r["Last Approved on"].ToString(), "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture);
                    //CCM duyet
                    if (!string.IsNullOrEmpty(r["CCM Approve"].ToString()))
                    {
                        if (r["CCM Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 10] = "x";
                        else if (r["CCM Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 9] = "x";
                    }
                    //Ke toan duyet
                    if (!string.IsNullOrEmpty(r["KT Approve"].ToString()))
                    {
                        if (r["KT Approve"].ToString() == "1")
                            oSheet.Cells[current_row, 12] = "x";
                        else if (r["KT Approve"].ToString() == "2")
                            oSheet.Cells[current_row, 11] = "x";
                    }
                    //Ngay quyet toan
                    if (!string.IsNullOrEmpty(r["U_BType"].ToString()))
                    {
                        if (r["U_BType"].ToString() == "3")
                        {
                            DateTime dt_qt = DateTime.Today;
                            DateTime.TryParse(r["U_DATETO"].ToString(), out dt_qt);
                            oSheet.Cells[current_row, 13] = dt_qt;
                        }
                    }
                    //Ghi chu
                    if (r["Canceled"].ToString() == "Y")
                    {
                        oSheet.Cells[current_row, 14] = "Rejected";
                    }
                    current_row++;
                    STT++;
                }
                //Total IV.2
                if (current_row - group_row2 > 1)
                {
                    oSheet.Cells[group_row2, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", group_row2 + 1, current_row - 1);
                    oSheet.Cells[group_row2, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", group_row2 + 1, current_row - 1);
                    oSheet.Cells[group_row2, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", group_row2 + 1, current_row - 1);
                }

                //Total IV
                if (current_row - group_row > 1)
                {
                    oSheet.Cells[group_row, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", group_row + 1, current_row - 1);
                    oSheet.Cells[group_row, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", group_row + 1, current_row - 1);
                    oSheet.Cells[group_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", group_row + 1, current_row - 1);
                }
                #endregion

                #region BCH
                //BCH
                oSheet.Cells[current_row, 1] = "V";
                oSheet.Cells[current_row, 2] = "CHI PHÍ BAN CHỈ HUY";
                oSheet.Range["A" + current_row, "N" + current_row].Interior.Color = System.Drawing.Color.FromArgb(191, 191, 191);
                oSheet.Range["A" + current_row, "N" + current_row].Font.Bold = true;
                group_row = current_row;
                current_row++;

                DataTable rs_bch = Get_Data_BCH(ComboBox0.Selected.Value, period);
                if (rs_bch.Rows.Count > 0)
                {
                    STT = 1;
                    foreach (DataRow r in rs_bch.Rows)
                    {
                        //STT
                        oSheet.Cells[current_row, 1] = STT;
                        //Tên CP
                        oSheet.Cells[current_row, 2] = r["TEN_CP"].ToString();
                        //Mã CP
                        oSheet.Cells[current_row, 3] = r["MA_CP"].ToString();
                        //GT Lũy kế đến kỳ này
                        oSheet.Cells[current_row, 4].Formula = string.Format("=E{0}+F{0}", current_row);
                        if (int.Parse(EditText0.Value) == period)
                        {
                            //GT thanh toan ky truoc
                            oSheet.Cells[current_row, 5].Value2 = r["CP_KY_TRUOC"].ToString();
                            //GT thanh toan ky nay
                            oSheet.Cells[current_row, 6].Value2 = r["CP_KY_NAY"].ToString();
                        }
                        else
                        {
                            decimal tmp1 = 0, tmp2 = 0;
                            decimal.TryParse(r["CP_KY_TRUOC"].ToString(), out tmp1);
                            decimal.TryParse(r["CP_KY_NAY"].ToString(), out tmp2);
                            oSheet.Cells[current_row, 5].Value2 = tmp1 + tmp2;
                        }
                        //Last Approved by
                        oSheet.Cells[current_row, 7] = r["Last Approved by"].ToString();
                        //Last Approved on
                        if (!string.IsNullOrEmpty(r["Last Approved on"].ToString()))
                            oSheet.Cells[current_row, 8] = DateTime.ParseExact(r["Last Approved on"].ToString(), "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture);
                        //CCM duyet
                        if (!string.IsNullOrEmpty(r["CCM Approve"].ToString()))
                        {
                            if (r["CCM Approve"].ToString() == "1")
                                oSheet.Cells[current_row, 10] = "x";
                            else if (r["CCM Approve"].ToString() == "2")
                                oSheet.Cells[current_row, 9] = "x";
                        }
                        //Ke toan duyet
                        if (!string.IsNullOrEmpty(r["KT Approve"].ToString()))
                        {
                            if (r["KT Approve"].ToString() == "1")
                                oSheet.Cells[current_row, 12] = "x";
                            else if (r["KT Approve"].ToString() == "2")
                                oSheet.Cells[current_row, 11] = "x";
                        }
                        STT++;
                        current_row++;
                    }
                }

                //Total BCH
                if (current_row - group_row > 1)
                {
                    oSheet.Cells[group_row, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", group_row + 1, current_row - 1);
                    oSheet.Cells[group_row, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", group_row + 1, current_row - 1);
                    oSheet.Cells[group_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", group_row + 1, current_row - 1);
                }
                current_row++;
                #endregion
                //Total
                oSheet.Range["A" + current_row, "C" + current_row].Merge();
                oSheet.Cells[current_row, 1] = "TỔNG CỘNG";
                oSheet.Range["A" + current_row, "N" + current_row].Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                oSheet.Range["A" + current_row, "N" + current_row].Font.Bold = true;

                oSheet.Cells[current_row, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1})", 7, current_row - 1);
                oSheet.Cells[current_row, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", 7, current_row - 1);
                oSheet.Cells[current_row, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", 7, current_row - 1);
                current_row++;
                
                //Border
                ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A6", "N" + (current_row - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                oXL.ActiveWindow.Activate();
            }
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

        private void Bill_Caculator(string FProject, int Period, string BPCode, string BGroup, string PUType, int BType, int DocEntry, DateTime Todate, out decimal KLTT_Kytruoc, out decimal KLTT_Kynay)
        {
            KLTT_Kytruoc = 0;
            KLTT_Kynay = 0;
            DataTable SPP = Load_Data_KLTT(FProject, "SPP", Period - 1, BPCode, BGroup, PUType, DocEntry, Todate);
            DataTable SPP_CURRENT = Load_Data_KLTT(FProject, "SPP", Period, BPCode, BGroup, PUType, DocEntry, Todate);
            DataTable AI = Load_Data_KLTT(FProject, "AI", Period, BPCode, BGroup, PUType, DocEntry, Todate);

            #region Thong tin chi tiet HD
            decimal GTHD = 0, PLTT = 0, PLT = 0;
            float PTBH = 0, PTTU = 0, PTHU = 0, PTGL = 0;
            string TTTU = "", HTBH = "";
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
                    }
                    else if (r["Type"].ToString() == "PLT")
                        PLT += tmp;
                    else if (r["Type"].ToString() == "PLTT")
                    {
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
                GTTT_KN = Math.Round((1 - (decimal)PTGL) * GTKN, 0) ;

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
                //String.Format("{0:n}", Math.Round((pp_ca * (1 - (decimal)PTGL)) + pp_tu_lastbill - pp_hu_lastbill, 0));
                else
                    KLTT_Kytruoc = Math.Round((pp_ca * (1 - (decimal)PTGL)), 0);
                //String.Format("{0:n}", Math.Round((pp_ca * (1 - (decimal)PTGL)), 0));
            }
            else if (BType == 3)
            {
                if (pp_tu_lastbill > 0)
                    KLTT_Kytruoc = Math.Round((pp_ca * (1 - (decimal)PTGL)) + pp_tu_lastbill - pp_hu_lastbill);
                //String.Format("{0:n}", Math.Round((pp_ca * (1 - (decimal)PTGL)) + pp_tu_lastbill - pp_hu_lastbill));
                else
                    KLTT_Kytruoc = Math.Round((pp_ca * (1 - (decimal)PTGL)));
                //String.Format("{0:n}", Math.Round((pp_ca * (1 - (decimal)PTGL))));
            }
            //GT thanh toan ky nay
            KLTT_Kynay = Math.Round(GTTT_KN - KLTT_Kytruoc, 0);
                //String.Format("{0:n}", Math.Round(decimal.Parse(EditText16.Value) - decimal.Parse(EditText18.Value), 0));
            #endregion
        }

        private SAPbouiCOM.EditText EditText0;
    }
}
