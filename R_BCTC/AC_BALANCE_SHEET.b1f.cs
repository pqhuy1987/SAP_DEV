using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;

namespace R_BCTC
{
    [FormAttribute("R_BCTC.AC_BALANCE_SHEET", "AC_BALANCE_SHEET.b1f")]
    class AC_BALANCE_SHEET : UserFormBase
    {
        public AC_BALANCE_SHEET()
        {
        }
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
        int Level = 0;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_frd").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txt_tod").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("txt_lvl").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_6").Specific));
            this.Grid0.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid0_DoubleClickAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_load").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
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
        }

        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button0;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                DateTime frdate = DateTime.Today;
                DateTime todate = DateTime.Today;
                if (string.IsNullOrEmpty(EditText0.Value))
                {
                    oApp.MessageBox("Please Enter From Date !");
                    return;
                }
                else
                {
                    frdate = DateTime.ParseExact(EditText0.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                }

                if (string.IsNullOrEmpty(EditText1.Value))
                {
                    oApp.MessageBox("Please Enter To Date !");
                    return;
                }
                else
                {
                    todate = DateTime.ParseExact(EditText1.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                }

                if (string.IsNullOrEmpty(EditText2.Value))
                {
                    oApp.MessageBox("Please Enter Level !");
                    return;
                }
                else
                {
                    int.TryParse(EditText2.Value.Trim(), out Level);
                }
                DataTable result = Get_Data(Level, frdate, todate);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);
                Grid0.AutoResizeColumns();
                for (int i = 0; i < Grid0.Columns.Count; i++)
                {
                    if (i >= 4)
                        Grid0.Columns.Item(i).RightJustified = true;
                }
                for (int i = 0; i < Grid0.Rows.Count; i++)
                {
                    int t_level = 0;
                    int.TryParse(Grid0.DataTable.GetValue("Cap", i).ToString(), out t_level);
                    if (t_level < Level)
                    {
                        //Dark Gray
                        Grid0.CommonSetting.SetRowBackColor(i + 1, 14277081);
                    }
                    else
                    {
                        //White Color
                        Grid0.CommonSetting.SetRowBackColor(i + 1, 16777215);
                    }
                }
                Grid0.Columns.Item(0).Visible = false;
                Grid0.Columns.Item(3).Visible = false;
                Grid0.AutoResizeColumns();
                oApp.SetStatusBarMessage("Load Data Completed !", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
            }

        }

        System.Data.DataTable Get_Data(int pLevel, DateTime pFrDate, DateTime pToDate)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BangCanDoiTaiKhoan", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@TuNgay", pFrDate);
                cmd.Parameters.AddWithValue("@DenNgay", pToDate);
                cmd.Parameters.AddWithValue("@Level", pLevel);
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

        System.Data.DataTable Get_MenuUID()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("GET_MENUUID", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@ReportName", "Sổ chi tiết tài khoản");
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

        private SAPbouiCOM.DataTable Convert_SAP_DataTable(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = oApp.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_BALS"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_BALS");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_BALS");
            }

            //Add column to DataTable
            foreach (System.Data.DataColumn c in pDataTable.Columns)
            {
                try
                {
                    if (c.DataType.ToString() == "System.DateTime")
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Date);
                    //else if (c.DataType.ToString() == "System.Int16")
                    //{
                    //    oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Integer);
                    //}
                    //else if (c.DataType.ToString() == "System.Decimal")
                    //{
                    //    oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Sum);
                    //}
                    else
                    {
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Text);
                    }
                }
                catch
                { }
            }
            //Add row to DataTable
            foreach (System.Data.DataRow r in pDataTable.Rows)
            {
                oDT.Rows.Add();
                //foreach (System.Data.DataColumn c in pDataTable.Columns)
                for (int i = 0 ; i < pDataTable.Columns.Count; i++)
                {
                    //oDT.SetValue(  c.ColumnName, oDT.Rows.Count - 1, r[c.ColumnName].ToString());
                    string Col_Name = pDataTable.Columns[i].ColumnName;
                    if (i > 3)
                    {
                        decimal tmp = 0;
                        decimal.TryParse(r[Col_Name].ToString(), out tmp);
                        if (tmp != 0)
                            oDT.SetValue(Col_Name, oDT.Rows.Count - 1, tmp.ToString("N0").Replace(',','.'));
                    }
                    else
                        oDT.SetValue(Col_Name, oDT.Rows.Count - 1, r[Col_Name].ToString());
                }
            }

            return oDT;
        }

        private bool CheckExistUniqueID(SAPbouiCOM.Form pForm, string pItemID)
        {
            if (pForm.DataSources.DataTables.Count > 0)
            {
                for (int i = 0; i < pForm.DataSources.DataTables.Count; i++)
                {
                    if (pForm.DataSources.DataTables.Item(i).UniqueID == pItemID)
                    {
                        return true;
                    }
                }
                return false;
            }
            else
            {
                return false;
            }
        }

        private void Grid0_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            if (Grid0.Rows.SelectedRows.Count == 1)
            {
                DataTable rs = Get_MenuUID();
                if (rs.Rows.Count > 0)
                {
                    oApp.ActivateMenuItem(rs.Rows[0]["MenuUID"].ToString());
                    SAPbouiCOM.Form act_frm = oApp.Forms.ActiveForm;
                    ((SAPbouiCOM.EditText)act_frm.Items.Item("1000003").Specific).Value = EditText0.Value;
                    ((SAPbouiCOM.EditText)act_frm.Items.Item("1000009").Specific).Value = EditText1.Value;
                    string STK = Grid0.DataTable.GetValue(1, Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                    ((SAPbouiCOM.EditText)act_frm.Items.Item("1000027").Specific).Value = STK;
                    act_frm.Items.Item("1").Click();
                }
            }
        }
    }
}
