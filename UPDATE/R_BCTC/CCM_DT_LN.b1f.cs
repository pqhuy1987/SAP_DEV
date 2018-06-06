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
    [FormAttribute("R_BCTC.CCM_DT_LN", "CCM_DT_LN.b1f")]
    class CCM_DT_LN : UserFormBase
    {
        public CCM_DT_LN()
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
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.OnCustomInitialize();

        }
        private SAPbouiCOM.Button Button0;
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

        System.Data.DataTable Get_Lst_BaseLine()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_BASELINE_LST", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
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

        private SAPbouiCOM.DataTable Convert_SAP_DataTable(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_BLList"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_BLList");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_BLList");
            }

            //Add column to DataTable
            foreach (System.Data.DataColumn c in pDataTable.Columns)
            {
                try
                {
                    if (c.DataType.ToString() == "System.DateTime")
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Date);
                    else
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Text);
                }
                catch
                { }

            }
            //Add row to DataTable
            foreach (System.Data.DataRow r in pDataTable.Rows)
            {
                oDT.Rows.Add();
                foreach (System.Data.DataColumn c in pDataTable.Columns)
                {
                    oDT.SetValue(c.ColumnName, oDT.Rows.Count - 1, r[c.ColumnName]);
                }
            }

            return oDT;
        }

        System.Data.DataTable Get_Data(DateTime pToDate)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_DT_LN_TONGHOP_LST", conn);
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

        System.Data.DataTable Get_Data_CURRENT(string pFProject, int pProjectID)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_CURRENT_A_INDEX", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FProject", pFProject);
                cmd.Parameters.AddWithValue("@ProjectID", pProjectID);
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

        System.Data.DataTable Get_Data_BASELINE(int pBASELINE_DocEntry, int pProjectID)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CCM_BASELINE_A_INDEX", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@BASELINE_DocEntry", pBASELINE_DocEntry);
                cmd.Parameters.AddWithValue("@ProjectID", pProjectID);
                //cmd.Parameters.AddWithValue("@ProjectID", pProjectID);
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
            if (!string.IsNullOrEmpty(EditText0.Value))
            {
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
                        oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\TMP_DNTL_TH.xlsx");
                        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                        int current_row = 5;
                        int current_group = 5;
                        int current_groupGDDA = 5;
                        int STT = 1;
                        int STT_GDDA = 1;
                        string GDDA = "";
                        string FProject = "";
                        oSheet.Cells[2, 3] = string.Format("Tháng {0}", todate.ToString("MM - yyyy"));

                        foreach (DataRow r in rs_detail.Rows)
                        {
                            if (GDDA != r["OWNER"].ToString())
                            {
                                //Print subtotal GDDA
                                if (FProject != "")
                                {
                                    oSheet.Cells[current_groupGDDA, 3].Formula = string.Format("=SUBTOTAL(9,C{0}:C{1}", current_groupGDDA + 1, current_row - 1);
                                    oSheet.Cells[current_groupGDDA, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1}", current_groupGDDA + 1, current_row - 1);
                                    oSheet.Cells[current_groupGDDA, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1}", current_groupGDDA + 1, current_row - 1);
                                    oSheet.Cells[current_groupGDDA, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1}", current_groupGDDA + 1, current_row - 1);
                                    oSheet.Cells[current_groupGDDA, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1}", current_groupGDDA + 1, current_row - 1);
                                    oSheet.Cells[current_groupGDDA, 9].Formula = string.Format("=SUBTOTAL(9,I{0}:I{1}", current_groupGDDA + 1, current_row - 1);
                                    oSheet.Cells[current_groupGDDA, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1}", current_groupGDDA + 1, current_row - 1);
                                }
                                //Print GDDA
                                current_groupGDDA = current_row;
                                oSheet.Cells[current_row, 1].Formula = string.Format("=ROMAN({0}",STT_GDDA);
                                oSheet.Cells[current_row, 2] = r["GDDA"];
                                oSheet.Range["A" + current_row, "M" + current_row].Font.Bold = true;
                                oSheet.Range["A" + current_row, "M" + current_row].Interior.Color = System.Drawing.Color.FromArgb(248, 203, 173);
                                current_row++;
                                STT = 1;
                                STT_GDDA++;
                                GDDA = r["OWNER"].ToString();


                            }
                            if (FProject != r["PrjCode"].ToString())
                            {
                                //Print subtotal PreGroup
                                if (FProject != "")
                                {
                                    oSheet.Cells[current_row, 3] = r["DTKehoach"];
                                    oSheet.Cells[current_group, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1}", current_group + 1, current_row - 1);
                                    oSheet.Cells[current_group, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1}", current_group + 1, current_row - 1);
                                    oSheet.Cells[current_group, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1}", current_group + 1, current_row - 1);
                                    oSheet.Cells[current_group, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1}", current_group + 1, current_row - 1);
                                    oSheet.Cells[current_group, 9].Formula = string.Format("=SUBTOTAL(9,I{0}:I{1}", current_group + 1, current_row - 1);
                                    oSheet.Cells[current_group, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1}", current_group + 1, current_row - 1);
                                }
                                //Print FProject Group
                                current_group = current_row;
                                oSheet.Cells[current_row, 1] = STT;
                                oSheet.Cells[current_row, 2] = r["PrjName"];
                                current_row++;
                                STT++;
                                FProject = r["PrjCode"].ToString();
                            }
                            //Ten goi thau
                            oSheet.Cells[current_row, 2] = r["NAME"];
                            //Doanh thu ke hoach nam
                            oSheet.Cells[current_row, 3] = 0;
                            //GTHD
                            oSheet.Cells[current_row, 4] = r["GTHD"];
                            //DT nam truoc
                            oSheet.Cells[current_row, 5] = r["DTNamtruoc"];
                            //DT trong nam
                            oSheet.Cells[current_row, 6] = r["DTthucte"];
                            //% Hoan thanh
                            oSheet.Cells[current_row, 7].Formula = string.Format("=F{0}/C{0}", current_row);
                            //DT con lai phai thu trong nam
                            oSheet.Cells[current_row, 8].Formula = string.Format("=C{0}-F{0}", current_row);
                            //DT ke hoach sang nam
                            oSheet.Cells[current_row, 9].Formula = string.Format("=D{0}-C{0}-E{0}", current_row);
                            
                            //Get A-Index Baseline
                            DataTable rs_baseline = Get_Data_BASELINE(-1, int.Parse(r["AbsEntry"].ToString()));
                            if (rs_baseline.Rows.Count >= 1)
                                oSheet.Cells[current_row, 10] = rs_baseline.Rows[0]["A-INDEX"];
                            //Get A-Index Current
                            DataTable rs_current = Get_Data_BASELINE(-1, int.Parse(r["AbsEntry"].ToString()));
                                //Get_Data_CURRENT(r["PrjCode"].ToString(), int.Parse(r["AbsEntry"].ToString()));
                            if (rs_current.Rows.Count >= 1)
                                oSheet.Cells[current_row, 11] = rs_current.Rows[0]["A-INDEX"];
                            //Loi nhuan 2
                            oSheet.Cells[current_row, 12].Formula = string.Format("=K{0}*C{0}", current_row);
                            oSheet.Range["A" + current_row, "B" + current_row].Font.Italic = true;
                            current_row++;
                        }
                        //Subtotal DA
                        oSheet.Cells[current_group, 3].Formula = string.Format("=SUBTOTAL(9,C{0}:C{1}", current_group + 1, current_row - 1);
                        oSheet.Cells[current_group, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1}", current_group + 1, current_row - 1);
                        oSheet.Cells[current_group, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1}", current_group + 1, current_row - 1);
                        oSheet.Cells[current_group, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1}", current_group + 1, current_row - 1);
                        oSheet.Cells[current_group, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1}", current_group + 1, current_row - 1);
                        oSheet.Cells[current_group, 9].Formula = string.Format("=SUBTOTAL(9,I{0}:I{1}", current_group + 1, current_row - 1);
                        oSheet.Cells[current_group, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1}", current_group + 1, current_row - 1);
                        //Subtotal GDDA
                        oSheet.Cells[current_groupGDDA, 3].Formula = string.Format("=SUBTOTAL(9,C{0}:C{1}", current_groupGDDA + 1, current_row - 1);
                        oSheet.Cells[current_groupGDDA, 4].Formula = string.Format("=SUBTOTAL(9,D{0}:D{1}", current_groupGDDA + 1, current_row - 1);
                        oSheet.Cells[current_groupGDDA, 5].Formula = string.Format("=SUBTOTAL(9,E{0}:E{1}", current_groupGDDA + 1, current_row - 1);
                        oSheet.Cells[current_groupGDDA, 6].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1}", current_groupGDDA + 1, current_row - 1);
                        oSheet.Cells[current_groupGDDA, 8].Formula = string.Format("=SUBTOTAL(9,H{0}:H{1}", current_groupGDDA + 1, current_row - 1);
                        oSheet.Cells[current_groupGDDA, 9].Formula = string.Format("=SUBTOTAL(9,I{0}:I{1}", current_groupGDDA + 1, current_row - 1);
                        oSheet.Cells[current_groupGDDA, 12].Formula = string.Format("=SUBTOTAL(9,L{0}:L{1}", current_groupGDDA + 1, current_row - 1);

                        //Border
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A4", "M" + (current_row - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        oXL.ActiveWindow.Activate();
                    }
                }
                catch (Exception ex)
                {
                    oApp.MessageBox(ex.Message);
                }
            }
            else
            {
                oApp.MessageBox("Please select from list !");
            }

        }

        private SAPbouiCOM.EditText EditText0;

    }
}
