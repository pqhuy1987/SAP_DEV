using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;

namespace U_DUTRU
{
    [FormAttribute("UDO_FT_CTG")]
    class CTG_20171220 : UDOFormBase
    {
        public CTG_20171220()
        {
        }
        SqlConnection conn = null;
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("1_U_G").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_dteq").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("2_U_G").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("bt_dtccm").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Matrix Matrix0;

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

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Matrix Matrix1;
        private SAPbouiCOM.Button Button1;

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //CCM
            if (pVal.FormMode == 1)
            {
                //Get DATA SUM CCM
                int DocEntry = int.Parse(((SAPbouiCOM.EditText)(this.GetItem("0_U_E").Specific)).Value);
                string Project = ((SAPbouiCOM.EditText)(this.GetItem("20_U_E").Specific)).Value;
                DataTable Sum_DUTRU = Load_Data_DUTRU(Project, DocEntry);

                oApp.ActivateMenuItem("DUTRU");
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                //Get DocNum CTG
                try
                {
                    ((SAPbouiCOM.EditText)oForm.Items.Item("20_U_E").Specific).Value = DocEntry.ToString();
                    try
                    {
                        //Select Type DUTRU
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("21_U_Cb").Specific).Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    catch (Exception ex)
                    {
                        oApp.MessageBox(ex.Message);
                    }
                    oForm.Freeze(true);
                    oForm.PaneLevel = 1;
                    SAPbouiCOM.Matrix oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;
                    for (int i = 0; i < Sum_DUTRU.Rows.Count; i++)
                    {
                        oApp.SetStatusBarMessage(string.Format("Creating DUTRU: {0}", (i + 1) + "/" + Sum_DUTRU.Rows.Count), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        DataRow r = Sum_DUTRU.Rows[i];
                        //Subproject Code
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_2").Cells.Item(i + 1).Specific).Value = r["High_Level_SUM"].ToString();
                        //Subproject Description
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_3").Cells.Item(i + 1).Specific).Value = r["High_Level_Name"].ToString();
                        //CP PRELIM
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_4").Cells.Item(i + 1).Specific).Value = r["CP_PRELIM"].ToString();
                        //CP TB
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_5").Cells.Item(i + 1).Specific).Value = r["CP_TB"].ToString();
                        //CP Khac
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_6").Cells.Item(i + 1).Specific).Value = r["CP_Khac"].ToString();
                        //CP NCC
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_7").Cells.Item(i + 1).Specific).Value = r["CP_NCC"].ToString();
                        //CP NTP
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_8").Cells.Item(i + 1).Specific).Value = r["CP_NTP"].ToString();
                        //CP DTC
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_9").Cells.Item(i + 1).Specific).Value = r["CP_DTC"].ToString();
                        //CP VTP
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_10").Cells.Item(i + 1).Specific).Value = r["CP_VTP"].ToString();
                        //CP VC
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_11").Cells.Item(i + 1).Specific).Value = r["CP_VC"].ToString();
                        //CP CN
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_12").Cells.Item(i + 1).Specific).Value = r["CP_CN"].ToString();
                        //CP DP
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_13").Cells.Item(i + 1).Specific).Value = r["CP_DP"].ToString();
                        //CP DP2
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_14").Cells.Item(i + 1).Specific).Value = r["CP_DP2"].ToString();



                        if (i < Sum_DUTRU.Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((SAPbouiCOM.EditText)oMtx.Columns.Item(1).Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                    oForm.Freeze(false);
                }
                catch (Exception ex)
                {
                }
                finally
                {
                    oForm.Freeze(false);
                }
            }
            else
            {
                Application.SBO_Application.MessageBox("Function works only on View Mode");
            }
        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //Thiet bi
            if (pVal.FormMode == 1)
            {
                //Get DATA SUM CCM
                int DocEntry = int.Parse(((SAPbouiCOM.EditText)(this.GetItem("0_U_E").Specific)).Value);
                string Project = ((SAPbouiCOM.EditText)(this.GetItem("20_U_E").Specific)).Value;
                DataTable Sum_DUTRU = Load_Data_DUTRU_TB(Project, DocEntry);

                oApp.ActivateMenuItem("DUTRU");
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                //Get DocNum CTG
                try
                {
                    ((SAPbouiCOM.EditText)oForm.Items.Item("20_U_E").Specific).Value = DocEntry.ToString();
                    try
                    {
                        //Select Type DUTRU
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("21_U_Cb").Specific).Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    catch (Exception ex)
                    {
                        oApp.MessageBox(ex.Message);
                    }
                    oForm.Freeze(true);
                    oForm.PaneLevel = 1;
                    SAPbouiCOM.Matrix oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;
                    for (int i = 0; i < Sum_DUTRU.Rows.Count; i++)
                    {
                        oApp.SetStatusBarMessage(string.Format("Creating DUTRU TB: {0}", (i + 1) + "/" + Sum_DUTRU.Rows.Count), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        DataRow r = Sum_DUTRU.Rows[i];
                        //Subproject Code
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_2").Cells.Item(i + 1).Specific).Value = r["High_Level_SUM"].ToString();
                        //Subproject Description
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_3").Cells.Item(i + 1).Specific).Value = r["High_Level_Name"].ToString();
                        //CP VC
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_11").Cells.Item(i + 1).Specific).Value = r["CP_VC"].ToString();
                        //CP Mua ban
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_15").Cells.Item(i + 1).Specific).Value = r["CP_MB"].ToString();
                        //CP Thue
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_16").Cells.Item(i + 1).Specific).Value = r["CP_THUE"].ToString();
                        //CP Van hanh
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_17").Cells.Item(i + 1).Specific).Value = r["CP_VH"].ToString();

                        if (i < Sum_DUTRU.Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((SAPbouiCOM.EditText)oMtx.Columns.Item(1).Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                    oForm.Freeze(false);
                }
                catch (Exception ex)
                {
                }
                finally
                {
                    oForm.Freeze(false);
                }
            }
            else
            {
                Application.SBO_Application.MessageBox("Function works only on View Mode");
            }
        }
        System.Data.DataTable Load_Data_DUTRU(string pFinancialProject, int pDocEntry)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CALCULATE_DUTRU", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                cmd.Parameters.AddWithValue("@DocEntry", pDocEntry);

                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.Message);
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }
            return result;
        }
        System.Data.DataTable Load_Data_DUTRU_TB(string pFinancialProject, int pDocEntry)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("CALCULATE_DUTRU_TB", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                cmd.Parameters.AddWithValue("@DocEntry", pDocEntry);

                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.Message);
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
