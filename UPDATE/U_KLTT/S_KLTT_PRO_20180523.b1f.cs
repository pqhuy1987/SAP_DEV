﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;

namespace U_KLTT
{
    [FormAttribute("UDO_FT_KLTT2")]
    class S_KLTT_PRO_20180523 : UDOFormBase
    {
        public S_KLTT_PRO_20180523()
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
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_tmp").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new SAPbouiCOM.Framework.FormBase.DataLoadAfterHandler(this.Form_DataLoadAfter);
            this.CloseBefore += new CloseBeforeHandler(this.Form_CloseBefore);

        }

        private void OnCustomInitialize()
        {
            this.oApp = (SAPbouiCOM.Application)SAPbouiCOM.Framework.Application.SBO_Application; //UI
            this.oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany(); //DI
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

        private SAPbouiCOM.EditText EditText0;

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            try
            {
                DataTable rs = Load_Data_KLTT();
                if (rs.Rows.Count == 1)
                {
                    double tmp = 0;
                    double.TryParse(rs.Rows[0]["SUM_CA_NOVAT"].ToString(), out tmp);
                    EditText0.Value = tmp.ToString("N2");
                }
            }
            catch
            { }

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

        System.Data.DataTable Load_Data_KLTT()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                int DocEntry = 0;
                int.TryParse(((SAPbouiCOM.EditText)(this.GetItem("0_U_E").Specific)).Value, out DocEntry);
                cmd = new SqlCommand("KLTT_GT_KYNAY", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry", DocEntry);
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
