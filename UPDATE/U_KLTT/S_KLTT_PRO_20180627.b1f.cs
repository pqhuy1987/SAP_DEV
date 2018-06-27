using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;
using System.Xml;

namespace U_KLTT
{
    [FormAttribute("UDO_FT_KLTT")]
    class S_KLTT_PRO_20180627 : UDOFormBase
    {
        public S_KLTT_PRO_20180627()
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
            this.DataDeleteBefore += new DataDeleteBeforeHandler(this.Form_DataDeleteBefore);

        }

        private SAPbouiCOM.EditText EditText0;
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

        private bool Check_Approve_KLTT(int pDocEntry)
        {
            bool kq = false;
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("KLTT_Check_Approve", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry", pDocEntry);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                if (result.Rows.Count > 0)
                {
                    int tmp = 0;
                    int.TryParse(result.Rows[0][0].ToString(), out tmp);
                    if (tmp > 0) kq = true;
                    else kq = false;
                }
                else
                {
                    kq = false;
                }
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
            return kq;
        }

        private void Form_DataDeleteBefore(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //Get ObjectKey has created
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.LoadXml(((SAPbouiCOM.BusinessObjectInfo)pVal).ObjectKey.Replace("Khối lượng thanh toán","KLTT"));
            XmlNodeList nodeList = xmldoc.GetElementsByTagName("DocEntry");
            string Object_Key = string.Empty;
            if (nodeList.Count > 0)
                Object_Key = nodeList.Item(0).InnerText;
            int DocEntry = 0;
            int.TryParse(Object_Key,out DocEntry);
            if (Check_Approve_KLTT(DocEntry))
            {
                oApp.SetStatusBarMessage("Bill was approved ! Delete failed !", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                BubbleEvent = false;
                
            }

        }
    }
}
