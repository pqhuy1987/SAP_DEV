
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Xml;
using System.Data.SqlClient;
using System.Data;

namespace S_FixedAsset
{

    [FormAttribute("65214", "Receipt from Production.b1f")]
    class Receipt_from_Production : SystemFormBase
    {
        public Receipt_from_Production()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataAddAfter += new DataAddAfterHandler(this.Form_DataAddAfter);

        }

        private void Form_DataAddAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            if (((SAPbouiCOM.BusinessObjectInfo)pVal).ActionSuccess)
            {
                SAPbobsCOM.Company oCom = ((SAPbobsCOM.Company)(Application.SBO_Application.Company.GetDICompany()));
                SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oR_RecordSet.DoQuery("Select * from [@ADDONCFG]");
                string uid = oR_RecordSet.Fields.Item("Code").Value.ToString();
                string pwd = oR_RecordSet.Fields.Item("Name").Value.ToString();
                //Get ObjectKey has created
                XmlDocument xmldoc = new XmlDocument();
                xmldoc.LoadXml(((SAPbouiCOM.BusinessObjectInfo)pVal).ObjectKey);
                XmlNodeList nodeList = xmldoc.GetElementsByTagName("DocEntry");
                string Object_Key = string.Empty;
                if (nodeList.Count > 0)
                    Object_Key = nodeList.Item(0).InnerText;

                SqlConnection conn = new SqlConnection(string.Format("Data Source={0}; Initial Catalog={1}; User id={2}; Password={3};", oCom.Server, oCom.CompanyDB, uid, pwd));
                SqlCommand cmd = new SqlCommand("DeleteJournalEntry_ReceiptfromProduction", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocNum", Object_Key);
                try
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.MessageBox("Error when delete JournalEntry: " + ex.Message);
                }
                finally
                {
                    conn.Close();
                    cmd.Dispose();
                }
            }
        }

        private void OnCustomInitialize()
        {

        }
    }
}
