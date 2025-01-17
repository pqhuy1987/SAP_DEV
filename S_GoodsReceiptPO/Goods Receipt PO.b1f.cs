
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Globalization;
using System.Xml;
using System.Data.SqlClient;
using System.Data;
namespace S_GoodsReceiptPO
{

    [FormAttribute("143", "Goods Receipt PO.b1f")]
    class Goods_Receipt_PO : SystemFormBase
    {
        public Goods_Receipt_PO()
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
        class Inventory_Item
        {
            public string Item_No;
            public string FixedAsset_ItemNo;
            public double Quantity;
            public double Unitprice;
            public string Whse;
            public int LineNum;
            public Inventory_Item()
            {
                Item_No = "";
                Quantity = 0;
                Unitprice = 0;
                Whse = "";
                LineNum = -1;
                FixedAsset_ItemNo = "";
            }
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
                
                SqlCommand cmd = null;
                SqlConnection conn = new SqlConnection(string.Format("Data Source={0}; Initial Catalog={1}; User id={2}; Password={3};", oCom.Server, oCom.CompanyDB, uid, pwd));

                //Delete JournalEntry GoodsReceipt PO
                try
                {
                    cmd = new SqlCommand("DeleteJournalEntry_GoodReceiptPO", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@DocNum", Object_Key);
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.MessageBox("Delete JournalEntry GoodsReceipt PO Error: " + ex.Message);
                }
                finally
                {
                    conn.Close();
                    cmd.Dispose();
                }

                //Get Info From Goods Receipt PO by UI
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                string p_post_date = ((SAPbouiCOM.EditText)oForm.Items.Item("10").Specific).Value;
                string p_document_date = ((SAPbouiCOM.EditText)oForm.Items.Item("46").Specific).Value;
                List<Inventory_Item> Inven_Lst = new List<Inventory_Item>();
                oForm.Freeze(true);
                oForm.PaneLevel = 1;
                try
                {
                    SAPbouiCOM.Matrix oMtx = ((SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific);
                    for (int i = 1; i <= oMtx.RowCount; i++)
                    {
                        if (!String.IsNullOrEmpty(((SAPbouiCOM.EditText)oMtx.Columns.Item(3).Cells.Item(i).Specific).Value))
                        {
                            Inventory_Item tmp = new Inventory_Item();
                            tmp.Item_No = ((SAPbouiCOM.EditText)oMtx.Columns.Item(3).Cells.Item(i).Specific).Value;
                            double.TryParse(((SAPbouiCOM.EditText)oMtx.Columns.Item(13).Cells.Item(i).Specific).Value, out tmp.Quantity);
                            string tmp_unit_price = ((SAPbouiCOM.EditText)oMtx.Columns.Item(20).Cells.Item(i).Specific).Value;
                            if (!String.IsNullOrEmpty(tmp_unit_price))
                            {
                                double.TryParse(tmp_unit_price.Split(',')[0].Replace('.', ','), out tmp.Unitprice);
                            }
                            tmp.Whse = ((SAPbouiCOM.EditText)oMtx.Columns.Item(32).Cells.Item(i).Specific).Value;
                            Inven_Lst.Add(tmp);
                        }
                    }
                }
                catch
                { }
                finally
                {
                    oForm.Freeze(false);
                }

                //Using DI Create Goods Receipt
                SAPbobsCOM.Documents oGrp = (SAPbobsCOM.Documents)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                oGrp.DocDate = DateTime.ParseExact(p_document_date, "yyyyMMdd", CultureInfo.InvariantCulture);
                oGrp.TaxDate = DateTime.ParseExact(p_post_date, "yyyyMMdd", CultureInfo.InvariantCulture);
                oGrp.Reference2 = Object_Key;
                int tmp_i = 0;
                int FixedAsset_Count = 0;
                foreach (Inventory_Item t in Inven_Lst)
                {
                    tmp_i++;
                    //string str_query = string.Format("Select count(*) as IsFixedAsset from OITM where ItemCode in (Select U_FA from OITM where ItemCode = '{0}')and ItemType ='F'", t.Item_No);
                    string str_query = string.Format("Select a.ItemCode,a.U_FA,(Select b.ItemType from OITM b where b.ItemCode = a.U_FA) as ItemType, a.ItmsGrpCod,(Select b.ItmsGrpCod from OITM b where b.ItemCode = a.U_FA) as ItmsGrpCod_FA  from OITM a where a.U_FA = '{0}'", t.Item_No);
                    oR_RecordSet.DoQuery(str_query);
                    if (oR_RecordSet.RecordCount > 0)
                    {
                        if (oR_RecordSet.Fields.Item("ItemType").Value.ToString() == "F" 
                            && ((oR_RecordSet.Fields.Item("ItmsGrpCod").Value.ToString() == "103" 
                            && oR_RecordSet.Fields.Item("ItmsGrpCod_FA").Value.ToString() == "103")
                            || (oR_RecordSet.Fields.Item("ItmsGrpCod").Value.ToString() == "105" 
                            && oR_RecordSet.Fields.Item("ItmsGrpCod_FA").Value.ToString() == "105")))
                        {
                            t.FixedAsset_ItemNo = oR_RecordSet.Fields.Item("ItemCode").Value.ToString();
                            if (!string.IsNullOrEmpty(t.FixedAsset_ItemNo))
                            {
                                oGrp.Lines.ItemCode = t.FixedAsset_ItemNo;
                                oGrp.Lines.Quantity = t.Quantity;
                                oGrp.Lines.UnitPrice = 0;
                                oGrp.Lines.Price = 0;
                                oGrp.Lines.WarehouseCode = t.Whse;
                                t.LineNum = FixedAsset_Count++;
                                if (tmp_i < Inven_Lst.Count)
                                    oGrp.Lines.Add();
                            }
                        }
                    }
                }
                if (FixedAsset_Count > 0)
                {
                    int RetVal = oGrp.Add();
                    if (RetVal == 0)
                    {
                        string New_Object_Key = oCom.GetNewObjectKey();
                        //Update Unit Price
                        cmd = new SqlCommand();
                        cmd.CommandType = CommandType.Text;
                        double receipt_total = 0;

                        //Update IGN1 SQL
                        foreach (Inventory_Item t in Inven_Lst)
                        {
                            if (t.LineNum >= 0)
                            {
                                double tmp_sum = t.Unitprice * t.Quantity;
                                receipt_total += tmp_sum;
                                string update_IGN1_query = string.Format("Update IGN1 set Price={0},LineTotal={1},OpenSum={2},PriceBefDi={3},TotalSumSy={4},OpenSumSys={5},INMPrice={6},StockPrice={7},StockSum={8},StockSumSc={9} where DocEntry={10} and ItemCode='{11}' and LineNum={12};"
                                    , t.Unitprice, tmp_sum, tmp_sum, t.Unitprice, tmp_sum, tmp_sum, t.Unitprice, t.Unitprice, tmp_sum, tmp_sum, New_Object_Key, t.FixedAsset_ItemNo, t.LineNum);
                                cmd.CommandText += update_IGN1_query;
                            }
                        }

                        //Update OIGN SQL
                        string update_OIGN_query = string.Format("Update OIGN set DocTotal={0},DocTotalSy={1},Max1099={2} where DocEntry={3};", receipt_total, receipt_total, receipt_total, New_Object_Key);
                        cmd.CommandText += update_OIGN_query;

                        try
                        {
                            cmd.Connection = conn;
                            conn.Open();
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            Application.SBO_Application.MessageBox(string.Format("Addon: Error when update GoodsRecipt: {0}", ex.Message));
                        }
                        finally
                        {
                            conn.Close();
                            cmd.Dispose();
                        }
                    }
                    else
                    {
                        int ErrCode;
                        string ErrMsg;
                        oCom.GetLastError(out ErrCode, out ErrMsg);
                        Application.SBO_Application.StatusBar.SetText(string.Format("Addon: Failed create Good Receipt from Good Receipt PO: {0}|{1}", ErrCode, ErrMsg), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
            }
        }

        private void OnCustomInitialize()
        {

        }
    }
}
