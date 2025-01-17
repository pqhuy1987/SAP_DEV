
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace S_FixedAsset
{

    [FormAttribute("1473000075", "Asset Master Data.b1f")]
    class Asset_Master_Data : SystemFormBase
    {
        public Asset_Master_Data()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_cpinv").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.GetItem("bt_cpinv").Top = this.GetItem("2").Top;
            this.GetItem("bt_cpinv").Left = ((this.GetItem("2").Left + this.GetItem("2").Width)
                        + 6);
            this.GetItem("bt_cpinv").Height = this.GetItem("2").Height;
            this.GetItem("bt_cpinv").FontSize = this.GetItem("2").FontSize;
            this.GetItem("bt_cpinv").Width = (this.GetItem("2").Width + 50);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Button Button0;

        private void OnCustomInitialize()
        {
        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //Get info of Fixed Asset
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
            try
            {
                #region Header
                //Item Group
                string p_ItemGrp = "";
                var cbo_ItemGrp = oForm.Items.Item("39").Specific;
                p_ItemGrp = ((SAPbouiCOM.ComboBox)cbo_ItemGrp).Selected.Value;
                if (p_ItemGrp == "103" || p_ItemGrp == "105")
                {
                    //Item Series
                    string p_Series = "";
                    var cbo_series = oForm.Items.Item("1320002059").Specific;
                    p_Series = ((SAPbouiCOM.ComboBox)cbo_series).Selected.Value;
                    //Item No
                    string p_ItemNo = "";
                    var txt_itemno = oForm.Items.Item("5").Specific;
                    p_ItemNo = ((SAPbouiCOM.EditText)txt_itemno).Value;
                    //Description
                    string p_Description = "";
                    var txt_des = oForm.Items.Item("7").Specific;
                    p_Description = ((SAPbouiCOM.EditText)txt_des).Value;
                    //Foreign Name
                    string p_ForeignName = "";
                    var txt_foreign = oForm.Items.Item("44").Specific;
                    p_ForeignName = ((SAPbouiCOM.EditText)txt_foreign).Value;

                    //UoM Group
                    string p_UoMGrp = "";
                    var cbo_UoMGrp = oForm.Items.Item("10002056").Specific;
                    p_UoMGrp = ((SAPbouiCOM.ComboBox)cbo_UoMGrp).Selected.Value;
                    //Price List
                    string p_PriceList = "";
                    var cbo_p_PriceList = oForm.Items.Item("24").Specific;
                    p_PriceList = ((SAPbouiCOM.ComboBox)cbo_p_PriceList).Selected.Value;
                    //Origin
                    string p_Origin = "";
                    var txt_origin = oForm.Items.Item("U_origin").Specific;
                    p_Origin = ((SAPbouiCOM.EditText)txt_origin).Value;
                    //BarCode
                    string p_BarCode = "";
                    var txt_barcode = oForm.Items.Item("107").Specific;
                    p_BarCode = ((SAPbouiCOM.EditText)txt_barcode).Value;
                    //Unit Price Flag
                    string p_UnitPrice_Flag = "";
                    var cbo_p_UnitPrice_Flag = oForm.Items.Item("1470002295").Specific;
                    p_UnitPrice_Flag = ((SAPbouiCOM.ComboBox)cbo_p_UnitPrice_Flag).Selected.Value;
                    //Unit Price Value
                    string p_UnitPrice = "";
                    var txt_unitprice = oForm.Items.Item("34").Specific;
                    p_UnitPrice = ((SAPbouiCOM.EditText)txt_unitprice).Value;
                    //Stock Item
                    bool p_stock = false;
                    var chk_stockitem = oForm.Items.Item("14").Specific;
                    p_stock = ((SAPbouiCOM.CheckBox)chk_stockitem).Checked;
                    //Sales Item
                    bool p_sales = false;
                    var chk_salesitem = oForm.Items.Item("13").Specific;
                    p_sales = ((SAPbouiCOM.CheckBox)chk_salesitem).Checked;
                    //Purchase Item
                    bool p_purchase = false;
                    var chk_purchaseitem = oForm.Items.Item("12").Specific;
                    p_purchase = ((SAPbouiCOM.CheckBox)chk_purchaseitem).Checked;
                    //Virtual Item
                    bool p_virtual = false;
                    var chk_virtualitem = oForm.Items.Item("234000008").Specific;
                    p_virtual = ((SAPbouiCOM.CheckBox)chk_virtualitem).Checked;
                #endregion
                    #region General Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 6;

                    //Do Not Apply Discount Group Check
                    bool p_dnadg = false;
                    p_dnadg = ((SAPbouiCOM.CheckBox)oForm.Items.Item("1470002294").Specific).Checked;
                    //Manufacturer
                    string p_manufacturer = "";
                    p_manufacturer = ((SAPbouiCOM.ComboBox)oForm.Items.Item("114").Specific).Selected.Value;
                    //Additional Identifier
                    string p_additional_id = "";
                    p_additional_id = ((SAPbouiCOM.EditText)oForm.Items.Item("186").Specific).Value;
                    //Shipping Type
                    //string p_shipping_type = "";
                    //p_shipping_type = ((SAPbouiCOM.ComboBox)oForm.Items.Item("35").Specific).Selected.Value;
                    //Manage Item by
                    string p_manage_item_by = "";
                    p_manage_item_by = ((SAPbouiCOM.ComboBox)oForm.Items.Item("162").Specific).Selected.Value;
                    //Advanced Rule Type
                    string p_adv_rule_type = "";
                    p_adv_rule_type = ((SAPbouiCOM.ComboBox)oForm.Items.Item("1470002293").Specific).Selected.Value;
                    //Linked to Resource
                    string p_linked2resource = "";
                    p_linked2resource = ((SAPbouiCOM.EditText)oForm.Items.Item("254000002").Specific).Value;

                    oForm.Freeze(false);
                    #endregion
                    #region Purchasing Data Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 1;

                    //Preferred Supplier
                    string p_prefer_supplier = "";
                    p_prefer_supplier = ((SAPbouiCOM.EditText)oForm.Items.Item("16").Specific).Value;
                    //Mfr Catalogue No.
                    string p_mfr_catalog_no = "";
                    p_mfr_catalog_no = ((SAPbouiCOM.EditText)oForm.Items.Item("18").Specific).Value;
                    //Purchasing UoM Name
                    string p_purchasing_uom_name = "";
                    p_purchasing_uom_name = ((SAPbouiCOM.EditText)oForm.Items.Item("20").Specific).Value;
                    //Item per Purchase Unit
                    string p_purchasing_unit = "";
                    p_purchasing_unit = ((SAPbouiCOM.EditText)oForm.Items.Item("22").Specific).Value;
                    //Packaging UoM Name
                    string p_purchasing_package_name = "";
                    p_purchasing_package_name = ((SAPbouiCOM.EditText)oForm.Items.Item("151").Specific).Value;
                    //Quantity per Package
                    string p_quantuty_package = "";
                    p_quantuty_package = ((SAPbouiCOM.EditText)oForm.Items.Item("153").Specific).Value;
                    //Duty Group
                    //string p_duty_grp = "";
                    //p_duty_grp = ((SAPbouiCOM.ComboBox)oForm.Items.Item("117").Specific).Selected.Value;
                    //Vat Code
                    //string p_vat_code = "";
                    //p_vat_code = ((SAPbouiCOM.ComboBox)oForm.Items.Item("149").Specific).Selected.Value;
                    //Length
                    string p_length = "";
                    p_length = ((SAPbouiCOM.EditText)oForm.Items.Item("10").Specific).Value;
                    //Width
                    string p_width = "";
                    p_width = ((SAPbouiCOM.EditText)oForm.Items.Item("99").Specific).Value;
                    //Height
                    string p_height = "";
                    p_height = ((SAPbouiCOM.EditText)oForm.Items.Item("38").Specific).Value;
                    //Volume
                    string p_volume = "";
                    p_volume = ((SAPbouiCOM.EditText)oForm.Items.Item("37").Specific).Value;
                    //Unit
                    string p_unit = "";
                    p_unit = ((SAPbouiCOM.ComboBox)oForm.Items.Item("51").Specific).Selected.Value;
                    //Weight
                    string p_weight = "";
                    p_weight = ((SAPbouiCOM.EditText)oForm.Items.Item("47").Specific).Value;
                    //Factor 1
                    //string factor1 = "";
                    //factor1 = ((SAPbouiCOM.EditText)oForm.Items.Item("132").Specific).Value;
                    //Factor 2
                    //string factor2 = "";
                    //factor2 = ((SAPbouiCOM.EditText)oForm.Items.Item("134").Specific).Value;
                    //Factor 3
                    //string factor3 = "";
                    //factor3 = ((SAPbouiCOM.EditText)oForm.Items.Item("139").Specific).Value;
                    //Factor 4
                    //string factor4 = "";
                    //factor4 = ((SAPbouiCOM.EditText)oForm.Items.Item("141").Specific).Value;

                    oForm.Freeze(false);
                    #endregion
                    #region Sales Data Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 2;
                    //VAT Code
                    string p_st_vat_code = "";
                    p_st_vat_code = ((SAPbouiCOM.ComboBox)oForm.Items.Item("69").Specific).Selected.Value;
                    //Sales UoM Name
                    string p_st_sales_uom_name = "";
                    p_st_sales_uom_name = ((SAPbouiCOM.EditText)oForm.Items.Item("55").Specific).Value;
                    //Items per Sales Unit
                    string p_st_sales_per_unit = "";
                    p_st_sales_per_unit = ((SAPbouiCOM.EditText)oForm.Items.Item("57").Specific).Value;
                    //Packaing UoM Name
                    string p_st_pck_uom_name = "";
                    p_st_pck_uom_name = ((SAPbouiCOM.EditText)oForm.Items.Item("155").Specific).Value;
                    //Quantity per Package
                    string p_st_quantity_per_pck = "";
                    p_st_quantity_per_pck = ((SAPbouiCOM.EditText)oForm.Items.Item("157").Specific).Value;
                    //Length
                    string p_st_length = "";
                    p_st_length = ((SAPbouiCOM.EditText)oForm.Items.Item("54").Specific).Value;
                    //Width
                    string p_st_width = "";
                    p_st_width = ((SAPbouiCOM.EditText)oForm.Items.Item("61").Specific).Value;
                    //Height
                    string p_st_height = "";
                    p_st_height = ((SAPbouiCOM.EditText)oForm.Items.Item("67").Specific).Value;
                    //Volume
                    string p_st_volume = "";
                    p_st_volume = ((SAPbouiCOM.EditText)oForm.Items.Item("66").Specific).Value;
                    //Unit
                    string p_st_unit = "";
                    p_st_unit = ((SAPbouiCOM.ComboBox)oForm.Items.Item("75").Specific).Selected.Value;
                    //Weight
                    string p_st_weight = "";
                    p_st_weight = ((SAPbouiCOM.EditText)oForm.Items.Item("71").Specific).Value;
                    //Factor 1
                    //string p_st_factor1 = "";
                    //p_st_factor1 = ((SAPbouiCOM.EditText)oForm.Items.Item("137").Specific).Value;
                    //Factor 2
                    //string p_st_factor2 = "";
                    //p_st_factor2 = ((SAPbouiCOM.EditText)oForm.Items.Item("138").Specific).Value;
                    //Factor 3
                    //string p_st_factor3 = "";
                    //p_st_factor3 = ((SAPbouiCOM.EditText)oForm.Items.Item("143").Specific).Value;
                    //Factor 4
                    //string p_st_factor4 = "";
                    //p_st_factor4 = ((SAPbouiCOM.EditText)oForm.Items.Item("144").Specific).Value;

                    oForm.Freeze(false);
                    #endregion
                    #region Stock Data Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 3;

                    //Set Stock Method By
                    string p_sht_stock_method = "";
                    p_sht_stock_method = ((SAPbouiCOM.ComboBox)oForm.Items.Item("80").Specific).Selected.Value;
                    //UoM Name
                    string p_sht_uom_name = "";
                    p_sht_uom_name = ((SAPbouiCOM.EditText)oForm.Items.Item("251").Specific).Value;
                    //Weight
                    string p_sht_weight = "";
                    p_sht_weight = ((SAPbouiCOM.EditText)oForm.Items.Item("234000002").Specific).Value;
                    //Valuation Method
                    string p_sht_valuation_method = "";
                    p_sht_valuation_method = ((SAPbouiCOM.ComboBox)oForm.Items.Item("248").Specific).Selected.Value;
                    //Item Cost
                    string p_sht_itm_cost = "";
                    if (p_sht_valuation_method == "Standard")
                    {
                        p_sht_itm_cost = ((SAPbouiCOM.EditText)oForm.Items.Item("64").Specific).Value;
                    }
                    //Manage Stock by Warehouse
                    bool p_sht_manage_stock = false;
                    p_sht_manage_stock = ((SAPbouiCOM.CheckBox)oForm.Items.Item("83").Specific).Checked;
                    string p_sht_required = "";
                    string p_sht_minumum = "";
                    string p_sht_maximum = "";
                    if (p_sht_manage_stock == false)
                    {
                        //Required
                        p_sht_required = ((SAPbouiCOM.EditText)oForm.Items.Item("88").Specific).Value;
                        //Minimum
                        p_sht_minumum = ((SAPbouiCOM.EditText)oForm.Items.Item("90").Specific).Value;
                        //Maximum
                        p_sht_maximum = ((SAPbouiCOM.EditText)oForm.Items.Item("213").Specific).Value;
                    }

                    oForm.Freeze(false);
                    #endregion
                    #region Planning Data Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 7;

                    //Planning Method
                    string p_pt_method = "";
                    p_pt_method = ((SAPbouiCOM.ComboBox)oForm.Items.Item("73").Specific).Selected.Value;
                    //Procurement Method
                    string p_pt_procurement_method = "";
                    p_pt_procurement_method = ((SAPbouiCOM.ComboBox)oForm.Items.Item("49").Specific).Selected.Value;
                    //Order Interval
                    string p_pt_order_interval = "";
                    p_pt_order_interval = ((SAPbouiCOM.ComboBox)oForm.Items.Item("36").Specific).Selected.Value;
                    //Order Multiple
                    string p_pt_order_multiple = "";
                    p_pt_order_multiple = ((SAPbouiCOM.EditText)oForm.Items.Item("97").Specific).Value;
                    //Minimum Order Qty
                    string p_pt_min_ord_qty = "";
                    p_pt_min_ord_qty = ((SAPbouiCOM.EditText)oForm.Items.Item("104").Specific).Value;
                    //Lead Time
                    string p_pt_leadtime = "";
                    p_pt_leadtime = ((SAPbouiCOM.EditText)oForm.Items.Item("123").Specific).Value;
                    //Tolerance Days
                    string p_pt_tolerance = "";
                    p_pt_tolerance = ((SAPbouiCOM.EditText)oForm.Items.Item("1320002074").Specific).Value;

                    oForm.Freeze(false);
                    #endregion
                    #region Production Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 15;

                    //Issue Method
                    string p_prt_issue_method = "";
                    p_prt_issue_method = ((SAPbouiCOM.ComboBox)oForm.Items.Item("249").Specific).Selected.Value;
                    //Production Std Cost
                    string p_prt_std_cost = "";
                    p_prt_std_cost = ((SAPbouiCOM.EditText)oForm.Items.Item("1880000011").Specific).Value;

                    oForm.Freeze(false);
                    #endregion
                    #region Remarks Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 5;

                    //Remarks
                    string p_remarks = "";
                    p_remarks = ((SAPbouiCOM.EditText)oForm.Items.Item("79").Specific).Value;

                    oForm.Freeze(false);
                    #endregion
                    //Active Form Inventory - Mode Add 
                    Application.SBO_Application.ActivateMenuItem("3073");
                    Application.SBO_Application.ActivateMenuItem("1282");
                    //Fill into Inventory Form
                    oForm = Application.SBO_Application.Forms.ActiveForm;
                    #region Header
                    //Item No
                    ((SAPbouiCOM.EditText)oForm.Items.Item("5").Specific).Value = "Z" + p_ItemNo;
                    //Description
                    ((SAPbouiCOM.EditText)oForm.Items.Item("7").Specific).Value = p_Description;
                    //Foreign Name
                    ((SAPbouiCOM.EditText)oForm.Items.Item("44").Specific).Value = p_ForeignName;
                    //Item Type
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("214").Specific).Select("I");
                    //Item Group
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("39").Specific).Select(p_ItemGrp);
                    //UoM Group
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("10002056").Specific).Select(p_UoMGrp);
                    //Price List
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("24").Specific).Select(p_PriceList);
                    //Origin
                    ((SAPbouiCOM.EditText)oForm.Items.Item("U_origin").Specific).Value = p_Origin;
                    //BarCode
                    ((SAPbouiCOM.EditText)oForm.Items.Item("107").Specific).Value = p_BarCode;
                    //Unit Price Flag
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("1470002295").Specific).Select(p_UnitPrice_Flag);
                    //Unit Price Value
                    ((SAPbouiCOM.EditText)oForm.Items.Item("34").Specific).Value = p_UnitPrice;
                    //Stock Item
                    ((SAPbouiCOM.CheckBox)oForm.Items.Item("14").Specific).Checked = true;
                    //Sales Item
                    ((SAPbouiCOM.CheckBox)oForm.Items.Item("13").Specific).Checked = true;
                    //Purchase Item
                    ((SAPbouiCOM.CheckBox)oForm.Items.Item("12").Specific).Checked = true;
                    //Fixed Asset Link
                    //((SAPbouiCOM.EditText)oForm.Items.Item("U_FA").Specific).Value = p_ItemNo;
                    #endregion
                    #region General Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 6;
                    //Do not Apply Discount Group
                    ((SAPbouiCOM.CheckBox)oForm.Items.Item("1470002294").Specific).Checked = p_dnadg;
                    //Manufacturer
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("114").Specific).Select(p_manufacturer);
                    //Additional Identifier
                    ((SAPbouiCOM.EditText)oForm.Items.Item("186").Specific).Value = p_additional_id;
                    //Shipping Type
                    //((SAPbouiCOM.ComboBox)oForm.Items.Item("35").Specific).Select(p_shipping_type);
                    //Manage Item by
                    //((SAPbouiCOM.ComboBox)oForm.Items.Item("162").Specific).Select(p_manage_item_by);
                    //Advanced Rule Type
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("1470002293").Specific).Select(p_adv_rule_type);
                    oForm.Freeze(false);
                    #endregion
                    #region Purchasing Data Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 1;

                    //Preferred Supplier
                    ((SAPbouiCOM.EditText)oForm.Items.Item("16").Specific).Value = p_prefer_supplier;
                    //Mfr Catalogue No.
                    ((SAPbouiCOM.EditText)oForm.Items.Item("18").Specific).Value = p_mfr_catalog_no;
                    if (p_UoMGrp == "-1")
                    {
                        //Purchasing UoM Name
                        ((SAPbouiCOM.EditText)oForm.Items.Item("20").Specific).Value = p_purchasing_uom_name;
                        //Item per Purchase Unit
                        ((SAPbouiCOM.EditText)oForm.Items.Item("22").Specific).Value = p_purchasing_unit;
                        //Packaging UoM Name
                        ((SAPbouiCOM.EditText)oForm.Items.Item("151").Specific).Value = p_purchasing_package_name;
                        //Quantity per Package
                        ((SAPbouiCOM.EditText)oForm.Items.Item("153").Specific).Value = p_quantuty_package;
                        //Factor 1
                        //((SAPbouiCOM.EditText)oForm.Items.Item("132").Specific).Value = factor1;
                        //Factor 2
                        //((SAPbouiCOM.EditText)oForm.Items.Item("134").Specific).Value = factor2;
                        //Factor 3
                        //((SAPbouiCOM.EditText)oForm.Items.Item("139").Specific).Value = factor3;
                        //Factor 4
                        //((SAPbouiCOM.EditText)oForm.Items.Item("141").Specific).Value = factor4;

                    }

                    //Duty Group
                    //((SAPbouiCOM.ComboBox)oForm.Items.Item("117").Specific).Select(p_duty_grp);
                    //Vat Code
                    //((SAPbouiCOM.ComboBox)oForm.Items.Item("149").Specific).Select(p_vat_code);
                    //Length
                    ((SAPbouiCOM.EditText)oForm.Items.Item("10").Specific).Value = p_length;
                    //Width
                    ((SAPbouiCOM.EditText)oForm.Items.Item("99").Specific).Value = p_width;
                    //Height
                    ((SAPbouiCOM.EditText)oForm.Items.Item("38").Specific).Value = p_height;
                    //Volume
                    ((SAPbouiCOM.EditText)oForm.Items.Item("37").Specific).Value = p_volume;
                    //Unit
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("51").Specific).Select(p_unit);
                    //Weight
                    ((SAPbouiCOM.EditText)oForm.Items.Item("47").Specific).Value = p_weight;

                    oForm.Freeze(false);
                    #endregion
                    #region Sales Data Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 2;
                    //VAT Code
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("69").Specific).Select(p_st_vat_code);
                    if (p_UoMGrp == "-1")
                    {
                        //Sales UoM Name
                        ((SAPbouiCOM.EditText)oForm.Items.Item("55").Specific).Value = p_st_sales_uom_name;
                        //Items per Sales Unit
                        ((SAPbouiCOM.EditText)oForm.Items.Item("57").Specific).Value = p_st_sales_per_unit;
                        //Packaing UoM Name
                        ((SAPbouiCOM.EditText)oForm.Items.Item("155").Specific).Value = p_st_pck_uom_name;
                        //Quantity per Package
                        ((SAPbouiCOM.EditText)oForm.Items.Item("157").Specific).Value = p_st_quantity_per_pck;
                        //Factor 1
                        //((SAPbouiCOM.EditText)oForm.Items.Item("137").Specific).Value = p_st_factor1;
                        //Factor 2
                        //((SAPbouiCOM.EditText)oForm.Items.Item("138").Specific).Value = p_st_factor2;
                        //Factor 3
                        //((SAPbouiCOM.EditText)oForm.Items.Item("143").Specific).Value = p_st_factor3;
                        //Factor 4
                        //((SAPbouiCOM.EditText)oForm.Items.Item("144").Specific).Value = p_st_factor4;
                    }
                    //Length
                    ((SAPbouiCOM.EditText)oForm.Items.Item("54").Specific).Value = p_st_length;
                    //Width
                    ((SAPbouiCOM.EditText)oForm.Items.Item("61").Specific).Value = p_st_width;
                    //Height
                    ((SAPbouiCOM.EditText)oForm.Items.Item("67").Specific).Value = p_st_height;
                    //Volume
                    ((SAPbouiCOM.EditText)oForm.Items.Item("66").Specific).Value = p_st_volume;
                    //Unit
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("75").Specific).Select(p_st_unit);
                    //Weight
                    ((SAPbouiCOM.EditText)oForm.Items.Item("71").Specific).Value = p_st_weight;

                    oForm.Freeze(false);
                    #endregion
                    #region Stock Data Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 3;

                    //Set Stock Method By
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("80").Specific).Select(p_sht_stock_method);
                    if (p_UoMGrp == "-1")
                    {
                        //UoM Name
                        ((SAPbouiCOM.EditText)oForm.Items.Item("251").Specific).Value = p_sht_uom_name;
                    }
                    //Weight
                    ((SAPbouiCOM.EditText)oForm.Items.Item("234000002").Specific).Value = p_sht_weight;
                    //Valuation Method
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("248").Specific).Select(p_sht_valuation_method);
                    //Item Cost
                    if (p_sht_valuation_method == "Standard")
                    {
                        ((SAPbouiCOM.EditText)oForm.Items.Item("64").Specific).Value = p_sht_itm_cost;
                    }
                    //Manage Stock by Warehouse
                    ((SAPbouiCOM.CheckBox)oForm.Items.Item("83").Specific).Checked = p_sht_manage_stock;
                    if (p_sht_manage_stock == false)
                    {
                        //Required
                        ((SAPbouiCOM.EditText)oForm.Items.Item("88").Specific).Value = p_sht_required;
                        //Minimum
                        ((SAPbouiCOM.EditText)oForm.Items.Item("90").Specific).Value = p_sht_minumum;
                        //Maximum
                        ((SAPbouiCOM.EditText)oForm.Items.Item("213").Specific).Value = p_sht_maximum;
                    }

                    oForm.Freeze(false);
                    #endregion
                    #region Planning Data Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 7;

                    //Planning Method
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("73").Specific).Select(p_pt_method);
                    //Procurement Method
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("49").Specific).Select(p_pt_procurement_method);
                    //Order Interval
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("36").Specific).Select(p_pt_order_interval);
                    //Order Multiple
                    ((SAPbouiCOM.EditText)oForm.Items.Item("97").Specific).Value = p_pt_order_multiple;
                    //Minimum Order Qty
                    ((SAPbouiCOM.EditText)oForm.Items.Item("104").Specific).Value = p_pt_min_ord_qty;
                    //Lead Time
                    ((SAPbouiCOM.EditText)oForm.Items.Item("123").Specific).Value = p_pt_leadtime;
                    //Tolerance Days
                    ((SAPbouiCOM.EditText)oForm.Items.Item("1320002074").Specific).Value = p_pt_tolerance;

                    oForm.Freeze(false);
                    #endregion
                    #region Production Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 15;

                    //Issue Method
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("249").Specific).Select(p_prt_issue_method);
                    //Production Std Cost
                    ((SAPbouiCOM.EditText)oForm.Items.Item("1880000011").Specific).Value = p_prt_std_cost;

                    oForm.Freeze(false);
                    #endregion
                    #region Remarks Tab
                    oForm.Freeze(true);
                    oForm.PaneLevel = 5;

                    //Remarks
                    ((SAPbouiCOM.EditText)oForm.Items.Item("79").Specific).Value = p_remarks;
                    //U_FA
                    ((SAPbouiCOM.EditText)oForm.Items.Item("U_FA").Specific).Value = p_ItemNo;

                    oForm.Freeze(false);
                    #endregion
                }
                else
                {
                    Application.SBO_Application.MessageBox("Only support for Item Group: 103|Thiết bị - Tài sản; 105|Tài sản công trường.");
                }
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                Application.SBO_Application.MessageBox(ex.Message);
            }
        }
    }
}
