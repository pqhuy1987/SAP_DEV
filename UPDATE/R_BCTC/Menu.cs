using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM.Framework;

namespace R_BCTC
{
    class Menu
    {
        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.Menus oMenuFolder = null;
            SAPbouiCOM.MenuItem oMenuItem = null;
            SAPbouiCOM.MenuItem oMenusFolderItem = null;
            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "RMS";
            oCreationPackage.String = "Excel Report System";
            oCreationPackage.Image = System.Windows.Forms.Application.StartupPath + "\\Excel_icon.png";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;

            oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception e)
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("RMS");
                oMenus = oMenuItem.SubMenus;
                //Create Folder
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "RMS.MM_FI_FOL";
                oCreationPackage.String = "Báo cáo tài chính";
                oMenus.AddEx(oCreationPackage);

                oMenusFolderItem = Application.SBO_Application.Menus.Item("RMS.MM_FI_FOL");
                oMenuFolder = oMenusFolderItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.MM_FI";
                oCreationPackage.String = "Báo cáo tài chính theo đối tượng";
                oMenuFolder.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.MM_CE";
                oCreationPackage.String = "Báo cáo dự trù tài chính";
                oMenuFolder.AddEx(oCreationPackage);

                //Create Folder
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "RMS.EQ_FOL";
                oCreationPackage.String = "Báo cáo thiết bị";
                oMenus.AddEx(oCreationPackage);

                oMenusFolderItem = Application.SBO_Application.Menus.Item("RMS.EQ_FOL");
                oMenuFolder = oMenusFolderItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.EQ_CE";
                oCreationPackage.String = "Báo cáo dự trù thiết bị theo công tác";
                oMenuFolder.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.EQ_CE_O";
                oCreationPackage.String = "Báo cáo dự trù thiết bị theo đối tượng";
                oMenuFolder.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.EQ_FR";
                oCreationPackage.String = "Báo cáo tài chính thiết bị";
                oMenuFolder.AddEx(oCreationPackage);

                //Create Folder
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "RMS.AC_FOL";
                oCreationPackage.String = "Báo cáo kế toán";
                oMenus.AddEx(oCreationPackage);

                oMenusFolderItem = Application.SBO_Application.Menus.Item("RMS.AC_FOL");
                oMenuFolder = oMenusFolderItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.AC_BALANCE_SHEET";
                oCreationPackage.String = "Bảng cân đối tài khoản";
                oMenuFolder.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.AC_SP_JV";
                oCreationPackage.String = "SupportJV";
                oMenuFolder.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.MM_ITEM_CON";
                oCreationPackage.String = "Báo cáo kiểm soát khối lượng";
                oMenus.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.CCM_SUMMARY_BILL";
                oCreationPackage.String = "Tổng hợp thanh toán kỳ";
                oMenus.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.CCM_SUMMARY_HD";
                oCreationPackage.String = "Thống kê duyệt hợp đồng";
                oMenus.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.CCM_APPROVE_BILL";
                oCreationPackage.String = "Thống kê duyệt khối lượng thanh toán";
                oMenus.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.CCM_HAOHUT_THEP";
                oCreationPackage.String = "Báo cáo hao hụt thép";
                oMenus.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.CCM_THEODOI_VP";
                oCreationPackage.String = "Báo cáo theo dõi chi tiết chi phí văn phòng";
                oMenus.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.CCM_TONGHOP_VP";
                oCreationPackage.String = "Báo cáo theo dõi tổng hợp chi phí văn phòng";
                oMenus.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.BASELINE_MAN";
                oCreationPackage.String = "BaseLine Management";
                oMenus.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.CCM_DT_LN_LIST";
                oCreationPackage.String = "Báo cáo thống kê doanh thu lợi nhuận";
                oMenus.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.CCM_DT_LN";
                oCreationPackage.String = "Báo cáo tổng hợp doanh thu lợi nhuận";
                oMenus.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.CCM_TK_SL";
                oCreationPackage.String = "Báo cáo thống kê sản lượng";
                oMenus.AddEx(oCreationPackage);

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "RMS.CCM_HQDP";
                oCreationPackage.String = "Báo cáo hiệu quả đàm phán";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception er)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.MM_FI")
                {
                    MM_FI activeForm = new MM_FI();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.MM_CE")
                {
                    MM_CE activeForm = new MM_CE();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.EQ_CE")
                {
                    EQ_CE activeForm = new EQ_CE();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.EQ_CE_O")
                {
                    EQ_CE_O activeForm = new EQ_CE_O();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.EQ_FR")
                {
                    EQ_FR activeForm = new EQ_FR();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.MM_ITEM_CON")
                {
                    MM_ITEM_CON activeForm = new MM_ITEM_CON();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.AC_BALANCE_SHEET")
                {
                    AC_BALANCE_SHEET activeForm = new AC_BALANCE_SHEET();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.CCM_SUMMARY_BILL")
                {
                    CCM_SUMMARY_BILL activeForm = new CCM_SUMMARY_BILL();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.CCM_SUMMARY_HD")
                {
                    CCM_SUMMARY_HD activeForm = new CCM_SUMMARY_HD();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.CCM_APPROVE_BILL")
                {
                    CCM_APPROVE_BILL activeForm = new CCM_APPROVE_BILL();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.CCM_HAOHUT_THEP")
                {
                    CCM_HAOHUT_THEP activeForm = new CCM_HAOHUT_THEP();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.CCM_THEODOI_VP")
                {
                    CCM_THEODOI_VP activeForm = new CCM_THEODOI_VP();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.CCM_TONGHOP_VP")
                {
                    CCM_TONGHOP_VP activeForm = new CCM_TONGHOP_VP();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.BASELINE_MAN")
                {
                    BASELINE_MAN activeForm = new BASELINE_MAN();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.CCM_DT_LN_LIST")
                {
                    CCM_DT_LN_LIST activeForm = new CCM_DT_LN_LIST();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.CCM_DT_LN")
                {
                    CCM_DT_LN activeForm = new CCM_DT_LN();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.CCM_TK_SL")
                {
                    CCM_SANLUONG activeForm = new CCM_SANLUONG();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.CCM_HQDP")
                {
                    CCM_HQDP activeForm = new CCM_HQDP();
                    activeForm.Show();
                }
                if (pVal.BeforeAction && pVal.MenuUID == "RMS.AC_SP_JV")
                {
                    AC_SP_JV activeForm = new AC_SP_JV();
                    activeForm.Show();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

    }
}
