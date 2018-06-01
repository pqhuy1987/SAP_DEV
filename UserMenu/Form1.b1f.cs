using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;

namespace UserMenu
{
    [FormAttribute("UserMenu.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
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
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
            SAPbobsCOM.CompanyService oCompServ = (SAPbobsCOM.CompanyService)oCompany.GetCompanyService();
            SAPbobsCOM.UserMenuService oUmServ= (SAPbobsCOM.UserMenuService)oCompServ.GetBusinessService(SAPbobsCOM.ServiceTypes.UserMenuService);

            //SAPbobsCOM.UserMenuItems oUserMenuItems = oUmServ.GetCurrentUserMenu();
            //SAPbobsCOM.UserMenuItem oUserMenuItem = oUserMenuItems.Item(1);

            SAPbobsCOM.UserMenuParams oUmPara = (SAPbobsCOM.UserMenuParams)oUmServ.GetDataInterface(SAPbobsCOM.UserMenuServiceDataInterfaces.umsdiUserMenuParams);
            oUmPara.UserID = 1;
            SAPbobsCOM.UserMenuItems oUserMenuItems = oUmServ.GetUserMenu(oUmPara);



        }

        private void OnCustomInitialize()
        {

        }
    }
}