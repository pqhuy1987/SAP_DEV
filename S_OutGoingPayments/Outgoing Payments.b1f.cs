
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace S_FixedAsset
{

    [FormAttribute("426", "Outgoing Payments.b1f")]
    class Outgoing_Payments : SystemFormBase
    {
        public Outgoing_Payments()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("U_TYPE").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.DataLoadAfter += new SAPbouiCOM.Framework.FormBase.DataLoadAfterHandler(this.Form_DataLoadAfter);
            this.DataUpdateBefore += new DataUpdateBeforeHandler(this.Form_DataUpdateBefore);

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            Application.SBO_Application.MessageBox("Form Loaded !");
            //Application.SBO_Application.Company.UserName;
            if (pVal.FormMode == 3)
            {
                this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("U_TYPE").Specific));
                this.ComboBox0.Select("VP");
                this.GetItem("5").Click();
                this.GetItem("U_TYPE").Enabled = false;
                //SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.;
                //SAPbouiCOM.ComboBox cbo_ptype = ((SAPbouiCOM.ComboBox)oForm.Items.Item("U_TYPE").Specific);
                
            }

        }

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.ComboBox ComboBox0;

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();

        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            //throw new System.NotImplementedException();
            this.GetItem("U_TYPE").Enabled = false;

        }

        private void Form_DataUpdateBefore(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            throw new System.NotImplementedException();

        }
    }
}
