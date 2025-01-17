
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace PurchaseBlanketAgreement
{

    [FormAttribute("1250000102", "Purchase Blanket Agreement - UI Edit Mode.b1f")]
    class Purchase_Blanket_Agreement___UI_Edit_Mode : SystemFormBase
    {
        public Purchase_Blanket_Agreement___UI_Edit_Mode()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
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
            if (pVal.FormMode == 1)
            {
                //Get Blanket No
                string p_BA_No = ((SAPbouiCOM.EditText)this.GetItem("1250000004").Specific).Value;
                PurchaseBlanketAgreement.frm_Approve frm = new PurchaseBlanketAgreement.frm_Approve(p_BA_No, "S", Application.SBO_Application.Forms.ActiveForm);
                frm.Show();
            }
            else
            {
                Application.SBO_Application.MessageBox("Approve function works on View Mode !");
            }

        }
    }
}
