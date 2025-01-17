
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace S_FixedAsset
{

    [FormAttribute("1250000100", "Sales Blanket Agreement.b1f")]
    class Sales_Blanket_Agreement : SystemFormBase
    {
        public Sales_Blanket_Agreement()
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
            this.Button0.Item.Top = this.GetItem("1250000002").Top;
            this.Button0.Item.Left = this.GetItem("1250000002").Left + this.GetItem("1250000002").Width + (this.GetItem("1250000002").Left - this.GetItem("1250000001").Left - this.GetItem("1250000001").Width);
            this.Button0.Item.Width = this.GetItem("1250000002").Width;
            this.Button0.Item.Height = this.GetItem("1250000002").Height;
        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.FormMode == 1)
            {
                //Get Blanket No
                string p_BA_No = ((SAPbouiCOM.EditText)this.GetItem("1250000004").Specific).Value;
                Blanket_Agreement_Approve.frm_Approve frm = new Blanket_Agreement_Approve.frm_Approve(p_BA_No, "C", Application.SBO_Application.Forms.ActiveForm);
                frm.Show();
            }
            else
            {
                Application.SBO_Application.MessageBox("Approve function works on View Mode !");
            }
        }
    }
}
