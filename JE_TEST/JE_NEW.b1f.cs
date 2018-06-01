using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace JE_TEST
{
    [FormAttribute("393", "JE_NEW.b1f")]
    class JE_NEW : SystemFormBase
    {
        public JE_NEW()
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
            this.DeactivateAfter += new DeactivateAfterHandler(this.Form_DeactivateAfter);

        }

        private void Form_DeactivateAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
            try
            {
                if (oForm.Title == "System Message")
                {
                    if (((SAPbouiCOM.StaticText)oForm.Items.Item(7).Specific).Caption == "Base amount changed manually. Continue?")
                    {
                        oForm.Items.Item(0).Click();
                    }
                }
            }
            catch
            { }
        }

        private void OnCustomInitialize()
        {

        }
    }
}
