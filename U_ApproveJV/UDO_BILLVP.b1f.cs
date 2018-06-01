using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace U_ApproveJV
{
    [FormAttribute("UDO_FT_BILLVP")]
    class UDO_BILLVP : UDOFormBase
    {
        public UDO_BILLVP()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("0_U_G").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Matrix Matrix0;

        private void OnCustomInitialize()
        {
            foreach (SAPbouiCOM.Column c in Matrix0.Columns)
            {
                if (c.Title.Trim() =="Giá trị (bao gồm VAT)")
                    c.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            }
           
        }
    }
}
