using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace PurchaseBlanketAgreement
{
    [FormAttribute("PurchaseBlanketAgreement.frm_Approve_List", "frm_Approve_List.b1f")]
    class frm_Approve_List : UserFormBase
    {
        public frm_Approve_List()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_0").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_1").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_3").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_5").Specific));
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

        }

        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;
    }
}
