using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;

namespace Excel_On_Web
{
    [FormAttribute("Excel_On_Web.Form1", "Form1.b1f")]
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
            //builder.CreateFile("xlsx");
            //a.CreateFile(

        }

        private void OnCustomInitialize()
        {

        }
    }
}