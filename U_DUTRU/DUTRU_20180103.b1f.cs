using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data;

namespace U_DUTRU
{
    [FormAttribute("UDO_FT_DUTRU")]
    class DUTRU_20180103 : UDOFormBase
    {
        public DUTRU_20180103()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_s1").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("bt_sall").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("1_U_G").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("bt_addl").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_addl").Specific));
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

        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.EditText EditText0;

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            DataTable tb_s = new DataTable();
            tb_s.Columns.Add("LineID", typeof(Int32));
            tb_s.Columns.Add("SubProject", typeof(String));
            tb_s.Columns.Add("SubProjectName", typeof(String));
            tb_s.Columns.Add("Split", typeof(Int32));
            try
            {
                this.UIAPIRawForm.Freeze(true);
                this.UIAPIRawForm.PaneLevel = 1;
                SAPbouiCOM.Matrix oMtx_S = ((SAPbouiCOM.Matrix)this.GetItem("0_U_G").Specific);
                if (oMtx_S.RowCount > 0)
                {
                    for (int i = 0; i < oMtx_S.RowCount; i++)
                    {
                        int lineid = int.Parse(((SAPbouiCOM.EditText)oMtx_S.Columns.Item("C_0_1").Cells.Item(i + 1).Specific).Value);
                        string subproject = ((SAPbouiCOM.EditText)oMtx_S.Columns.Item("C_0_2").Cells.Item(i + 1).Specific).Value;
                        string subproject_name = ((SAPbouiCOM.EditText)oMtx_S.Columns.Item("C_0_3").Cells.Item(i + 1).Specific).Value;
                        int split = int.Parse(((SAPbouiCOM.EditText)oMtx_S.Columns.Item("C_0_18").Cells.Item(i + 1).Specific).Value);
                        if (!string.IsNullOrEmpty(subproject))
                        {
                            DataRow r = tb_s.NewRow();
                            r["LineID"] = lineid;
                            r["SubProject"] = subproject;
                            r["SubProjectName"] = subproject_name;
                            r["Split"] = split;
                            tb_s.Rows.Add(r);
                        }
                    }

                    //Split to Details
                    this.UIAPIRawForm.PaneLevel = 2;
                    SAPbouiCOM.Matrix oMtx_D = ((SAPbouiCOM.Matrix)this.GetItem("1_U_G").Specific);
                    int Total_Line = int.Parse(tb_s.Compute("SUM(Split)", "").ToString());
                    int MaxLine = 0;
                    int.TryParse(((SAPbouiCOM.EditText)oMtx_D.Columns.Item("C_1_1").Cells.Item(oMtx_D.RowCount).Specific).Value, out MaxLine);
                    int LineID = 1;
                    if (MaxLine > 0 && this.UIAPIRawForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        oMtx_D.AddRow();
                        ((SAPbouiCOM.EditText)oMtx_D.Columns.Item("C_1_1").Cells.Item(MaxLine + LineID).Specific).Value = (MaxLine + LineID).ToString();
                    }
                    else
                    {
                        MaxLine = 0;
                    }
                    foreach (DataRow t in tb_s.Rows)
                    {
                        int split2 = (int)t["Split"];
                        for (int j = 0; j < split2; j++)
                        {
                            //Sum LINE ID
                            ((SAPbouiCOM.EditText)oMtx_D.Columns.Item("C_1_2").Cells.Item(MaxLine + LineID).Specific).Value = t["LineID"].ToString();
                            //Subproject Code
                            ((SAPbouiCOM.EditText)oMtx_D.Columns.Item("C_1_3").Cells.Item(MaxLine + LineID).Specific).Value = t["SubProject"].ToString();
                            //Subproject Description
                            ((SAPbouiCOM.EditText)oMtx_D.Columns.Item("C_1_4").Cells.Item(MaxLine + LineID).Specific).Value = t["SubProjectName"].ToString();
                            LineID++;
                            if (LineID < Total_Line + 1)
                            {
                                oMtx_D.AddRow();
                                ((SAPbouiCOM.EditText)oMtx_D.Columns.Item("C_1_1").Cells.Item(MaxLine + LineID).Specific).Value = (MaxLine + LineID).ToString();
                            }
                        }
                    }
                }
                this.UIAPIRawForm.Freeze(false);
            }
            catch (Exception ex)
            {
                this.UIAPIRawForm.Freeze(false);
                Application.SBO_Application.MessageBox(ex.Message);
            }

        }

        private void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            int LineNumber = 0;
            int.TryParse(this.EditText0.Value.ToString(), out LineNumber);
            if (LineNumber > 0)
            {
                //Add Line
                this.UIAPIRawForm.PaneLevel = 2;
                SAPbouiCOM.Matrix oMtx_D = ((SAPbouiCOM.Matrix)this.GetItem("1_U_G").Specific);
                int MaxLine = 0;
                int.TryParse(((SAPbouiCOM.EditText)oMtx_D.Columns.Item("C_1_1").Cells.Item(oMtx_D.RowCount).Specific).Value, out MaxLine);
                for (int i = 0; i < LineNumber; i++)
                {
                    if (MaxLine + i + 1 < MaxLine + LineNumber + 1)
                    {
                        oMtx_D.AddRow();
                        ((SAPbouiCOM.EditText)oMtx_D.Columns.Item("C_1_1").Cells.Item(MaxLine + i + 1).Specific).Value = (MaxLine + i + 1).ToString();
                    }
                }

            }

        }
    }
}
