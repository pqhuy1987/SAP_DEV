using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data;

namespace U_DUTRU
{
    [FormAttribute("UDO_FT_DUTRU")]
    class DUTRU_NEW : UDOFormBase
    {
        public DUTRU_NEW()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_sall").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("21_U_Cb").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Button Button0;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
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
                        int split = int.Parse(((SAPbouiCOM.EditText)oMtx_S.Columns.Item("C_0_15").Cells.Item(i + 1).Specific).Value);
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
                        //if (i < Sum_DUTRU.Rows.Count - 1)
                        //{
                        //    oMtx.AddRow();
                        //    ((SAPbouiCOM.EditText)oMtx.Columns.Item(1).Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        //}
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

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.ComboBox ComboBox0;

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                this.UIAPIRawForm.PaneLevel = 1;
                SAPbouiCOM.Matrix oMtx_S = ((SAPbouiCOM.Matrix)this.GetItem("0_U_G").Specific);
                Visible_Column_Matrix(oMtx_S,int.Parse(this.ComboBox0.Selected.Value));
            }
            catch
            {

            }
            finally
            {
 
            }

        }
        private void Visible_Column_Matrix(SAPbouiCOM.Matrix pMatrixItem, int pType)
        {
            try
            {
                foreach (SAPbouiCOM.Column c in pMatrixItem.Columns)
                {
                    if (pType == 1)
                    {
                        //CCM
                        if (c.Title == "Chi phí Mua bán" || c.Title == "Chi phí Thuê" || c.Title == "Chi phí Vận hành")
                        {
                            if (c.Visible == true)
                                c.Visible = false;
                        }
                        else
                        {
                            if (c.Visible == false)
                                c.Visible = true;
                        }
                    }
                    else if (pType == 2)
                    {
                        //DT
                        if (c.Title == "Chi phí Mua bán" || c.Title == "Chi phí Thuê" || c.Title == "Chi phí Vận hành" || c.Title == "LineId"
                            || c.Title == "Mã hạng mục" || c.Title == "Tên hạng mục" || c.Title == "Chi phí Vận chuyển" || c.Title == "Split to")
                        {
                            if (c.Visible == false)
                                c.Visible = true;
                        }
                        else
                        {
                            if (c.Visible == true)
                                c.Visible = false;
                        }
                    }
                }
            }
            catch
            { }
        }
    }
}
