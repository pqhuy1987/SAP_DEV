using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;

namespace R_BCTC
{
    [FormAttribute("R_BCTC.BASELINE", "BASELINE.b1f")]
    class BASELINE : UserFormBase
    {
        public BASELINE()
        {
        }

        public BASELINE(string pFProject)
        {
            FProject = pFProject;
            this.EditText0.Value = FProject;
            this.EditText1.Value = DateTime.Today.ToString("yyyyMMdd");
            this.EditText2.Item.Click();
            this.EditText0.Item.Enabled = false;
            this.EditText1.Item.Enabled = false;
        }
        string FProject = "";
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_fprj").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txt_dt").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("txt_note").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_7").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {
            this.oApp = (SAPbouiCOM.Application)SAPbouiCOM.Framework.Application.SBO_Application;
            this.oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            //Create Connection SQL
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery("Select * from [@ADDONCFG]");
            if (oR_RecordSet.RecordCount > 0)
            {
                string uid = oR_RecordSet.Fields.Item("Code").Value.ToString();
                string pwd = oR_RecordSet.Fields.Item("Name").Value.ToString();
                conn = new SqlConnection(string.Format("Data Source={0}; Initial Catalog={1}; User id={2}; Password={3};", oCompany.Server, oCompany.CompanyDB, uid, pwd));
            }
            else
            {
                oApp.MessageBox("Can't connect DB !");
            }
        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.Button Button0;

        System.Data.DataTable Get_Approve_Process_BASELINE()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_GetList_Approve", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Usr", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }
            return result;
        }
        void Add_Data(int pDocEntry_BaseLine)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BASELINE_Add_Data", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@BaseLine_DocEntry", pDocEntry_BaseLine);
                conn.Open();
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }
        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService sCmp = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralData oChild = null;
            SAPbobsCOM.GeneralDataCollection oChildren = null;

            sCmp = oCompany.GetCompanyService();
            oGeneralService = sCmp.GetGeneralService("BaseLine");
            //Create data for new row in main UDO
            oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
            //Financial Project
            oGeneralData.SetProperty("U_FProject", FProject);
            //Date
            oGeneralData.SetProperty("U_BaseDate", DateTime.Today);
            //Note
            oGeneralData.SetProperty("U_Note", EditText2.Value.Trim());
            //Add Quy trinh duyet
            DataTable rs = Get_Approve_Process_BASELINE();
            if (rs.Rows.Count > 0)
            {
                oChildren = oGeneralData.Child("BASELINE_APPR");
                foreach (DataRow r in rs.Rows)
                {
                    oChild = oChildren.Add();
                    oChild.SetProperty("U_Level", r["LEVEL"].ToString());
                    oChild.SetProperty("U_Posistion", r["Position"].ToString());
                    oChild.SetProperty("U_DeptName", r["DeptName"].ToString());
                    oChild.SetProperty("U_PosName", r["PosName"].ToString());
                }
            }
            oGeneralParams = oGeneralService.Add(oGeneralData);
            
            if (!string.IsNullOrEmpty(oGeneralParams.GetProperty("DocEntry").ToString()))
            {
                int BaseLine_DocEntry = 0;
                int.TryParse(oGeneralParams.GetProperty("DocEntry").ToString(),out BaseLine_DocEntry);
                //Add BaseLine Data
                Add_Data(BaseLine_DocEntry);
                oApp.SetStatusBarMessage("Add Completed - DocEntry: " + oGeneralParams.GetProperty("DocEntry").ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                this.UIAPIRawForm.Close();
            }

        }

    }
}
