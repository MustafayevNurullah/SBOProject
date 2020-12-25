using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using SAPbouiCOM;
using System.Windows.Forms;
using SBOProject;
using SAPbobsCOM;
using System.Threading.Tasks;

namespace SBOProject
{
    [FormAttribute("SBOProject.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {

        }
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        /// 
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Connection").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Grid3 = ((SAPbouiCOM.Grid)(this.GetItem("Item_4").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_0").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.OnCustomInitialize();

        }
        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.DataTable dt;
        private Grid Grid3;

         string Query="CALL \"BMS::ALL_USER\"";
        void Grid_Data_Write()
        {
            oForm = (SAPbouiCOM.Form)SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            oForm.DataSources.DataTables.Add("dt");
            oForm.DataSources.DataTables.Item(0).ExecuteQuery(Query);
            Grid3.DataTable = oForm.DataSources.DataTables.Item("dt");
        }
        void Matrix_Data_Write()
        {

            SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)SAPbouiCOM.Framework.Application.SBO_Application.Company.GetDICompany();
            SAPbobsCOM.Recordset oRset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRset.DoQuery(Query);
            
            if (oRset.RecordCount > 0)
            {
                for (int i = 0; i < oRset.RecordCount; i++)
                {
                    Matrix0.AddRow();
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Name").Cells.Item(i + 1).Specific).Value = oRset.Fields.Item("lastName").Value.ToString();
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Surname").Cells.Item(i + 1).Specific).Value = oRset.Fields.Item("firstName").Value.ToString();
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FatherName").Cells.Item(i + 1).Specific).Value = oRset.Fields.Item("middleName").Value.ToString();
                    oRset.MoveNext();
                }
            }
        }
        private void Button0_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //connection Grid
            Grid_Data_Write();
            Matrix_Data_Write();

        }

        private Matrix Matrix0;

        public void OnInitializeFormEvents()
        {

        }

        private StaticText StaticText0;
        private StaticText StaticText1;
    }
}