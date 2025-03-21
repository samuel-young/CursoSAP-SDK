
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace DocumentosIntercompany
{

    [FormAttribute("DocumentosIntercompany.DocInter_b1f", "DocInter.b1f")]
    class DocInter_b1f : UserFormBase
    {
        SAPbouiCOM.EditText Txt_DocEntry;
        SAPbouiCOM.EditText Txt_DocNum;
        SAPbouiCOM.EditText Txt_Nombre;
        SAPbouiCOM.Form oForm;
        string idForm;

        public DocInter_b1f()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            try
            {
                this.oForm = (SAPbouiCOM.Form)this.UIAPIRawForm;
                this.idForm = this.oForm.UniqueID;

                this.Txt_DocEntry = (SAPbouiCOM.EditText)this.oForm.Items.Item("T_1").Specific;
                this.Txt_DocNum = (SAPbouiCOM.EditText)this.oForm.Items.Item("T_2").Specific;
                this.Txt_Nombre = (SAPbouiCOM.EditText)this.oForm.Items.Item("T_3").Specific;

                //(2:add; 1:update / ok; -1:all; 4:find; 8:view )
                this.Txt_DocEntry.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                this.Txt_DocEntry.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_Default);
                this.Txt_DocEntry.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                this.Txt_DocNum.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                this.Txt_DocNum.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_Default);
                this.Txt_DocNum.Item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            }
            catch(Exception ex)
            {
                Application.SBO_Application.MessageBox("Error: " + ex.Message);
            }
        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }




    }
}
