using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// 
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;
//
using Ingr.SP3D.Common.Client;
using Ingr.SP3D.Common.Client.Services;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
//
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
namespace MyNameSpace
{
    public class SavePartsXml:BaseModalCommand
    {
        frmPartXml frm;
        public override void OnStart(int commandID, object argument)
        {
            if (frm==null)
            {
                frm = new frmPartXml();
                frm.FormClosed += Frm_FormClosed;
                frm.ShowDialog();
            }
            else
            {
                frm.Activate();
            }
            
        }

        private void Frm_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            frm = null;
        }
    }
}
