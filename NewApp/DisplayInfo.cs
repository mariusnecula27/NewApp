using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NewApp
{
    public class DisplayInfo
    {
        public string allInfos;
        public string imporantInfos;

        public void comboBoxAllWriter(string infos, System.Windows.Forms.TextBox textInfos, System.Windows.Forms.ComboBox combo, bool imporant)
        {
            if (allInfos == null)
                allInfos += infos;
            else
                allInfos += "\r\n" + infos;

            if(imporant is true)
            {
                if (imporantInfos == null)
                    imporantInfos += infos;
                else
                    imporantInfos += "\r\n" + infos;
            }

            if (combo.Text == "All")
            {
                textInfos.Text = allInfos;

            }else if (combo.Text == "Important")
            {
                textInfos.Text = imporantInfos;
            }
               
        }
    }
}
