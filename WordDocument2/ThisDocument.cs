using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocument2
{
    public partial class ThisDocument
    {
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            this.bookmark1.Range.InsertAfter("STEVE");
            this.Tables[1].Rows.Add();
            this.Tables[1].Rows[2].Range.Bold = 0;
            this.Tables[1].Rows.Add();
            this.Tables[1].Cell(2, 1).Merge(this.Tables[1].Cell(3, 1));
            this.Tables[1].Cell(2, 1).Range.Text = "1";
            this.Tables[1].Cell(2, 2).Merge(this.Tables[1].Cell(3, 2));
            this.Tables[1].Cell(2, 2).Range.Text = "Rntl";
            this.Tables[1].Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            this.Tables[1].Cell(2, 3).Range.Text = "General Ledger?";
            this.Tables[1].Cell(2, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            this.Tables[1].Cell(2, 5).Range.Text = "N35714";
            this.Tables[1].Cell(2, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            this.Tables[1].Cell(2, 6).Range.Text = "1.00";
            this.Tables[1].Cell(2, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            this.Tables[1].Cell(2, 7).Range.Text = "Hours";
            this.Tables[1].Cell(2, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            this.Tables[1].Cell(2, 8).Range.Text = "45.00";
            this.Tables[1].Cell(2, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            this.Tables[1].Cell(2, 9).Range.Text = "45.00";
            this.Tables[1].Cell(3, 4).Merge(this.Tables[1].Cell(3, 8));
            this.Tables[1].Cell(3, 4).FitText = true;
            this.Tables[1].Cell(3, 4).Range.Text = "Hobbs Out: 1067.50 In: 1068.10   Tach Out: 6596.56 In: 6596.99";
            
        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion
    }
}
