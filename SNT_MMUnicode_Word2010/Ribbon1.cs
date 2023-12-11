using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.IO;

namespace SNT_MMUnicode_Word2010
{
    public partial class Ribbon1
    {
        string output, input;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ZGToUni_Click(object sender, RibbonControlEventArgs e)
        {

            /*
            if (Globals.ThisAddIn.Application.Documents.Count > 0)
            {
                Microsoft.Office.Interop.Word.Document nativeDocument =
                Globals.ThisAddIn.Application.ActiveDocument;
                Microsoft.Office.Tools.Word.Document vstoDocument =
                Globals.Factory.GetVstoObject(nativeDocument);
            }
             * */
            Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            input = currentRange.Text;
            /*
             foreach (char s in input) {
               output+=Rabbit.Zg2Uni(s.ToString());
                
            
             }*/

            output = Rabbit.Zg2Uni(input);
            currentRange.Text = output;
            currentRange.Font.Name = "Pyidaungsu"; 
        }
        private void U_Zaw_Click(object sender, RibbonControlEventArgs e)
        {

            /*
            if (Globals.ThisAddIn.Application.Documents.Count > 0)
            {
                Microsoft.Office.Interop.Word.Document nativeDocument =
                Globals.ThisAddIn.Application.ActiveDocument;
                Microsoft.Office.Tools.Word.Document vstoDocument =
                Globals.Factory.GetVstoObject(nativeDocument);
            }
             * */
            Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            input = currentRange.Text;
            /*
             foreach (char s in input) {
               output+=Rabbit.Zg2Uni(s.ToString());
                
            
             }*/

            output = Rabbit.Uni2Zg(input);
            currentRange.Text = output;
            currentRange.Font.Name = "Zawgyi-One";
        }

        private void btnSavePdf_Click(object sender, RibbonControlEventArgs e)
        {
            //string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //string fileName = "QuickExport.pdf";
            // SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "PDF Files|*.pdf";
            saveFileDialog1.Title = "Save a PDF File";
            saveFileDialog1.InitialDirectory = @"D:\";
            saveFileDialog1.ShowDialog();
            var file = new FileInfo(saveFileDialog1.FileName);
            Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(
                Path.Combine(file.DirectoryName, saveFileDialog1.FileName),
                Word.WdExportFormat.wdExportFormatPDF,
                OpenAfterExport: true);
        }

        private void btnXPS_Click(object sender, RibbonControlEventArgs e)
        {
            //string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //string fileName = "QuickExport.xps";
            saveFileDialog1.Filter = "XPS Files|*.XPS";
            saveFileDialog1.Title = "Save a XPS File";
            saveFileDialog1.InitialDirectory = @"D:\";
            saveFileDialog1.ShowDialog();
            var file = new FileInfo(saveFileDialog1.FileName);
            Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(
                Path.Combine(file.DirectoryName, saveFileDialog1.FileName),
                Word.WdExportFormat.wdExportFormatXPS,
                OpenAfterExport: true);

        }
    }
}
