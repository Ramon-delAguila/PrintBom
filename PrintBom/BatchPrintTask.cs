using log4net;
using Print_Bom;
using PrintBom.Properties;
using QRCoder;
using SolidEdgeCommunity;
using SolidEdgeDraft;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidEdgeFramework;
using Application = SolidEdgeFramework.Application;
using System.Threading.Tasks;
using System.Security.Principal;

namespace PrintBom
{
    public class BatchPrintTask : IsolatedTaskProxy
    {
        private static readonly ILog _log = LogManager.GetLogger(typeof(BatchPrintTask));
        
        public void Print(ListViewItem item, DraftPrintUtilityOptions options, List<string> parametros)
        {
           InvokeSTAThread(PrintInternal, item, options, parametros);
        }
        private void  PrintInternal(ListViewItem item, DraftPrintUtilityOptions options, List<string> parametros)
        {
            SolidEdgeFramework.Application application = null;
            SolidEdgeDraft.DraftDocument draftDocument = null;
            SolidEdgeDraft.DraftPrintUtility draftPrintUtility = null;
            using var docManager = new SolidEdgeDocumentManager(false);

            try
            {// Register with OLE to handle concurrency issues on the current thread.
                SolidEdgeCommunity.OleMessageFilter.Register();               
                application = docManager.SEApplication;
                // Get a reference to the DraftPrintUtility.
                draftPrintUtility = (SolidEdgeDraft.DraftPrintUtility)application.GetDraftPrintUtility();

                var settings = new PrinterSettings();
                // Copy all of the settings from DraftPrintUtilityOptions to the DraftPrintUtility object
                CopyOptions(draftPrintUtility, options);

                ConfigureDraftPrintUtility(draftPrintUtility, settings);

                var dftFullName = item.SubItems[2].Text;
                if (string.IsNullOrEmpty(dftFullName)) return;

                draftDocument = (SolidEdgeDraft.DraftDocument)application.Documents.Open(dftFullName);

                WriteQuantityOF(draftDocument, item, parametros[0]);
                Extensions.Qrcode(draftDocument, item, parametros[0]);

                HandlePrintAndPDF(draftDocument, draftPrintUtility, item, parametros);
                // guardar los desplegados de las chapas.
                PrintBom.PSMtoDXFCreation.Process_DftEX(draftDocument);

            }

            catch (Exception ex)
            {
                LogException(ex, "Error during printing process");
                throw;
            }

            finally
            {
                // Make sure we close the document.
                draftDocument?.Close(false);
                application?.DoIdle();
                //docManager.Dispose();
                SolidEdgeCommunity.OleMessageFilter.Revoke();          

            }
        }
        private void HandlePrintAndPDF(SolidEdgeDraft.DraftDocument draftDocument, SolidEdgeDraft.DraftPrintUtility draftPrintUtility, ListViewItem item, List<string> parametros)
        {
            var cmbValue = parametros[1];

            if (cmbValue == "1" || cmbValue == "3")
            {
                PrintSheets(draftDocument, draftPrintUtility);
            }

            if (cmbValue == "2" || cmbValue == "3")
            {
                var name = Extensions.OutfilenamePDFBOM(item.SubItems[2].Text, item.SubItems[0].Text, ".pdf", parametros[3], parametros[2]);
                Extensions.CreatePDF(draftDocument, name, parametros[4]);
            }
        }
        private static void ConfigureDraftPrintUtility(SolidEdgeDraft.DraftPrintUtility draftPrintUtility, PrinterSettings settings)
        {
            draftPrintUtility.Printer = settings.PrinterName;
            draftPrintUtility.PrintAsBlack = true;
            draftPrintUtility.BestFit = false;
            draftPrintUtility.AutoOrient = true;
            draftPrintUtility.UsePrinterClipping = false;
            draftPrintUtility.UsePrinterMargins = false;
        }
        private void CopyOptions(SolidEdgeDraft.DraftPrintUtility draftPrintUtility, DraftPrintUtilityOptions options)
        {
            var fromType = typeof(DraftPrintUtilityOptions);
            var toType = typeof(SolidEdgeDraft.DraftPrintUtility);
            var properties = toType.GetProperties().Where(x => x.CanWrite).ToArray();

            // Copy all of the properties from DraftPrintUtility to this object.
            foreach (var toProperty in properties)
            {
                // Some properties may throw an exception if options are incompatible.
                // For instance, if PrintToFile = false, setting PrintToFileName = "" will cause an exception.
                // Mostly irrelevant but handle it as you see fit.
                try
                {
                    var fromProperty = fromType.GetProperty(toProperty.Name);
                    if (fromProperty != null)
                    {
                        var val = fromProperty.GetValue(options);

                        toType.InvokeMember(toProperty.Name, BindingFlags.SetProperty, null, draftPrintUtility, new object[] { val });

                    }
                }
                catch (Exception ex)
                {
                    _log.Warn($"Failed to copy property {toProperty.Name}", ex);
                }
            }
        }
        private void WriteQuantityOF(SolidEdgeDraft.DraftDocument draftDocument, ListViewItem item, string of)
        {

            foreach (var block in draftDocument.Blocks.OfType<SolidEdgeDraft.Block>().Where(b => b.Name.ToUpper().Contains("CAJETIN")))
            {
                var vistaBloque = block.DefaultView;
                foreach (var balloon in vistaBloque.Balloons.OfType<SolidEdgeFrameworkSupport.Balloon>().Where(b => b.BalloonText.ToUpper().Trim() == "%{CANTIDAD|R1}"))
                {
                    balloon.BalloonText = item.SubItems[1].Text.Contains("+") ? $"{item.SubItems[1].Text}+{item.SubItems[1].Text}" : item.SubItems[1].Text.ToString();
                }
                

                foreach (var textbox in vistaBloque.TextBoxes.OfType<SolidEdgeFrameworkSupport.TextBox>().Where(t => t.Text == "OF"))
                {
                    try
                    {
                        textbox.Text = of;
                        textbox.Width = 1;
                        textbox.Edit.TextSize = 0.0035;
                    }
                    catch (Exception ex)
                    {
                        _log.Warn("Failed to update textbox", ex);
                    }
                }
            }

        }
        private void PrintSheets(SolidEdgeDraft.DraftDocument draftDocument, SolidEdgeDraft.DraftPrintUtility draftPrintUtility)
        {
            foreach (var sheet in draftDocument.Sections.WorkingSection.Sheets.OfType<SolidEdgeDraft.Sheet>())
            {
                var setup = sheet.Background.SheetSetup;
                draftPrintUtility.PaperSize = setup.SheetWidth.Equals(0.21) || (setup.SheetWidth.Equals(0.297) && setup.SheetHeight.Equals(0.21))
                    ? DraftPrintPaperSizeConstants.igDraftPrintPaperSize_A4
                    : DraftPrintPaperSizeConstants.igDraftPrintPaperSize_A3;

                draftPrintUtility.AddSheet(sheet);
            }

            //draftPrintUtility.PrintOut();
            // Aquí también usamos Task.Run para evitar bloquear el hilo principal
            draftPrintUtility.PrintOut();
            draftPrintUtility.RemoveAllDocuments();
        }
        private static void LogException(Exception ex, string customMessage)
        {
            _log.Error(customMessage, ex);
        }

    }

}
