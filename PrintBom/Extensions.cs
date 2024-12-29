
using log4net;
using QRCoder;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using SolidEdgeDraft;
using SolidEdgeAssembly;
using Path = System.IO.Path;
using System.Threading.Tasks;
using System.Threading;


namespace PrintBom

{
    [Serializable]
    public class PrintItem
    {
        // obviously you find meaningful names of the 2 properties

        public string FilePath { get; set; }
        public int Quantity { get; set; }
        public string DocumentNumber { get; set; }

        public string DraftFilePath { get; set; }

        /// <summary>
        /// Initializes a new instance of the PrintItem class.
        /// </summary>
        /// <param name="filePath">Full path of the file.</param>
        /// <param name="quantity">Quantity of items to print.</param>
        /// <param name="documentNumber">Document identification number.</param>
        /// <param name="draftFilePath">Full path of the draft file associated with the item.</param>
        public PrintItem(string filePath, int quantity, string documentNumber, string draftFilePath)
        {
            this.FilePath = filePath;
            this.Quantity = quantity;
            this.DocumentNumber = documentNumber;
            this.DraftFilePath = draftFilePath;
        }


    }

    public static class Extensions
    {
        internal static void PopulateBOM(dynamic occurrences, List<PrintItem> items)
        {
            foreach (var occurrence in occurrences)
            {
                // Ignorar componentes internos o archivos faltantes
                if (occurrence.FileMissing() || occurrence.IsInternalComponent)
                    continue;

                // Obtener el nombre del archivo y verificar si existe su plano
                string occurrenceFilenamePart = occurrence is SolidEdgeAssembly.Occurrence
                    ? GetFilenameWithoutExclamation(occurrence.OccurrenceFileName.ToLower())
                    : GetFilenameWithoutExclamation(occurrence.SubOccurrenceFileName.ToLower());

                string dftFullName = Path.ChangeExtension(occurrenceFilenamePart, ".dft");
                bool dftExists = File.Exists(dftFullName);

                // Validar si debe incluirse en el BOM
                if (!ShouldIncludeInBOM(occurrence, dftExists))
                    continue;

                // Agregar o actualizar el ítem en la lista
                AddOrUpdateItem(items, occurrence, occurrenceFilenamePart, dftFullName);

                // Procesar subensamblajes recursivamente
                if (occurrence.Subassembly)
                {
                    PopulateBOM(occurrence.SubOccurrences, items);
                }
            }
        }
        private static string GetFilenameWithoutExclamation(string filename)
        {
            int exclamationIndex = filename.IndexOf('!');
            return exclamationIndex >= 0 ? filename.Substring(0, exclamationIndex) : filename;
        }
        private static bool ShouldIncludeInBOM(dynamic occurrence, bool dftExists)
        {
            bool includeInBom;

            if (occurrence is SolidEdgeAssembly.Occurrence)
            {
                includeInBom = occurrence.IncludeInBom;
            }
            else if (occurrence is SolidEdgeAssembly.SubOccurrence)
            {
                includeInBom = occurrence.ThisAsOccurrence.IncludeInBom;
            }
            else
            {
                throw new ArgumentException("Unsupported occurrence type.");
            }

            //return (dftExists || !IsInLibrary(occurrence)) && includeInBom && occurrence.Visible;
            return dftExists && includeInBom && occurrence.Visible;
        }
        private static bool IsInLibrary(dynamic occurrence)
        {
            string occurrenceDir;

            // Verificar si es Occurrence o SubOccurrence
            if (occurrence is SolidEdgeAssembly.Occurrence)
            {
                occurrenceDir = Path.GetDirectoryName(occurrence.OccurrenceFileName.ToLower());
            }
            else if (occurrence is SolidEdgeAssembly.SubOccurrence)
            {
                occurrenceDir = Path.GetDirectoryName(occurrence.SubOccurrenceFileName.ToLower());
            }
            else
            {
                throw new ArgumentException("Unsupported occurrence type.");
            }

            return PrintBom.Mainform.Library.Any(path => occurrenceDir.Contains(path));
        }

        //pasamos de los archivos de libreria solo los imprimimos sin tienen dft sea como sean.
        private static void AddOrUpdateItem(List<PrintItem> items, dynamic occurrence, string filePath, string draftFilePath)
        {
            var existingItem = items.FirstOrDefault(x => x.FilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase));
            if (existingItem == null)
            {
                SolidEdgeFramework.SolidEdgeDocument document;
                int quantity;

                // Diferenciamos entre Occurrence y SubOccurrence
                if (occurrence is SolidEdgeAssembly.SubOccurrence subOccurrence)
                {
                    document = (SolidEdgeFramework.SolidEdgeDocument)subOccurrence.SubOccurrenceDocument;
                    quantity = GetQuantity(subOccurrence.ThisAsOccurrence, document);
                }
                else if (occurrence is SolidEdgeAssembly.Occurrence standardOccurrence)
                {
                    document = (SolidEdgeFramework.SolidEdgeDocument)standardOccurrence.OccurrenceDocument;
                    quantity = GetQuantity(standardOccurrence, document);
                }
                else
                {
                    throw new InvalidOperationException("Unsupported occurrence type.");
                }

                var summaryInfo = (SolidEdgeFramework.SummaryInfo)document.SummaryInfo;
                items.Add(new PrintItem(filePath, quantity, summaryInfo.DocumentNumber, draftFilePath));
            }
            else
            {
                // Usamos la cantidad según el tipo
                if (occurrence is SolidEdgeAssembly.SubOccurrence subOccurrence)
                {
                    existingItem.Quantity += subOccurrence.ThisAsOccurrence.Quantity;
                }
                else
                {
                    existingItem.Quantity += occurrence.Quantity;
                }
            }
        }
        private static int GetQuantity(dynamic occurrence, SolidEdgeFramework.SolidEdgeDocument document)
        {
            try //si hay un error en la cantidad es porque la cantidad está sobreescrita posiblemente
            {
                return occurrence.Quantity;
            }
            catch
            {
                return GetCustomQuantity(document);
            }
        }
        private static int GetCustomQuantity(SolidEdgeFramework.SolidEdgeDocument document)
        {
            try
            {
                var customProperties = ((SolidEdgeFramework.PropertySets)document.Properties).Item("Custom");
                var assemblyQuantityProperty = (SolidEdgeFramework.Property)customProperties.Item("SE_ASSEMBLY_QUANTITY_OVERRIDE");
                return (int)((dynamic)assemblyQuantityProperty).Value;
            }
            catch
            {
                Mainform.LogError("Error retrieving custom quantity.");
                return 1; // Default quantity
            }
        }
        private static void LogError(string message)
        {
            // Use log4net or other logging framework to log errors
        }
        internal static string OutfilenamePDFBOM(string filename, string document_number, string extension, string _path, string assdocnumber)
        {
            //string _path = System.IO.Path.GetDirectoryName(filename);
            string _filename = System.IO.Path.GetFileName(filename);

            //if (!System.IO.Directory.Exists(_path + "\\PDF"))
            //{
            //    System.IO.Directory.CreateDirectory(_path + "\\PDF");
            //}

            if (!System.IO.Directory.Exists(_path + "\\PDF's_" + assdocnumber))
            {
                System.IO.Directory.CreateDirectory(_path + "\\PDF's_" + assdocnumber);
            }

            string outfile;
            if (document_number == "")
            {
                outfile = System.IO.Path.ChangeExtension(_filename, extension);
            }
            else
            {
                outfile = document_number + extension;

            }
            return _path + "\\PDF's_" + assdocnumber + "\\" + outfile;
        }
       
        public static void CreatePDF(SolidEdgeDraft.DraftDocument draftDocument, string outfilePDF, string overwriteIfExists)
        {
            //ILog _log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            try
            {
                if (File.Exists(outfilePDF))
                {
                    var fileModifiedTime = File.GetLastWriteTime(outfilePDF);
                    var draftModifiedTime = File.GetLastWriteTime(draftDocument.FullName);

                    if (draftModifiedTime <= fileModifiedTime && !Convert.ToBoolean(overwriteIfExists))
                    {
                        outfilePDF = NextAvailableFilename(outfilePDF);
                    }
                    else if (Convert.ToBoolean(overwriteIfExists))
                    {
                        File.Delete(outfilePDF);
                    }
                }

                draftDocument.SaveAs(outfilePDF, FileFormat: false);
            }
            catch (Exception ex)
            {
                Mainform.LogError($"Error creating PDF at {outfilePDF}: {ex.Message}");
            }
        }
        private static void RunInSTA(Action action)
        {
            var thread = new Thread(() =>
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    Mainform._log.Error("Error in STA thread execution", ex);
                    throw;
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(); // Esperar a que el hilo termine, ya que es un método sincrónico
        }
        public static string NextAvailableFilename(string path)
        {
            // Short-cut if already available
            if (!File.Exists(path))
                return path;
            // If path has extension then insert the number pattern just before the extension and return next filename
            // Otherwise just append the pattern to the path and return next filename
            string pattern = Path.HasExtension(path) ? path.Insert(path.LastIndexOf(Path.GetExtension(path)), " ({0})") : path + " ({0})";
            return GetNextFilename(pattern);
        }
        private static string GetNextFilename_(string pattern)
        {

            string tmp = string.Format(pattern, 1);
            if (tmp == pattern)
            {
                throw new ArgumentException("The pattern must include an index place-holder", "pattern");
            }

            if (!File.Exists(tmp))
            {
                return tmp; // short-circuit if no matches
            }

            int min = 1, max = 2; // min is inclusive, max is exclusive/untested

            while (File.Exists(string.Format(pattern, max)))
            {
                min = max;
                max *= 2;
            }

            while (max != min + 1)
            {
                int pivot = (max + min) / 2;
                if (File.Exists(string.Format(pattern, pivot)))
                {
                    min = pivot;
                }
                else
                {
                    max = pivot;
                }
            }

            return string.Format(pattern, max);
        }
        private static string GetNextFilename(string pattern)
        {
            int index = 1;
            while (File.Exists(string.Format(pattern, index)))
            {
                index++;
            }

            return string.Format(pattern, index);
        }
        public static void Do<TControl>(this TControl control, Action<TControl> action)
          where TControl : Control
        {
            if (control.InvokeRequired)
            {
                control.BeginInvoke(action, control);
            }
            else
            {
                action(control);
            }
        }
        public static bool HasWritePermission(string tempfilepath)
        {
            try
            {
                System.IO.File.Create(tempfilepath + "temp.txt").Close();
                System.IO.File.Delete(tempfilepath + "temp.txt");
            }
            catch (System.UnauthorizedAccessException)
            {

                return false;
            }
            return true;
        }
        public static string PathFilename(string filename, string document_number, string extension)
        {
            // Obtener la ruta de la carpeta y el nombre del archivo de la ruta completa.
            string _path = System.IO.Path.GetDirectoryName(filename); // Ruta de la carpeta.
            string _filename = System.IO.Path.GetFileName(filename); // Nombre del archivo.

            // Crear la carpeta "DWG-DXF" si no existe.
            if (!System.IO.Directory.Exists(_path + "\\DWG-DXF"))
            {
                System.IO.Directory.CreateDirectory(_path + "\\DWG-DXF");
            }

            string outfile;

            if (string.IsNullOrEmpty(document_number))
            {
                // Si el número de documento no existe, usar el nombre del archivo con la extensión proporcionada.
                outfile = System.IO.Path.ChangeExtension(_filename, extension);
            }
            else
            {
                // Usar el número de documento y la extensión proporcionada.
                outfile = document_number + extension;
            }

            // Devolver la ruta completa del archivo.
            return _path + "\\DWG-DXF\\" + outfile;
        }
        public static void AddUpdateAppSettings(string key, string value)
        {
            try
            {
                var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;
                if (settings[key] == null)
                {
                    settings.Add(key, value);
                }
                else
                {
                    settings[key].Value = value;
                }
                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error writing app settings");
            }
        }
        public static void Qrcode(SolidEdgeDraft.DraftDocument draftdocument, ListViewItem item, string of)
        {
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            string _string = ("OF " + of + " - Nº " + item.SubItems[0].Text + " - C=" + item.SubItems[1].Text).ToString();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(_string, QRCodeGenerator.ECCLevel.Q);
            //QRCodeData qrCodeData = qrGenerator.CreateQrCode("https://www.shopdirect-online.es/vitrina-46-sin-iluminacin-movible-l46xf46xa182cmpuerta-con-cierre-a-llave/", QRCodeGenerator.ECCLevel.Q);

            QRCode qrCode = new QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(20);

            //vamos a ver si le metemos un qr


            qrCodeImage.Save(System.IO.Path.GetTempPath() + "qrcode.bmp");

            SolidEdgeDraft.Sections _sections = draftdocument.Sections;
            SolidEdgeDraft.Section _section = _sections.WorkingSection;

            double ancho = 0.010; //tamaño en metros del qr


            foreach (SolidEdgeDraft.Sheet _sheet in _section.Sheets)
            {
                SolidEdgeFrameworkSupport.Image2d imagen = _sheet.Images2d.AddImage(false, System.IO.Path.GetTempPath() + "qrcode.bmp");

                PaperSizeConstants tamaño = _sheet.SheetSetup.SheetSizeOption;

                double altura = _sheet.SheetSetup.SheetHeight;
                double originx = 0.040;
                double originy = 0.040;
                double offset = 0.001;

                switch (tamaño)
                {
                    case PaperSizeConstants.igIsoA4Tall:

                        originx = 0.020 - ancho - offset;
                        originy = altura - ancho - 0.005;

                        break;

                    case PaperSizeConstants.igIsoA3Wide:

                        originx = 0.015 - ancho - offset;
                        originy = altura - ancho - 0005;

                        break;

                    case PaperSizeConstants.igIsoA2Wide:

                        originx = 0.015 - ancho - offset;
                        originy = altura - ancho - 0.005;

                        break;

                    case PaperSizeConstants.igIsoA1Wide:

                        originx = 0.015 - ancho - offset;
                        originy = altura - ancho - 0.005;

                        break;

                    case PaperSizeConstants.igIsoA0Wide:

                        originx = 0.015 - ancho - offset;
                        originy = altura - ancho - 0.005;

                        break;

                }

                imagen.Width = ancho;
                imagen.ShowBorder = false;
                imagen.SetOrigin(originx, originy);

            }


        }

    }




}

