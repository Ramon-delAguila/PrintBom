
using log4net;
using log4net.Config;
using log4net.Repository.Hierarchy;
using Print_Bom;
using SolidEdgeCommunity;
using SolidEdgeCommunity.Extensions;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Marshal = System.Runtime.InteropServices.Marshal;

namespace PrintBom
{
    public partial class Mainform : Form
    {
        private SolidEdgeFramework.Application SeApplication;
        private SolidEdgeDocumentManager docManager;
        private SolidEdgeAssembly.AssemblyDocument assemblyDocument;
        private DraftPrintUtilityOptions options;
        private readonly List<PrintItem> _items = new List<PrintItem>();
        public static ILog _log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static string[] Library = ConfigurationManager.AppSettings["Library_folder"].Split(',').Select(n =>Convert.ToString(n).ToLower()).ToArray();
        public static string selectedFolderPath;
        public static List<string> _params = new List<string>();
        public static string pdf_folder = ConfigurationManager.AppSettings["pdf_folder"];
        public static string ass_doc_number;

        static Mainform()
        {
            XmlConfigurator.Configure();
        }

        public Mainform() => InitializeComponent();
        private void MainForm_Load(object sender, EventArgs e)
        { BeginInvoke(new MethodInvoker(MainForm_Load_Async)); }
        private void MainForm_Load_Async()
        {
            try

            {
            SetUpLogging();

            ConnectToSolidEdge();

            docManager.ToggleAddIns(false);

            SolidEdgeFramework.SummaryInfo summaryInfo;

            if (assemblyDocument == null)
            {
            ShowErrorAndExit("No hay ningún archivo de conjunto activo.");
            return;
            }

            options = CreatePrintOptions();
            summaryInfo = (SolidEdgeFramework.SummaryInfo)assemblyDocument.SummaryInfo;
            ass_doc_number = summaryInfo.DocumentNumber;

            LoadAssemblyDetails(summaryInfo);

            ConfigureUI();

            this.Show();

            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
            HandleError(ex);
                
            }
        }
        private void Frm_FormClosing(object sender, FormClosingEventArgs e)
        {
            ReleaseResources();
        }
        private void ReleaseResources()
        {
            if (SeApplication != null)
            {   
                docManager.ToggleAddIns(true);
                docManager.SetAndRestoreParameters(false);
                docManager.Dispose();
                OleMessageFilter.Revoke();
                Marshal.ReleaseComObject(SeApplication);
            }
        }
        private bool ChooseFolder(string assemblyDocNumber)
        {
            using (var folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = $"Seleccionar la ruta donde crear todos los PDF´s. \r\n Los archivos se colocarán bajo una carpeta nueva que se creará con el nombre PDF´s_{assemblyDocNumber}";
                folderBrowserDialog.ShowNewFolderButton = false;

                if (!string.IsNullOrEmpty(pdf_folder))
                {
                    folderBrowserDialog.SelectedPath = pdf_folder;
                }

                var res = folderBrowserDialog.ShowDialog();

                if (res == DialogResult.OK)
                {
                    // Get the selected folder path
                    selectedFolderPath = folderBrowserDialog.SelectedPath;

                    // Guardamos la configuración
                    Extensions.AddUpdateAppSettings("pdf_folder", selectedFolderPath);

                    return Extensions.HasWritePermission(selectedFolderPath);
                }
                else
                {
                    EnableUIState();
                    return false;
                }
            }
        }

        //CUANDO ABANDONAMOS EL FOCO SE MULTIPLICAN LAS CANTIDADES 
        private void textBoxCantidad_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {

            //System.Windows.Forms.TextBox textb = (System.Windows.Forms.TextBox)sender;

            //if (int.TryParse(textb.Text, out int cantidad) && cantidad >= 1)
            if (sender is System.Windows.Forms.TextBox textb && int.TryParse(textb.Text, out int cantidad) && cantidad >= 1)

            {
                foreach (ListViewItem itemRow in listViewFiles.Items)
                {
                    //itemRow.SubItems[1].Text = (Convert.ToInt32(itemRow.SubItems[3].Text) * cantidad).ToString();

                    if (int.TryParse(itemRow.SubItems[3].Text, out int subItemValue))
                    {
                        itemRow.SubItems[1].Text = (subItemValue * cantidad).ToString();
                    }

                }
                e.Cancel = false;
            }
            else
            {
                MessageBox.Show("Por favor introduce una cantidad válida", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //errorProvider1.SetError(textBoxCantidad, "Por favor introduce una cantidad válida");
                e.Cancel = true;
            }
        }
        private void textBoxOF_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            System.Windows.Forms.TextBox textb = (System.Windows.Forms.TextBox)sender;

            //quité la validación para poder poner repuestos
            //if (!((int.TryParse(textb.Text, out int fof) && fof >= 1) || textb.Text.TrimStart() == ""))
            //{
            //    MessageBox.Show("Por favor introduce un número de orden correcto", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    e.Cancel = true;
            //}
            //else
            //{
            //    e.Cancel = false;
            //}
        }
        private void DisableUIState()
        {
            this.btnPrint.Enabled = false;
            this.comboBox_tasks.Enabled = false;
        }
        private void EnableUIState()
        {
            this.btnPrint.Enabled = true;
            this.comboBox_tasks.Enabled = true;
        }
        private void textBoxCantidad_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = e.SuppressKeyPress = true;
                //textBoxCantidad_Validating(textBoxCantidad, new CancelEventArgs());
                btnPrint.Focus();
            }
        }
        private void textBoxOF_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = e.SuppressKeyPress = true;
                // textBoxOF_Validating(textBoxOF, new CancelEventArgs());
                btnPrint.Focus();
            }
        }
        private  void Print_Click(object sender, EventArgs e)

        {
            if (string.IsNullOrWhiteSpace(textBoxOF.Text))
            {
                var result = MessageBox.Show("¿Estás seguro que quieres imprimir los planos sin Orden de fabricación?", "Advertencia", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.No) return;
                
            }
            DisableUIState();
            var cmbvalue = (int)comboBox_tasks.SelectedValue; //valor del combo de imprimir planos

            _params.Add(textBoxOF.Text);
            _params.Add(cmbvalue.ToString());
            _params.Add(ass_doc_number);
   
            if (cmbvalue == 2 || cmbvalue == 3)
            {
                if (!ChooseFolder(ass_doc_number)) return;
                _params.Add(selectedFolderPath);            
            }

            _params.Add(checkBoxPDF.Checked.ToString());
     
            Stopwatch stopWatch = Stopwatch.StartNew();

            //var opts = new ParallelOptions { MaxDegreeOfParallelism = Convert.ToInt32(Math.Ceiling((Environment.ProcessorCount * 0.35) * 2.0)) };
            //Parallel.ForEach(_items, opts, _PrintItem =>           
            //foreach (var _PrintItem in _items)

            // Optional settings you may tweak for performance improvements. Results may vary.
            docManager.ToggleConfiguration("disable");

            foreach (ListViewItem item in listViewFiles.Items)
            {
                try
                {
                    if (string.IsNullOrEmpty(item.SubItems[2].Text)) return;

                    //imprimir planos
                    if (cmbvalue == 1 || cmbvalue == 3)
                    {
                        _log.Info($"Imprimiendo {item.SubItems[2].Text}");
                    }

                    if (cmbvalue == 2 || cmbvalue == 3)
                    {
                        _log.Info($"Creando PDF {item.SubItems[2].Text}");
                    }

                    using (var task = new IsolatedTask<BatchPrintTask>())
                    {
                        task.Proxy.Print(item, options, _params);
                    }

                }
                catch (Exception ex)
                {
                    //Console.WriteLine(ex.Message);
                    _log.Error($"Error al imprimir {item.SubItems[2].Text}: {ex.Message}");

                }

            }

            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;
            // Format and display the TimeSpan value.
            string elapsedTime = $"{ts:hh\\:mm\\:ss\\.ff}";
            _log.Info($"Tiempo en imprimir {elapsedTime}");
            //this.Close();
            docManager.ToggleConfiguration("enable");

        }
        private void SetUpLogging()
        {
            var textBoxAppender = LogManager.GetRepository().GetAppenders().OfType<TextBoxAppender>().FirstOrDefault();
            if (textBoxAppender != null)
            {
                textBoxAppender.TextBox = outputTextBox;
            }
        }
        public static void LogError(string message)
        {
            if (_log.IsErrorEnabled)
            {
                _log.Error(message);
            }
        }
        private void ConnectToSolidEdge()
        {
            OleMessageFilter.Register();
            docManager = new SolidEdgeDocumentManager(false);
            SeApplication = docManager.SEApplication;
            docManager.SetAndRestoreParameters(true);
            assemblyDocument = (SolidEdgeAssembly.AssemblyDocument)docManager.ActiveDocument();
            assemblyDocument.ActivateAll();
        }
        private void ShowErrorAndExit(string message)
        {
            MessageBox.Show(message, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            Close();
        }
        private DraftPrintUtilityOptions CreatePrintOptions()
        {
            var settings = new PrinterSettings();
            return new DraftPrintUtilityOptions(SeApplication)
            {
                Printer = settings.PrinterName,
                PrintAsBlack = true,
                BestFit = false,
                AutoOrient = true,
                UsePrinterClipping = true,
                UsePrinterMargins = true
            };
        }
        private void  LoadAssemblyDetails(SolidEdgeFramework.SummaryInfo summaryInfo)
        {
            var assemblyFullName = assemblyDocument.FullName.Split('!');
            var dftFullName = System.IO.Path.ChangeExtension(assemblyFullName[0], ".dft");

            if (!System.IO.File.Exists(dftFullName))
            {
                dftFullName = "No existe DFT";
            }

            PrintItem _item = new PrintItem(assemblyDocument.FullName, 1, summaryInfo.DocumentNumber, dftFullName);
            _items.Add(_item);  

            _log.Info("Creando listado...");

            // Begin the recursive extraction process
            Extensions.PopulateBOM(assemblyDocument.Occurrences, _items);

            PopulateListView();

            _log.Info("Listado creado.");
        }
        private void PopulateListView()
        {
            foreach (var item in _items)
            {
                var lvitem = new ListViewItem(item.DocumentNumber)
                {
                    SubItems = { item.Quantity.ToString(), item.DraftFilePath, item.Quantity.ToString() }
                };
                listViewFiles.Items.Add(lvitem);
                
            }
            ResizeListViewColumns();
        }
        private void ResizeListViewColumns()
        {
            if (listViewFiles.Items.Count > 0)
            {
                listViewFiles.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
                listViewFiles.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
                listViewFiles.Columns[listViewFiles.Columns.Count - 1].Width = -2;
            }
        }
        private void HandleError(Exception ex)
        {
            if (SeApplication == null)
            {
                ShowErrorAndExit("Solid Edge no está funcionando");
            }
            else if (assemblyDocument == null)
            {
                ShowErrorAndExit("No hay ningún archivo de conjunto abierto");
            }
            else
            {
                MessageBox.Show(ex.StackTrace, ex.Message);
            }
        }
        private void ConfigureUI()
        {

            // Configuración inicial de los controles en la interfaz de usuario 
            this.checkBoxPDF.Checked = false;  // Establece el valor inicial de CheckBox como no seleccionado
            this.textBoxOF.Clear();  // Limpia el TextBox para la Orden de Fabricación

            // Bind combobox to a dictionary.
            Dictionary<int, string> test = new Dictionary<int, string>();
            test.Add(1, "Imprimir planos");
            test.Add(2, "Imprimir PDF´s");
            test.Add(3, "Ambos");
            comboBox_tasks.DataSource = new BindingSource(test, null);
            comboBox_tasks.DisplayMember = "Value";
            comboBox_tasks.ValueMember = "Key";
            comboBox_tasks.SelectedIndex = 2; // Establece el índice predeterminado del ComboBox

            // Habilitar o deshabilitar controles según la situación
            this.btnPrint.Enabled = _items.Any();  // Solo habilitar el botón de impresión si hay elementos en la lista
            this.comboBox_tasks.Enabled = true;  // Asegúrate de que el ComboBox esté habilitado
            this.checkBoxPDF.Enabled = true;  // Habilita el CheckBox PDF


        }


    }
}
