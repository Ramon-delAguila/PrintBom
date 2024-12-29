using System;
using SolidEdgeFramework;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using System.Linq;
using SolidEdgeCommunity.Extensions;
using System.Runtime.CompilerServices;
using System.Collections.Generic;

namespace Print_Bom
{
    public class SolidEdgeDocumentManager:IDisposable
    {
        private Application _seApplication;
        private SolidEdgeDocument _document;
        private readonly Dictionary<string, bool> _dictAddins;
        // private Dictionary<string, bool> _dict_Addins = new Dictionary<string, bool>();
        object objVal1;
        object objVal2;

        public Application SEApplication => _seApplication;
        public SolidEdgeDocument Document => _document;
        public Dictionary<string, bool> DictAddins => _dictAddins;
       
        public SolidEdgeDocumentManager(bool startIfNotRunning)
        {
            _dictAddins = new Dictionary<string, bool>();

            try
            {
                // Intentar conectar a una instancia en ejecución de Solid Edge.
                _seApplication = (Application)Marshal.GetActiveObject("SolidEdge.Application");
            }
            catch (COMException ex) when (ex.ErrorCode == -2147221021 /* MK_E_UNAVAILABLE */)
            {
                // No se pudo conectar.
                if (startIfNotRunning)
                {
                    // Iniciar Solid Edge.
                    _seApplication = (Application)Activator.CreateInstance(Type.GetTypeFromProgID("SolidEdge.Application"));
                }
                else
                {
                    throw new Exception("Solid Edge is not running.", ex);
                }
            }
        }

        public SolidEdgeDocument ActiveDocument()
        {
            var  documents = _seApplication.Documents;
            if (documents.Count > 0)
            {
                // Activar un documento en Solid Edge
                _document = (SolidEdgeDocument)_seApplication.ActiveDocument;
            }
            else
            {
                _document = null;
                // Aquí puedes lanzar una excepción o log para indicar que no hay documento activo
            }

            return _document;
        }

        public void ToggleAddIns(bool activate)
        {
            if (!activate)
            {
                foreach (SolidEdgeFramework.AddIn _addin in _seApplication.AddIns)
                {
                    if (!_dictAddins.ContainsKey(_addin.GUID))
                    {
                        _dictAddins.Add(_addin.GUID, _addin.Connect);
                    }
                    _addin.Connect = activate;
                }
            }
            else
            {
                foreach (SolidEdgeFramework.AddIn _addin in _seApplication.AddIns)
                {
                    if (_dictAddins.TryGetValue(_addin.GUID, out bool originalState))
                    {
                        _addin.Connect = originalState;
                    }
                }
            }
        }


        public void ToggleConfiguration(string action)
        {

            if (_seApplication != null)
            {
                bool enable = action.Equals("enable", StringComparison.OrdinalIgnoreCase);

                if (!enable) //si disable hacemos que no se almacene el archivo en la lista de recientes
                {
                    _seApplication.SuspendMRU();
                }
                else
                {
                    _seApplication.ResumeMRU();
                }
                _seApplication.DisplayAlerts = enable;
                _seApplication.DelayCompute = !enable;
                _seApplication.Interactive = enable;
                _seApplication.ScreenUpdating = enable;

            }
            
        }

        public void SetAndRestoreParameters(bool save)
        {
            if (save)
            {
                try
                {
                    objVal1 = ApplicationExtensions.GetGlobalParameter(_seApplication, SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFSaveAllColorsBlack);
                    objVal2 = ApplicationExtensions.GetGlobalParameter(_seApplication, SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFSheetOptions);

                }
                catch ( Exception ex)
                {
                    // Manejar la excepción adecuadamente
                    Console.WriteLine("Error al obtener parámetros: " + ex.Message);
                }
                // Establecer nuevos valores para los parámetros
                _seApplication.SetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFSaveAllColorsBlack, true);
                _seApplication.SetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFSheetOptions, SolidEdgeConstants.DraftSaveAsPDFSheetOptionsConstants.seDraftSaveAsPDFSheetOptionsConstantsAllSheets);
            }
            else
            {
                // Aquí puedes realizar las operaciones que necesiten los nuevos valores de los parámetros

                // Restaurar los valores originales de los parámetros
                _seApplication.SetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFSaveAllColorsBlack, objVal1);
                _seApplication.SetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFSheetOptions, objVal2);
            }
        }

        public void Dispose()
        {
            if (_seApplication != null)
            {
                Marshal.ReleaseComObject(_seApplication);
                _seApplication = null;
            }
        }
    }
}