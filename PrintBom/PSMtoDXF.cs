
using Print_Bom;
using SolidEdgeDraft;
using SolidEdgeFramework;
using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;


namespace PrintBom
{
    public static class PSMtoDXFCreation
    {
        static SolidEdgePart.SheetMetalDocument sheetMetalDocument = null;
        static SolidEdgePart.Models _models = null;
        //static SolidEdgeGeometry.Face RefFace = null;
        static object RefFace = null;
        //static SolidEdgeGeometry.Edge RefEdge = null;
        static object RefEdge = null;
        //static SolidEdgeGeometry.Vertex RefVertex = null;
        static object RefVertex = null;

        static void Process_Dft(SolidEdgeDraft.DraftDocument draftDocument)

        {
            // Declare variables
            SolidEdgeFramework.Application SEApplication = null;
            ModelLinks oModelLinks = null;
            //ModelLink oModelLink = null;
            List<SolidEdgeDocument> Model_Links_List = new List<SolidEdgeDocument>();
            //SolidEdgeDocument Model;
            List<string> ModelLinks_List_NotFound = new List<string>();
            string TempOutfile = null;
            string TempDftOutfile = null;
            //int ModelLinksCount = 0;
            int i = 1;
            bool Existingflat = false;
            bool HasBeenAutoflatten = false;
            SummaryInfo SumInfo;
            DateTime LastSavedDateDFT;
            //System.IO.FileInfo infoReaderPDF, infoReaderDXF;
            //bool DXFUpdated, PDFUpdated;
            SolidEdgeFileProperties.PropertySets objPropertySets = new SolidEdgeFileProperties.PropertySets();
            SolidEdgeFileProperties.Properties objProperties;
            SolidEdgeFileProperties.Property objProperty = null;
            //string _path = System.IO.Path.GetDirectoryName(draftDocument.FullName); // Get the directory path of the DFT file
            bool _seekaccess;
            //FlattenItem Part_to_flat;
            //Tuple<bool, bool, string> tup;
            SheetMetalDocument sheetmetalcopy = null;
            string TempfileModelCopy = null;
            string PSMfileModelCopy = null;
            string Errs = null;

            //SEApplication = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);
            SEApplication = (SolidEdgeFramework.Application)
                  Marshal.GetActiveObject("SolidEdge.Application");

            SolidEdgeDocument oModeldocument = null;
            oModelLinks = draftDocument.ModelLinks;

            foreach (ModelLink oModelLink in oModelLinks)
            {
                try
                {
                    oModeldocument = (SolidEdgeDocument)oModelLink.ModelDocument;
                    if (oModeldocument.Type == SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument &&
                        !Model_Links_List.Contains(oModeldocument))
                    {
                        Model_Links_List.Add(oModeldocument);
                    }
                }
                catch
                {
                    // If an error occurs, it means the Modeldocument is not found. Add the file to the list.
                    if (!ModelLinks_List_NotFound.Contains(oModelLink.FileName))
                    {
                        ModelLinks_List_NotFound.Add(oModelLink.FileName);
                    }
                }
            }

            int MLinks_number = Model_Links_List.Count;

            switch (ModelLinks_List_NotFound.Count)
            {
                case 1:
                    MessageBox.Show($"No se ha encontrado el archivo {ModelLinks_List_NotFound[0]}. Como el plano solo tiene un modelo, no se generará ni PDF ni DXF.", "Información EFATOOLS: Creación de chapas desarrolladas", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;

                case 2:
                    string files = string.Join(", ", ModelLinks_List_NotFound);
                    MessageBox.Show($"No se han encontrado los archivos {files}. No se generarán ni PDFs ni DXFs de estos modelos.", "Información EFATOOLS: Creación de chapas desarrolladas", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
            }
            if (MLinks_number > 1)
            {
                // Crear un archivo temporal
                TempOutfile = System.IO.Path.GetTempFileName();

                // Cambiar la extensión del archivo temporal a "dft"
                TempDftOutfile = System.IO.Path.ChangeExtension(TempOutfile, "dft");

                // Guardar una copia del documento de plano en el archivo temporal
                draftDocument.SaveCopyAs(TempDftOutfile);
            }


            foreach (var model in Model_Links_List)
            {
                var Part_to_flat = new Flattenitem(model); // Crear un objeto FlattenItem para el modelo actual.

                var docs = SEApplication.Documents; // Obtener la colección de documentos abiertos en Solid Edge.

                foreach (SolidEdgeDocument opendoc in docs)
                {
                    if (opendoc.FullName == Part_to_flat.Fullname) // Comprobar si el nombre completo del documento coincide con el nombre completo del modelo.
                    {
                        Part_to_flat.Isopen = true; // Marcar el modelo como abierto.

                        // En lugar de cerrar todos los documentos y guardarlos, vamos a quitarles el acceso de escritura para poder leer sus propiedades y solo cerrar los que no se puedan poner en escritura.
                        opendoc.SeekWriteAccess(out _seekaccess);

                        if (!_seekaccess) // Si no se puede acceder al documento en modo de escritura, significa que está en un estado que lo impide (por ejemplo, abierto por otra persona o marcado como solo lectura).
                        {
                            string fulname = opendoc.FullName;
                            Marshal.FinalReleaseComObject(opendoc); // Liberar el objeto COM del documento.
                            docs.CloseDocument(fulname, false, null, null, true); // Cerrar el documento sin guardar cambios.
                            fulname = null;
                            Part_to_flat.IsReadOnlyAfterForce = true; // Establecer la bandera para poder abrir nuevamente el archivo cerrado.
                        }
                        else
                        {
                            Part_to_flat.IsReadOnlyAfterForce = false; // El documento se puede abrir en modo de escritura.
                            opendoc.SeekReadOnlyAccess(out _seekaccess); // Cambiar al modo de solo lectura para acceder a las propiedades del documento.
                            // Como el documento está abierto, se asigna directamente como un documento de chapa metálica.
                            var sheetMetalDocument = (SolidEdgePart.SheetMetalDocument)opendoc;
                        }

                        break; // Salir del bucle una vez que se ha encontrado el documento correspondiente al modelo.
                    }
                }

                // Extraer las fechas del DFT y PSM
                SumInfo = (SummaryInfo)draftDocument.SummaryInfo;
                LastSavedDateDFT = (DateTime)SumInfo.SaveDate; // Última vez que se guardó el DFT

                // Verificar primero si existen los archivos PDF y DXF
                // Luego comparar la fecha del PDF con la del DFT para saber si es necesario actualizarlo
                // Finalmente, comparar la fecha del DXF con la del PSM para saber si es necesario actualizarlo
                bool PDFUpdated = false; // Por defecto, consideramos que hay que crear ambos archivos
                bool DXFUpdated = false;

                if (File.Exists(Part_to_flat.OutfilePDF))
                {
                    FileInfo infoReaderPDF = new FileInfo(Part_to_flat.OutfilePDF);
                    // Si existe el PDF y se creó después del DFT, está actualizado
                    if (infoReaderPDF.LastWriteTime > LastSavedDateDFT)
                        PDFUpdated = true;
                }

                if (File.Exists(Part_to_flat.OutfileDXF))
                {
                    FileInfo infoReaderDXF = new FileInfo(Part_to_flat.OutfileDXF);
                    // Si existe el DXF y se creó después del PSM, está actualizado
                    if (infoReaderDXF.LastWriteTime > Part_to_flat.LastSavedDatePSM)
                        DXFUpdated = true;
                }

                if (DXFUpdated == false)
                {

                    switch (Part_to_flat.Status)
                    {
                        case SolidEdgeFramework.DocumentStatus.igStatusBaselined:
                        case SolidEdgeFramework.DocumentStatus.igStatusObsolete:
                        case SolidEdgeFramework.DocumentStatus.igStatusReleased:
                        case SolidEdgeFramework.DocumentStatus.igStatusInWork:
                            // Since the file is closed or in read-only mode, we can change its status from the properties and uncheck it before opening it.
                            objPropertySets.Open(Part_to_flat.Fullname, false);
                            objProperties = (SolidEdgeFileProperties.Properties)objPropertySets["ExtendedSummaryInformation"];
                            objProperty = (SolidEdgeFileProperties.Property)objProperties[2];


                            objProperty.Value = SolidEdgeConstants.DocumentStatus.igStatusAvailable;
                            objPropertySets.Save();
                            objPropertySets.Close();
                            break;
                    }

                    if (!Part_to_flat.Isopen || Part_to_flat.IsReadOnlyAfterForce)
                    {
                        SEApplication.DisplayAlerts = false;
                        SEApplication.ScreenUpdating = false;
                        // MyAddIn.Instance.Application.Interactive = false;
                        // MyAddIn.Instance.Application.DelayCompute = true;

                        sheetMetalDocument = (SheetMetalDocument)SEApplication.Documents.Open(Part_to_flat.Fullname, DocRelationAutoServer: 8);
                    }

                    if (sheetMetalDocument == null)
                    {
                        throw new System.Exception("No hay ningún documento de chapa activo.");

                    }
                    else
                    {
                        try
                        {
                            // Tenemos que ganar acceso de escritura sobre la chapa. Si la chapa está abierta, estará en solo lectura. Si está cerrada, podría estar emitida y haber dejado de estarlo o estar abierta por otro programa. Por eso usamos SeekWriteAccess y lo comprobamos.
                            sheetMetalDocument.SeekWriteAccess(out _seekaccess);

                            if (_seekaccess)
                            {
                                // Aquí es donde está todo el código para trabajar con la chapa, antes de llegar aquí.
                                var tup = FlatSheet(sheetMetalDocument, Autoflatten: true, Externalerrors: true);
                                Existingflat = tup.Item1;
                                HasBeenAutoflatten = tup.Item2;
                                Errs = tup.Item3;
                            }
                            else
                            {
                                // Si no hay acceso de escritura, la única forma de poder hacer el desarrollo aunque sea temporalmente es abrir una copia del archivo.
                                TempfileModelCopy = System.IO.Path.GetTempFileName();
                                PSMfileModelCopy = System.IO.Path.ChangeExtension(TempfileModelCopy, "psm");
                                sheetMetalDocument.SaveCopyAs(PSMfileModelCopy); // Creamos un temporal

                                sheetmetalcopy = (SheetMetalDocument)SEApplication.Documents.Open(PSMfileModelCopy, DocRelationAutoServer: 8);
                                var tup = FlatSheet(sheetmetalcopy, Autoflatten: true);
                                Existingflat = tup.Item1;
                                HasBeenAutoflatten = tup.Item2;
                                Errs = tup.Item3;
                            }

                            if (Errs != null)
                            {
                                MessageBox.Show(Errs, "Información EFATOOLS", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Manejar cualquier excepción aquí
                            MessageBox.Show(ex.StackTrace, ex.Message);
                        }

                        if (Existingflat)
                        {
                            try
                            {
                                RefFace = (object)null;
                                RefEdge = (object)null;
                                RefVertex = (object)null;

                                _models.SaveAsFlatDXFEx(Part_to_flat.OutfileDXF, RefFace, RefEdge, RefVertex, true);
                            }
                            catch
                            {

                                MessageBox.Show("No se ha podido guardar el DXF", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                // en caso de que no se pueda guardar el dxf buscamos si existe tanto dxf como pdf y los borramos para no introducir errores
                                if (File.Exists(Part_to_flat.OutfileDXF))
                                {

                                    try
                                    {
                                        File.Delete(Part_to_flat.OutfileDXF);
                                    }
                                    catch
                                    {
                                        MessageBox.Show("Al intentar crear el DXF no se pudo guardar, bien porque está abierto o en solo lectura, y existe una versión anterior que no se ha podido borrar. Si tienes el DXF abierto cierralo e intenta imprimir el documento de nuevo. Deberías borrar el DXF manualmente para no dar lugaar a errores.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    }

                                }

                                if (File.Exists(Part_to_flat.OutfilePDF))
                                {
                                    try
                                    {
                                        File.Delete(Part_to_flat.OutfilePDF);
                                    }
                                    catch
                                    {
                                        MessageBox.Show("Al intentar crear el PDF no se pudo guardar, bien porque está abierto o en solo lectura, y existe una versión anterior que no se ha podido borrar. Si tienes el DXF abierto cierralo e intenta imprimir el documento de nuevo. Deberías borrar el DXF manualmente para no dar lugaar a errores.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    }
                                }
                            }
                        }

                        else
                        {
                            MessageBox.Show("No se ha podido guardar el DXF porque la pieza " + sheetMetalDocument.FullName + " no tiene desarrollo. Deberías editarla y generar el desarrollo para que la próxima vez que se imprima el plano se pueda guardar", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                        if (HasBeenAutoflatten & File.Exists(Part_to_flat.OutfileDXF))
                        {
                            // abrimos dxf y metemos texto
                            Part_to_flat.AddTitletoDXF();
                        }

                        // 'antes de cerrar el archivo debemos poner el mismo estado que tenía si es que había sido cambiado
                        switch (Part_to_flat.Status)
                        {

                            case var @case when @case == SolidEdgeFramework.DocumentStatus.igStatusBaselined:
                            case var case1 when case1 == SolidEdgeFramework.DocumentStatus.igStatusObsolete:
                            case var case2 when case2 == SolidEdgeFramework.DocumentStatus.igStatusReleased:
                            case var case3 when case3 == SolidEdgeFramework.DocumentStatus.igStatusInWork:
                                {

                                    SolidEdgeFramework.PropertySets _objpropertysets = (SolidEdgeFramework.PropertySets)sheetMetalDocument.Properties;
                                    SolidEdgeFramework.Properties _objProperties = (SolidEdgeFramework.Properties)_objpropertysets.Item("ExtendedSummaryInformation");
                                    SolidEdgeFramework.Property _objproperty = default;

                                    try
                                    {
                                        _objproperty = (SolidEdgeFramework.Property)_objProperties.Item("Status");
                                    }
                                    catch
                                    {
                                        _objproperty = (SolidEdgeFramework.Property)_objProperties.Item("Estado");
                                    }

                                    _objproperty.set_Value(Part_to_flat.Status);
                                    sheetMetalDocument.Save();
                                    _objpropertysets = default;
                                    break;
                                }
                        }

                        // Si el fichero estaba abierto pero se le ha cambiado el estado, se ha cerrado, y ahora tenemos que traerlo a primer plano como estaba antes.
                        if (Part_to_flat.IsReadOnlyAfterForce)
                        {
                            SolidEdgeFramework.Window win = (SolidEdgeFramework.Window)sheetMetalDocument.Windows.Item("1");
                            win.WindowState = 2;
                        }

                        if (Part_to_flat.Isopen & Part_to_flat.IsReadOnly)
                        {
                            sheetMetalDocument.SeekReadOnlyAccess(out _seekaccess);
                        }

                        if (!Part_to_flat.Isopen)
                        {
                            string fulname = Part_to_flat.Fullname;
                            Marshal.FinalReleaseComObject(sheetMetalDocument);
                            sheetMetalDocument = null;
                            SEApplication.Documents.CloseDocument(fulname, false, Missing.Value, Missing.Value, true);
                            fulname = null;
                            // sheetMetalDocument.Close()
                            SEApplication.DoIdle();
                        }

                        if (sheetmetalcopy != null)
                        {
                            sheetmetalcopy.Close();              // cerramos el archivo copiado
                                                                 // borramos el archivo copiado
                            if (File.Exists(TempfileModelCopy))
                            {
                                File.Delete(TempfileModelCopy);
                            }
                            if (File.Exists(PSMfileModelCopy))
                            {
                                File.Delete(PSMfileModelCopy);
                            }
                            sheetmetalcopy = null;
                        }

                        draftDocument.Activate();

                    }
                }
                else
                {

                    //si no se actualiza el dxf las propieades han quedado abiertas y hay que cerrarlas, sino las cerramos cuando cambiamos el estado 
                    objPropertySets.Close();
                }

                if (PDFUpdated == false & File.Exists(Part_to_flat.OutfileDXF)) // si no hay pdf o este está desactualizado se crea el pdf siemore que exista el dxf
                {
                    try
                    {
                        CreatePDF(draftDocument, Part_to_flat.Fullname, Part_to_flat.OutfilePDF, TempDftOutfile, MLinks_number);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.StackTrace, ex.Message);
                    }
                }

                i += 1;

            }

            if (MLinks_number > 1)
            {
                // se borra el fichero temporal cuando se alcanza el numero de archivos contenidos
                if (File.Exists(TempDftOutfile))
                {
                    File.Delete(TempDftOutfile);
                }
                if (File.Exists(TempOutfile))
                {
                    File.Delete(TempOutfile);
                }
            }

            {
                var withBlock = SEApplication;
                withBlock.DisplayAlerts = true;
                withBlock.ScreenUpdating = true;
                // .Interactive = True
                // .DelayCompute = False
            }


            objPropertySets = null;
            objProperties = null;
            objProperty = null;
            sheetMetalDocument = null;
            RefFace = null;
            RefEdge = null;
            RefVertex = null;
            _models = null;

            Marshal.ReleaseComObject(SEApplication);
        }


        public static Tuple<bool, bool, string> FlatSheet(SheetMetalDocument sheetMetalDocument, bool ColorFlat = false, bool Autoflatten = true, bool Externalerrors = false, bool Save = true)
        {
            Model _model = null;
            SolidEdgeGeometry.Body _body = null;
            Shells _shells;
            Shell _shell;
            Faces _faces = null;
            Face objBiggestFace = null;
            Edges _edges = null;
            SolidEdgeGeometry.Edge LargestEdge = null;
            string Errors = null;

            FlatPatternModels flatPatternModels = null;
            FlatPatternModel flatPatternModel = null;
            FlatPattern flatPattern = null;
            bool Existingflat = false, Autoflat = false;

            SolidEdgeGeometry.Body _flat_body = null;
            Faces _flat_faces = null;

            Array RefKeyFace = Array.CreateInstance(typeof(byte), 0);
            Array RefKeyEdge = Array.CreateInstance(typeof(byte), 0);
            Array RefKeyVertex = Array.CreateInstance(typeof(byte), 0);

            double[] SizeXY = new double[2];

            string strRegex = string.Empty;
            double oMaxCutSizeX = 0.0, oMaxCutSizeY = 0.0;
            //bool oShowRangeBox = false;
            bool oAlarmOnX = false, oAlarmOnY = false;
            //bool oUseDefaultValues = true;

            try
            {
                Models _models = (Models)sheetMetalDocument.Models; // Get a reference to the Models collection.

                // Chequea si hay geometría en el modelo.
                if (_models.Count == 0)
                {
                    throw new System.Exception("No hay geometría en el modelo.");
                }

                foreach (Model model in _models)
                {
                    // Bucle que busca el modelo de diseño entre todos.

                    if (model.Name == "Design Model")
                    {
                        // Si el modelo es el de diseño.

                        _body = (SolidEdgeGeometry.Body)model.Body; // Obtiene una referencia al cuerpo del modelo.
                        _shells = (Shells)_body.Shells;
                        _shell = (Shell)_shells.Item(0);
                        _faces = (Faces)_shell.Faces; // Referencia todas las caras del modelo de diseño.
                        _model = model;

                        // Ejecutamos una función que nos devuelve la cara más grande.
                        objBiggestFace = GetLargestFace(_faces);

                        // Ejecutamos una función que nos devuelve la arista más grande de la cara más grande.
                        _edges = (Edges)objBiggestFace.Edges;
                        LargestEdge = GetLargestEdge(_edges);
                        break;
                    }
                    // Fin de la condición si el modelo es el de diseño.
                }

                CaraVista_Style(sheetMetalDocument);                    //generamos un estilo de cara vista si no existe
                flatPatternModels = sheetMetalDocument.FlatPatternModels;    // Get a reference to the FlatPatternModels collection,

                if (flatPatternModels.Count > 0)
                {
                    // Número de modelos de desarrollo, al menos tiene que haber uno.
                    // Asignamos el primero de los modelos, aunque solo puede haber uno.
                    flatPatternModel = (FlatPatternModel)flatPatternModels.Item(1);

                    if (!flatPatternModel.IsUpToDate)
                    {
                        flatPatternModel.Update(); // Si no está actualizado, se actualiza.
                    }

                    // Fin de condición de al menos un desarrollo.
                    if (flatPatternModel.FlatPatterns.Count > 0) // Tiene que haber un desarrollo
                    {
                        flatPattern = flatPatternModel.FlatPatterns.Item(1); // Asignamos el desarrollo


                        object description;
                        // Chequeo si es desarrollo está OK
                        if (flatPattern.Status[out description] == FeatureStatusConstants.igFeatureOK)
                        {

                            Existingflat = true;

                            // Cambiamos el nombre de todos los desplegados excepto los que coinciden con la cadena
                            strRegex = "AUTODESARROLLO|CHAPA DESARROLLADA";
                            if (MatchString(strRegex, flatPattern.Name) == "NOMATCH")
                            {
                                flatPattern.Name = "CHAPA DESARROLLADA";
                            }
                            else if (MatchString(strRegex, flatPattern.Name) == "AUTODESARROLLO")
                            {
                                Autoflat = true;
                            }

                            flatPatternModel.MakeActive();       // Necesario para obtener coordenadas modelo desarrollado

                            //Application.DoIdle();
                            sheetMetalDocument.Parent.DoIdle();

                            _flat_body = (SolidEdgeGeometry.Body)flatPatternModel.Body;
                            var FaceType = SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryPlane;
                            _flat_faces = (Faces)_flat_body.Faces[FaceType];

                            if (!ColorFlat)
                                ColorFaces(_flat_faces); // Buen momento para colorear las caras externas
                                                         // si se pasar colorflat como false es que no están coloreadas y se deben colorear
                            bool CutSizeOK = true;

                            // función para obtener la cara de referencia
                            RefKeyFace = GetReferenceFaceKey(flatPattern, _flat_faces, _faces, _model, flatPatternModel);
                            sheetMetalDocument.BindKeyToObject(RefKeyFace, out RefFace);         // asignamos la cara de referencia

                            try
                            {
                                flatPatternModel.GetCutSize(out SizeXY[0], out SizeXY[1]);               // Intentamos sacar las medidas de desarrollo
                                                                                                         // si no funciona aplicamos setsize
                            }
                            catch
                            {
                                try
                                {
                                    flatPattern.SetCutSizeValues(MaxCutSizeX: oMaxCutSizeX, MaxCutSizeY: oMaxCutSizeY, ShowRangeBox: true, AlarmOnX: oAlarmOnX, AlarmOnY: oAlarmOnY, UseDefaultValues: false);
                                }
                                catch (Exception)                                 // si tampoco funciona no se generan las medidas
                                {
                                    sheetMetalDocument.Application.StatusBar = "Problema al obtener las medidas de chapa desarrollada. Se va a proceder a rehacer el patrón de desarrollo intentando usar los mismos elementos de referencia que el existente";
                                    CutSizeOK = false;
                                }
                            }

                            // Rehacemos desarrollo por no conseguir las medidas de corte
                            if (CutSizeOK == false)
                            {

                                EdgeVertex RefEdgeVertex = GetReferenceEdgeANDVertex((Face)RefFace);

                                // función para obtener el eje de referencia
                                RefKeyEdge = RefEdgeVertex.ReferEdge;
                                sheetMetalDocument.BindKeyToObject(RefKeyEdge, out RefEdge);         // asignamos la arista de referencia

                                RefKeyVertex = RefEdgeVertex.ReferVertex;
                                sheetMetalDocument.BindKeyToObject(RefKeyVertex, out RefVertex);     // vertice de referencia

                                // para rehacer patrones que solo tienen aristas de referencia si borramos el patrón recuperamos la cara en el modelo desarrollado, pero primero tenemos que tener su reference key por eso lo hemos sacado del cutsize

                                try
                                {
                                    // Regeneramos el patrón
                                    flatPatternModel.FlatPatterns.Add(ReferenceEdge: RefEdge, ReferenceFace: RefFace, ReferenceVertex: RefVertex, ModelType: SolidEdgeConstants.FlattenPatternModelTypeConstants.igFlattenPatternModelTypeFlattenAnything);

                                    flatPattern.Delete();  // borro el anterior
                                    flatPattern = flatPatternModel.FlatPatterns.Item(1);
                                    flatPattern.Name = "DESARROLLO REGENERADO";

                                    flatPattern.SetCutSizeValues(MaxCutSizeX: oMaxCutSizeX, MaxCutSizeY: oMaxCutSizeY, ShowRangeBox: true, AlarmOnX: oAlarmOnX, AlarmOnY: oAlarmOnY, UseDefaultValues: false);
                                }
                                catch (Exception)
                                {
                                    string err = "La regeneración del patrón ha dado error \r\n";
                                    if (Externalerrors)
                                    {
                                        Errors += err;
                                    }
                                    else
                                    {
                                        MessageBox.Show(err, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    }
                                }
                            }
                        }                                         // Final de condición de chequeo del desarrollo

                        if (!Existingflat)
                        {
                            flatPattern.Delete();                        // como el patrón tiene un error lo borramos para generarlo de nuevo
                        }
                    }

                    else
                    {
                        Existingflat = false;
                    }                 // si hay modelo desarrollado pero no hay desarrollo pues hay que generarlo
                }
                else
                {
                    // No existen modelos de desarrollo.
                    Existingflat = false;
                }


                if (!Existingflat & Autoflatten == true)         // si no hay modelo desarrollado y elegimos desarrollar
                {

                    if (0 == flatPatternModels.Count)         // no existe ningún modelo de desarrollo por tanto habrá que crearlo
                    {
                        flatPatternModel = sheetMetalDocument.FlatPatternModels.Add(sheetMetalDocument.Models.Item(1));
                    }
                    else // por el contrario existe un modelo desarrollado pero no existe el desarrollo
                    {
                        flatPatternModel = sheetMetalDocument.FlatPatternModels.Item(1);
                    }

                    // usamos dos metodos distintos dependiendo de si la arista es lineal o no
                    switch (true)
                    {
                        case object _ when LargestEdge.Type == GNTTypePropertyConstants.igLine:
                            {
                                flatPatternModel.FlatPatterns.Add(ReferenceEdge: LargestEdge, ReferenceFace: objBiggestFace, ReferenceVertex: LargestEdge.EndVertex, ModelType: FlattenPatternModelTypeConstants.igFlattenPatternModelTypeFlattenAnything);
                                break;
                            }

                        default:
                            {
                                flatPatternModel.FlatPatterns.Add(ReferenceEdge: LargestEdge, ReferenceFace: objBiggestFace, ReferenceVertex: LargestEdge, ModelType: FlattenPatternModelTypeConstants.igFlattenPatternModelTypeFlattenAnything);
                                break;
                            }
                    }

                    // aquí ahora deberíamos pintar la cara de referencia a ver si así se pintan las dos

                    flatPattern = flatPatternModel.FlatPatterns.Item(1);
                    flatPattern.Name = "AUTODESARROLLO " + DateTime.Now.ToString();
                    Existingflat = true;
                    Autoflat = true;

                    try
                    {
                        flatPattern.SetCutSizeValues(MaxCutSizeX: oMaxCutSizeX, MaxCutSizeY: oMaxCutSizeY, ShowRangeBox: true, AlarmOnX: oAlarmOnX, AlarmOnY: oAlarmOnY, UseDefaultValues: false);
                    }
                    catch (Exception)
                    {
                        string err = "La colocación automática de las cotas del tamaño de corte ha fallado. Deberás mostrar tamaño de corte y dimensiones manualmente \r\n";
                        if (Externalerrors)
                        {
                            Errors += err;
                        }
                        else
                        {
                            MessageBox.Show("", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }

                }                                     // fin condición de no hay modelo desarrollado

                _model.MakeActive();

                if (Save)
                {

                    try
                    {
                        sheetMetalDocument.Save();
                    }
                    catch
                    {
                        string err = "No se ha podido guardar en el fichero del modelo la actualización del desarrollo porque el archivo está en solo lectura \r\n";
                        if (Externalerrors)
                        {
                            Errors += err;
                        }
                        else
                        {
                            MessageBox.Show(err, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }

                }


                return new Tuple<bool, bool, string>(Existingflat, Autoflat, Errors);

                //oUseDefaultValues = true;
                objBiggestFace = null;
                LargestEdge = null;
                oMaxCutSizeX = 0.0;
                oMaxCutSizeY = 0.0;
                //oShowRangeBox = false;
                oAlarmOnX = false;
                oAlarmOnY = false;
                _body = null;
                _shell = null;
                _shells = null;
                _flat_body = null;
                _model = null;
                _body = null;
                _shells = null;
                _shell = null;
                _faces = null;
                _edges = null;
                flatPatternModels = null;
                flatPatternModel = null;
                flatPattern = null;
                Existingflat = false;
                _flat_body = null;
                //_flat_shells = null;
                //_flat_shell = null;
                _flat_faces = null;
                RefKeyFace = null;
                SizeXY[0] = 0.0;
                SizeXY[1] = 0.0;
                strRegex = string.Empty;
                RefKeyEdge = null;
                RefKeyVertex = null;

            }


            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

            return new Tuple<bool, bool, string>(Existingflat, Autoflat, Errors);

        }
        internal static Face GetLargestFace(Faces ofaces)
        {
            double DLargest = 0.0; // Área máxima de la cara más grande.
            Face GetLargestFace = null;
            foreach (Face oface in ofaces)
            {
                // Recorremos todas las caras del shell buscando la de mayor superficie.
                if (oface.Area > DLargest)
                {
                    DLargest = oface.Area;
                    GetLargestFace = oface;
                }
            }
            return GetLargestFace;
        }
        internal static Edge GetLargestEdge(Edges oedges)
        {
            double minParam = 0.0, maxParam = 0.0, Length = 0.0;
            double dLargest_igline = 0.0;
            double dLargest_igcurve = 0.0;
            Edge LargestEdge_igline = null;
            Edge LargestEdge_igcurve = null;
            Edge GetLargestEdge = null;

            // Buscamos primero entre todas las aristas lineales.
            foreach (Edge oEdge in oedges)
            {
                // Recorremos todas las aristas buscando un vértice.

                oEdge.GetParamExtents(MinParam: out minParam, MaxParam: out maxParam);
                oEdge.GetLengthAtParam(FromParam: minParam, ToParam: maxParam, Length: out Length);

                switch (oEdge.Geometry)
                {
                    case SolidEdgeConstants.GNTTypePropertyConstants.igLine:
                        if (Length > dLargest_igline)
                        {
                            // Cuando una arista lineal es > que la acumulada, nos quedamos con esa.
                            LargestEdge_igline = oEdge;
                            dLargest_igline = Length;
                        }
                        break;

                    default:
                        if (Length > dLargest_igcurve)
                        {
                            // Cuando una arista no lineal es > que la acumulada, nos quedamos con esa.
                            LargestEdge_igcurve = oEdge;
                            dLargest_igcurve = Length;
                        }
                        break;
                }
            }

            // Si no hay ninguna arista lineal, usamos la mayor de las curvas.
            if (LargestEdge_igline == null)
            {
                GetLargestEdge = (Edge)LargestEdge_igcurve;
            }
            else
            {
                GetLargestEdge = (Edge)LargestEdge_igline;
            }

            return GetLargestEdge;
        }
        internal static void CaraVista_Style(SolidEdgePart.SheetMetalDocument oSheetmetaldocument)
        {
            // Necesitamos crear un estilo de cara nuevo porque no todos los archivos tienen los mismos estilos.
            bool flag = false;
            SolidEdgeFramework.FaceStyles _facestyles = (SolidEdgeFramework.FaceStyles)oSheetmetaldocument.FaceStyles;
            foreach (SolidEdgeFramework.FaceStyle style in _facestyles)
            {
                if (style.StyleName == "Cara vista")
                {
                    flag = true;
                    break;
                }
            }
            if (!flag)
            {
                SolidEdgeFramework.FaceStyle seFaceStyle = _facestyles.Add("Cara vista", "");
                seFaceStyle.AcceptsShadows = 1;
                seFaceStyle.CastsShadows = 1;
                seFaceStyle.AmbientBlue = 0.43f;
                seFaceStyle.AmbientGreen = 0.68f;
                seFaceStyle.AmbientRed = 0.4f;
                seFaceStyle.BumpmapFileName = "";
                seFaceStyle.BumpmapHeight = 1;
                seFaceStyle.BumpmapInvert = 0;
                seFaceStyle.BumpmapMirrorX = 0;
                seFaceStyle.BumpmapMirrorY = 0;
                seFaceStyle.BumpmapOffsetX = 0;
                seFaceStyle.BumpmapOffsetY = 0;
                seFaceStyle.BumpmapRotation = 0;
                seFaceStyle.BumpmapScaleX = 1;
                seFaceStyle.BumpmapScaleY = 1;
                seFaceStyle.BumpmapUnits = 1;
                seFaceStyle.DiffuseBlue = 0.62f;
                seFaceStyle.DiffuseGreen = 0.2054f;
                seFaceStyle.DiffuseRed = 0.79f;
                seFaceStyle.EmissionBlue = 0;
                seFaceStyle.EmissionGreen = 0;
                seFaceStyle.EmissionRed = 0;
                seFaceStyle.Flags = 1;
                seFaceStyle.LineWidth = 1;
                seFaceStyle.Opacity = 1;
                seFaceStyle.Reflectivity = 0;
                seFaceStyle.Refraction = 1;
                //seFaceStyle.ShaderType = 0;
                seFaceStyle.Shininess = 0.1f;
                // seFaceStyle.SkyboxAltitude = 0;
                // seFaceStyle.SkyboxAzimuth = 0;
                // seFaceStyle.SkyboxConeAngle = 60;
                // seFaceStyle.SkyboxRoll = 0;
                // seFaceStyle.SkyboxType = -1;
                seFaceStyle.SpecularBlue = 0.36f;
                seFaceStyle.SpecularGreen = 0.036f;
                seFaceStyle.SpecularRed = 0.2368f;
                seFaceStyle.StipplePattern = 65535;
                seFaceStyle.StippleScale = 1;
                seFaceStyle.TextureMirrorX = 0;
                seFaceStyle.TextureMirrorY = 0;
                seFaceStyle.TextureOffsetX = 0;
                seFaceStyle.TextureOffsetY = 0;
                seFaceStyle.TextureRotation = 0;
                seFaceStyle.TextureScaleX = 1;
                seFaceStyle.TextureScaleY = 1;
                seFaceStyle.TextureTransparent = 0;
                seFaceStyle.TextureTransparentColorBlue = 0;
                seFaceStyle.TextureTransparentColorGreen = 0;
                seFaceStyle.TextureTransparentColorRed = 0;
                seFaceStyle.TextureUnits = 1;
                seFaceStyle.TextureWeight = 1;
                seFaceStyle.Type = 1;
                seFaceStyle.WidthSpace = 0;
                seFaceStyle.WireframeColorBlue = 0;
                seFaceStyle.WireframeColorGreen = 0;
                seFaceStyle.WireframeColorRed = 0.6f;
            }
        }
        internal static string MatchString(string strregex, string ostring)
        {
            string MatchStringRet = default;
            MatchStringRet = "";
            Regex myRegex;
            Match myMatch;
            myRegex = new Regex(strregex, RegexOptions.None);
            myMatch = myRegex.Match(ostring);
            if (myMatch.Success)
            {
                MatchStringRet = myMatch.Value;
            }
            else
            {
                MatchStringRet = "NOMATCH";
            }
            return MatchStringRet;
        }
        internal static void ColorFaces(Faces ofaces)
        {
            foreach (Face oface in ofaces)
            {
                if (TopFace(oface))
                {
                    //oface.Style =sheetMetalDocument.FaceStyles("Cara vista");
                    oface.Style.StyleName = "Cara vista";

                }
            }

        }
        internal static bool TopFace(Face face)
        {

            var dblMinRange = new double[3];
            var dblMaxRange = new double[3];
            var dblNormal = new double[4];

            // Variables para saber las caras que son dobleces
            var Bendexists = default(bool);
            object Endpoints = null;
            object AttrbuteVersion = null;
            object BendRadius = null;
            object BendAngle = null;
            object BendSweepAngle = null;
            object BendOrientation = null;

            face.GetParamRange(MinParam: dblMinRange, MaxParam: dblMaxRange);   // Get the parametric range of the face.

            face.GetNormal(NumParams: 1, Params: dblMinRange, Normals: dblNormal); // Get the parametric range of the face.

            // Chequear si la cara es de doblado
            face.GetBendCenterLineAttributesEx(out Bendexists, out Endpoints, out AttrbuteVersion, out BendRadius, out BendAngle, out BendSweepAngle, out BendOrientation);

            // Check to see if it's pointed in the positive Z direction.
            if (dblNormal[2] > 0.9d & !Bendexists)
            {

                return true;  // one of the top faces and not bend face
            }
            else
            {
                return false;
            }
        }
        internal static Array GetReferenceFaceKey(FlatPattern oFlatpattern, Faces oFlatFaces, Faces oFaces, Model model, FlatPatternModel oflatpatternmodel)
        {
            // Inicialización de variables
            Face oFlatFace = default;
            var Reference_FaceKey = Array.CreateInstance(typeof(byte), 0);
            Edges reference_edges = default; // Aristas de referencia
            Face reference_face = oFlatpattern.Reference as Face; // Cara de referencia
            object keysize;
            // Comprobación de la existencia de la cara de referencia
            if (reference_face != null)
            {
                // Si existe, obtenemos su clave de referencia
                reference_face.GetReferenceKey(ReferenceKey: Reference_FaceKey, out keysize);
            }
            else // Si no existe la cara de referencia
            {
                try
                {
                    // Asignamos las aristas de referencia
                    reference_edges = (Edges)oFlatpattern.Reference;

                    var num = default(int);
                    bool Found = false;
                    var _faces = Array.CreateInstance(typeof(Face), 0);

                    // Recorremos las aristas de referencia
                    foreach (Edge refedge in reference_edges)
                    {
                        //model.Parent.Application.DoIdle();

                        refedge.GetFaces(NumFaces: out num, Faces: _faces);

                        // Buscamos la cara de referencia en las caras asociadas a la arista
                        foreach (Face currentOFlatFace in _faces)
                        {
                            oFlatFace = currentOFlatFace;
                            if (TopFace(oFlatFace))
                            {
                                Found = true;
                                break;
                            }
                        }
                        if (Found)
                            break;
                        oFlatFace = default;
                    }

                    // Si encontramos la cara de referencia en el desplegado
                    if (oFlatFace != null)
                    {
                        // Obtenemos su clave de referencia
                        oFlatFace.GetReferenceKey(ReferenceKey: Reference_FaceKey, out keysize);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("No existe ni cara de referencia ni aristas de referencia", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

            // Devolvemos la clave de referencia de la cara
            return Reference_FaceKey;
        }
        internal static EdgeVertex GetReferenceEdgeANDVertex(Face oRefFace)
        {
            EdgeVertex GetReferenceEdgeANDVertexRet = default;
            // vamos a devolver la arista como un referencekey, solid edge no almacena la arista como referencia, solo podemos sacarla sabiendo en el desplegado cual está alineada con el eje X

            var StartPoint = Array.CreateInstance(typeof(double), 0);       // coordenadas del inicio
            var EndPoint = Array.CreateInstance(typeof(double), 0);         // coordenadas del final
            var Reference_EdgeKey = Array.CreateInstance(typeof(byte), 0);
            var Reference_VertexKey = Array.CreateInstance(typeof(byte), 0);
            var RefVertex = new Vertex[3];
            Edges RefEdges = default;
            var RefEdge = new Edge[3];

            RefEdges = (Edges)oRefFace.Edges;               // Asociamos todas las aristas de la cara de referencia

            foreach (Edge oedge in RefEdges)
            {
                if (oedge.Type == GNTTypePropertyConstants.igLine)               // solo considera aristas lineales 
                {

                    oedge.GetEndPoints(ref StartPoint, ref EndPoint);

                    bool exitFor = false;
                    switch (true)
                    {

                        // buscamos aquella arista alineada con el eje x y que pasa por 0, aquella que tiene la y a 0
                        case object _ when Round((double)StartPoint.GetValue(1)) == 0 & Round((double)EndPoint.GetValue(1)) == 0:
                            {
                                RefEdge[0] = oedge;

                                if (Round((double)StartPoint.GetValue(0)) == 0)                                        // punto de inicio en origen?
                                {
                                    RefVertex[0] = (Vertex)oedge.StartVertex;
                                }
                                else
                                {
                                    RefVertex[0] = (Vertex)oedge.EndVertex;
                                }                // Buscamos el vertice en el origen

                                exitFor = true;
                                break;
                            }

                        // Si no pasa por el origen al menos nos quedamos con la arista alineada
                        case object _ when Round((double)StartPoint.GetValue(1)) == Round((double)EndPoint.GetValue(1)):
                            {
                                RefEdge[1] = oedge;
                                // Como no tenemos arista que pase por origen elegiremos el vertice más cerca del origen

                                if (Round((double)StartPoint.GetValue(0)) < Round((double)EndPoint.GetValue(0)))                                 // elegimos el vertice mas a la izquierda
                                {
                                    RefVertex[1] = (Vertex)oedge.StartVertex;
                                }
                                else
                                {
                                    RefVertex[1] = (Vertex)oedge.EndVertex;
                                    // si ni una cosa ni la otra o la arista esta orientada con y o z o  esta inclinada
                                }

                                break;
                            }
                        default:
                            {
                                // no queremos aristas inclinadas pero si nos interesan la demas por si no tenemos nada para orientar
                                // así que como no está orientadda con y que al menos esté con x o z
                                if (Round((double)StartPoint.GetValue(1)) == Round((double)EndPoint.GetValue(1)) | Round((double)StartPoint.GetValue(2)) == Round((double)EndPoint.GetValue(2)))
                                {
                                    RefEdge[2] = oedge;
                                    RefVertex[2] = (Vertex)oedge.StartVertex;
                                }
                                break;
                            }
                    }

                    if (exitFor)
                    {
                        break;
                    }
                }
            }

            if (RefVertex is null)                                                      // para aquellos casos que no hay vertice porque todas las aristas son curvas o están en diagonal
            {
                RefEdge[0] = (Edge)RefEdges.Item(1);
                RefVertex[0] = (Vertex)RefEdge[0].StartVertex;
            }
            object keysize = 0;
            switch (true)
            {

                case object _ when RefEdge[0] != null:
                    {
                        RefEdge[0].GetReferenceKey(ReferenceKey: Reference_EdgeKey, out keysize);              // Referencekey de la arista
                        RefVertex[0].GetReferenceKey(ReferenceKey: Reference_VertexKey, out keysize);          // Referencekey del vertice
                        break;
                    }
                case object _ when RefEdge[1] != null:
                    {
                        RefEdge[1].GetReferenceKey(ReferenceKey: Reference_EdgeKey, out keysize);              // Referencekey de la arista
                        RefVertex[1].GetReferenceKey(ReferenceKey: Reference_VertexKey, out keysize);          // Referencekey del vertice
                        break;
                    }
                case object _ when RefEdge[2] != null:
                    {
                        RefEdge[2].GetReferenceKey(ReferenceKey: Reference_EdgeKey, out keysize);              // Referencekey de la arista
                        RefVertex[2].GetReferenceKey(ReferenceKey: Reference_VertexKey, out keysize);          // Referencekey del vertice
                        break;
                    }

            }

            GetReferenceEdgeANDVertexRet.ReferEdge = Reference_EdgeKey;
            GetReferenceEdgeANDVertexRet.ReferVertex = Reference_VertexKey;

            return GetReferenceEdgeANDVertexRet;
        }
        internal partial struct EdgeVertex
        {
            public Array ReferVertex;
            public Array ReferEdge;
        }
        internal static double Round(double value)
        {
            return Math.Round(value * 1000d, 3);
        }

        internal static void CreatePDF(SolidEdgeDraft.DraftDocument odraftDocument, string File, string outfilePDF, string tempoutfile, int number_of_files)
        {

            SolidEdgeDraft.DraftDocument tempDraftDocument = default;
            SolidEdgeDocument oModeldocument = default;
            SolidEdgeDraft.Section oSection;

            if (number_of_files > 1)
            {
                // abrimos el temporal
                try
                {
                    tempDraftDocument = (SolidEdgeDraft.DraftDocument)odraftDocument.Parent.Documents.Open(tempoutfile);
                }
                catch
                {
                    MessageBox.Show("No se  ha podido abrir el fichero temporal para crear el PDF", "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                foreach (SolidEdgeDraft.ModelLink oModellink in tempDraftDocument.ModelLinks) // itera sobre todos los archivos del plano
                {
                    try
                    {
                        oModeldocument = (SolidEdgeDocument)oModellink.ModelDocument;    // si hay algún archivo que no encuentra da error, nunca va a passar por aqui porque esos ficheros ni se almacenan en la lista
                    }
                    catch
                    {
                        oModeldocument = default;
                        oModellink.Delete(); // borramos la vista del archivo que no encuentra
                    }

                    if (oModeldocument != null)            // para prevenir cuando no se encuentra el archivo
                    {
                        if (oModeldocument.FullName != File)
                        {
                            oModellink.Delete();
                        }
                    }

                    // tambien borramos las hojas que se han quedado vacias
                    oSection = tempDraftDocument.Sections.Item("Section1");
                    foreach (SolidEdgeDraft.Sheet osheet in oSection.Sheets)
                    {
                        // si hay una hoja vacia la eliminamos
                        if (((SolidEdgeDraft.DrawingViews)osheet.DrawingViews).Count == 0)
                        {
                            osheet.Delete();
                        }
                    }
                }
                tempDraftDocument.SaveAs(outfilePDF);
                tempDraftDocument.Close(false);
            }

            else
            {

                odraftDocument.SaveAs(outfilePDF);

            }


            odraftDocument.Parent.DoIdle();
        }

        public static string Process_DftEX(SolidEdgeDraft.DraftDocument draftDocument)

        {
            // Declare variables
            SolidEdgeFramework.Application SEApplication = null;
            ModelLinks oModelLinks = null;
            //ModelLink oModelLink = null;
            List<SolidEdgeDocument> ModelDocument_List = new List<SolidEdgeDocument>();
            List<string> ModelDocuments_List_NotFound = new List<string>();
            int i = 1;
            SolidEdgeFileProperties.PropertySets objPropertySets = new SolidEdgeFileProperties.PropertySets();
            string Errors = null;
            SolidEdgePart.Models _models = null;
            DateTime LastSavedDateDFT;
            SummaryInfo SumInfo;
            string TempDftOutfile = null;

            //conectamos con la aplicación;
            //SEApplication = PrintBom.Mainform.ConnectToSolidEdge(false);
            SolidEdgeCommunity.OleMessageFilter.Register();
            var docManager = new SolidEdgeDocumentManager(false);
            SEApplication = docManager.SEApplication;

            SolidEdgeDocument oModeldocument = null;
            oModelLinks = draftDocument.ModelLinks;

            //Hace dos listas de enlaces a chapas, una con enlaces hayados y otra con los que no se pueden encontrar          
            foreach (ModelLink oModelLink in oModelLinks)
            {
                try
                {
                    oModeldocument = (SolidEdgeDocument)oModelLink.ModelDocument;
                    if (oModeldocument.Type == SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument &&
                        !ModelDocument_List.Contains(oModeldocument))
                    { ModelDocument_List.Add(oModeldocument); }
                }
                catch
                {
                    // If an error occurs, it means the Modeldocument is not found. Add the file to the list.
                    if (!ModelDocuments_List_NotFound.Contains(oModelLink.FileName))
                    {
                        ModelDocuments_List_NotFound.Add(oModelLink.FileName);
                    }
                }
            }

            int MLinks_number = ModelDocument_List.Count;

            //en el caso de que haya archivos desenlazados
            switch (ModelDocuments_List_NotFound.Count)
            {
                case 1:
                    Errors += "No se ha encontrado el archivo" + ModelDocuments_List_NotFound[0] + ". Como el plano solo tiene un modelo, no se generará ni PDF ni DXF.\r\n";
                    PrintBom.Mainform._log.Info(String.Format(Errors));
                    break;

                case int n when n >= 1:
                    string files = string.Join(", ", ModelDocuments_List_NotFound);
                    Errors += "No se han encontrado los archivos " + files + ". No se generarán ni PDFs ni DXFs de estos modelos.";
                    PrintBom.Mainform._log.Info(String.Format(Errors));
                    break;

            }

            //iteramos sobre todos los modelos del plano
            foreach (var modeldocument in ModelDocument_List)
            {
                var Part_to_flat = new Flattenitem(modeldocument); // Crear un objeto FlattenItem para el modelo actual.

                // Extraer las fechas del DFT y PSM
                SumInfo = (SummaryInfo)draftDocument.SummaryInfo;
                LastSavedDateDFT = (DateTime)SumInfo.SaveDate; // Última vez que se guardó el DFT

                // Verificar primero si existen los archivos PDF y DXF
                // Luego comparar la fecha del PDF con la del DFT para saber si es necesario actualizarlo
                // Finalmente, comparar la fecha del DXF con la del PSM para saber si es necesario actualizarlo
                
                bool PDFUpdated = false; // Por defecto, consideramos que hay que crear ambos archivos
                bool DXFUpdated = false;

                if (File.Exists(Part_to_flat.OutfilePDF))
                {
                    FileInfo infoReaderPDF = new FileInfo(Part_to_flat.OutfilePDF);
                    // Si existe el PDF y se creó después del DFT, está actualizado
                    if (infoReaderPDF.LastWriteTime > LastSavedDateDFT)
                        PDFUpdated = true;
                }

                if (File.Exists(Part_to_flat.OutfileDXF))
                {
                    FileInfo infoReaderDXF = new FileInfo(Part_to_flat.OutfileDXF);
                    // Si existe el DXF y se creó después del PSM, está actualizado
                    if (infoReaderDXF.LastWriteTime > Part_to_flat.LastSavedDatePSM)
                        DXFUpdated = true;
                }

                // si el dxf no está creado
                if (DXFUpdated == false)
                {

                    try
                    {
                        //abrimos el archivo en modo oculto.
                        sheetMetalDocument = (SheetMetalDocument)SEApplication.Documents.Open(Part_to_flat.Fullname, DocRelationAutoServer: 8);

                        if (isFlatten(sheetMetalDocument))
                        {
                            // Si hay un modelo desarrollado guardamos el dxf, pero tendríamos que borrar el dxf si está ya guardado para sobreescribirlo siempre que la fecha del dft sea mas moderna que la del dxf
                            RefFace = (object)null;
                            RefEdge = (object)null;
                            RefVertex = (object)null;
                            _models = sheetMetalDocument.Models;

                            if (File.Exists(Part_to_flat.OutfileDXF))
                            {
                                try
                                {
                                    File.Delete(Part_to_flat.OutfileDXF);

                                    _models.SaveAsFlatDXFEx(Part_to_flat.OutfileDXF, RefFace, RefEdge, RefVertex, true);
                                }
                                catch
                                {
                                    PrintBom.Mainform._log.Error("Al intentar crear el DXF no se pudo guardar, bien porque está abierto o en solo lectura, y existe una versión anterior que no se ha podido borrar");

                                }

                            }
                            else
                            {
                                _models.SaveAsFlatDXFEx(Part_to_flat.OutfileDXF, RefFace, RefEdge, RefVertex, true);
                            }



                            if (File.Exists(Part_to_flat.OutfilePDF))
                            {
                                try
                                {
                                    File.Delete(Part_to_flat.OutfilePDF);
                                }
                                catch
                                {
                                    PrintBom.Mainform._log.Error("Al intentar crear el PDF no se pudo guardar, bien porque está abierto o en solo lectura, y existe una versión anterior que no se ha podido borrar.");
                                }
                            }



                            SEApplication.Documents.CloseDocument(Part_to_flat.Fullname, false, Missing.Value, Missing.Value, true);
                            SEApplication.DoIdle();
                        }

                    }

                    catch
                    {
                        PrintBom.Mainform._log.Error(String.Format("No se puede abrir {0}", Part_to_flat.Fullname));
                    }

                }

                if (PDFUpdated == false & File.Exists(Part_to_flat.OutfileDXF)) // si no hay pdf o este está desactualizado se crea el pdf siemore que exista el dxf
                {
                    try
                    {//OJO CON LOS PROBLEMAS QUE DE VEZ EN CUANDO DA
                        CreatePDF(draftDocument, Part_to_flat.Fullname, Part_to_flat.OutfilePDF, TempDftOutfile, MLinks_number);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.StackTrace, ex.Message);
                    }
                } 

            }


            objPropertySets = null;
            sheetMetalDocument = null;
            RefFace = null;
            RefEdge = null;
            RefVertex = null;
            _models = null;

            Marshal.ReleaseComObject(SEApplication);
            SolidEdgeCommunity.OleMessageFilter.Revoke();

            return Errors;
        }


        public static bool isFlatten(SheetMetalDocument sheetMetalDocument)
        {

            FlatPatternModels flatPatternModels;
            FlatPatternModel flatPatternModel;
            FlatPattern flatPattern;
            bool Existingflat = false;

            try
            {
                // referencia a la coleccion de modelos de desarrollo
                flatPatternModels = sheetMetalDocument.FlatPatternModels;

                if (flatPatternModels.Count > 0)
                {
                    // Número de modelos de desarrollo, al menos tiene que haber uno.
                    // Asignamos el primero de los modelos, aunque solo puede haber uno.
                    flatPatternModel = (FlatPatternModel)flatPatternModels.Item(1);

                    if (!flatPatternModel.IsUpToDate) flatPatternModel.Update(); // Si no está actualizado, se actualiza.

                    if (flatPatternModel.FlatPatterns.Count > 0) // Tiene que haber algún desarrollo
                    {
                        flatPattern = flatPatternModel.FlatPatterns.Item(1); // Asignamos el desarrollo
                        object description;

                        // Si el desarrollo está ok devuelvo que hay un desarrollo
                        if (flatPattern.Status[out description] == FeatureStatusConstants.igFeatureOK)
                        {
                            Existingflat = true;

                            //flatPatternModel.MakeActive();       // Necesario para obtener coordenadas modelo desarrollado

                            //sheetMetalDocument.Parent.DoIdle();
                        }

                    }

                    else
                    {
                        Existingflat = false;
                    }
                }
                else
                {
                    // No existen modelos de desarrollo.
                    Existingflat = false;
                }


            }


            catch

            {
                Existingflat = false;
            }

            return Existingflat;

        }
    }



}
