using SolidEdgeFramework;
using System;
using System.Windows.Forms;

namespace PrintBom
{
    public class Flattenitem
    {

        private SolidEdgeDocument ModelDocument; // Declaración de una variable privada para el documento de Solid Edge.

        public Flattenitem(SolidEdgeDocument oModelDocument)
        {

            // Constructor que recibe un objeto SolidEdgeDocument como parámetro.
            // Asigna el objeto recibido a la variable ModelDocument.
            this.ModelDocument = oModelDocument;

            // Obtener información sobre el documento.
            IsReadOnly = oModelDocument.ReadOnly; // Indica si el documento está en modo de solo lectura.
            Status = oModelDocument.Status; // Estado del documento.
            Fullname = ModelDocument.FullName; // Nombre completo del documento.

            // Obtener información adicional del resumen del documento.
            var summaryInfo = (SolidEdgeFramework.SummaryInfo)oModelDocument.SummaryInfo;
            DocumentNumber = summaryInfo.DocumentNumber; // Número de documento.
            LastSavedDatePSM = Convert.ToDateTime(summaryInfo.SaveDate); // Fecha de la última vez que se guardó el documento.
        }

        public string Fullname { get; set; } // Propiedad para el nombre completo del documento.
        public bool Isopen { get; set; } = false; // Propiedad para indicar si el documento está abierto.
        public bool IsReadOnlyAfterForce { get; set; } = false; // Propiedad para indicar si el documento está en modo de solo lectura después de forzarlo.
        public bool IsReadOnly { get; set; } // Propiedad para indicar si el documento está en modo de solo lectura.
        public DocumentStatus Status { get; set; } // Propiedad para el estado del documento.
        public string DocumentNumber { get; set; } // Propiedad para el número de documento.
        public DateTime LastSavedDatePSM { get; set; } // Propiedad para la fecha de la última vez que se guardó el documento.
        public string OutfilePDF
        {
            get
            {
                // Devuelve la ruta completa del archivo PDF utilizando el nombre completo del documento y el número de documento.
                return Extensions.PathFilename(Fullname, DocumentNumber.ToString(), ".pdf");
            }
            set
            {
                // No se realiza ninguna acción en el setter.
            }
        }
        public string OutfileDXF
        {
            get
            {
                // Devuelve la ruta completa del archivo DXF utilizando el nombre completo del documento y el número de documento.
                return Extensions.PathFilename(Fullname, DocumentNumber.ToString(), ".dxf");
            }
            set
            {
                // No se realiza ninguna acción en el setter.
            }
        }
        public void AddTitletoDXF()
        {
            // Cargar el archivo DXF.
            var dxf = netDxf.DxfDocument.Load(this.OutfileDXF);
            int n = 0;
            netDxf.Header.DxfVersion version = netDxf.DxfDocument.CheckDxfFileVersion(this.OutfileDXF, out bool boolValue);

            try
            {
                n = 2;
                // IReadOnlyList<netDxf.Entities.Line> _lines;
                n = 4;
                // _lines = dxf.Lines;
                n = 5;
                double MinX = double.MaxValue;
                double MinY = double.MaxValue;
                double MaxX = double.MinValue;
                double MaxY = double.MinValue;
                double BasepointX;
                double BasepointY;

                n = 3;

                foreach (var line in dxf.Entities.Lines)
                {
                    // Calculate the bounding box of the lines.
                    if (MaxX < line.StartPoint.X) MaxX = line.StartPoint.X;
                    if (MaxX < line.EndPoint.X) MaxX = line.EndPoint.X;

                    if (MinX > line.StartPoint.X) MinX = line.StartPoint.X;
                    if (MinX > line.StartPoint.X) MinX = line.StartPoint.X;

                    if (MaxY < line.StartPoint.Y) MaxY = line.StartPoint.Y;
                    if (MaxY < line.EndPoint.Y) MaxY = line.EndPoint.Y;

                    if (MinY > line.StartPoint.Y) MinY = line.StartPoint.Y;
                    if (MinY > line.StartPoint.Y) MinY = line.StartPoint.Y;
                }

                BasepointX = (MinX + MaxX) / 2;
                BasepointY = (MinY + MaxY) / 2;
                n = 4;

                double largo = Math.Abs(MinX) + Math.Abs(MaxX);
                double ancho = Math.Abs(MinY) + Math.Abs(MaxY);

                var _text = new netDxf.Entities.Text();
                var vector = new netDxf.Vector3(BasepointX, BasepointY, 0.0);
                _text.Position = vector;
                _text.Alignment = netDxf.Entities.TextAlignment.MiddleCenter;

                if (ancho > (3 * largo))
                {
                    _text.Rotation = 90;
                    if (ancho > 40)
                    {
                        _text.Height = Math.Round(Math.Abs(ancho / 60));
                    }
                    else
                    {
                        _text.Height = 1;
                    }
                }
                else
                {
                    _text.Rotation = 0;
                    if (largo > 40)
                    {
                        _text.Height = Math.Round(Math.Abs(largo / 60));
                    }
                    else
                    {
                        _text.Height = 1;
                    }
                }
                _text.Value = "AUTODESARROLLO - CARA VISTA SIN COMPROBAR";
                dxf.Entities.Add(_text);
                n = 5;
                dxf.Save(this.OutfileDXF);
                n = 6;
            }
            catch (Exception)
            {
                // Handle exceptions.
                MessageBox.Show("Hay un problema a la hora de acceder al archivo DXF. Error " + n.ToString() + "  " + version, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }



    }
}
