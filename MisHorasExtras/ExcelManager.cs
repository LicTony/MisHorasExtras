using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MisHorasExtras
{

    /// <summary>
    /// Clase para manipular archivos Excel usando ClosedXML
    /// </summary>
    public class ExcelManager : IDisposable
    {
        private readonly XLWorkbook _workbook;
        private string _rutaArchivo=string.Empty;
        private bool _disposed = false;

        /// <summary>
        /// Constructor que abre un archivo Excel existente
        /// </summary>
        /// <param name="rutaArchivo">Ruta del archivo Excel</param>
        public ExcelManager(string rutaArchivo)
        {
            if (string.IsNullOrEmpty(rutaArchivo))
                throw new ArgumentException("La ruta del archivo no puede estar vacía", nameof(rutaArchivo));

            if (!File.Exists(rutaArchivo))
                throw new FileNotFoundException($"El archivo no existe: {rutaArchivo}");

            _rutaArchivo = rutaArchivo;
            _workbook = new XLWorkbook(rutaArchivo);
        }

        /// <summary>
        /// Constructor para crear un nuevo archivo Excel
        /// </summary>
        public ExcelManager()
        {
            _workbook = new XLWorkbook();
        }

        /// <summary>
        /// Verifica si existe un tab con el nombre especificado
        /// </summary>
        /// <param name="nombreTab">Nombre del tab a buscar</param>
        /// <returns>True si el tab existe, false en caso contrario</returns>
        public bool ExisteTab(string nombreTab)
        {
            ValidarDisposed();

            if (string.IsNullOrEmpty(nombreTab))
                throw new ArgumentException("El nombre del tab no puede estar vacío", nameof(nombreTab));

            return _workbook.TryGetWorksheet(nombreTab, out _);
        }

        /// <summary>
        /// Obtiene una hoja de trabajo por su nombre
        /// </summary>
        /// <param name="nombreTab">Nombre del tab a obtener</param>
        /// <returns>La IXLWorksheet correspondiente</returns>
        public IXLWorksheet ObtenerHojaDeTrabajo(string nombreTab)
        {
            ValidarDisposed();

            if (string.IsNullOrEmpty(nombreTab))
                throw new ArgumentException("El nombre del tab no puede estar vacío", nameof(nombreTab));

            if (!_workbook.TryGetWorksheet(nombreTab, out IXLWorksheet worksheet))
                throw new ArgumentException($"El tab '{nombreTab}' no existe");

            return worksheet;
        }

        /// <summary>
        /// Obtiene el valor de una celda específica
        /// </summary>
        /// <param name="nombreTab">Nombre del tab</param>
        /// <param name="celda">Dirección de la celda (ej: "A1", "B5")</param>
        /// <returns>El valor de la celda como string, null si la celda está vacía</returns>
        public string? ObtenerValorCelda(string nombreTab, string celda)
        {
            ValidarDisposed();

            if (!_workbook.TryGetWorksheet(nombreTab, out IXLWorksheet worksheet))
                throw new ArgumentException($"El tab '{nombreTab}' no existe");

            if (string.IsNullOrEmpty(celda))
                throw new ArgumentException("La dirección de celda no puede estar vacía", nameof(celda));

            var cell = worksheet.Cell(celda);

            if (cell == null)
                return null;

            return cell.IsEmpty() ? null : cell.GetValue<string>();
        }

        /// <summary>
        /// Obtiene el valor de una celda usando coordenadas de fila y columna
        /// </summary>
        /// <param name="nombreTab">Nombre del tab</param>
        /// <param name="fila">Número de fila (empezando en 1)</param>
        /// <param name="columna">Número de columna (empezando en 1)</param>
        /// <returns>El valor de la celda como string</returns>
        public string? ObtenerValorCelda(string nombreTab, int fila, int columna)
        {
            ValidarDisposed();

            if (!_workbook.TryGetWorksheet(nombreTab, out IXLWorksheet worksheet))
                throw new ArgumentException($"El tab '{nombreTab}' no existe");

            if (fila < 1 || columna < 1)
                throw new ArgumentException("Los números de fila y columna deben ser mayores a 0");

            var cell = worksheet.Cell(fila, columna);

            if (cell == null)
                return null;

            return cell.IsEmpty() ? null : cell.GetValue<string>();
        }

        /// <summary>
        /// Establece el valor de una celda específica
        /// </summary>
        /// <param name="nombreTab">Nombre del tab</param>
        /// <param name="celda">Dirección de la celda (ej: "A1", "B5")</param>
        /// <param name="valor">El valor a establecer</param>
        public void EstablecerValorCelda(string nombreTab, string celda, object valor)
        {
            ValidarDisposed();

            if (!_workbook.TryGetWorksheet(nombreTab, out IXLWorksheet worksheet))
                throw new ArgumentException($"El tab '{nombreTab}' no existe");

            if (string.IsNullOrEmpty(celda))
                throw new ArgumentException("La dirección de celda no puede estar vacía", nameof(celda));

            AsignarValorCelda(worksheet.Cell(celda), valor);
        }

        /// <summary>
        /// Establece el valor de una celda usando coordenadas de fila y columna
        /// </summary>
        /// <param name="nombreTab">Nombre del tab</param>
        /// <param name="fila">Número de fila (empezando en 1)</param>
        /// <param name="columna">Número de columna (empezando en 1)</param>
        /// <param name="valor">El valor a establecer</param>
        public void EstablecerValorCelda(string nombreTab, int fila, int columna, object valor)
        {
            ValidarDisposed();

            if (!_workbook.TryGetWorksheet(nombreTab, out IXLWorksheet worksheet))
                throw new ArgumentException($"El tab '{nombreTab}' no existe");

            if (fila < 1 || columna < 1)
                throw new ArgumentException("Los números de fila y columna deben ser mayores a 0");

            AsignarValorCelda(worksheet.Cell(fila, columna), valor);
        }

        /// <summary>
        /// Inserta una fila de datos en el tab especificado
        /// </summary>
        /// <param name="nombreTab">Nombre del tab</param>
        /// <param name="fila">Número de fila donde insertar (empezando en 1)</param>
        /// <param name="datos">Array de objetos con los datos a insertar</param>
        public void InsertarFila(string nombreTab, int fila, params object[] datos)
        {
            ValidarDisposed();

            if (!_workbook.TryGetWorksheet(nombreTab, out IXLWorksheet worksheet))
                throw new ArgumentException($"El tab '{nombreTab}' no existe");

            if (fila < 1)
                throw new ArgumentException("El número de fila debe ser mayor a 0");

            if (datos == null || datos.Length == 0)
                throw new ArgumentException("Debe proporcionar al menos un dato");

            // Insertar los datos en la fila especificada
            for (int i = 0; i < datos.Length; i++)
            {
                AsignarValorCelda(worksheet.Cell(fila, i + 1), datos[i]);
            }
        }

        /// <summary>
        /// Agrega una fila de datos al final del contenido existente
        /// </summary>
        /// <param name="nombreTab">Nombre del tab</param>
        /// <param name="datos">Array de objetos con los datos a agregar</param>
        public void AgregarFila(string nombreTab, params object[] datos)
        {
            ValidarDisposed();

            if (!_workbook.TryGetWorksheet(nombreTab, out IXLWorksheet worksheet))
                throw new ArgumentException($"El tab '{nombreTab}' no existe");

            // Encontrar la primera fila vacía
            int ultimaFila = worksheet.LastRowUsed()?.RowNumber() ?? 0;
            InsertarFila(nombreTab, ultimaFila + 1, datos);
        }

        /// <summary>
        /// Cambia el ancho de una columna específica
        /// </summary>
        /// <param name="nombreTab">Nombre del tab</param>
        /// <param name="columna">Letra de la columna (ej: "A", "B", "AA") o número</param>
        /// <param name="ancho">Ancho de la columna en puntos</param>
        public void CambiarAnchoColumna(string nombreTab, string columna, double ancho)
        {
            ValidarDisposed();

            if (!_workbook.TryGetWorksheet(nombreTab, out IXLWorksheet worksheet))
                throw new ArgumentException($"El tab '{nombreTab}' no existe");

            if (string.IsNullOrEmpty(columna))
                throw new ArgumentException("La columna no puede estar vacía");

            if (ancho <= 0)
                throw new ArgumentException("El ancho debe ser mayor a 0");

            worksheet.Column(columna).Width = ancho;
        }

        /// <summary>
        /// Cambia el ancho de una columna usando su número
        /// </summary>
        /// <param name="nombreTab">Nombre del tab</param>
        /// <param name="numeroColumna">Número de columna (empezando en 1)</param>
        /// <param name="ancho">Ancho de la columna en puntos</param>
        public void CambiarAnchoColumna(string nombreTab, int numeroColumna, double ancho)
        {
            ValidarDisposed();

            if (!_workbook.TryGetWorksheet(nombreTab, out IXLWorksheet worksheet))
                throw new ArgumentException($"El tab '{nombreTab}' no existe");

            if (numeroColumna < 1)
                throw new ArgumentException("El número de columna debe ser mayor a 0");

            if (ancho <= 0)
                throw new ArgumentException("El ancho debe ser mayor a 0");

            worksheet.Column(numeroColumna).Width = ancho;
        }

        /// <summary>
        /// Ajusta automáticamente el ancho de una columna según su contenido
        /// </summary>
        /// <param name="nombreTab">Nombre del tab</param>
        /// <param name="columna">Letra de la columna</param>
        public void AjustarAnchoColumna(string nombreTab, string columna)
        {
            ValidarDisposed();

            if (!_workbook.TryGetWorksheet(nombreTab, out IXLWorksheet worksheet))
                throw new ArgumentException($"El tab '{nombreTab}' no existe");

            worksheet.Column(columna).AdjustToContents();
        }

        /// <summary>
        /// Limpia todas las celdas de un tab específico
        /// </summary>
        /// <param name="nombreTab">Nombre del tab a limpiar</param>
        /// <param name="mantenerFormato">Si es true, mantiene el formato de las celdas</param>
        public void LimpiarTab(string nombreTab, bool mantenerFormato = false)
        {
            ValidarDisposed();

            if (!_workbook.TryGetWorksheet(nombreTab, out IXLWorksheet worksheet))
                throw new ArgumentException($"El tab '{nombreTab}' no existe");

            if (mantenerFormato)
            {
                // Solo limpia el contenido, mantiene formato
                worksheet.Clear(XLClearOptions.Contents);
            }
            else
            {
                // Limpia: contenido y formato
                worksheet.Clear(XLClearOptions.All);
            }
        }

        /// <summary>
        /// Crea un nuevo tab con el nombre especificado
        /// </summary>
        /// <param name="nombreTab">Nombre del nuevo tab</param>
        /// <returns>True si se creó exitosamente, false si ya existía</returns>
        public bool CrearTab(string nombreTab)
        {
            ValidarDisposed();

            if (string.IsNullOrEmpty(nombreTab))
                throw new ArgumentException("El nombre del tab no puede estar vacío");

            if (ExisteTab(nombreTab))
                return false;

            _workbook.Worksheets.Add(nombreTab);
            return true;
        }

        /// <summary>
        /// Obtiene la lista de nombres de todos los tabs
        /// </summary>
        /// <returns>Lista con los nombres de los tabs</returns>
        public List<string> ObtenerNombresTabs()
        {
            ValidarDisposed();

            var nombres = new List<string>();
            foreach (var worksheet in _workbook.Worksheets)
            {
                nombres.Add(worksheet.Name);
            }
            return nombres;
        }

        /// <summary>
        /// Guarda los cambios en el archivo
        /// </summary>
        public void Guardar()
        {
            ValidarDisposed();

            if (!string.IsNullOrEmpty(_rutaArchivo))
            {
                _workbook.Save();
            }
            else
            {
                throw new InvalidOperationException("No se puede guardar sin especificar una ruta. Use GuardarComo()");
            }
        }

        /// <summary>
        /// Guarda el archivo en la ruta especificada
        /// </summary>
        /// <param name="rutaArchivo">Ruta donde guardar el archivo</param>
        public void GuardarComo(string rutaArchivo)
        {
            ValidarDisposed();

            if (string.IsNullOrEmpty(rutaArchivo))
                throw new ArgumentException("La ruta del archivo no puede estar vacía");

            _workbook.SaveAs(rutaArchivo);
            _rutaArchivo = rutaArchivo;
        }

        /// <summary>
        /// Asigna un valor a una celda manejando diferentes tipos de datos
        /// </summary>
        /// <param name="cell">La celda donde asignar el valor</param>
        /// <param name="valor">El valor a asignar</param>
        private static void AsignarValorCelda(IXLCell cell, object valor)
        {
            if (valor == null)
            {
                cell.Clear();
                return;
            }

            cell.Value = valor switch
            {
                string s => (XLCellValue)s,
                int i => (XLCellValue)i,
                double d => (XLCellValue)d,
                decimal dec => (XLCellValue)dec,
                float f => (XLCellValue)f,
                long l => (XLCellValue)l,
                DateTime dt => (XLCellValue)dt,
                TimeSpan ts => (XLCellValue)ts,
                bool b => (XLCellValue)b,
                _ => (XLCellValue)valor.ToString(),// Para cualquier otro tipo, convertir a string
            };
        }

        /// <summary>
        /// Valida que el objeto no haya sido disposed
        /// </summary>
        private void ValidarDisposed()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        /// <summary>
        /// Implementación del patrón Dispose
        /// </summary>
        /// <param name="disposing">True si se está llamando desde Dispose(), false si desde el finalizer</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Liberar recursos administrados
                    _workbook?.Dispose();
                }

                // Liberar recursos no administrados (si los hubiera)
                // En este caso no tenemos recursos no administrados

                _disposed = true;
            }
        }

        /// <summary>
        /// Libera los recursos utilizados
        /// </summary>
        public void Dispose()
        {
            // Llamar al método Dispose con disposing = true
            Dispose(true);
            // Suprimir la finalización ya que ya liberamos los recursos
            GC.SuppressFinalize(this);
        }
    }

}
