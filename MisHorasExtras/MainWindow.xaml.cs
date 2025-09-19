using System.Globalization;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace MisHorasExtras
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            SetStatus("");
        }

        private void SetStatus(string message, bool isError = false)
        {
            LblStatus.Content = message;
            LblStatus.Foreground = isError ? System.Windows.Media.Brushes.Red : System.Windows.Media.Brushes.Black;
        }

        /// <summary>
        /// Ejecuta las acciones sobre el excel MisHorasExtras.xlsm
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnEjecutar_Click(object sender, RoutedEventArgs e)
        {
            string excelFileName = "MisHorasExtras.xlsm";
            string executablePath = AppDomain.CurrentDomain.BaseDirectory;
            string excelFilePath = System.IO.Path.Combine(executablePath, excelFileName);

            if (!System.IO.File.Exists(excelFilePath))
            {
                SetStatus($"Error: No se encontró el archivo '{excelFileName}'.", true);
                return;
            }

            try
            {
                using var excelManager = new ExcelManager(excelFilePath);
                SetStatus("Archivo Excel abierto correctamente.");

                bool entradaExiste = excelManager.ExisteTab("Entrada");
                bool salidaExiste = excelManager.ExisteTab("Salida");

                if (entradaExiste && salidaExiste)
                {
                    SetStatus("Pestañas 'Entrada' y 'Salida' encontradas.");
                
                    // Obtener la hoja de trabajo 'Entrada'
                    var hojaEntrada = excelManager.ObtenerHojaDeTrabajo("Entrada");
                
                    // Contar el número de filas utilizadas en la hoja 'Entrada'
                    int rowCount = hojaEntrada.LastRowUsed()?.RowNumber() ?? 0;

                    if (rowCount <= 1) // Si hay al menos una fila (encabezado)
                    {
                        SetStatus("Pestaña 'Entrada' está vacía o solo tiene encabezado.");
                        return; // Salir del método si no hay datos para procesar
                    }

                    SetStatus($"Procesando {rowCount - 1} registros en la pestaña 'Entrada'...");

                    // Recorrer la columna A desde la fila 2
                    for (int fila = 2; fila <= rowCount; fila++)
                    {
                        string celdaAFecha = excelManager.ObtenerValorCelda("Entrada", fila, 1)??""; // Columna A
                        string celdaBHoraDesde = excelManager.ObtenerValorCelda("Entrada", fila, 2) ?? ""; // Columna B
                        string celdaCHoraHasta = excelManager.ObtenerValorCelda("Entrada", fila, 3) ?? ""; // Columna C

                        if (DateTime.TryParse(celdaAFecha, CultureInfo.CurrentCulture, out DateTime fecha))
                        {
                            // Es una fecha válida, colocar el día de la semana en la columna E
                            string diaSemana = fecha.ToString("dddd"); // Formato de día de la semana completo
                            excelManager.EstablecerValorCelda("Entrada", fila, 5, diaSemana); // Columna E
                        }
                        else
                        {
                            // No es una fecha válida, colocar "Fecha invalida" en la columna E
                            excelManager.EstablecerValorCelda("Entrada", fila, 5, "Fecha invalida"); // Columna E
                        }
                        SetStatus($"Procesando fila {fila - 1} de {rowCount - 1}...");

                        bool horaDesdeValida = TimeOnly.TryParse(celdaBHoraDesde, CultureInfo.CurrentCulture, out TimeOnly horaDesde);
                        bool horaHastaValida = TimeOnly.TryParse(celdaCHoraHasta, CultureInfo.CurrentCulture, out TimeOnly horaHasta);

                        if (!horaDesdeValida)
                        {
                            excelManager.EstablecerValorCelda("Entrada", fila, 6, "Hora desde invalida"); // Columna F
                        }
                        else if (!horaHastaValida)
                        {
                            excelManager.EstablecerValorCelda("Entrada", fila, 6, "Hora hasta invalida"); // Columna F
                        }
                        else if (horaDesde >= horaHasta)
                        {
                            excelManager.EstablecerValorCelda("Entrada", fila, 6, "Hora desde debe ser menor a Hora hasta"); // Columna F
                        }
                        else
                        {
                            excelManager.EstablecerValorCelda("Entrada", fila, 6, ""); // Limpiar la celda de error si no hubo errores
                        }


                        if (horaDesdeValida && horaHastaValida) 
                        { 
                            //todo validar que la frajan  hora desde a hora hasta sea menor o mayor a la franja horaria 09:00 a 16:42                        
                        }
                    }

                    // Guardar los cambios en el archivo Excel
                    excelManager.Guardar();
                    SetStatus("Proceso completado. Archivo Excel guardado con los días de la semana y validaciones.");

                }
                else
                {
                    string mensaje = "Error: Pestañas no encontradas. ";
                    if (!entradaExiste) mensaje += "Falta 'Entrada'. ";
                    if (!salidaExiste) mensaje += "Falta 'Salida'.";
                    SetStatus(mensaje, true);
                }
            }
            catch (System.IO.IOException ex)
            {
                // Verificar si el error es porque el archivo está en uso
                // Esto es una heurística, ya que el código de error exacto puede variar
                // o no ser directamente accesible en todas las versiones de .NET o sistemas operativos.
                // Un mensaje común para archivo en uso es 'The process cannot access the file because it is being used by another process.'
                if (ex.Message.Contains("being used by another process") || ex.HResult == -2147024864) // 0x80070020 ERROR_SHARING_VIOLATION
                {
                    SetStatus("Error: El archivo Excel está abierto en otra aplicación. Ciérrelo e intente de nuevo.", true);
                }
                else
                {
                    SetStatus($"Error de E/S al procesar el archivo: {ex.Message}", true);
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Error inesperado al procesar el archivo: {ex.Message}", true);
            }
        }
    }
}