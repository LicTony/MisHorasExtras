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
            LblStatus.Content = "";
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
                LblStatus.Content = $"Error: No se encontró el archivo '{excelFileName}'.";
                return;
            }

            try
            {
                using var excelManager = new ExcelManager(excelFilePath);
                LblStatus.Content = "Archivo Excel abierto correctamente.";

                bool entradaExiste = excelManager.ExisteTab("Entrada");
                bool salidaExiste = excelManager.ExisteTab("Salida");

                if (entradaExiste && salidaExiste)
                {
                    LblStatus.Content = "Pestañas 'Entrada' y 'Salida' encontradas.";
                
                    // Obtener la hoja de trabajo 'Entrada'
                    var hojaEntrada = excelManager.ObtenerHojaDeTrabajo("Entrada");
                
                    // Contar el número de filas utilizadas en la hoja 'Entrada'
                    int rowCount = hojaEntrada.LastRowUsed()?.RowNumber() ?? 0;

                    if (rowCount <= 1) // Si hay al menos una fila (encabezado)
                    {
                        LblStatus.Content = "Pestaña 'Entrada' está vacía o solo tiene encabezado.";
                        return; // Salir del método si no hay datos para procesar
                    }

                    LblStatus.Content = $"Procesando {rowCount - 1} registros en la pestaña 'Entrada'...";

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
                        LblStatus.Content = $"Procesando fila {fila - 1} de {rowCount - 1}...";


                        //todo validar que la variable celdaBHoraDesde sea un fomatato de hora valido
                        //todo validar que la variable celdaCHoraHasta sea un fomatato de hora valido
                        //todo validar que celdaBHoraDesde sea menor a celdaCHoraHasta

                    }

                    // Guardar los cambios en el archivo Excel
                    excelManager.Guardar();
                    LblStatus.Content = "Proceso completado. Archivo Excel guardado con los días de la semana.";
                }
                else
                {
                    string mensaje = "Error: Pestañas no encontradas. ";
                    if (!entradaExiste) mensaje += "Falta 'Entrada'. ";
                    if (!salidaExiste) mensaje += "Falta 'Salida'.";
                    LblStatus.Content = mensaje;
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
                    LblStatus.Content = "Error: El archivo Excel está abierto en otra aplicación. Ciérrelo e intente de nuevo.";
                }
                else
                {
                    LblStatus.Content = $"Error de E/S al procesar el archivo: {ex.Message}";
                }
            }
            catch (Exception ex)
            {
                LblStatus.Content = $"Error inesperado al procesar el archivo: {ex.Message}";
            }
        }
    }
}