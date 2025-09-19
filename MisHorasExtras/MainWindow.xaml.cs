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
using System.Collections.Generic; // Added for List<Entrada>

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
            List<Entrada> entradas = []; 

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

                    // Limpiar columnas E, F y G antes de procesar
                    for (int fila = 2; fila <= rowCount; fila++)
                    {
                        excelManager.EstablecerValorCelda("Entrada", fila, 5, ""); // Columna E
                        excelManager.EstablecerValorCelda("Entrada", fila, 6, ""); // Columna F
                        excelManager.EstablecerValorCelda("Entrada", fila, 7, ""); // Columna G
                    }

                    // Recorrer la columna A desde la fila 2
                    for (int fila = 2; fila <= rowCount; fila++)
                    {
                        string celdaAFecha = excelManager.ObtenerValorCelda("Entrada", fila, 1)?? ""; // Columna A
                        string celdaBHoraDesde = excelManager.ObtenerValorCelda("Entrada", fila, 2) ?? ""; // Columna B
                        string celdaCHoraHasta = excelManager.ObtenerValorCelda("Entrada", fila, 3) ?? ""; // Columna C

                        StringBuilder erroresFila = new();
                        

                        // 1. Validación de fecha (Columna A -> Columna E y errores a F)
                        bool fechaValida = DateTime.TryParse(celdaAFecha, CultureInfo.CurrentCulture, out DateTime fecha);
                        if (fechaValida)
                        {
                            string diaSemana = fecha.ToString("dddd");
                            excelManager.EstablecerValorCelda("Entrada", fila, 5, diaSemana); // Columna E
                        }
                        else
                        {
                            erroresFila.Append("Fecha invalida; ");
                        }

                        // 2. Validación de horas (Columnas B y C -> errores a F)
                        bool horaDesdeValida = TimeOnly.TryParse(celdaBHoraDesde, CultureInfo.CurrentCulture, out TimeOnly horaDesde);
                        bool horaHastaValida = TimeOnly.TryParse(celdaCHoraHasta, CultureInfo.CurrentCulture, out TimeOnly horaHasta);

                        if (!horaDesdeValida)
                        {
                            erroresFila.Append("Hora desde invalida; ");
                        }
                        else if (!horaHastaValida)
                        {
                            erroresFila.Append("Hora hasta invalida; ");
                        }
                        else if (horaDesde >= horaHasta)
                        {
                            erroresFila.Append("Hora desde debe ser menor a Hora hasta; ");
                        }

                        // 3. Validación de franja horaria (si las horas son válidas y horaDesde < horaHasta -> errores a F)
                        // Solo aplicar esta validación si la fecha es válida y es un día de semana (Lunes a Viernes)
                        if (fechaValida && // Use fechaValida here
                            fecha.DayOfWeek != DayOfWeek.Saturday &&
                            fecha.DayOfWeek != DayOfWeek.Sunday &&
                            horaDesdeValida && horaHastaValida && horaDesde < horaHasta)
                        {
                            TimeOnly jornadaInicio = new(9, 0);
                            TimeOnly jornadaFin = new(16, 42);

                            if (horaDesde < jornadaFin && horaHasta > jornadaInicio)
                            {
                                erroresFila.Append("Dentro de la franja horaria (09:00 - 16:42); ");
                            }
                        }

                        // Si la fila no tiene errores hasta ahora (erroresFila.Length == 0) agregar los datos a la lista entradas
                        if (erroresFila.Length == 0)
                        {
                            entradas.Add(new Entrada
                            {
                                Fecha = DateOnly.FromDateTime(fecha),
                                HoraDesde = new DateTime(fecha.Year, fecha.Month, fecha.Day, 
                                    horaDesde.Hour, horaDesde.Minute, horaDesde.Second,DateTimeKind.Local),
                                HoraHasta = new DateTime(fecha.Year, fecha.Month, fecha.Day, 
                                    horaHasta.Hour, horaHasta.Minute, horaHasta.Second, DateTimeKind.Local)
                            });
                        }

                        //todo validar que en entradas si una fecha se repite no esten solapados sus horarios y que no haya un hueco  entre el horario anterior a
                        // 9 a 16:42 y tampoco haya un huevo despues



                        // Escribir todos los errores concatenados en la Columna F
                        if (erroresFila.Length > 0)
                        {
                            excelManager.EstablecerValorCelda("Entrada", fila, 6, erroresFila.ToString().TrimEnd(' ', ';'));
                        }
                        else
                        {
                            excelManager.EstablecerValorCelda("Entrada", fila, 6, ""); // Limpiar la celda F si no hay errores
                        }

                        SetStatus($"Procesando fila {fila - 1} de {rowCount - 1}...");
                    }

                    // Guardar los cambios en el archivo Excel
                    excelManager.Guardar();
                    SetStatus("Proceso completado. Archivo Excel guardado con los días de la semana y validaciones.");

                    // Abrir el archivo Excel para que el usuario lo vea
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(excelFilePath) { UseShellExecute = true });

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
