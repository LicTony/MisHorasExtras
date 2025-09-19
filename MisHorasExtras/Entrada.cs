using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MisHorasExtras
{
    public class Entrada
    {
        public int RenglonDelExcel { get; set; }

        public DateOnly Fecha { get; set; }

        public DateTime HoraDesde { get; set; }

        public DateTime HoraHasta { get; set; }

        public string Detalle { get; set; } = string.Empty;
    }
}
