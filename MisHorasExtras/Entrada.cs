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

        public TimeOnly HoraDesde { get; set; }

        public TimeOnly HoraHasta { get; set; }

        public string Detalle { get; set; } = string.Empty;
    }
}
