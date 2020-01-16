using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DSIGE.Modelo
{
  public  class ImportarDiasTrabajo
    {

        public int id_Sector { get; set; }

        public string Sector { get; set; }

        public string fecha { get; set; }

        public string estado { get; set; }

        public int usuario_creacion { get; set; }

        public string fecha_creacion { get; set; }

        public int usuario_edicion { get; set; }

        public string fecha_edicion { get; set; }

        public int id_Local { get; set; }


    }
}
