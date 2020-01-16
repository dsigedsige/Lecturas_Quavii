using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Dato;
using DSIGE.Modelo;
using System.Data;

namespace DSIGE.Negocio
{
  public  class NimportarDiasTrabajo
    {
                   
             public List<Local> getLocales(){
                 return new DImportarDiasTrabajo().getLocales();
             }
             public List<ImportarDiasTrabajo> getSector_DiasTrabajo( int idLocal, string fechaIni, string fechaFin){
                 return new DImportarDiasTrabajo().getSector_DiasTrabajo(idLocal,fechaIni,fechaFin);
             }
             public bool ActualizarSector_DiasTrabajo(ImportarDiasTrabajo objDiasTrabajo)
             {
                 return new DImportarDiasTrabajo().ActualizarSector_DiasTrabajo(objDiasTrabajo);
             }
             public bool RegistrarSector_DiasTrabajo(List<ImportarDiasTrabajo> ListDiasTrabajo)
             {
                 return new DImportarDiasTrabajo().RegistrarSector_DiasTrabajo(ListDiasTrabajo);
             }
             public DataTable ListaExcel(string fileLocation)
             {
                 return new DImportarDiasTrabajo().ListaExcel(fileLocation);
             }
    }
}
