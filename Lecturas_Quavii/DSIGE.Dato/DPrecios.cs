using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using Microsoft.Practices.EnterpriseLibrary.Data;

namespace DSIGE.Dato
{
   public class DPrecios
    {
       public DPrecios()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

       public List<Precios> ListarPrecios()
       {
           try
           {
               List<Precios> listaPrecios = new List<Precios>();

               using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_ListarPrecios"))
               {
                   if (iDr != null)
                   {
                       while (iDr.Read())
                       {
                           listaPrecios.Add(
                                   new Precios()
                                   {
                                       id_Precio = Convert.ToInt32(iDr["id_Precio"]),
                                       Precio = Convert.ToString(iDr["Precio"]),
                                       Concepto = Convert.ToString(iDr["Concepto"]),
                                       anio = Convert.ToString(iDr["Año"]),
                                       id_Local = Convert.ToString(iDr["id_Local"]),
                                       nombre_local = Convert.ToString(iDr["nombre_local"]),
                                       estado = Convert.ToString(iDr["estado"]),
                                   }
                               );
                       }
                   }
               }
               return listaPrecios;
           }
           catch (Exception e)
           {
               throw e;
           }
       }

       public bool RegistrarPrecios(Precios objPrecio)
       {
           try
           {
               DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_RegistrarPrecios",
                   objPrecio.Precio,
                   objPrecio.Concepto,
                   objPrecio.anio,
                   objPrecio.id_Local,
                   objPrecio.estado,
                   objPrecio.usuario_creacion);
               
               return true;
           }
           catch (Exception e)
           {
               throw e;
           }
       }

       public bool ActualizarPrecios(Precios objPrecio)
       {
           try
           {
               DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_ActualizarPrecios",
                  objPrecio.Precio,
                  objPrecio.Concepto,
                  objPrecio.anio,
                  objPrecio.id_Local,
                  objPrecio.estado,
                  objPrecio.usuario_edicion,
                  objPrecio.id_Precio);
               return true;
           }
           catch (Exception e)
           {
               throw e;
           }
       }
    }
}
