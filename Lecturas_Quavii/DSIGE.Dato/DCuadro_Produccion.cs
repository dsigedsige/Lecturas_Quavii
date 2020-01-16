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
  public  class DCuadro_Produccion
    {
      public DCuadro_Produccion()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

      public List<Cuadro_Produccion> GetOperarios(string fechaInicial, string fechaFinal,int idlocal, string valor, string lista)
      {
          try
          {
              List<Cuadro_Produccion> listaOperarios = new List<Cuadro_Produccion>();

              using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_Reporte_Lecturas_Operarios",fechaInicial,fechaFinal,idlocal,valor, lista))
              {
                  if (iDr != null)
                  {
                      while (iDr.Read())
                      {
                          listaOperarios.Add(
                                  new Cuadro_Produccion()
                                  {
                                      id_Operario = Convert.ToInt32(iDr["id_Operario"]),
                                      usuario_Operario = Convert.ToString(iDr["usuario_Operario"]),
                                      nombre_Operario = Convert.ToString(iDr["nombre_Operario"]),
                                      ing_Operario = Convert.ToString(iDr["Ing"]),
                                      cargo_Operario = Convert.ToString(iDr["Cargo"]),
                                      Abre_Operario = Convert.ToString(iDr["Abre"]),
                                      Precio = Convert.ToDecimal(iDr["Precio"]),
                                      PrecioUnit = Convert.ToDecimal(iDr["PrecioUnit"]),
                                  }
                              );
                      }
                  }
              }
              return listaOperarios;
          }
          catch (Exception e)
          {
              throw e;
          }
      }

      public List<Cuadro_Produccion> GetFechas(string fechaInicial, string fechaFinal, int idlocal, string valor, string lista)
      {
          try
          {
              List<Cuadro_Produccion> listaFechas = new List<Cuadro_Produccion>();

              using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_Reporte_Lecturas_Operarios", fechaInicial, fechaFinal, idlocal, valor,lista))
              {
                  if (iDr != null)
                  {
                      while (iDr.Read())
                      {
                          DateTime fecha = Convert.ToDateTime(iDr["fecha"]);
                          listaFechas.Add(
                                  new Cuadro_Produccion()
                                  {

                                      fechaAsignacion_Lectura = fecha.ToString("dd/MM/yyyy")                            
                                 }
                              );
                      }
                  }
              }
              return listaFechas;
          }
          catch (Exception e)
          {
              throw e;
          }
      }

      public List<Cuadro_Produccion> GetDatos(string fechaInicial, string fechaFinal, int idlocal, string valor, string lista)
      {
          try
          {
              List<Cuadro_Produccion> listaDatos = new List<Cuadro_Produccion>();

              using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_Reporte_Lecturas_Operarios", fechaInicial, fechaFinal, idlocal, valor, lista))
              {
                  if (iDr != null)
                  {
                      while (iDr.Read())
                      {
                          listaDatos.Add(
                                  new Cuadro_Produccion()
                                  {
                                      id_Operario = Convert.ToInt32(iDr["id_Operario"]),
                                      nombre_Operario = Convert.ToString(iDr["operario"]),
                                      fechaAsignacion_Lectura = Convert.ToString(iDr["fechaAsignacion_Lectura"]),
                                      cantidad = Convert.ToInt32(iDr["Cantidad"]),                                  
                                  }
                              );
                      }
                  }
              }
              return listaDatos;
          }
          catch (Exception e)
          {
              throw e;
          }
      }
    }
}
