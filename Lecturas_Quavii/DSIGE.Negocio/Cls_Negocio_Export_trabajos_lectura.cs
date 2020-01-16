 
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Dato;
using DSIGE.Modelo;

namespace DSIGE.Negocio
{
  public   class Cls_Negocio_Export_trabajos_lectura
    {

      public List<Cls_Entidad_Export_trabajos_lectura> Capa_Negocio_Get_ListaLecturas(string fechaAsigna, int TipoServicio)
        {
            try
            {
                Cls_Dato_Export_trabajos_lectura Objeto_Dato = new Cls_Dato_Export_trabajos_lectura();
                return Objeto_Dato.Capa_Dato_Get_ListandoLecturas(fechaAsigna, TipoServicio);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


      public List<Cls_Entidad_Export_trabajos_lectura> Capa_Negocio_Get_ListaLecturas_Excel(string fechaAsigna, int TipoServicio)
      {
          try
          {
              Cls_Dato_Export_trabajos_lectura Objeto_Dato = new Cls_Dato_Export_trabajos_lectura();
              return Objeto_Dato.Capa_Dato_Get_ListandoLecturas_Excel(fechaAsigna, TipoServicio);
          }
          catch (Exception e)
          {
              throw e;
          }
      }


        public DataTable Capa_Negocio_Get_ListaCortesReconexion_Excel(int Local, string fechaAsigna, int TipoServicio)
        {
            try
            {
                Cls_Dato_Export_trabajos_lectura Objeto_Dato = new Cls_Dato_Export_trabajos_lectura();
                return Objeto_Dato.Capa_Dato_Get_ListaCortesReconexion_Excel(Local, fechaAsigna, TipoServicio);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string Capa_Negocio_InsertandoUpdate_LecturaMeses(int user, string fechaAsigna)
      {
          try
          {



              Cls_Dato_Export_trabajos_lectura Objeto_Dato = new Cls_Dato_Export_trabajos_lectura();
              return Objeto_Dato.Capa_Dato_InsertandoUpdate_LecturaMeses(user, fechaAsigna);
          }
          catch (Exception e)
          {
              throw e;
          }

      }

      public string Capa_Negocio_Validacion_LecturaMeses_lectura(string fechaAsignacion)
      {
          try
          {


              Cls_Dato_Importacion_Lecturas Objeto_Dato01 = new Cls_Dato_Importacion_Lecturas();

              return Objeto_Dato01.Capa_Dato_ValidacionLectura(fechaAsignacion);
              



          }
          catch (Exception e)
          {
              throw e;
          }

      }



    }
}
