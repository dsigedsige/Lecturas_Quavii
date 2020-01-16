using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DSIGE;
using DSIGE.Modelo;
using DSIGE.Dato;

namespace DSIGE.Negocio
{
    public class NPermisos
    {
        public List<Permisos> GetListaPermisos(string id)
        {
            return new Dpermisos().GetListaPermisos(id);
        }
        public List<Permisos> GetListaPermisosxUsuario(parametersPermisos obj)
        {
            return new Dpermisos().GetListaPermisosxUsuario(obj);
        }

        public bool InsertaAccesoOpciones(Permisos objPermiso)
        {
            return new  Dpermisos().InsertaAccesoOpciones(objPermiso);           
        }

        public bool ActualizarAccesoOpciones(Permisos objPermiso)
        {
            return new Dpermisos().ActualizarAccesoOpciones(objPermiso);   
        }
        public List<Perfil> getPerfiles()
        {
            return new Dpermisos().getPerfiles();
        }
        public bool InsertaPerfil_AccesoOpciones(Permisos objPermiso)
        {
            return new Dpermisos().InsertaPerfil_AccesoOpciones(objPermiso);
        }
        public bool ActualizarPerfil_AccesoOpciones(Permisos objPermiso)
        {
            return new Dpermisos().ActualizarPerfil_AccesoOpciones(objPermiso);
        }
        public List<Permisos> GetListaPermisosxPerfil(parametersPermisos obj)
        {
            return new Dpermisos().GetListaPermisosxPerfil(obj);
        }
        public bool EliminarAccesoOpciones(Permisos objPermiso)
        {
            return new Dpermisos().EliminarAccesoOpciones(objPermiso);
        }
        public bool EliminarPerfil_AccesoOpciones(Permisos objPermiso)
        {
            return new Dpermisos().EliminarPerfil_AccesoOpciones(objPermiso);
        }
    }
}
