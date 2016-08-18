using GenericExcelExport.Misc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenericExcelExport
{
    class Program
    {
        public static void Main(string[] args)
        {
            #region Parameters => the description of each one is commented in the method ExportToExcel
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("Nombre Pers.", "nombre");
            dic.Add("Edad", "edad");
            dic.Add("Sexo", "sexo");
            List<string> reportInfo = new List<string>();
            reportInfo.Add("EFLORES");
            reportInfo.Add("COMPANY NAME");
            reportInfo.Add("REP00001");
            #endregion

            UtilExcel.ExportToExcel("GenericExcel", GetList(), dic, reportInfo);
        }

        public static List<Persona> GetList()
        {
            return new List<Persona>() { 
                new Persona { id = 1, nombre = "Edgar Flores Palma", edad = 28, sexo = "Masculino" } ,
                new Persona { id = 2, nombre = "Jean Montenegro", edad = 33, sexo = "Masculino" } ,
                new Persona { id = 3, nombre = "Julissa Soza", edad = 26, sexo = "Femenino" } ,
                new Persona { id = 4, nombre = "Diego Quispe", edad = 48, sexo = "Masculino" } 
            };
        }
    }
}
