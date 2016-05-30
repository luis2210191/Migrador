using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MigradorXls
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Menu());           
        }
    }

    /// <summary>
    /// Tipo de operación. Es el tipo de data a migrar y ya tienen predefinido 
    /// las acciones a realizar y los requerimientos necesarios para proceder
    /// con la acción, ejemplo: cliente, proveedor, inventario.
    /// </summary>
    public class TypeOperation
    {
        // Titulo de la operacion, esta se visualizará en el combo de tipo de operaciones
        public String Title { get; set; }
        // Número minimo de archivos que debe tener para proceder con la acción
        public int Loaded { get; set; }
        // Mensaje que se mostrará al elegir la operacion en el combo, mostrara los requerimientos de la operación
        public String Menssage { get; set; }
        // Ejecutará todas las acciones del tipo de operación, se ejecutará despues de hacer clic en el botón MIGRAR
        public Action Action { get; set; }
    }
}
