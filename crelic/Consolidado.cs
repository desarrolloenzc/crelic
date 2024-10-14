using crelic;
using System;
using System.Windows;


namespace crelic
{
    public class Consolidado
    {
        public string nombre;
        public string empresa;
        public string sap;
        public bool esta;
        public int col;
        public bool hc;

        public Consolidado(string nombre,string empresa,string sap)
        {
            this.nombre = nombre;
            this.empresa = empresa;
            this.sap = sap;
            this.esta = false;
            this.col = 0;
            this.hc = false;
        }
    }
}
