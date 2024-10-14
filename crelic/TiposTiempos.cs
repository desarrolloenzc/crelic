using crelic;
using System;
using System.Windows;


namespace crelic
{
    public class TiposTiempos
    {
        public string tipo;
        public string tiempo;
        public int columna;
        public double tasa;
        public int columnaEmp;

        public TiposTiempos(string tipo, string tiempo, int columna)
        {
            this.tipo = tipo;
            this.tiempo = tiempo;
            this.columna = columna;
        }
    }
}
