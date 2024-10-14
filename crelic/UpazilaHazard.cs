using crelic;
using System;
using System.Windows;


namespace crelic
{
    public class UpazilaHazard
    {
        public string upazila;
        public string hazard;
        public string padre;
        public int level = 0;
        public bool leido = false;

        public UpazilaHazard(string upazila, string hazard, string padre, int level)
        {
            this.upazila = upazila;
            this.hazard = hazard;
            this.padre = padre;
            this.level = level;
        }
    }
}
