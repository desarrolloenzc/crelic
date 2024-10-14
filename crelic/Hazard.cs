using crelic;
using System;
using System.Windows;


namespace crelic
{
    public class Hazard
    {
        public string hazard;
        public float veryhigh = -1;
        public float high = -1;
        public float medium = -1;
        public float low = -1;
        public float verylow = -1;

        public Hazard(string hazard)
        {
            this.hazard = hazard;
        }
    }
}
