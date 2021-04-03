using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace autocad_dynamic
{
    class Program
    {
        static void Main(string[] args)
        {
            dynamic autocadType = Type.GetTypeFromProgID("AutoCAD.Application");
            var AcadApp = Activator.CreateInstance(autocadType);
            
            AcadApp.Visible = true;

            double[] CenterOfCircle = new double[3];
            CenterOfCircle[0] = 0;
            CenterOfCircle[1] = 0;
            CenterOfCircle[2] = 0;

            double RadiusOfCircle = 100;

            var Circle = AcadApp.ActiveDocument.ModelSpace.AddCircle(CenterOfCircle, RadiusOfCircle);
            Circle.color = 3;

            AcadApp.ZoomExtents();

            Console.WriteLine("Done!");
            Console.ReadLine();

        }
    }
}
