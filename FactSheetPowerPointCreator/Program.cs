using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace FactSheetPowerPointCreator
{
    class Program
    {



        static void Main(string[] args)
        {

            Application objApp;
            Presentations objPresSet;
            _Presentation objPres;
            SlideShowWindows objSSWs;
            SlideShowSettings objSSS;

            objApp = new Application();
            objPresSet = objApp.Presentations;
            objPres = objPresSet.Open(@"D:\Projects\Work\FactSheetPowerPointCreator\FactSheetPowerPointCreator\bin\Debug\FS_Standart.ppt");// что открываем            
            objSSS = objPres.SlideShowSettings;
            objSSS.StartingSlide = 1;

            Slide s = objPres.Slides[1];

            int a = 0, b = 0, c = 0, d = 0; 


            foreach(Microsoft.Office.Interop.PowerPoint.Shape shape in s.Shapes)
            {
                a++;
                try
                {
                    
                    try
                    {
                        b = 0;
                        foreach(Microsoft.Office.Interop.PowerPoint.TextRange tr in shape.TextFrame.TextRange)
                        {
                            b++;
                            tr.Text = a.ToString() + " test "+ b.ToString();
                        }
                    }
                    catch
                    {

                    }
                  
                   
                }
                catch {
                    
                }
                //}

            }



            objSSS.Run();
            Console.ReadLine();

          //

           objPres.Close();
            objApp.Quit();//показ окончен
            GC.Collect();
        }
    }
}
