using Inventor;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Media3D;
using System.Windows.Shapes;
using Point = Inventor.Point;

namespace Nesting
{
    class IdentificazioneEntita
    {
        public static Edges main(Edges oEdges, Inventor.Application iApp)
        {
            SketchEntity result = null;

            IDictionary<CurveTypeEnum, int> counter = new Dictionary<CurveTypeEnum, int>();

            foreach (Edge e in oEdges)
            {
                if (!counter.ContainsKey(e.GeometryType))
                {
                    counter[e.GeometryType] = 0;
                }
                counter[e.GeometryType] = counter[e.GeometryType] + 1;
            }

            IdentificazioneEntita.gestiscoRisultati(counter, oEdges, iApp);

            return oEdges;
        }

        public static void gestiscoRisultati(IDictionary<CurveTypeEnum, int> data, Edges oEdges, Inventor.Application iApp)
        {
            SketchEntity result = null;

            // ! ASOLA
            if (data.ContainsKey(CurveTypeEnum.kLineSegmentCurve) && data.ContainsKey(CurveTypeEnum.kCircularArcCurve) 
                &&
                data[CurveTypeEnum.kLineSegmentCurve] == 2 && data[CurveTypeEnum.kCircularArcCurve] == 2
                )
            {
                // ? centro
                Point centro = null;

                // ? SweepAngle 
                Double sweepAngle = 0;

                // ? Width
                Double width = 0;

                Edge tmpEdge = null;
                foreach (Edge e in oEdges)
                {
                    if (e.GeometryType == CurveTypeEnum.kLineSegmentCurve)
                    {
                        if (tmpEdge == null)
                        {
                            tmpEdge = e;
                        }
                        else
                        {
                            width = iApp.MeasureTools.GetMinimumDistance(tmpEdge, e);
                            break;
                        }
                    }
                }

                // ? startP
                Point startPoint = null;

                foreach (Edge e in oEdges)
                {
                    if (e.GeometryType == CurveTypeEnum.kCircularArcCurve)
                    {
                        PartDocument oDoc = (PartDocument)iApp.ActiveDocument;
                        SheetMetalComponentDefinition oComp = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;

                        Vertex v = e.StartVertex;

                        WorkPoint oWp = oComp.WorkPoints.AddByPoint(v);
                        //oWp.SetByPoint();
                        
                        break;
                    }
                }
            }
        }
    }
}
