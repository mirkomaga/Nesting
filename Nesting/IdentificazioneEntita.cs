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
        public static List<Lavorazione> main(EdgeLoops oEdgeLoops, Inventor.Application iApp)
        {
            List<Lavorazione> result = new List<Lavorazione>();

            foreach (EdgeLoop oEdgeLoop in oEdgeLoops)
            {
                string nameLav = whois(oEdgeLoop.Edges);

                if (!string.IsNullOrEmpty(nameLav))
                {
                    EdgeCollection oEdgeColl = iApp.TransientObjects.CreateEdgeCollection();

                    foreach (Edge oEdge in oEdgeLoop.Edges)
                    {
                        oEdgeColl.Add(oEdge);
                    }

                    result.Add(new Lavorazione(nameLav, oEdgeColl));
                }
            }

            return result;
        }
        public static string whois(Edges oEdges)
        {
            string nameLav = null;

            IDictionary<CurveTypeEnum, int> typeOfEdge = new Dictionary<CurveTypeEnum, int>();

            foreach (Edge oEdge in oEdges)
            {
                if (!typeOfEdge.ContainsKey(oEdge.GeometryType))
                {
                    typeOfEdge.Add(oEdge.GeometryType, 0);
                }

                typeOfEdge[oEdge.GeometryType] ++;
            }

            // ! Foro
            if (typeOfEdge.ContainsKey(CurveTypeEnum.kCircleCurve) && typeOfEdge[CurveTypeEnum.kCircleCurve] == 1)
            {
                nameLav = "foro";
            }

            // ! Asola
            if (typeOfEdge.ContainsKey(CurveTypeEnum.kCircularArcCurve) && typeOfEdge.ContainsKey(CurveTypeEnum.kLineSegmentCurve) &&
                typeOfEdge[CurveTypeEnum.kLineSegmentCurve] == 2 && typeOfEdge[CurveTypeEnum.kCircularArcCurve] == 2)
            {
                nameLav = "asola";
            }

            return nameLav;
        }
    }
    public struct Lavorazione
    {
        public Lavorazione(string nameLav, EdgeCollection oEdgeColl)
        {
            nameLav_ = nameLav;
            oEdgeColl_ = oEdgeColl;
        }
        public string nameLav_ { get; private set; }
        public EdgeCollection oEdgeColl_ { get; private set; }
    }
}