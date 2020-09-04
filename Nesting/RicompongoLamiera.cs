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
    class RicompongoLamiera
    {
        public static Inventor.Application iApp = null;
        public static void main(PartDocument oDoc)
        {
            // ! prendo instanza Inventor
            getIstance();

            oDoc = (PartDocument) iApp.ActiveDocument;

            // ! Imposto SheetMetalDocument
            //setSheetMetalDocument(oDoc);

            // ! Imposto thickness
            //setThickness(oDoc);

            // ! Elimino raggiature
            //List<string> faceCollToKeep = deleteFillet(oDoc);

            // ! Elimino facce
            //deleteFace(oDoc, faceCollToKeep);

            // ! Creo Profilo
            createProfile(oDoc);
        }
        public static void getIstance()
        {
            try
            {
                iApp = (Inventor.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
                iApp.Visible = true;
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                iApp = (Inventor.Application)System.Activator.CreateInstance(System.Type.GetTypeFromProgID("Inventor.Application"));
                iApp.Visible = true;
                //System.Windows.Forms.MessageBox.Show("Assicurarsi di aver aperto inventor");
            }
        }
        public static void setSheetMetalDocument(PartDocument oDoc)
        {
            if (oDoc.SubType != "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                oDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}";
            }
        }
        public static void setThickness(PartDocument oDoc)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;

            double thikness = 0;
            int count = 0;

            foreach (Edge e in oCompDef.SurfaceBodies[1].Edges)
            {
                double distance = Math.Round(iApp.MeasureTools.GetMinimumDistance(e.StartVertex.Point, e.StopVertex.Point) * 100);
                
                if (e.GeometryType == CurveTypeEnum.kLineSegmentCurve && distance != 0 && distance <= 50)
                {
                    thikness = distance/10;
                    count++;
                }
            }

            if (count > 1)
            {
                string toSelect = null;

                switch (thikness)
                {
                    case 0.6:
                        toSelect = "Aluminium THK 0.6mm";
                        break;
                    case 0.8:
                        toSelect = "Aluminium THK 0.8mm";
                        break;
                    case 1.0:
                        toSelect = "Aluminium THK 1.0mm";
                        break;
                    case 1.2:
                        toSelect = "Aluminium THK 1.2mm";
                        break;
                    case 1.5:
                        toSelect = "Aluminium THK 1.5mm";
                        break;
                    case 2.0:
                        toSelect = "Aluminium THK 2.0mm";
                        break;
                    case 2.5:
                        toSelect = "Aluminium THK 2.5mm";
                        break;
                    case 3.0:
                        toSelect = "Aluminium THK 3.0mm";
                        break;
                    case 4.0:
                        toSelect = "Aluminium THK 4.0mm";
                        break;
                    case 5.0:
                        toSelect = "Aluminium THK 5.0mm";
                        break;
                    default:
                        //toSelect = "Default_mm";
                        Console.WriteLine("SONO ENTRATO IN DEFAULT");
                        break;
                }

                if (!string.IsNullOrEmpty(toSelect))
                {
                    SheetMetalStyle oStyle = oCompDef.SheetMetalStyles[toSelect];
                    oStyle.Activate();
                }
                else
                {
                    throw new Exception("Nessuno spessore trovato");
                }
            }
        }
        public static void coloroEntita(PartDocument oDoc, byte r, byte g, byte b, dynamic e)
        {
            HighlightSet oHS2 = oDoc.HighlightSets.Add();
            oHS2.Color = iApp.TransientObjects.CreateColor(r, g, b);
            oHS2.AddItem(e);
        }
        public static List<string> deleteFillet(PartDocument oDoc)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;
            
            List<string> faceCollToKeep = new List<string>();

            if (oCompDef.Bends.Count > 0)
            {
                Bend oBend = oCompDef.Bends[1];

                FaceCollection oFaceColl = oBend.FrontFaces[1].TangentiallyConnectedFaces;
                oFaceColl.Add(oBend.FrontFaces[1]);

                foreach (Face oFace in oFaceColl)
                {
                    faceCollToKeep.Add(oFace.InternalName);
                }

                NonParametricBaseFeature oBaseFeature = oCompDef.Features.NonParametricBaseFeatures[1];

                oBaseFeature.Edit();

                SurfaceBody basebody = oBaseFeature.BaseSolidBody;

                ObjectCollection oColl = iApp.TransientObjects.CreateObjectCollection();

                foreach (Face f in basebody.Faces)
                {
                    if (faceCollToKeep.Contains(f.InternalName))
                    {
                        if (f.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
                        {
                            oColl.Add(f);
                        }
                    }
                }
                oBaseFeature.DeleteFaces(oColl);
                oBaseFeature.ExitEdit();
            }

            return faceCollToKeep;
        }
        public static void deleteFace(PartDocument oDoc, List<string> faceCollToKeep)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;

            FaceCollection faceCollToDelete = iApp.TransientObjects.CreateFaceCollection();

            foreach (Face oFace in oCompDef.SurfaceBodies[1].Faces)
            {
                if (!faceCollToKeep.Contains(oFace.InternalName))
                {
                    faceCollToDelete.Add(oFace);
                }
            }

            PartFeatures oFeat = oCompDef.Features;
            oFeat.DeleteFaceFeatures.Add(faceCollToDelete);
        }
        public static void createProfile(PartDocument oDoc)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;

            FaceCollection oFaceColl = iApp.TransientObjects.CreateFaceCollection();

            foreach (Face f in oCompDef.SurfaceBodies[1].Faces)
            {
                oFaceColl.Add(f);
            }

            PartFeatures oFeat = oCompDef.Features;

            oFeat.ThickenFeatures.Add(oFaceColl, oCompDef.Thickness.Value, PartFeatureExtentDirectionEnum.kNegativeExtentDirection, PartFeatureOperationEnum.kJoinOperation, true);
        }
    }
}
