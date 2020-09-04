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
            setSheetMetalDocument(oDoc);

            // ! Imposto thickness
            setThickness(oDoc);

            // ! Elimino raggiature
            List<string> faceCollToKeep = deleteFillet(oDoc);

            // ! Elimino facce
            deleteFace(oDoc, faceCollToKeep);

            // ! Creo Profilo
            createProfile(oDoc);

            // ! Creo Raggiature
            createFillet(oDoc);

            // ! Cerco le lavorazioni
            IDictionary<Face, List<Lavorazione>> lavorazione = detectLavorazioni(oDoc);

            // ! Creo sketch lavorazioni
            List<string>  nomeSketch = createSketchLavorazione(oDoc, lavorazione);

            // ! Elimino con direct le lavorazioni
            deleteLavorazione(oDoc);

            // ! Cut lavorazione
            createCutLavorazione(oDoc, nomeSketch);

            // ! Aggiungo piano nel mezzo
            WorkPlane oWpReference = addPlaneInTheMiddleOfBox(oDoc);

            // ! Aggiungo proiezione cut
            addProjectCut(oDoc, oWpReference);

            // ! Coloro lato bello
            setTexture(oDoc);
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
        public static void createFillet(PartDocument oDoc)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;

            SheetMetalFeatures sFeatures = (SheetMetalFeatures)oCompDef.Features;

            foreach (Edge oEdge in oCompDef.SurfaceBodies[1].ConcaveEdges)
            {
                int tmpCount = oCompDef.SurfaceBodies[1].ConcaveEdges.Count;

                try
                {
                    EdgeCollection oBendEdges = iApp.TransientObjects.CreateEdgeCollection();

                    oBendEdges.Add(oEdge);

                    BendDefinition oBendDef = sFeatures.BendFeatures.CreateBendDefinition(oBendEdges);

                    BendFeature oBendFeature = sFeatures.BendFeatures.Add(oBendDef);

                    if (tmpCount != oCompDef.SurfaceBodies[1].ConcaveEdges.Count)
                    {
                        createFillet(oDoc);
                        break;
                    }
                }
                catch { }
            }
        }
        public static IDictionary<Face, List<Lavorazione>> detectLavorazioni(PartDocument oDoc)
        {
            IDictionary<Face, List<Lavorazione>> result = new Dictionary<Face, List<Lavorazione>>();

            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;

            FaceCollection oFaceColl = oCompDef.Bends[1].FrontFaces[1].TangentiallyConnectedFaces;
            oFaceColl.Add(oCompDef.Bends[1].FrontFaces[1]);

            foreach (Face f in oFaceColl)
            {
                if (f.EdgeLoops.Count > 1)
                {
                    List<Lavorazione> lavorazione = IdentificazioneEntita.main(f.EdgeLoops, iApp);

                    if (lavorazione.Count > 0)
                    {
                        result.Add(f, lavorazione);
                    }
                }
            }

            return result;
        }
        public static List<string> createSketchLavorazione(PartDocument oDoc, IDictionary<Face, List<Lavorazione>> lavorazione)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;

            List<string> result = new List<string>();

            int index = 1;
            foreach (var x in lavorazione) 
            {
                Face oFace = x.Key;
                List<Lavorazione> lavList = x.Value;

                foreach (Lavorazione lav in lavList)
                {
                    PlanarSketch oSketch = oCompDef.Sketches.Add(oFace, false);

                    foreach(Edge oEdge in lav.oEdgeColl_)
                    {
                        oSketch.AddByProjectingEntity(oEdge);
                        oSketch.Visible = false;
                    }

                    string nameSketch = lav.nameLav_ + "_" + index;
                    oSketch.Name = nameSketch;
                    result.Add(nameSketch);

                    index++;
                }
            }

            return result;
        }
        public static void deleteLavorazione(PartDocument oDoc)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;
            int numeroFacceTan = oCompDef.Bends.Count * 2;

            NonParametricBaseFeature oBaseFeature = oCompDef.Features.NonParametricBaseFeatures[1];

            oBaseFeature.Edit();

            SurfaceBody basebody = oBaseFeature.BaseSolidBody;

            foreach (Face f in basebody.Faces)
            {
                if(f.TangentiallyConnectedFaces.Count == numeroFacceTan)
                {
                    string nameFace = f.InternalName;

                    ObjectCollection oColl = iApp.TransientObjects.CreateObjectCollection();

                    foreach (EdgeLoop oEdgeLoops in f.EdgeLoops)
                    {
                        Edges oEdges = oEdgeLoops.Edges;

                        string lav = IdentificazioneEntita.whois(oEdges);

                        if (!string.IsNullOrEmpty(lav))
                        {
                            Edge oEdge = oEdges[1];

                            Faces oFaceColl = oEdge.Faces;

                            foreach (Face oFaceLav in oFaceColl)
                            {
                                if (oFaceLav.InternalName != nameFace)
                                {
                                    oColl.Add(oFaceLav);

                                    foreach (Face oFaceLavTang in oFaceLav.TangentiallyConnectedFaces)
                                    {
                                        oColl.Add(oFaceLavTang);
                                    }
                                }
                            }
                        }
                    }

                    if (oColl.Count > 0)
                    {
                        try
                        {
                            oBaseFeature.DeleteFaces(oColl);
                        }
                        catch
                        {

                        }
                    }
                }
            }
            oBaseFeature.ExitEdit();
        }
        public static void createCutLavorazione(PartDocument oDoc, List<string> nomeSketch)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;

            foreach (string nameS in nomeSketch)
            {
                try
                {
                    PlanarSketch oSketch = oCompDef.Sketches[nameS];

                    SheetMetalFeatures oSheetMetalFeatures = (SheetMetalFeatures)oCompDef.Features;

                    Profile oProfile = oSketch.Profiles.AddForSolid();

                    CutDefinition oCutDefinition = oSheetMetalFeatures.CutFeatures.CreateCutDefinition(oProfile);

                    //oCutDefinition.SetThroughAllExtent(PartFeatureExtentDirectionEnum.kNegativeExtentDirection);

                    CutFeature oCutFeature = oSheetMetalFeatures.CutFeatures.Add(oCutDefinition);
                }
                catch
                {
                    throw new Exception("Nome sketch non esiste: " + nameS);
                }
            }
        }
        public static WorkPlane addPlaneInTheMiddleOfBox(PartDocument oDoc)
        {
            SheetMetalComponentDefinition oComp = (SheetMetalComponentDefinition) oDoc.ComponentDefinition;

            Box oRb = oComp.SurfaceBodies[1].RangeBox;

            TransientBRep oTransientBRep = iApp.TransientBRep;

            SurfaceBody oBody = oTransientBRep.CreateSolidBlock(oRb);

            NonParametricBaseFeature oBaseFeature = oComp.Features.NonParametricBaseFeatures.Add(oBody);

            FaceCollection oFaceColl = iApp.TransientObjects.CreateFaceCollection();

            foreach (Face f in oBaseFeature.SurfaceBodies[1].Faces)
            {
                WorkPlane tmpWp = oComp.WorkPlanes.AddByPlaneAndOffset(f, 0);

                if (tmpWp.Plane.IsParallelTo[oComp.WorkPlanes[1].Plane])
                {
                    oFaceColl.Add(f);
                }

                tmpWp.Delete();
            }

            WorkPlane wpWork = null;
            if (oFaceColl.Count >= 2)
            {
                WorkPlane wp1 = oComp.WorkPlanes.AddByPlaneAndOffset(oFaceColl[1], 0);
                WorkPlane wp2 = oComp.WorkPlanes.AddByPlaneAndOffset(oFaceColl[2], 0);

                wpWork = oComp.WorkPlanes.AddByTwoPlanes(wp1, wp2);
                wpWork.Name = "wpWorkReference";
                wpWork.Grounded = true;
                wpWork.Visible = false;

                oBaseFeature.Delete(false, true, true);

                wp1.Delete();
                wp2.Delete();
            }

            return wpWork;
        }
        public static void addProjectCut(PartDocument oDoc, WorkPlane oWpReference)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;

            WorkPlane oWpWork = oCompDef.WorkPlanes.AddByPlaneAndOffset(oWpReference, 0);
            oWpWork.Name = "wpWork";
            oWpWork.Visible = false;

            PlanarSketch oSketch = oCompDef.Sketches.Add(oWpWork);
            ProjectedCut oProjectCut = oSketch.ProjectedCuts.Add();

            // TODO RICORDARSI DI FARE SE DIVERSO DA 2
            int tmpSegmThk = countThicknessSegment(oProjectCut.SketchEntities);

            List<ObjectCollection> dataLine = splittoLinea(oProjectCut.SketchEntities);

            ObjectCollection linea = lengthPerimetro();

            TransientGeometry oTransGeom = iApp.TransientGeometry;
            UnitVector oNormalVector = oSketch.PlanarEntityGeometry.Normal;
            UnitVector2d oLineDir = linea[1].Geometry.Direction;
            UnitVector oLineVector = (UnitVector)oTransGeom.CreateUnitVector(oLineDir.X, oLineDir.Y, 0);
            UnitVector oOffsetVector = (UnitVector)oLineVector.CrossProduct(oNormalVector);
            UnitVector oDesiredVector = (UnitVector)oTransGeom.CreateUnitVector(0, 1, 0);

            bool bNaturalOffsetDir;

            if (oOffsetVector.IsEqualTo(oDesiredVector))
            {
                bNaturalOffsetDir = true;
            }
            else
            {
                bNaturalOffsetDir = false;
            }

            SketchEntitiesEnumerator oSSketchEntitiesEnum = oSketch.OffsetSketchEntitiesUsingDistance(linea, 0.5, bNaturalOffsetDir, false);
            oProjectCut.Delete();

            coloroLinea(oSketch.SketchEntities);
            
            int countThicknessSegment(SketchEntitiesEnumerator oSketchEntities)
            {
                int result = 0;

                foreach (SketchEntity e in oSketchEntities)
                {
                    if (e.Type == ObjectTypeEnum.kSketchLineObject)
                    {
                        SketchLine oSketchLine = (SketchLine) e;

                        double length = Math.Round(oSketchLine.Length * 100)/100;

                        if (length == oCompDef.Thickness.Value)
                        {
                            result++;
                        }
                    }
                }

                return result;
            }
            List<ObjectCollection> splittoLinea(SketchEntitiesEnumerator oSketchEntities)
            {
                List<ObjectCollection> tmp = new List<ObjectCollection>();
                tmp.Add(iApp.TransientObjects.CreateObjectCollection());
                tmp.Add(iApp.TransientObjects.CreateObjectCollection());
                tmp.Add(iApp.TransientObjects.CreateObjectCollection());

                int indice = 0;
                foreach (SketchEntity oSketchEntity in oSketchEntities)
                {
                    if (oSketchEntity.Type != ObjectTypeEnum.kSketchPointObject)
                    {
                        if (oSketchEntity.Type == ObjectTypeEnum.kSketchLineObject) 
                        {
                            SketchLine oSketchLine = (SketchLine)oSketchEntity;
                            
                            if (Math.Round(oSketchLine.Length * 100) / 100 == oCompDef.Thickness.Value)
                            {
                                if(indice == 0)
                                {
                                    indice = 1;
                                }
                                else
                                {
                                    indice = 2;
                                }
                            }
                            else
                            {
                                tmp[indice].Add(oSketchEntity);
                            }
                        }

                        if (oSketchEntity.Type == ObjectTypeEnum.kSketchArcObject) 
                        {
                            tmp[indice].Add(oSketchEntity);
                        }
                    }
                }

                foreach (SketchEntity oSketchEntity in tmp[0])
                {
                    tmp[2].Add(oSketchEntity);
                }

                List<ObjectCollection> result = new List<ObjectCollection>();
                result.Add(tmp[1]);
                result.Add(tmp[2]);

                return result;
            }
            void coloroLinea(SketchEntitiesEnumerator oSketchEntities)
            {
                foreach(SketchEntity oSe in oSketchEntities)
                {
                    // TODO coloro
                    Console.WriteLine("COLORO");
                }
            }
            ObjectCollection lengthPerimetro()
            {
                double l0 = 0;
                double l1 = 0;

                foreach (SketchEntity oSE in dataLine[0])
                {
                    if (oSE.Type == ObjectTypeEnum.kSketchLineObject)
                    {
                        SketchLine se = (SketchLine)oSE;

                        l0 = l0 + se.Length;
                    }
                    else if (oSE.Type == ObjectTypeEnum.kSketchArcObject)
                    {
                        SketchArc se = (SketchArc)oSE;

                        l0 = l0 + se.Length;
                    }
                }
                foreach (SketchEntity oSE in dataLine[1])
                {
                    if (oSE.Type == ObjectTypeEnum.kSketchLineObject)
                    {
                        SketchLine se = (SketchLine)oSE;

                        l1 = l1 + se.Length;
                    }
                    else if (oSE.Type == ObjectTypeEnum.kSketchArcObject)
                    {
                        SketchArc se = (SketchArc)oSE;

                        l1 = l1 + se.Length;
                    }
                }

                if (l0>l1)
                {
                    return dataLine[0];
                }
                else
                {
                    return dataLine[1];
                }
            }
        }
        public static void setTexture(PartDocument oDoc)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;

            FaceCollection fcFront = oCompDef.Bends[1].FrontFaces[1].TangentiallyConnectedFaces;
            fcFront.Add(oCompDef.Bends[1].FrontFaces[1]);

            FaceCollection fcBack = oCompDef.Bends[1].BackFaces[1].TangentiallyConnectedFaces;
            fcBack.Add(oCompDef.Bends[1].BackFaces[1]);

            double area0 = 0;
            foreach (Face oFace in fcFront)
            {
                area0 = area0+oFace.Evaluator.Area;
            }

            double area1 = 0;
            foreach (Face oFace in fcBack)
            {
                area1 = area1 + oFace.Evaluator.Area;
            }

            Asset oAsset;

            try
            {
                oAsset = oDoc.Assets["RawSide"];
            }
            catch (System.ArgumentException e)
            {
                Assets oAssets = oDoc.Assets;

                AssetLibrary oAssetsLib = iApp.AssetLibraries["3D_Pisa_Col"];
                Asset oAssetLib = oAssetsLib.AppearanceAssets["RawSide"];

                oAsset = oAssetLib.CopyTo(oDoc);
            }

            FaceCollection fc;

            if (area0>area1)
            {
                fc = fcFront;
            }
            else
            {
                fc = fcBack;
            }

            foreach (Face f in fc)
            {
                f.Appearance = oAsset;
            }
        }
    }
}
