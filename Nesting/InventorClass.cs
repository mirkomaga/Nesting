using Inventor;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Media3D;
using System.Windows.Shapes;

namespace Nesting
{
    class InventorClass
    {
        public static Inventor.Application iApp = null;
        private static Inventor.ApplicationEvents iAppE = null;
        public static void CreateSM(List<Rettangolo> oggetto, ToolStripProgressBar bar, ListView lv, string path)
        {
            getIstance();

            if (iApp != null)
            {
                bar.Maximum = oggetto.Count;
                int num = 0;
                foreach (Rettangolo rett in oggetto)
                {
                    num += 1;

                    PartDocument oDoc = createNewPart();
                    drawRectangle(oDoc, rett.sviluppo, rett.lunghezza);
                    try
                    {
                        saveDocument(@path + "//" + rett.filename + ".ipt", oDoc);
                        string[] row = { rett.filename, "OK" };
                        var listViewItem = new ListViewItem(row);
                        lv.Items.Add(listViewItem);
                        lv.Refresh();
                    }
                    catch
                    {
                        string[] row = { rett.filename, "Non salvato" };
                        var listViewItem = new ListViewItem(row);
                        lv.Items.Add(listViewItem);
                        lv.Refresh();
                    }

                    lv.EnsureVisible(lv.Items.Count - 1);

                    oDoc.Close();
                    bar.Value = num;
                }
            }
            System.Windows.Forms.MessageBox.Show("Operazione conclusa.");
        }
        public static void drawRectangle(PartDocument oSheetMetalDoc, double b, double h)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oSheetMetalDoc.ComponentDefinition;

            SheetMetalFeatures oSheetMetalFeatures = (SheetMetalFeatures)oCompDef.Features;

            PlanarSketch oSketch = default(PlanarSketch);

            oSketch = oCompDef.Sketches.Add(oCompDef.WorkPlanes[3]);

            TransientGeometry oTg = (TransientGeometry)iApp.TransientGeometry;

            ObjectCollection oCollPnts = iApp.TransientObjects.CreateObjectCollection();

            oCollPnts.Add(oTg.CreatePoint2d(0, 0));
            oCollPnts.Add(oTg.CreatePoint2d(b / 10, 0));
            oCollPnts.Add(oTg.CreatePoint2d(b / 10, h / 10));
            oCollPnts.Add(oTg.CreatePoint2d(0, h / 10));

            Polyline2d polilinea = oTg.CreatePolyline2d(oCollPnts);

            SketchLine[] oLines = new SketchLine[polilinea.PointCount];

            oLines[0] = oSketch.SketchLines.AddByTwoPoints(polilinea.PointAtIndex[1], polilinea.PointAtIndex[2]);
            oLines[1] = oSketch.SketchLines.AddByTwoPoints(oLines[0].EndSketchPoint, polilinea.PointAtIndex[3]);
            oLines[2] = oSketch.SketchLines.AddByTwoPoints(oLines[1].EndSketchPoint, polilinea.PointAtIndex[4]);
            oLines[3] = oSketch.SketchLines.AddByTwoPoints(oLines[2].EndSketchPoint, oLines[0].StartSketchPoint);


            //oSketch.SketchLines.AddAsTwoPointRectangle(
            //    oTransGeom.CreatePoint2d(0, 0),
            //    oTransGeom.CreatePoint2d(h, b)
            //);


            Profile oProfile = (Profile)oSketch.Profiles.AddForSolid();

            FaceFeatureDefinition oFaceFeatureDefinition = (FaceFeatureDefinition)oSheetMetalFeatures.FaceFeatures.CreateFaceFeatureDefinition(oProfile);

            FaceFeature oFaceFeature = default(FaceFeature);

            oFaceFeature = oSheetMetalFeatures.FaceFeatures.Add(oFaceFeatureDefinition);

            oCompDef.Unfold();
            oCompDef.FlatPattern.ExitEdit();

            //Face oFrontFace = default(Face);

            //oFrontFace = oFaceFeature.Faces[6];

            //PlanarSketch oFoldLineSketch = (PlanarSketch)oCompDef.Sketches.Add(oFrontFace);

            //Inventor.Point oEdge1MidPoint = (Inventor.Point)oFrontFace.Edges[1].Geometry.MidPoint;


            //Point2d oSketchPoint1 =

            //   (Point2d)oFoldLineSketch.

            //         ModelToSketchSpace(oEdge1MidPoint);





            //Inventor.Point oEdge2MidPoint =

            //    (Inventor.Point)oFrontFace.Edges[3].

            //                          Geometry.MidPoint;



            //Point2d oSketchPoint2 =

            //    (Point2d)oFoldLineSketch.

            //         ModelToSketchSpace(oEdge2MidPoint);



            // Create the fold line between the midpoint

            // of two opposite edges on the face

            //SketchLine oFoldLine =

            //    (SketchLine)oFoldLineSketch.SketchLines.

            // AddByTwoPoints(oSketchPoint1, oSketchPoint2);



            //FoldDefinition oFoldDefinition =

            //    (FoldDefinition)oSheetMetalFeatures.

            //    FoldFeatures.CreateFoldDefinition

            //                      (oFoldLine, "60 deg");



            //// Create a fold feature

            //FoldFeature oFoldFeature =

            //    (FoldFeature)oSheetMetalFeatures.

            //          FoldFeatures.Add(oFoldDefinition);


            iApp.ActiveView.GoHome();
        }
        private static void getIstance()
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
        private static PartDocument createNewPart()
        {
            PartDocument oSheetMetalDoc = (PartDocument)iApp.Documents.Add(
                DocumentTypeEnum.kPartDocumentObject, iApp.FileManager.GetTemplateFile
                 (DocumentTypeEnum.kPartDocumentObject,
                   SystemOfMeasureEnum.kDefaultSystemOfMeasure,
                    DraftingStandardEnum.kDefault_DraftingStandard,
                     "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"));

            return (PartDocument)oSheetMetalDoc;
        }
        private static void saveDocument(string name, PartDocument doc)
        {
            doc.SaveAs(@name, false);
        }
        public static Polyline2d creoPoliLinea(double lunghezza, double sviluppo)
        {
            getIstance();

            TransientGeometry oTG = iApp.TransientGeometry;

            ObjectCollection oColl = iApp.TransientObjects.CreateObjectCollection();

            oColl.Add(oTG.CreatePoint2d(0, 0));
            oColl.Add(oTG.CreatePoint2d(lunghezza / 10, 0));
            oColl.Add(oTG.CreatePoint2d(lunghezza / 10, sviluppo / 10));
            oColl.Add(oTG.CreatePoint2d(0, sviluppo / 10));

            return oTG.CreatePolyline2d(oColl);
        }
        public static void changeThksSheetsMtl(string path, ListView lvThks, System.Windows.Forms.ProgressBar pbThk)
        {
            getIstance();

            if (iApp != null)
            {

                string[] listFiles = System.IO.Directory.GetFiles(@path, "*.ipt");

                pbThk.Maximum = listFiles.Length;

                int counter = 0;
                foreach (string file in listFiles)
                {
                    counter++;
                    pbThk.Value = counter;

                    int percent = (int)(((double)pbThk.Value / (double)pbThk.Maximum) * 100);
                    pbThk.Refresh();
                    pbThk.CreateGraphics().DrawString(percent.ToString() + "%",
                        new System.Drawing.Font("Arial", (float)8.25, FontStyle.Regular),
                        Brushes.Black,
                        new PointF(pbThk.Width / 2 - 10, pbThk.Height / 2 - 7));


                    if (System.IO.Path.GetExtension(file) == ".ipt")
                    {
                        PartDocument oSheetMetalDoc = (PartDocument)iApp.Documents.Open(@file);

                        iApp.ActiveView.GoHome();

                        // conversione in sm
                        try
                        {
                            oSheetMetalDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}";
                        }
                        catch { }

                        SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oSheetMetalDoc.ComponentDefinition;

                        List<Face> faces = findFace(oSheetMetalDoc);

                        IDictionary<double, Face> facesPart = new Dictionary<double, Face>();

                        foreach (Face fc in faces)
                        {
                            if (!facesPart.ContainsKey(fc.Evaluator.Area))
                                facesPart.Add(fc.Evaluator.Area, fc);
                        }

                        var ordered = facesPart.OrderByDescending(x => x.Key);

                        bool changed = false;

                        // parto dall'area più grande
                        foreach (var fc in ordered)
                        {
                            double dist = getDistanceFromFace(fc.Value);
                            if (dist > 0.0)
                            {
                                try
                                {
                                    changeSMProperti(dist, oCompDef);
                                    changed = true;
                                    break;
                                }
                                catch
                                {
                                    // ! se non trova la thck passa alla seconda area più grande
                                    Console.WriteLine("IMPO");
                                }
                            }
                        }

                        if (changed)
                        {
                            ListViewItem item1 = new ListViewItem(System.IO.Path.GetFileName(oSheetMetalDoc.FullFileName), 0);
                            item1.SubItems.Add("Ok");
                            lvThks.Items.AddRange(new ListViewItem[] { item1 });
                        }
                        else
                        {
                            ListViewItem item1 = new ListViewItem(System.IO.Path.GetFileName(oSheetMetalDoc.FullFileName), 0);
                            item1.SubItems.Add("ATTENZIONEE");
                            lvThks.Items.AddRange(new ListViewItem[] { item1 });
                        }

                        lvThks.EnsureVisible(lvThks.Items.Count - 1);

                        oCompDef.Unfold();
                        iApp.ActiveView.GoHome();
                        oSheetMetalDoc.Save();
                        oSheetMetalDoc.Close();
                    }
                }
            }

            MessageBox.Show("Operazione copletata", "Info");
        }
        private static double getDistanceFromFace(Face fc)
        {
            Inventor.Point origin = fc.PointOnFace;

            double[] pt = new double[3] { origin.X, origin.Y, origin.Z };

            double[] n = new double[3];

            fc.Evaluator.GetNormalAtPoint(pt, ref n);

            UnitVector normal = iApp.TransientGeometry.CreateUnitVector(-n[0], -n[1], -n[2]);

            SurfaceBody body = fc.Parent;

            ObjectsEnumerator objects;
            ObjectsEnumerator pts;

            body.FindUsingRay(origin, normal, 0.001, out objects, out pts, true);

            double dist = iApp.MeasureTools.GetMinimumDistance(origin, objects[2]) * 10;

            return dist;
        }
        private static void changeSMProperti(double sps, SheetMetalComponentDefinition oCompDef)
        {
            string toSelect = null;

            sps = Math.Round(sps);

            switch (sps)
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
                    Console.WriteLine("SONO ENTRATO IN DEFAULT");
                    break;
            }

            if (!string.IsNullOrEmpty(toSelect))
            {
                SheetMetalStyle oStyle = oCompDef.SheetMetalStyles[toSelect];
                oStyle.Activate();

                Console.WriteLine("Impostato: " + toSelect);
            }
            else
            {
                throw new Exception("Nessuno spessore trovato");
            }
        }
        public static List<Face> findFace(PartDocument oSheetMetalDoc)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oSheetMetalDoc.ComponentDefinition;
            List<Face> faces = new List<Face>();

            foreach (Face partFace in oCompDef.SurfaceBodies[1].Faces)
            {
                if (partFace.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                {
                    faces.Add(partFace);
                }
            }

            return faces;
        }
        public static void disegnaProfiloBello(string path)
        {
            getIstance();
            if (iApp != null)
            {

                string[] listFiles = System.IO.Directory.GetFiles(@path, "*.ipt");

                int counter = 0;
                foreach (string file in listFiles)
                {
                    counter++;

                    if (System.IO.Path.GetExtension(file) == ".ipt")
                    {
                        PartDocument oSheetMetalDoc = (PartDocument)iApp.Documents.Open(@file);

                        //PartDocument oSheetMetalDoc = (PartDocument)iApp.Documents.Open(@"C:\\Users\\edgelocal\\Desktop\\testL.ipt");
                        InventorClass.clearAllSup(oSheetMetalDoc);

                        iApp.ActiveView.GoHome();

                        // conversione in sm
                        if(oSheetMetalDoc.SubType != "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                        {
                            oSheetMetalDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}";
                        }

                        SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oSheetMetalDoc.ComponentDefinition;
                        
                        double thk = oCompDef.Thickness.Value;

                        List<Face> faces = findFace(oSheetMetalDoc);

                        FaceCollection facesPart = iApp.TransientObjects.CreateFaceCollection();

                        foreach (Face fc in faces)
                        {
                            facesPart.Add(fc);
                        }

                        //prendo la faccia più grande
                        Face fcBig = InventorClass.getBiggestPlanarFaceOfColl(facesPart);

                        //prendo lo spigolo più lungo
                        Edge egBig = InventorClass.getLongestEdgeFace(fcBig);


                        WorkPoint oWpoint = oCompDef.WorkPoints.AddByMidPoint(egBig);
                        oWpoint.Name = "wPntReference";
                        oWpoint.Visible = false;

                        WorkPlane oWpReference = oCompDef.WorkPlanes.AddByNormalToCurve(egBig, oWpoint);
                        oWpReference.Name = "wpReference";
                        oWpReference.Visible = false;

                        WorkPlane oWpWork = oCompDef.WorkPlanes.AddByPlaneAndOffset(oWpReference, 0);
                        oWpWork.Name = "wpWork";
                        oWpWork.Visible = true;

                        PlanarSketch oSketch = oCompDef.Sketches.Add(oWpWork);
                        ProjectedCut oProjectCut = oSketch.ProjectedCuts.Add();

                        int nTmpLinee = oProjectCut.SketchEntities.Count / 2;
                        int nLinee = InventorClass.getLineeCutCompleto(oCompDef);

                        bool errLinee = false;
                        
                        int offset = 1;

                        int loop = 0;
                        while (nLinee != nTmpLinee)
                        {
                            // Devo spostare il piano se ci sono cose nel mezzo
                            oWpWork.SetByPlaneAndOffset(oWpReference, offset);

                            nTmpLinee = oProjectCut.SketchEntities.Count / 2;
                            nLinee = InventorClass.getLineeCutCompleto(oCompDef);
                            loop++;

                            offset += offset;
                            
                            if (loop == 20)
                            {
                                errLinee = true;
                                break;
                            }
                        }

                        if (!errLinee)
                        {
                            IDictionary<string, List<SketchEntity>> dati = InventorClass.processoEntita(oSketch, thk);

                            InventorClass.offsetSketch(oSketch, dati["lunga"]);
                            //InventorClass.offsetListaLinee(dati["lunga"]);
                        }

                        iApp.ActiveView.GoHome();
                        //oSheetMetalDoc.Save();
                        oSheetMetalDoc.Close();
                    }
                }
            }
        }
        public static void offsetSketch(PlanarSketch oSketch, List<SketchEntity> listaLinee)
        {
            TransientGeometry oTransGeom = iApp.TransientGeometry;
            UnitVector oNormalVector = oSketch.PlanarEntityGeometry.Normal;

            UnitVector2d oLineDir = null;
            foreach (SketchEntity e in listaLinee) 
            {
                if (e.Type == ObjectTypeEnum.kSketchLineObject)
                {
                    SketchLine ent = (SketchLine)e;
                    oLineDir = (UnitVector2d) ent.Geometry.Direction;
                    break;
                }
            }

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

            ObjectCollection ll = iApp.TransientObjects.CreateObjectCollection();

            foreach(SketchEntity e in listaLinee)
            {
                if(e.Type == ObjectTypeEnum.kSketchLineObject)
                {
                    SketchLine skl = (SketchLine)e;
                    ll.Add(skl);
                }
            }
            foreach (SketchEntity e in listaLinee)
            {
                if (e.Type == ObjectTypeEnum.kSketchArcObject)
                {
                    SketchArc ska = (SketchArc)e;
                    ll.Add(ska);
                }
            }

            SketchEntitiesEnumerator oSSketchEntitiesEnum = oSketch.OffsetSketchEntitiesUsingDistance(ll, 1, bNaturalOffsetDir, false);
        }
        public static void offsetListaLinee(List<SketchEntity> lineeList)
        {
            foreach (SketchEntity e in lineeList)
            {
                if (e.Type == ObjectTypeEnum.kSketchLineObject)
                {
                    SketchLine ent = (SketchLine)e;
                    
                }
                if (e.Type == ObjectTypeEnum.kSketchArcObject)
                {
                    SketchArc ent = (SketchArc)e;
                }
            }
        }
        public static IDictionary<string, List<SketchEntity>> processoEntita(PlanarSketch oSketch, double Thickness)
        {
            IDictionary<string, List<SketchEntity>> results = new Dictionary<string, List<SketchEntity>>();
            results.Add("prima", new List<SketchEntity>());
            results.Add("seconda", new List<SketchEntity>());

            Inventor.Color oColorRed = iApp.TransientObjects.CreateColor(255, 0, 0);
            Inventor.Color oColorGreen = iApp.TransientObjects.CreateColor(0, 255, 0);
            Inventor.Color oColorBlue = iApp.TransientObjects.CreateColor(0, 0, 255);

            string selettore = "prima";
            bool add = true;
            
            foreach (SketchEntity entita in oSketch.SketchEntities)
            {
                if (entita.Type == ObjectTypeEnum.kSketchLineObject)
                {
                    SketchLine ent =(SketchLine)entita;
                    double lunghezzaL = Math.Round(ent.Length, 1);
                                        
                    if(Thickness == lunghezzaL)
                    {
                        // ho trovato entita lunga quanto spessore
                        add = false;
                        ent.OverrideColor = oColorRed;
                        if(selettore == "prima")
                        {
                            selettore = "seconda";
                        }
                        else
                        {
                            selettore = "prima";
                        }
                    }
                    else
                    {
                        // ho trovato entita normale
                        add = true;
                        ent.OverrideColor = oColorGreen;
                    }

                }

                else if (entita.Type == ObjectTypeEnum.kSketchArcObject)
                {
                    SketchArc ent = (SketchArc)entita;
                    add = true;
                    ent.OverrideColor = oColorGreen;
                }

                if((entita.Type == ObjectTypeEnum.kSketchLineObject || entita.Type == ObjectTypeEnum.kSketchArcObject) && add)
                {
                    results[selettore].Add(entita);
                }
            }

            double countPol1 = InventorClass.perimetroLines(results["prima"]);
            double countPol2 = InventorClass.perimetroLines(results["seconda"]);
            
            IDictionary<string, List<SketchEntity>> dati = new Dictionary<string, List<SketchEntity>>();
            if (countPol1> countPol2)
            {
                dati.Add("lunga", results["prima"]);
                dati.Add("corta", results["seconda"]);
            }
            else
            {
                dati.Add("lunga", results["seconda"]);
                dati.Add("corta", results["prima"]);
            }

            foreach (SketchEntity e in dati["lunga"])
            {
                if (e.Type == ObjectTypeEnum.kSketchLineObject)
                {
                    SketchLine ent = (SketchLine)e;
                    ent.OverrideColor = oColorBlue;
                }
                if (e.Type == ObjectTypeEnum.kSketchArcObject)
                {
                    SketchArc ent = (SketchArc)e;
                    ent.OverrideColor = oColorBlue;
                }
            }

            return dati;
        }
        public static double perimetroLines(List<SketchEntity> linee)
        {
            double result = 0;
            foreach (SketchEntity e in linee)
            {
                if(e.Type == ObjectTypeEnum.kSketchLineObject)
                {
                    SketchLine ent = (SketchLine) e;
                    result += ent.Length;
                }
                if (e.Type == ObjectTypeEnum.kSketchArcObject)
                {
                    SketchArc ent = (SketchArc) e;
                    result += ent.Length;
                }
            }
            return result;
        }
        public static int getLineeCutCompleto(SheetMetalComponentDefinition oCompDef)
        {
            int nPieghe = (oCompDef.Bends.Count *4) + 4;

            return nPieghe;
        }
        public static Face getBiggestPlanarFaceOfColl(FaceCollection fcs)
        {
            Face fc = null;

            foreach(Face x in fcs)
            {
                if (fc == null || x.Evaluator.Area > fc.Evaluator.Area)
                {
                    fc = x;
                }
            }

            return fc;
        }
        public static Edge getLongestEdgeFace(Face oFace)
        {
            Edge edge = null;

            double highest = 0;

            foreach (Edge tmpEdge in oFace.Edges)
            {
                if(tmpEdge.GeometryType == CurveTypeEnum.kLineSegmentCurve)
                {
                    double tmpLength = iApp.MeasureTools.GetMinimumDistance(tmpEdge.StartVertex, tmpEdge.StopVertex);

                    if (tmpLength > highest)
                    {
                        highest = tmpLength;
                        edge = tmpEdge;
                    }
                }
            }

            return (Edge) edge;
        }
        public static void coplanarFace(IDictionary<int, Face> facesPart)
        {

            double[] points = new double[3];

            if(facesPart.Count > 0)
            {
                Face oFace;
                facesPart.TryGetValue(facesPart.Keys.First(), out oFace);

                //oFace.Evaluator.GetPointAtParam(params, points);

                //Point oPoint = iApp.TransientGeometry.CreatePoint(points(0), points(1), points(2))
            }
        }
        public static void clearAllSup(PartDocument oDoc)
        {
            //SheetMetalComponentDefinition oComp = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;

            //SheetMetalFeatures oSheetFeat = (SheetMetalFeatures)oComp.Features;

            //foreach (object obj in oComp.AppearanceOverridesObjects)
            //{
            //    if (obj is SurfaceBodies)
            //    {
            //        Console.WriteLine("dsa");
            //    }
            //    else if (obj is PartFeature)
            //    {
            //        Console.WriteLine("dsa");

            //    }
            //    else if (obj is Face)
            //    {
            //        obj.AppearanceSourceType = ObjectsEnumerator.kFeatureAppearance;
            //    }
            //    else
            //    {
            //        Console.WriteLine("Obj sconosciuto");
            //    }
            //}
        }
    }
}