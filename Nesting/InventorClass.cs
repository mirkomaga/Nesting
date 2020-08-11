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
        public static void changeThksOpenFile(PartDocument oSheetMetalDoc)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oSheetMetalDoc.ComponentDefinition;
            List<Face> faces = InventorClass.findFace(oSheetMetalDoc);

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
            if (!changed)
            {
                throw new Exception("Spessore non trovato.");
            }
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

            sps = Math.Truncate(sps*100)/100;

            Console.WriteLine(sps);

            sps = Math.Round(sps,1);

            Console.WriteLine(sps);

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
                    //toSelect = "Default_mm";
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

                        //PartDocument oSheetMetalDoc = (PartDocument)iApp.Documents.Open(@"C:\\Users\\Mirko Magalotti\\Documents\\tttttttipt.ipt");
                        InventorClass.clearAllSup(oSheetMetalDoc);

                        iApp.ActiveView.GoHome();

                        // conversione in sm
                        if(oSheetMetalDoc.SubType != "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
                        {
                            oSheetMetalDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}";
                        }
                        
                        InventorClass.ricostruiscoLamiera(oSheetMetalDoc);

                        try
                        {
                            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oSheetMetalDoc.ComponentDefinition;
                            WorkPlane checkPlane = oCompDef.WorkPlanes["wpReference"];

                            iApp.ActiveView.GoHome();
                            oSheetMetalDoc.Save();
                            oSheetMetalDoc.Close();
                        }
                        catch
                        {

                            //imposto spessore
                            try
                            {
                                InventorClass.changeThksOpenFile(oSheetMetalDoc);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e);
                            }

                            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oSheetMetalDoc.ComponentDefinition;
                            oCompDef.UseSheetMetalStyleThickness = false;
                            Console.WriteLine("Bends Count :" + oCompDef.Bends.Count);
                            
                            List<Face> faces = findFace(oSheetMetalDoc);

                            FaceCollection facesPart = iApp.TransientObjects.CreateFaceCollection();

                            foreach (Face fc in faces)
                            {
                                facesPart.Add(fc);
                            }

                            //prendo la faccia più grande
                            Face fcBig = InventorClass.getBiggestPlanarFaceOfColl(facesPart);

                            WorkPlane oWpReference = InventorClass.getLongestEdgeFace(fcBig, oCompDef, oSheetMetalDoc);
                            oWpReference.Name = "wpReference";
                            oWpReference.Visible = false;

                            if (!oWpReference.Plane.IsParallelTo[oCompDef.WorkPlanes[1].Plane])
                            {
                                //throw new Exception("Piano non in asse");
                            }

                            WorkPlane oWpWork = oCompDef.WorkPlanes.AddByPlaneAndOffset(oWpReference, 0);
                            oWpWork.Name = "wpWork";
                            oWpWork.Visible = false;

                            PlanarSketch oSketch = oCompDef.Sketches.Add(oWpWork);
                            ProjectedCut oProjectCut = oSketch.ProjectedCuts.Add();

                            double thk = oCompDef.Thickness.Value;
                            int nTmpLinee = InventorClass.getoSketchEntityThickness(oProjectCut.SketchEntities, thk);
                            //int nLinee = InventorClass.getLineeCutCompleto(oCompDef);
                            int nLinee = 2;

                            //if (nLinee == 0)
                            //{
                            //    throw new Exception("Nessuna piega trovata.");
                            //}

                            bool errLinee = false;
                        
                            int offset = 1;

                            int loop = 0;
                            try
                            {

                                while (nLinee != nTmpLinee)
                                {
                                    // Devo spostare il piano se ci sono cose nel mezzo
                                    oWpWork.SetByPlaneAndOffset(oWpReference, offset);

                                    nTmpLinee = InventorClass.getoSketchEntityThickness(oProjectCut.SketchEntities, thk);
                                    loop++;

                                    offset += offset;
                            
                                    if (loop == 20)
                                    {
                                        errLinee = true;
                                        //throw new Exception("Numero massimo offset piano.");
                                    }
                                }
                            }
                            catch
                            {
                                errLinee = true;
                            }

                                                
                            if (!errLinee)
                            {
                                //oCompDef.SurfaceBodies[1].Visible = false;
                                oProjectCut.Delete();
                                
                                oProjectCut = oSketch.ProjectedCuts.Add();

                                Console.WriteLine("THIK: "+ thk);
                                IDictionary<string, List<SketchEntity>> dati = InventorClass.processoEntita(oSketch, thk);

                                //oCompDef.SurfaceBodies[1].Visible = true;

                                TransientGeometry oTransGeom = iApp.TransientGeometry;

                                InventorClass.offsetSketch(oSketch, dati["lunga"]);
                            
                                oProjectCut.Delete();

                                try
                                {
                                    IDictionary<string, FaceCollection> facce = InventorClass.getFacceInterneEsterne(oCompDef, oSheetMetalDoc);

                                    Asset oAsset;

                                    try
                                    {
                                        oAsset = oSheetMetalDoc.Assets["RawSide"];
                                    }
                                    catch (System.ArgumentException e)
                                    {
                                        Assets oAssets = oSheetMetalDoc.Assets;

                                        AssetLibrary oAssetsLib = iApp.AssetLibraries["3D_Pisa_Col"];
                                        Asset oAssetLib = oAssetsLib.AppearanceAssets["RawSide"];

                                        oAsset = oAssetLib.CopyTo(oSheetMetalDoc);
                                    }

                                    InventorClass.settingAssetsToFaces(facce["lunga"], oAsset);
                                }
                                catch
                                {
                                    throw new Exception("Impossibile colorare facce.");
                                }
                            }

                            iApp.ActiveView.GoHome();
                            oSheetMetalDoc.Save();
                            oSheetMetalDoc.Close();
                        }
                    }
                }
            }
        }
        public static int getoSketchEntityThickness(SketchEntitiesEnumerator seC, double Thickness)
        {
            int counter = 0;
            foreach (SketchEntity se in seC)
            {
                if(se.Type == ObjectTypeEnum.kSketchLineObject)
                {
                    SketchLine sl = (SketchLine)se;
                    double slLength = Math.Round(sl.Length * 100)/100;
                    if (slLength == Thickness)
                    {
                        counter += 1;
                    }
                }
            }
            return counter;
        }
        public static void settingAssetsToFaces(FaceCollection fc, Asset oAsset)
        {
            foreach(Face f in fc)
            {
                f.Appearance = oAsset;
            }
        }
        public static IDictionary<string, FaceCollection> getFacceInterneEsterne(SheetMetalComponentDefinition oCompDef, PartDocument oSheetMetalDoc)
        {
            IDictionary<string, FaceCollection> results = new Dictionary<string, FaceCollection>();
            
            results.Add("lunga", iApp.TransientObjects.CreateFaceCollection());
            results.Add("corta", iApp.TransientObjects.CreateFaceCollection());

            Bend piega = oCompDef.Bends[1];

            Face fInterna = piega.FrontFaces[1];
            Face fEsterna = piega.BackFaces[1];

            FaceCollection interne = fInterna.TangentiallyConnectedFaces;
            interne.Add(fInterna);

            FaceCollection esterne = fEsterna.TangentiallyConnectedFaces;
            esterne.Add(fEsterna);

            double sommaAreaI = InventorClass.sommaArea(interne);
            double sommaAreaE = InventorClass.sommaArea(esterne);

            if (sommaAreaI > sommaAreaE)
            {
                results["lunga"] = interne;
                results["corta"] = esterne;
            }
            else if (sommaAreaI < sommaAreaE)
            {
                results["lunga"] = esterne;
                results["corta"] = interne;
            }
            else
            {
                MessageBox.Show("Aree uguali", "Attenzione");
            }

            return results;
        }
        public static double sommaArea(FaceCollection coll)
        {
            double result = 0;

            foreach(Face f in coll)
            {
                result += f.Evaluator.Area;
            }

            return result;
        }
        public static void offsetSketch(PlanarSketch oSketch, List<SketchEntity> dati)
        {

            ObjectCollection oCollSide = iApp.TransientObjects.CreateObjectCollection();
            foreach (SketchEntity e in dati)
            {
                oCollSide.Add(e);
            }

            TransientGeometry oTransGeom = iApp.TransientGeometry;
            UnitVector oNormalVector = oSketch.PlanarEntityGeometry.Normal;
            UnitVector2d oLineDir = oCollSide[1].Geometry.Direction;
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

            SketchEntitiesEnumerator oSSketchEntitiesEnum = oSketch.OffsetSketchEntitiesUsingDistance(oCollSide, 0.5, bNaturalOffsetDir, false);
        }
        public static IDictionary<string, List<SketchEntity>> processoEntita(PlanarSketch oSketch, double Thickness)
        {
            IDictionary<string, List<SketchEntity>> results = new Dictionary<string, List<SketchEntity>>();

            results.Add("prima", new List <SketchEntity>());
            results.Add("seconda", new List <SketchEntity>());

            ObjectCollection cleanedSketch = iApp.TransientObjects.CreateObjectCollection();

            foreach(SketchEntity entita in oSketch.SketchEntities)
            {
                if (entita.Type == ObjectTypeEnum.kSketchLineObject || entita.Type == ObjectTypeEnum.kSketchArcObject)
                {
                    cleanedSketch.Add(entita);
                }
            }

            List<SketchEntity> tmpListSE = new List<SketchEntity>();

            // ! trovo i punti di interruzzione
            List<int> listIndexThk = new List<int>();
            int c = 1;
            foreach (SketchEntity entita in cleanedSketch)
            {
                if (entita.Type == ObjectTypeEnum.kSketchLineObject)
                {
                    SketchLine sl = (SketchLine) entita;

                    double tolleranza = 0.01;

                    //Console.WriteLine(Thickness - tolleranza);
                    //Console.WriteLine(sl.Length);
                    //Console.WriteLine(Thickness + tolleranza);
                    //Console.WriteLine("--------------------");

                    if (sl.Length <= Thickness + tolleranza && sl.Length >= Thickness - tolleranza)
                    {
                        listIndexThk.Add(c);
                    }
                }

                tmpListSE.Add(entita);
                c++;
            }

            SketchLine linea1 = (SketchLine)cleanedSketch[listIndexThk[0]];
            SketchLine linea2 = (SketchLine)cleanedSketch[listIndexThk[1]];

            linea1.OverrideColor = iApp.TransientObjects.CreateColor(255, 0, 0);
            linea2.OverrideColor = iApp.TransientObjects.CreateColor(255, 0, 0);

            c = 1;
            
            List<SketchEntity> tmpQueue = new List<SketchEntity>();

            foreach (SketchEntity entita in cleanedSketch)
            {
                if (c != listIndexThk[0] && c != listIndexThk[1])
                {
                    if (c < listIndexThk[0])
                    {
                        tmpQueue.Add(entita);
                    }
                    else if ( c > listIndexThk[1])
                    {
                        results["prima"].Add(entita);
                    }
                    else
                    {
                        results["seconda"].Add(entita);
                    }
                }
                c++;
            }
            foreach(SketchEntity entita in tmpQueue)
            {
                results["prima"].Add(entita);
            }

            double countPol1 = InventorClass.perimetroLines(results["prima"]);
            double countPol2 = InventorClass.perimetroLines(results["seconda"]);

            IDictionary<string, List<SketchEntity>> dati = new Dictionary<string, List<SketchEntity>>();
            if (countPol1 > countPol2)
            {
                dati.Add("lunga", results["prima"]);
                dati.Add("corta", results["seconda"]);
            }
            else
            {
                dati.Add("lunga", results["seconda"]);
                dati.Add("corta", results["prima"]);
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
            int segmentiTot;

            if (oCompDef.Bends.Count != 0)
            {
                segmentiTot = (oCompDef.Bends.Count * 4) + 4;
            }
            else
            {
                segmentiTot = 0;
            }

            return segmentiTot;
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
        public static WorkPlane getLongestEdgeFace(Face oFace, SheetMetalComponentDefinition oCompDef, PartDocument oDoc)
        {
            FaceCollection oFaceColl = iApp.TransientObjects.CreateFaceCollection();

            foreach (Face f in oCompDef.SurfaceBodies[1].Faces)
            {
                if (f.TangentiallyConnectedFaces.Count == 0 && f.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                {
                    WorkPlane tmpWp = oCompDef.WorkPlanes.AddByPlaneAndOffset(f, 0);
                    if (tmpWp.Plane.IsParallelTo[oCompDef.WorkPlanes[1].Plane])
                    {
                        oDoc.SelectSet.Select(f);
                        oFaceColl.Add(f);
                    }
                    tmpWp.Delete();
                }
            }

            // ? cerco le due facce più lontane

            double tmpDist = 0;

            FaceCollection ofc = iApp.TransientObjects.CreateFaceCollection();

            if (oFaceColl.Count >= 2)
            {
                foreach (Face f in oFaceColl)
                {
                    foreach (Face f1 in oFaceColl)
                    {
                        double dist = iApp.MeasureTools.GetMinimumDistance(f.PointOnFace, f1.PointOnFace);

                        if (tmpDist <= dist)
                        {
                            tmpDist = dist;
                            ofc = iApp.TransientObjects.CreateFaceCollection();
                            ofc.Add(f1);
                            ofc.Add(f);
                        }
                    }
                }

                WorkPlane wp = oCompDef.WorkPlanes.AddByTwoPlanes(ofc[1], ofc[2]);
                return (WorkPlane)wp;
            }
            else
            {
                Edge edge = null;

                oFaceColl = iApp.TransientObjects.CreateFaceCollection();

                foreach (Face f in oCompDef.SurfaceBodies[1].Faces)
                {
                    if (f.TangentiallyConnectedFaces.Count == 0 && f.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                    {
                        oDoc.SelectSet.Select(f);
                        oFaceColl.Add(f);
                    }
                }

                double highest = 0;

                foreach (Face tmpFace in oFaceColl)
                {
                    foreach (Edge tmpEdge in oFace.Edges)
                    {
                        if (tmpEdge.GeometryType == CurveTypeEnum.kLineSegmentCurve)
                        {
                            double tmpLength = iApp.MeasureTools.GetMinimumDistance(tmpEdge.StartVertex, tmpEdge.StopVertex);

                            if (tmpLength > highest)
                            {
                                highest = tmpLength;
                                edge = tmpEdge;
                            }
                        }
                    }
                }

                WorkPoint oWpoint = oCompDef.WorkPoints.AddByMidPoint(edge);
                oWpoint.Name = "wPntReference";
                oWpoint.Visible = false;

                WorkPlane oWpReference = oCompDef.WorkPlanes.AddByNormalToCurve(edge, oWpoint);
                oWpReference.Name = "wpReference";
                oWpReference.Visible = true;

                return oWpReference;
            }
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
        public static void ricostruiscoLamiera(PartDocument oDoc)
        {
            SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDoc.ComponentDefinition;

            oCompDef.SurfaceBodies[1].Visible = true;

            FaceCollection fFaces = oCompDef.Bends[1].FrontFaces[1].TangentiallyConnectedFaces;
            fFaces.Add(oCompDef.Bends[1].FrontFaces[1]);

            FaceCollection bFaces = oCompDef.Bends[1].BackFaces[1].TangentiallyConnectedFaces;
            bFaces.Add(oCompDef.Bends[1].BackFaces[1]);

            HighlightSet oSet = oDoc.CreateHighlightSet();

            foreach (Face f in bFaces)
            {
                oSet.AddItem(f);
            }

            PartFeatures oFeat = oCompDef.Features;

            oFeat.DeleteFaceFeatures.Add(fFaces);

            oFeat.ThickenFeatures.Add(bFaces, 2, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kNewBodyOperation, false);

            oSet.Color = iApp.TransientObjects.CreateColor(255, 0, 0);
        }
    }
}