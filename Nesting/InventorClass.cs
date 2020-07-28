using Inventor;
using System;
using System.Collections.Generic;
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
                        saveDocument(@path+"//"+rett.filename+".ipt", oDoc);
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

            TransientGeometry oTransGeom = (TransientGeometry)iApp.TransientGeometry;

            oSketch.SketchLines.AddAsTwoPointRectangle(
                oTransGeom.CreatePoint2d(0, 0),
                oTransGeom.CreatePoint2d(h, b)
            );

            Profile oProfile = (Profile)oSketch.Profiles.AddForSolid();

            FaceFeatureDefinition oFaceFeatureDefinition = (FaceFeatureDefinition)oSheetMetalFeatures.FaceFeatures.CreateFaceFeatureDefinition(oProfile);

            FaceFeature oFaceFeature = default(FaceFeature);
            oFaceFeature = oSheetMetalFeatures.FaceFeatures.Add(oFaceFeatureDefinition);

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
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                System.Windows.Forms.MessageBox.Show("Assicurarsi di aver aperto inventor");
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
    }
}
