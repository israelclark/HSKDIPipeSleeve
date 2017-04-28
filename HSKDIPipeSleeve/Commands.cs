using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;


namespace HSKDIProject
{
    public class Commands
    {
        [CommandMethod("Sleeve")]
        public void Sleeve()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            
                                    
            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                BlockTableRecord mspace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Record tranformation matrix between current ucs & wcs for restoration at end of function
                Matrix3d ucs = ed.CurrentUserCoordinateSystem;                   
                ed.WriteMessage("UCS:{0}\n", ucs);

                // Select the pipe
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect Polyine to sleeve.");
                peo.SetRejectMessage("You must select a single polyline.");
                peo.AddAllowedClass(typeof(Polyline), false);
                peo.AllowObjectOnLockedLayer = true;                         

                PromptEntityResult per = ed.GetEntity(peo);

                if (per.Status != PromptStatus.OK)
                    return;

                // Select the sleeve end points
                Polyline pipe = (Polyline)tr.GetObject(per.ObjectId, OpenMode.ForRead);
                
                
                PromptPointOptions ppo = new PromptPointOptions("\nSelect Sleeve start point.");
                ppo.UseBasePoint = false;
                ppo.UseDashedLine = false;                    
                PromptPointResult ppr = ed.GetPoint(ppo);

                if (ppr.Status != PromptStatus.OK)
                    return;

                Point3d startPt = HSKDICommon.Commands.ClosestPtOnSegment(ppr.Value, pipe, ucs);
                
                ppo.Message = "\nSelect Sleeve end point.";
                ppo.UseBasePoint = true;
                ppo.BasePoint = startPt;
                ppo.UseDashedLine = true;
                ppr = ed.GetPoint(ppo);

                if (ppr.Status != PromptStatus.OK)
                    return;

                Point3d endPt = HSKDICommon.Commands.ClosestPtOnSegment(ppr.Value, pipe, ucs);
                               
                Polyline sleeveMidline = new Polyline();
                double sleeveWidth = 0.02075 * HSKDICommon.Commands.getdimscale();
                double sleeveOffsetDist = 0.04 * HSKDICommon.Commands.getdimscale();
                int startSeg;
                int endSeg;

                GetIntersectSegments(pipe, ucs, startPt, endPt, out startSeg, out endSeg);

                //build sleeve centerline

                sleeveMidline = (Polyline)pipe.Clone();
                                
                for (int i = sleeveMidline.NumberOfVertices - 1; i >= 0; i--)
                {
                    if (i < startSeg)
                    {
                        sleeveMidline.RemoveVertexAt(i);
                    }                    
                    if (i > endSeg + 1) sleeveMidline.RemoveVertexAt(i);
                }      
                // now we have the relevent segments only. We need to trim the line at the start & end points while keeping the bulges correct
                Polyline untrimmedSleeveMidline = (Polyline)sleeveMidline.Clone();
                                
                sleeveMidline.SetPointAt(0, new Point2d(startPt.X, startPt.Y));
                sleeveMidline.SetPointAt(sleeveMidline.NumberOfVertices - 1, new Point2d(endPt.X, endPt.Y));                
                sleeveMidline.TransformBy(ucs);
                
                for (int i = 0; i <= untrimmedSleeveMidline.NumberOfVertices - 1; i++)
                {
                    double segmentBulge;
                    switch (untrimmedSleeveMidline.GetSegmentType(i))
                    {
                        case SegmentType.Arc:
                            CircularArc3d arc = untrimmedSleeveMidline.GetArcSegmentAt(i);
                            Point3d arcCenter = arc.Center;
                            double arcIncludedAngle = arcCenter.GetVectorTo(startPt).DotProduct(arcCenter.GetVectorTo(endPt));
                            segmentBulge = Math.Tan(arcIncludedAngle / 4);
                            break;
                        default:
                            segmentBulge = 0;
                            break;
                    }
                    sleeveMidline.SetBulgeAt(i, segmentBulge);
                }
                
                // offset sleeve curves

                DBObjectCollection sleeves = new DBObjectCollection();
                DBObjectCollection sleeves2 = new DBObjectCollection();

                if (sleeveMidline.NumberOfVertices > 1)
                {
                    sleeves = sleeveMidline.GetOffsetCurves(sleeveOffsetDist);
                    sleeves2 = sleeveMidline.GetOffsetCurves(-sleeveOffsetDist);
                }
                else ed.WriteMessage("\nSleeving Failed - Bad Geometry.");

                foreach (DBObject dbo in sleeves2)
                    sleeves.Add(dbo);

                HSKDICommon.Commands.LayerFree("HSKDI-Pipe-Sleeve");

                foreach (Entity sleeve in sleeves)
                {
                    // Add each offset object
                    sleeve.Layer = "HSKDI-Pipe-Sleeve";
                    Polyline sleevePL = (Polyline)sleeve;
                    sleevePL.ColorIndex = 256; // byLayer
                    //sleeve.TransformBy(ucs);
                                        
                    for (int i = 0; i < sleevePL.NumberOfVertices; i++)
                    {
                        sleevePL.SetStartWidthAt(i, sleeveWidth);
                        sleevePL.SetEndWidthAt(i, sleeveWidth);                          
                    }
                    
                    mspace.AppendEntity(sleeve);
                    tr.AddNewlyCreatedDBObject(sleeve, true);
                }

                // restore previous UCS
                ed.CurrentUserCoordinateSystem = ucs;
                /* FOR TESTING
                DBPoint startPtEnt = new DBPoint(startPt);
                startPtEnt.TransformBy(ucs);
                startPtEnt.ColorIndex = 1;
                DBPoint endPtEnt = new DBPoint(endPt);
                endPtEnt.TransformBy(ucs); 
                endPtEnt.ColorIndex = 3;
                mspace.AppendEntity(startPtEnt);
                mspace.AppendEntity(endPtEnt);
                tr.AddNewlyCreatedDBObject(startPtEnt, true);
                tr.AddNewlyCreatedDBObject(endPtEnt, true);
                */

                tr.Commit();
            }
        }

        private static void GetIntersectSegments(Polyline pipe, Matrix3d ucs, Point3d startPt, Point3d endPt, out int startSeg, out int endSeg)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;            
            Editor ed = doc.Editor;

            bool reverseCurve = false;
            startSeg = -1;
            endSeg = -1;                       

            try
            {
                do
                {
                    reverseCurve = false;
                    startSeg = -1;
                    endSeg = -1;
                    if (pipe.NumberOfVertices == 2)
                    {
                        startSeg = 0;
                        endSeg = 0;
                    }

                    for (int i = 0; i < pipe.NumberOfVertices - 1; i++)
                    {
                        if (pipe.GetSegmentType(i) == SegmentType.Arc)
                        {
                            CircularArc3d arcS = pipe.GetArcSegmentAt(i);
                            arcS.TransformBy(ucs.Inverse());
                            if (arcS.IsOn(startPt, new Tolerance(.1, .1)))
                                startSeg = i;
                            if (arcS.IsOn(endPt, new Tolerance(.1, .1)))
                                endSeg = i;
                        }
                        else if (pipe.GetSegmentType(i) == SegmentType.Line)
                        {
                            LineSegment3d lS = pipe.GetLineSegmentAt(i);
                            lS.TransformBy(ucs.Inverse());
                            if (lS.IsOn(startPt, new Tolerance(.1, .1)))
                                startSeg = i;
                            if (lS.IsOn(endPt, new Tolerance(.1, .1)))
                                endSeg = i;
                        }
                        else break;
                    }
                    if (startSeg < 0 || endSeg < 0) break; // somehow it didn't catch an endpoint

                    if (startSeg > endSeg)
                    {
                        reverseCurve = true;
                        pipe.UpgradeOpen();
                        pipe.ReverseCurve();
                        pipe.DowngradeOpen();
                    }
                } while (reverseCurve == true);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("\nException: {0}.", ex.Message);
            }
        }
    }
}

