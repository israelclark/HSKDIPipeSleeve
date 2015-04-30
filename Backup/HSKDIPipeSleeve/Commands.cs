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
        private static Point3d ClosestPtOnSegment(Point3d pt, Polyline pl)
        {            
            Point3d pNearest = new Point3d();            

            for (int i = 0; i < pl.NumberOfVertices - 1; i++)
            {
                Line lnNormal = new Line();
                Point3d nearestPtOnSegment = new Point3d();
                Point3dCollection intersects = new Point3dCollection();                

                if (pl.GetSegmentType(i) == SegmentType.Line)
                {
                    LineSegment3d ls = pl.GetLineSegmentAt(i);
                    Line ln = new Line(ls.StartPoint, ls.EndPoint);

                    if (ln != null) // is a line
                    {
                        Vector3d vect = new Vector3d(ln.EndPoint.X - ln.StartPoint.X,
                                                     ln.EndPoint.Y - ln.StartPoint.Y,
                                                     ln.EndPoint.Z - ln.StartPoint.Z);
                        Vector3d uNormal = vect.GetPerpendicularVector();
                        lnNormal.StartPoint = pt.Add(uNormal * (Math.Max(pt.DistanceTo(ln.StartPoint), pt.DistanceTo(ln.EndPoint))));
                        
                        lnNormal.EndPoint = pt.Add(-uNormal*(Math.Max(pt.DistanceTo(ln.StartPoint), pt.DistanceTo(ln.EndPoint))));

                        lnNormal.IntersectWith(ln, Intersect.OnBothOperands, intersects, 0, 0);
                        if (intersects.Count == 1) nearestPtOnSegment = intersects[0];
                        else if (intersects.Count == 0) nearestPtOnSegment = pt.DistanceTo(ln.StartPoint) < pt.DistanceTo(ln.EndPoint) ? ln.StartPoint : ln.EndPoint;
                        else
                        {
                            // shouldn't be possible                        
                        }
                    }                
                }
                else if (pl.GetSegmentType(i) == SegmentType.Arc)
                {              
                    CircularArc3d arcs = pl.GetArcSegmentAt(i);
                    Arc arc = new Arc();
                    try
                    {
                        arc.Center = arcs.Center;
                        arc.Radius = arcs.Radius;
                        arc.StartAngle = arcs.StartAngle;
                        arc.EndAngle = arcs.EndAngle;
                    }
                    catch
                    {
                        // not an arc
                        arc = null;
                    }

                    

                    if (arc != null)
                    {
                        // normal to a curve will always pass through curve center
                        lnNormal.StartPoint = pt.Add(10*arcs.Center.GetVectorTo(pt));
                        lnNormal.EndPoint = arc.Center;


                        lnNormal.IntersectWith(arc, Intersect.OnBothOperands, intersects, 0, 0);
                        if (intersects.Count == 1) nearestPtOnSegment = intersects[0];
                        else if (intersects.Count == 0) nearestPtOnSegment = pt.DistanceTo(arc.StartPoint) < pt.DistanceTo(arc.EndPoint) ? arc.StartPoint : arc.EndPoint;
                        else
                        {
                            // we are ending in the arc center so should not be possible.
                        }
                    }
                }                
                                
                if (pt.DistanceTo(nearestPtOnSegment) < pt.DistanceTo(pNearest)) pNearest = nearestPtOnSegment;
            }
            return pNearest;
        }        

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

                // Select the valve
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect Polyine to sleeve.");
                peo.SetRejectMessage("You must select a single polyline");
                peo.AddAllowedClass(typeof(Polyline), false);
                peo.AllowObjectOnLockedLayer = true;

                PromptEntityResult per = ed.GetEntity(peo);

                if (per.Status != PromptStatus.OK)
                    return;

                Polyline pipe = (Polyline)tr.GetObject(per.ObjectId, OpenMode.ForRead);

                PromptPointOptions ppo = new PromptPointOptions("\nSelect Sleeve start point.");
                ppo.UseBasePoint = false;
                ppo.UseDashedLine = false;
                PromptPointResult ppr = ed.GetPoint(ppo);

                if (ppr.Status != PromptStatus.OK)
                    return;

                Point3d startPt = ClosestPtOnSegment(ppr.Value, pipe);

                ppo.Message = "\nSelect Sleeve end point.";
                ppo.UseBasePoint = true;
                ppo.BasePoint = startPt;
                ppo.UseDashedLine = true;
                ppr = ed.GetPoint(ppo);

                if (ppr.Status != PromptStatus.OK)
                    return;

                Point3d endPt = ClosestPtOnSegment(ppr.Value, pipe);
                

                Polyline sleeveMidline = new Polyline();
                double sleeveWidth = 0.02075 * HSKDICommon.Commands.getdimscale();
                double sleeveOffsetDist = 0.04 * HSKDICommon.Commands.getdimscale();
                int startSeg;
                int endSeg;

                GetIntersectSegments(pipe, startPt, endPt, out startSeg, out endSeg);

                //build sleeve centerline

                sleeveMidline = (Polyline)pipe.Clone();
                                                
                int j = sleeveMidline.NumberOfVertices - 1;
                for (int i = j; i >= 0; i--)
                {
                    if (i < startSeg)
                    {
                        sleeveMidline.RemoveVertexAt(i);
                    }                    
                    if (i > endSeg + 1) sleeveMidline.RemoveVertexAt(i);
                }      
                // now we have the relevent segments only. We need to trim the line at the start & end points
                sleeveMidline.SetPointAt(0, new Point2d(startPt.X, startPt.Y));
                sleeveMidline.SetPointAt(sleeveMidline.NumberOfVertices - 1, new Point2d(endPt.X, endPt.Y));


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
                                        
                    for (int i = 0; i < sleevePL.NumberOfVertices; i++)
                    {
                        sleevePL.SetStartWidthAt(i, sleeveWidth);
                        sleevePL.SetEndWidthAt(i, sleeveWidth);                          
                    }
                    
                    mspace.AppendEntity(sleeve);
                    tr.AddNewlyCreatedDBObject(sleeve, true);
                }                
                tr.Commit();
            }
        }

        private static void GetIntersectSegments(Polyline pipe, Point3d startPt, Point3d endPt, out int startSeg, out int endSeg)
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
                            CircularArc3d arcs = pipe.GetArcSegmentAt(i);
                            if (arcs.IsOn(startPt, new Tolerance(.1, .1)))
                                startSeg = i;
                            if (arcs.IsOn(endPt, new Tolerance(.1, .1)))
                                endSeg = i;
                        }
                        else if (pipe.GetSegmentType(i) == SegmentType.Line)
                        {
                            LineSegment3d ls = pipe.GetLineSegmentAt(i);
                            if (ls.IsOn(startPt, new Tolerance(.1, .1)))
                                startSeg = i;
                            if (ls.IsOn(endPt, new Tolerance(.1, .1)))
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

