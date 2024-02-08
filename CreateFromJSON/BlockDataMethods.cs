using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace GMEPUtilities
{
    public class BlockDataMethods
    {
        public static void CreateObjectGivenData(
            List<Dictionary<string, Dictionary<string, object>>> data,
            Editor ed,
            Point3d selectedPoint
        )
        {
            foreach (var objData in data)
            {
                var objectType = objData.Keys.First();

                switch (objectType)
                {
                    case "polyline":
                        selectedPoint = CreatePolyline(ed, selectedPoint, objData);
                        break;

                    case "line":
                        selectedPoint = CreateLine(ed, selectedPoint, objData);
                        break;

                    case "mtext":
                        selectedPoint = CreateMText(ed, selectedPoint, objData);
                        break;

                    case "circle":
                        selectedPoint = CreateCircle(ed, selectedPoint, objData);
                        break;

                    case "solid":
                        selectedPoint = CreateSolid(ed, selectedPoint, objData);
                        break;

                    default:
                        break;
                }
            }
        }

        private static void SetMTextStyleByName(MText mtext, string styleName)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                TextStyleTable textStyleTable =
                    tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                if (textStyleTable.Has(styleName))
                {
                    TextStyleTableRecord textStyle =
                        tr.GetObject(textStyleTable[styleName], OpenMode.ForRead)
                        as TextStyleTableRecord;
                    mtext.TextStyleId = textStyle.ObjectId;
                }
                tr.Commit();
            }
        }

        private static Point3d CreateCircle(
            Editor ed,
            Point3d selectedPoint,
            Dictionary<string, Dictionary<string, object>> objData
        )
        {
            var circleData = objData["circle"] as Dictionary<string, object>;
            var circle = new Circle();

            if (!LayerExists(circleData["layer"].ToString()))
            {
                CreateLayer(circleData["layer"].ToString(), 4);
            }

            circle.Layer = circleData["layer"].ToString();

            var centerData = JsonConvert.DeserializeObject<Dictionary<string, double>>(
                circleData["center"].ToString()
            );

            var centerX = Convert.ToDouble(centerData["x"]) + selectedPoint.X;
            var centerY = Convert.ToDouble(centerData["y"]) + selectedPoint.Y;
            var centerZ = Convert.ToDouble(centerData["z"]) + selectedPoint.Z;
            circle.Center = new Point3d(centerX, centerY, centerZ);

            circle.Radius = Convert.ToDouble(circleData["radius"]);

            // Add circle to the drawing
            using (var transaction = ed.Document.Database.TransactionManager.StartTransaction())
            {
                var blockTable =
                    transaction.GetObject(ed.Document.Database.BlockTableId, OpenMode.ForRead)
                    as BlockTable;
                var blockTableRecord =
                    transaction.GetObject(
                        blockTable[BlockTableRecord.PaperSpace],
                        OpenMode.ForWrite
                    ) as BlockTableRecord;
                blockTableRecord.AppendEntity(circle);
                transaction.AddNewlyCreatedDBObject(circle, true);
                transaction.Commit();
            }

            return selectedPoint;
        }

        private static Point3d CreateMText(
            Editor ed,
            Point3d selectedPoint,
            Dictionary<string, Dictionary<string, object>> objData
        )
        {
            var mtextData = objData["mtext"] as Dictionary<string, object>;
            var mtext = new MText();

            if (!LayerExists(mtextData["layer"].ToString()))
            {
                CreateLayer(mtextData["layer"].ToString(), 2);
            }

            mtext.Layer = mtextData["layer"].ToString();
            SetMTextStyleByName(mtext, mtextData["style"].ToString());
            mtext.Attachment = (AttachmentPoint)
                Enum.Parse(typeof(AttachmentPoint), mtextData["justification"].ToString());
            mtext.Contents = mtextData["text"].ToString();
            mtext.TextHeight = Convert.ToDouble(mtextData["height"]);

            if (mtext.Contents.Contains("STRING"))
            {
                mtext.TextHeight = 0.185;
            }
            else if (mtext.Contents.Contains("SOLAR"))
            {
                mtext.TextHeight = 0.1;
            }
            else if (mtext.Contents.Contains("MPPT"))
            {
                mtext.TextHeight = 0.075;
            }

            mtext.LineSpaceDistance = Convert.ToDouble(mtextData["lineSpaceDistance"]);

            var locationData = JsonConvert.DeserializeObject<Dictionary<string, double>>(
                mtextData["location"].ToString()
            );

            var locX = Convert.ToDouble(locationData["x"]) + selectedPoint.X;
            var locY = Convert.ToDouble(locationData["y"]) + selectedPoint.Y;
            var locZ = Convert.ToDouble(locationData["z"]) + selectedPoint.Z;
            mtext.Location = new Point3d(locX, locY, locZ);

            // Add mtext to the drawing
            using (var transaction = ed.Document.Database.TransactionManager.StartTransaction())
            {
                var blockTable =
                    transaction.GetObject(ed.Document.Database.BlockTableId, OpenMode.ForRead)
                    as BlockTable;
                var blockTableRecord =
                    transaction.GetObject(
                        blockTable[BlockTableRecord.PaperSpace],
                        OpenMode.ForWrite
                    ) as BlockTableRecord;
                blockTableRecord.AppendEntity(mtext);
                transaction.AddNewlyCreatedDBObject(mtext, true);
                transaction.Commit();
            }

            return selectedPoint;
        }

        private static void CreateLayer(string layerName, short colorIndex)
        {
            if (layerName.Contains("SYM"))
            {
                colorIndex = 6;
            }

            if (layerName.Contains("Inverter"))
            {
                colorIndex = 2;
            }

            using (
                var transaction =
                    HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            )
            {
                var layerTable =
                    transaction.GetObject(
                        HostApplicationServices.WorkingDatabase.LayerTableId,
                        OpenMode.ForRead
                    ) as LayerTable;

                if (!layerTable.Has(layerName))
                {
                    var layerTableRecord = new LayerTableRecord();
                    layerTableRecord.Name = layerName;
                    layerTableRecord.Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
                    layerTableRecord.LineWeight = LineWeight.LineWeight050;
                    layerTableRecord.IsPlottable = true;

                    layerTable.UpgradeOpen();
                    layerTable.Add(layerTableRecord);
                    transaction.AddNewlyCreatedDBObject(layerTableRecord, true);
                    transaction.Commit();
                }
            }
        }

        private static bool LayerExists(string layerName)
        {
            using (
                var transaction =
                    HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            )
            {
                var layerTable =
                    transaction.GetObject(
                        HostApplicationServices.WorkingDatabase.LayerTableId,
                        OpenMode.ForRead
                    ) as LayerTable;

                return layerTable.Has(layerName);
            }
        }

        private static Point3d CreateLine(
            Editor ed,
            Point3d selectedPoint,
            Dictionary<string, Dictionary<string, object>> objData
        )
        {
            var lineData = objData["line"] as Dictionary<string, object>;
            var line = new Line();

            if (!LayerExists(lineData["layer"].ToString()))
            {
                CreateLayer(lineData["layer"].ToString(), 4);
            }

            line.Layer = lineData["layer"].ToString();

            if (lineData.ContainsKey("linetype"))
            {
                if (!LinetypeExists(lineData["linetype"].ToString()))
                {
                    CreateLinetype(lineData["linetype"].ToString());
                }

                line.Linetype = lineData["linetype"].ToString();
            }

            var startPointData = JsonConvert.DeserializeObject<Dictionary<string, double>>(
                lineData["startPoint"].ToString()
            );

            var startPtX = Convert.ToDouble(startPointData["x"]) + selectedPoint.X;
            var startPtY = Convert.ToDouble(startPointData["y"]) + selectedPoint.Y;
            var startPtZ = Convert.ToDouble(startPointData["z"]) + selectedPoint.Z;
            line.StartPoint = new Point3d(startPtX, startPtY, startPtZ);

            var endPointData = JsonConvert.DeserializeObject<Dictionary<string, double>>(
                lineData["endPoint"].ToString()
            );

            var endPtX = Convert.ToDouble(endPointData["x"]) + selectedPoint.X;
            var endPtY = Convert.ToDouble(endPointData["y"]) + selectedPoint.Y;
            var endPtZ = Convert.ToDouble(endPointData["z"]) + selectedPoint.Z;
            line.EndPoint = new Point3d(endPtX, endPtY, endPtZ);

            // Add line to the drawing
            using (var transaction = ed.Document.Database.TransactionManager.StartTransaction())
            {
                var blockTable =
                    transaction.GetObject(ed.Document.Database.BlockTableId, OpenMode.ForRead)
                    as BlockTable;
                var blockTableRecord =
                    transaction.GetObject(
                        blockTable[BlockTableRecord.PaperSpace],
                        OpenMode.ForWrite
                    ) as BlockTableRecord;
                blockTableRecord.AppendEntity(line);
                transaction.AddNewlyCreatedDBObject(line, true);
                transaction.Commit();
            }

            return selectedPoint;
        }

        private static void CreateLinetype(string linetypeName)
        {
            Document acDoc = Autodesk
                .AutoCAD
                .ApplicationServices
                .Application
                .DocumentManager
                .MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            using (
                var transaction =
                    HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            )
            {
                var linetypeTable =
                    transaction.GetObject(
                        HostApplicationServices.WorkingDatabase.LinetypeTableId,
                        OpenMode.ForRead
                    ) as LinetypeTable;

                if (!linetypeTable.Has(linetypeName))
                {
                    acCurDb.LoadLineTypeFile(linetypeName, "acad.lin");
                    transaction.Commit();
                }
            }
        }

        private static bool LinetypeExists(string linetypeName)
        {
            using (
                var transaction =
                    HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            )
            {
                var linetypeTable =
                    transaction.GetObject(
                        HostApplicationServices.WorkingDatabase.LinetypeTableId,
                        OpenMode.ForRead
                    ) as LinetypeTable;

                return linetypeTable.Has(linetypeName);
            }
        }

        private static Point3d CreatePolyline(
            Editor ed,
            Point3d selectedPoint,
            Dictionary<string, Dictionary<string, object>> objData
        )
        {
            var polylineData = objData["polyline"] as Dictionary<string, object>;
            var polyline = new Polyline();

            if (!LayerExists(polylineData["layer"].ToString()))
            {
                CreateLayer(polylineData["layer"].ToString(), 4);
            }

            polyline.Layer = polylineData["layer"].ToString();

            if (polylineData.ContainsKey("linetype"))
            {
                if (!LinetypeExists(polylineData["linetype"].ToString()))
                {
                    CreateLinetype(polylineData["linetype"].ToString());
                }

                polyline.Linetype = polylineData["linetype"].ToString();
            }

            var vertices = JArray
                .Parse(polylineData["vertices"].ToString())
                .ToObject<List<Dictionary<string, double>>>();

            foreach (var vertex in vertices)
            {
                var x = vertex["x"] + selectedPoint.X;
                var y = vertex["y"] + selectedPoint.Y;
                var z = vertex["z"] + selectedPoint.Z;
                polyline.AddVertexAt(polyline.NumberOfVertices, new Point2d(x, y), z, 0, 0);
            }

            var closed = Convert.ToBoolean(polylineData["isClosed"]);

            polyline.Closed = closed;

            // Add polyline to the drawing
            using (var transaction = ed.Document.Database.TransactionManager.StartTransaction())
            {
                var blockTable =
                    transaction.GetObject(ed.Document.Database.BlockTableId, OpenMode.ForRead)
                    as BlockTable;
                var blockTableRecord =
                    transaction.GetObject(
                        blockTable[BlockTableRecord.PaperSpace],
                        OpenMode.ForWrite
                    ) as BlockTableRecord;
                blockTableRecord.AppendEntity(polyline);
                transaction.AddNewlyCreatedDBObject(polyline, true);
                transaction.Commit();
            }

            return selectedPoint;
        }

        private static Point3d CreateSolid(
            Editor ed,
            Point3d selectedPoint,
            Dictionary<string, Dictionary<string, object>> objData
        )
        {
            var solidData = objData["solid"];
            var solid = new Solid();
            short i = 0;

            if (!LayerExists(solidData["layer"].ToString()))
            {
                CreateLayer(solidData["layer"].ToString(), 2);
            }

            solid.Layer = solidData["layer"].ToString();

            var points = JArray
                .Parse(solidData["vertices"].ToString())
                .ToObject<List<Dictionary<string, double>>>();

            foreach (var point in points)
            {
                var x = point["x"] + selectedPoint.X;
                var y = point["y"] + selectedPoint.Y;
                var z = point["z"] + selectedPoint.Z;
                var point3d = new Point3d(x, y, z);
                solid.SetPointAt(i, point3d);
                i++;
            }

            // Add solid to the drawing
            using (var transaction = ed.Document.Database.TransactionManager.StartTransaction())
            {
                var blockTable =
                    transaction.GetObject(ed.Document.Database.BlockTableId, OpenMode.ForRead)
                    as BlockTable;
                var blockTableRecord =
                    transaction.GetObject(
                        blockTable[BlockTableRecord.PaperSpace],
                        OpenMode.ForWrite
                    ) as BlockTableRecord;
                blockTableRecord.AppendEntity(solid);
                transaction.AddNewlyCreatedDBObject(solid, true);
                transaction.Commit();
            }

            return selectedPoint;
        }
    }
}
