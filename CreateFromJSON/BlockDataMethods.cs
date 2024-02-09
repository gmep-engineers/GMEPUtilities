using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
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

                    case "arc":
                        selectedPoint = CreateArc(ed, selectedPoint, objData);
                        break;

                    default:
                        break;
                }
            }
        }

        private static Point3d CreateArc(
            Editor ed,
            Point3d selectedPoint,
            Dictionary<string, Dictionary<string, object>> objData
        )
        {
            var arcData = objData["arc"];
            var arc = new Arc();

            if (!LayerExists(arcData["layer"].ToString()))
            {
                CreateLayer(arcData["layer"].ToString(), 4);
            }

            arc.Layer = arcData["layer"].ToString();

            var centerData = JsonConvert.DeserializeObject<Dictionary<string, double>>(
                arcData["center"].ToString()
            );

            var centerX = Convert.ToDouble(centerData["x"]) + selectedPoint.X;
            var centerY = Convert.ToDouble(centerData["y"]) + selectedPoint.Y;
            var centerZ = Convert.ToDouble(centerData["z"]) + selectedPoint.Z;
            arc.Center = new Point3d(centerX, centerY, centerZ);

            arc.Radius = Convert.ToDouble(arcData["radius"]);
            arc.StartAngle = Convert.ToDouble(arcData["startAngle"]);
            arc.EndAngle = Convert.ToDouble(arcData["endAngle"]);

            //var startPointData = JsonConvert.DeserializeObject<Dictionary<string, double>>(
            //    arcData["startPoint"].ToString()
            //);

            //var startPtX = Convert.ToDouble(startPointData["X"]) + selectedPoint.X;
            //var startPtY = Convert.ToDouble(startPointData["Y"]) + selectedPoint.Y;
            //var startPtZ = Convert.ToDouble(startPointData["Z"]) + selectedPoint.Z;
            //arc.StartPoint = new Point3d(startPtX, startPtY, startPtZ);

            //var endPointData = JsonConvert.DeserializeObject<Dictionary<string, double>>(
            //    arcData["endPoint"].ToString()
            //);

            //var endPtX = Convert.ToDouble(endPointData["X"]) + selectedPoint.X;
            //var endPtY = Convert.ToDouble(endPointData["Y"]) + selectedPoint.Y;
            //var endPtZ = Convert.ToDouble(endPointData["Z"]) + selectedPoint.Z;
            //arc.EndPoint = new Point3d(endPtX, endPtY, endPtZ);

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
                blockTableRecord.AppendEntity(arc);
                transaction.AddNewlyCreatedDBObject(arc, true);
                transaction.Commit();
            }

            return selectedPoint;
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

        public static List<Dictionary<string, Dictionary<string, object>>> GetData(string path)
        {
            var dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var jsonPath = Path.Combine(dllPath, path);
            var json = File.ReadAllText(jsonPath);

            return JArray
                .Parse(json)
                .ToObject<List<Dictionary<string, Dictionary<string, object>>>>();
        }

        public static string GetUnparsedJSONData(string path)
        {
            var dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var jsonPath = Path.Combine(dllPath, path);
            var json = File.ReadAllText(jsonPath);

            return json;
        }

        public static void CreateFilledCircleInPaperSpace(Point3d center, double radius)
        {
            Document acDoc = Autodesk
                .AutoCAD
                .ApplicationServices
                .Application
                .DocumentManager
                .MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl;
                BlockTableRecord acBlkTblRec;

                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                acBlkTblRec =
                    acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite)
                    as BlockTableRecord;

                // Create a circle
                using (Circle acCircle = new Circle())
                {
                    acCircle.Center = center;
                    acCircle.Radius = radius;

                    if (!LayerExists("E-CONDUIT"))
                    {
                        CreateLayer("E-CONDUIT", 4);
                    }

                    acCircle.Layer = "E-CONDUIT";

                    acCircle.SetDatabaseDefaults();
                    acBlkTblRec.AppendEntity(acCircle);
                    acTrans.AddNewlyCreatedDBObject(acCircle, true);

                    using (Hatch acHatch = new Hatch())
                    {
                        acBlkTblRec.AppendEntity(acHatch);
                        acTrans.AddNewlyCreatedDBObject(acHatch, true);
                        acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                        acHatch.Associative = true;
                        acHatch.Layer = "E-CONDUIT";
                        acHatch.AppendLoop(
                            HatchLoopTypes.Outermost,
                            new ObjectIdCollection() { acCircle.ObjectId }
                        );
                        acHatch.EvaluateHatch(true);
                    }
                }

                acTrans.Commit();
            }
        }
    }

    public class HelperMethods
    {
        public static void GetUserToClick(out Editor ed, out PromptPointResult pointResult)
        {
            ed = Autodesk
                .AutoCAD
                .ApplicationServices
                .Application
                .DocumentManager
                .MdiActiveDocument
                .Editor;
            PromptPointOptions pointOptions = new PromptPointOptions("Select a point: ");
            pointResult = ed.GetPoint(pointOptions);
        }

        public static void SaveDataToJsonFile(object data, string fileName)
        {
            var filePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                fileName
            );
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(
                data,
                Newtonsoft.Json.Formatting.Indented
            );
            File.WriteAllText(filePath, json);
        }

        public static bool IsInModel()
        {
            if (
                Autodesk
                    .AutoCAD
                    .ApplicationServices
                    .Core
                    .Application
                    .DocumentManager
                    .MdiActiveDocument
                    .Database
                    .TileMode
            )
                return true;
            else
                return false;
        }

        public static bool IsInLayout()
        {
            return !IsInModel();
        }

        public static bool IsInLayoutPaper()
        {
            Autodesk.AutoCAD.ApplicationServices.Document doc = Autodesk
                .AutoCAD
                .ApplicationServices
                .Core
                .Application
                .DocumentManager
                .MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            if (db.TileMode)
                return false;
            else
            {
                if (db.PaperSpaceVportId == ObjectId.Null)
                    return false;
                else if (ed.CurrentViewportObjectId == ObjectId.Null)
                    return false;
                else if (ed.CurrentViewportObjectId == db.PaperSpaceVportId)
                    return true;
                else
                    return false;
            }
        }
    }

    public class RetrieveObjectData
    {
        public static List<Dictionary<string, object>> HandleSolid(
            Solid solid,
            List<Dictionary<string, object>> data,
            Point3d origin
        )
        {
            var solidData = new Dictionary<string, object> { { "layer", solid.Layer } };

            var vertices = new List<object>();
            for (short i = 0; i < 4; i++)
            {
                var vertex = new Dictionary<string, object>
                {
                    { "x", solid.GetPointAt(i).X - origin.X },
                    { "y", solid.GetPointAt(i).Y - origin.Y },
                    { "z", solid.GetPointAt(i).Z - origin.Z }
                };
                vertices.Add(vertex);
            }

            solidData.Add("vertices", vertices);

            var encapsulate = new Dictionary<string, object> { { "solid", solidData } };

            data.Add(encapsulate);

            return data;
        }

        public static List<Dictionary<string, object>> HandleArc(
            Arc arc,
            List<Dictionary<string, object>> data,
            Point3d origin
        )
        {
            var arcData = new Dictionary<string, object> { { "layer", arc.Layer } };

            var center = new Dictionary<string, object>
            {
                { "x", arc.Center.X - origin.X },
                { "y", arc.Center.Y - origin.Y },
                { "z", arc.Center.Z - origin.Z }
            };

            arcData.Add("center", center);

            arcData.Add("radius", arc.Radius);
            arcData.Add("startAngle", arc.StartAngle);
            arcData.Add("endAngle", arc.EndAngle);

            arcData.Add("startPoint", arc.StartPoint);
            arcData.Add("endPoint", arc.EndPoint);

            var encapsulate = new Dictionary<string, object> { { "arc", arcData } };

            data.Add(encapsulate);

            return data;
        }

        public static List<Dictionary<string, object>> HandlePolyline(
            Autodesk.AutoCAD.DatabaseServices.Polyline polyline,
            List<Dictionary<string, object>> data,
            Point3d origin
        )
        {
            var polylineData = new Dictionary<string, object>();
            polylineData.Add("layer", polyline.Layer);

            var vertices = new List<object>();
            for (int i = 0; i < polyline.NumberOfVertices; i++)
            {
                var vertex = new Dictionary<string, object>
                {
                    { "x", polyline.GetPoint3dAt(i).X - origin.X },
                    { "y", polyline.GetPoint3dAt(i).Y - origin.Y },
                    { "z", polyline.GetPoint3dAt(i).Z - origin.Z }
                };
                vertices.Add(vertex);
            }

            polylineData.Add("vertices", vertices);

            polylineData.Add("linetype", polyline.Linetype);

            if (polyline.Closed)
            {
                polylineData.Add("isClosed", true);
            }
            else
            {
                polylineData.Add("isClosed", false);
            }

            var encapsulate = new Dictionary<string, object>();
            encapsulate.Add("polyline", polylineData);

            data.Add(encapsulate);

            return data;
        }

        public static List<Dictionary<string, object>> HandleLine(
            Autodesk.AutoCAD.DatabaseServices.Line line,
            List<Dictionary<string, object>> data,
            Point3d origin
        )
        {
            var lineData = new Dictionary<string, object>();
            lineData.Add("layer", line.Layer);

            var startPoint = new Dictionary<string, object>
            {
                { "x", line.StartPoint.X - origin.X },
                { "y", line.StartPoint.Y - origin.Y },
                { "z", line.StartPoint.Z - origin.Z }
            };
            lineData.Add("startPoint", startPoint);

            var endPoint = new Dictionary<string, object>
            {
                { "x", line.EndPoint.X - origin.X },
                { "y", line.EndPoint.Y - origin.Y },
                { "z", line.EndPoint.Z - origin.Z }
            };
            lineData.Add("endPoint", endPoint);

            lineData.Add("linetype", line.Linetype);

            var encapsulate = new Dictionary<string, object>();
            encapsulate.Add("line", lineData);

            data.Add(encapsulate);

            return data;
        }

        public static List<Dictionary<string, object>> HandleMText(
            Autodesk.AutoCAD.DatabaseServices.MText mtext,
            List<Dictionary<string, object>> data,
            Point3d origin
        )
        {
            var mtextData = new Dictionary<string, object>();
            mtextData.Add("layer", mtext.Layer);
            mtextData.Add("style", mtext.TextStyleName);
            mtextData.Add("justification", mtext.Attachment.ToString());
            mtextData.Add("text", mtext.Contents);
            mtextData.Add("height", mtext.TextHeight);
            mtextData.Add("lineSpaceDistance", mtext.LineSpaceDistance);

            var location = new Dictionary<string, object>
            {
                { "x", mtext.Location.X - origin.X },
                { "y", mtext.Location.Y - origin.Y },
                { "z", mtext.Location.Z - origin.Z }
            };
            mtextData.Add("location", location);

            var encapsulate = new Dictionary<string, object>();
            encapsulate.Add("mtext", mtextData);

            data.Add(encapsulate);

            return data;
        }

        public static List<Dictionary<string, object>> HandleCircle(
            Autodesk.AutoCAD.DatabaseServices.Circle circle,
            List<Dictionary<string, object>> data,
            Point3d origin
        )
        {
            var circleData = new Dictionary<string, object>();
            circleData.Add("layer", circle.Layer);

            var center = new Dictionary<string, object>
            {
                { "x", circle.Center.X - origin.X },
                { "y", circle.Center.Y - origin.Y },
                { "z", circle.Center.Z - origin.Z }
            };
            circleData.Add("center", center);

            circleData.Add("radius", circle.Radius);

            var encapsulate = new Dictionary<string, object>();
            encapsulate.Add("circle", circleData);

            data.Add(encapsulate);

            return data;
        }
    }
}
