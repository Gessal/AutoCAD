using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Xml;
using Autodesk.AutoCAD.Geometry;
using System;
using System.IO;

[assembly: CommandClass(typeof(AutoCAD_PointsReader.MyCommands))]

namespace AutoCAD_PointsReader
{
    public class MyCommands
    {
        /// <summary>
        /// Команда строит полилинию по координатам
        /// </summary>
        [CommandMethod("CreateProfile", CommandFlags.Modal)]
        public void MyCreateProfile()
        {
            Database acCurDb = MyPlugin.doc.Database;

            // Старт транзакции
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Открытие таблицы Блоков для чтения
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                // Открытие записи таблицы Блоков пространства Модели для записи
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                //загружаю XML
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load("profiles.xml");
                XmlElement xRoot = xDoc.DocumentElement;
                //профили

                PromptStringOptions pStrOpts = new PromptStringOptions("\nProfile name: ");
                pStrOpts.AllowSpaces = true;
                PromptResult profileName = MyPlugin.doc.Editor.GetString(pStrOpts);
                if (profileName.Status == PromptStatus.OK)
                {
                    string profStr = profileName.StringResult;
                        foreach (XmlNode profile in xRoot.ChildNodes)
                        {
                            if (profile.Attributes.GetNamedItem("name").InnerText.Equals(profStr))
                            {
                                DBObjectCollection outRegionColl = new DBObjectCollection();
                                DBObjectCollection inRegionColl = new DBObjectCollection();

                                /*Строю внешний контур*/
                                XmlNode outNode = profile.FirstChild;
                                XmlNodeList outLines = outNode.ChildNodes;
                                foreach (XmlNode line in outLines)
                                {
                                    if (line.Name.Equals("LWPolyLine"))
                                    {
                                        Polyline polyline = new Polyline();
                                        polyline.SetDatabaseDefaults();
                                        int n = 0;
                                        for (int i = 0; i < line.ChildNodes.Count; i = i + 4)
                                        {
                                            double[] xy = new double[2] { Double.Parse(line.ChildNodes.Item(i).InnerText)/* + ptStart.X*/,
                                                                          Double.Parse(line.ChildNodes.Item(i + 1).InnerText)/* + /*ptStart.Y*/ };
                                            polyline.AddVertexAt(n, new Point2d(xy), Double.Parse(line.ChildNodes.Item(i + 3).InnerText), 0, 0);
                                            n++;
                                        }
                                        polyline.Closed = true;
                                        // Добавление контуров в массив объектов
                                        DBObjectCollection acDBObjColl = new DBObjectCollection();
                                        acDBObjColl.Add(polyline);
                                        // Вычисление областей из каждого замкнутого контуров
                                        outRegionColl.Add(Region.CreateFromCurves(acDBObjColl)[0]);
                                        // Очистка объектов в памяти не добавленных в базу данных
                                        polyline.Dispose();
                                    }

                                    if (line.Name.Equals("Circle"))
                                    {
                                        Circle circle = new Circle(new Point3d(Double.Parse(line.ChildNodes.Item(0).InnerText)/* + ptStart.X*/,
                                                                               Double.Parse(line.ChildNodes.Item(1).InnerText)/* + ptStart.Y*/,
                                                                               Double.Parse(line.ChildNodes.Item(2).InnerText)/* + ptStart.Z*/),
                                                                   new Vector3d(Double.Parse(line.ChildNodes.Item(3).InnerText),
                                                                                Double.Parse(line.ChildNodes.Item(4).InnerText),
                                                                                Double.Parse(line.ChildNodes.Item(5).InnerText)),
                                                                   Double.Parse(line.ChildNodes.Item(6).InnerText));
                                        circle.SetDatabaseDefaults();
                                        // Добавление контуров в массив объектов
                                        DBObjectCollection acDBObjColl = new DBObjectCollection();
                                        acDBObjColl.Add(circle);
                                        // Вычисление областей из каждого замкнутого контуров
                                        outRegionColl.Add(Region.CreateFromCurves(acDBObjColl)[0]);
                                        // Очистка объектов в памяти не добавленных в базу данных
                                        circle.Dispose();
                                    }

                                    if (line.Name.Equals("Ellipse"))
                                    {
                                        Ellipse ellipse = new Ellipse(new Point3d(Double.Parse(line.ChildNodes.Item(0).InnerText)/* + ptStart.X*/,
                                                                               Double.Parse(line.ChildNodes.Item(1).InnerText)/* + ptStart.Y*/,
                                                                               Double.Parse(line.ChildNodes.Item(2).InnerText)/* + ptStart.Z*/),
                                                                   new Vector3d(Double.Parse(line.ChildNodes.Item(3).InnerText),
                                                                                Double.Parse(line.ChildNodes.Item(4).InnerText),
                                                                                Double.Parse(line.ChildNodes.Item(5).InnerText)),
                                                                   new Vector3d(Double.Parse(line.ChildNodes.Item(6).InnerText),
                                                                                Double.Parse(line.ChildNodes.Item(7).InnerText),
                                                                                Double.Parse(line.ChildNodes.Item(8).InnerText)),
                                                                   Double.Parse(line.ChildNodes.Item(9).InnerText),
                                                                   Double.Parse(line.ChildNodes.Item(10).InnerText),
                                                                   Double.Parse(line.ChildNodes.Item(11).InnerText));
                                        ellipse.SetDatabaseDefaults();
                                        // Добавление контуров в массив объектов
                                        DBObjectCollection acDBObjColl = new DBObjectCollection();
                                        acDBObjColl.Add(ellipse);
                                        // Вычисление областей из каждого замкнутого контуров
                                        outRegionColl.Add(Region.CreateFromCurves(acDBObjColl)[0]);
                                        // Очистка объектов в памяти не добавленных в базу данных
                                        ellipse.Dispose();
                                    }
                                }

                                /*Строю внутренний контур*/
                                XmlNode inNode = profile.LastChild;
                                XmlNodeList inLines = inNode.ChildNodes;
                                foreach (XmlNode line in inLines)
                                {
                                    if (line.Name.Equals("LWPolyLine"))
                                    {
                                        Polyline polyline = new Polyline();
                                        polyline.SetDatabaseDefaults();
                                        int n = 0;
                                        for (int i = 0; i < line.ChildNodes.Count; i = i + 4)
                                        {
                                            double[] xy = new double[2] { Double.Parse(line.ChildNodes.Item(i).InnerText)/* + ptStart.X*/,
                                                                          Double.Parse(line.ChildNodes.Item(i + 1).InnerText)/* + ptStart.Y*/ };
                                            polyline.AddVertexAt(n, new Point2d(xy), Double.Parse(line.ChildNodes.Item(i + 3).InnerText), 0, 0);
                                            n++;
                                        }
                                        polyline.Closed = true;
                                        // Добавление контуров в массив объектов
                                        DBObjectCollection acDBObjColl = new DBObjectCollection();
                                        acDBObjColl.Add(polyline);
                                        // Вычисление областей из каждого замкнутого контуров
                                        inRegionColl.Add(Region.CreateFromCurves(acDBObjColl)[0]);
                                        // Очистка объектов в памяти не добавленных в базу данных
                                        polyline.Dispose();
                                    }

                                    if (line.Name.Equals("Circle"))
                                    {
                                        Circle circle = new Circle(new Point3d(Double.Parse(line.ChildNodes.Item(0).InnerText)/* + ptStart.X*/,
                                                                               Double.Parse(line.ChildNodes.Item(1).InnerText)/* + ptStart.Y*/,
                                                                               Double.Parse(line.ChildNodes.Item(2).InnerText)/* + ptStart.Z*/),
                                                                   new Vector3d(Double.Parse(line.ChildNodes.Item(3).InnerText),
                                                                                Double.Parse(line.ChildNodes.Item(4).InnerText),
                                                                                Double.Parse(line.ChildNodes.Item(5).InnerText)),
                                                                   Double.Parse(line.ChildNodes.Item(6).InnerText));
                                        circle.SetDatabaseDefaults();
                                        // Добавление контуров в массив объектов
                                        DBObjectCollection acDBObjColl = new DBObjectCollection();
                                        acDBObjColl.Add(circle);
                                        // Вычисление областей из каждого замкнутого контуров
                                        inRegionColl.Add(Region.CreateFromCurves(acDBObjColl)[0]);
                                        // Очистка объектов в памяти не добавленных в базу данных
                                        circle.Dispose();
                                    }

                                    if (line.Name.Equals("Ellipse"))
                                    {
                                        Ellipse ellipse = new Ellipse(new Point3d(Double.Parse(line.ChildNodes.Item(0).InnerText)/* + ptStart.X*/,
                                                                               Double.Parse(line.ChildNodes.Item(1).InnerText)/* + ptStart.Y*/,
                                                                               Double.Parse(line.ChildNodes.Item(2).InnerText)/* + ptStart.Z*/),
                                                                   new Vector3d(Double.Parse(line.ChildNodes.Item(3).InnerText),
                                                                                Double.Parse(line.ChildNodes.Item(4).InnerText),
                                                                                Double.Parse(line.ChildNodes.Item(5).InnerText)),
                                                                   new Vector3d(Double.Parse(line.ChildNodes.Item(6).InnerText),
                                                                                Double.Parse(line.ChildNodes.Item(7).InnerText),
                                                                                Double.Parse(line.ChildNodes.Item(8).InnerText)),
                                                                   Double.Parse(line.ChildNodes.Item(9).InnerText),
                                                                   Double.Parse(line.ChildNodes.Item(10).InnerText),
                                                                   Double.Parse(line.ChildNodes.Item(11).InnerText));
                                        ellipse.SetDatabaseDefaults();
                                        // Добавление контуров в массив объектов
                                        DBObjectCollection acDBObjColl = new DBObjectCollection();
                                        acDBObjColl.Add(ellipse);
                                        // Вычисление областей из каждого замкнутого контуров
                                        inRegionColl.Add(Region.CreateFromCurves(acDBObjColl)[0]);
                                        // Очистка объектов в памяти не добавленных в базу данных
                                        ellipse.Dispose();
                                    }
                                }
                                /* Пока можно работать только с одной внешней областью */
                                // Вычитание меньшей области из большей
                                Region outRegion = outRegionColl[0] as Region;
                                foreach (DBObject inObject in inRegionColl)
                                {
                                    Region inRegion = inObject as Region;
                                    outRegion.BooleanOperation(BooleanOperationType.BoolSubtract, inRegion);
                                    inRegion.Dispose();
                                }
                                PromptEntityOptions peo2 = new PromptEntityOptions("\nВыберите траекторию выдавливания: ");
                                peo2.SetRejectMessage("\nНедопустимая траектория!");
                                peo2.AddAllowedClass(typeof(Curve), false);
                                PromptEntityResult per = MyPlugin.doc.Editor.GetEntity(peo2);
                                if (per.Status != PromptStatus.OK)
                                    return;
                                ObjectId splId = per.ObjectId;
                                Curve pathEnt = acTrans.GetObject(splId, OpenMode.ForRead) as Curve;

                                SweepOptionsBuilder sob = new SweepOptionsBuilder();
                                sob.Align = SweepOptionsAlignOption.AlignSweepEntityToPath;
                                sob.BasePoint = new Point3d(0.0, 0.0, 0.0);
                                sob.Bank = true;
                                Solid3d sol = new Solid3d();
                                sol.CreateSweptSolid(outRegion, pathEnt, sob.ToSweepOptions());
                                outRegion.Dispose();

                                acBlkTblRec.AppendEntity(sol);
                                acTrans.AddNewlyCreatedDBObject(sol, true);
                            }
                        }
                }
                acTrans.Commit();
            }
        }

        //---------------------------------------------------
        [CommandMethod("getExDict", CommandFlags.Modal)]
        public void getExDict()
        {
            
        }

        /*************************************************************************/
        [CommandMethod("GetPoints", CommandFlags.Modal | CommandFlags.UsePickSet)]
        public void GetPoints()
        {
            Utils.Utils.GetPoints();
        }

    }
}
