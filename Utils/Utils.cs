using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace AutoCAD_PointsReader.Utils
{
    class Utils
    {
        /// <summary>
        /// Команда выводит в консоль DXF имя выбранного элемента
        /// </summary>
        [CommandMethod("GetDXFName", CommandFlags.Modal)]
        public void GetDXFName()
        {
            Editor ed;
            if (MyPlugin.doc != null)
            {
                ed = MyPlugin.doc.Editor;
                PromptSelectionResult result = null;
                PromptSelectionOptions psoOptions = new PromptSelectionOptions();
                psoOptions.SingleOnly = false;
                psoOptions.SinglePickInSpace = true;
                result = ed.GetSelection(psoOptions);
                if (result.Status != PromptStatus.OK)
                    return;
                SelectionSet ss = result.Value;
                ObjectId[] ids = ss.GetObjectIds();
                for (int i = 0; i < ids.Length; i++)
                {
                    ed.WriteMessage("\n" + ids[i].ObjectClass.DxfName);
                }
            }

        }

        /// <summary>
        /// Команда выводит список точек построения объекта
        /// </summary>
        public static void GetPoints()
        {
            Editor ed;
            if (MyPlugin.doc != null)
            {
                ed = MyPlugin.doc.Editor;

                /*считываем внешние контуры*/
                PromptSelectionResult outResult = null;
                PromptSelectionOptions psoOptions = new PromptSelectionOptions();
                psoOptions.SingleOnly = false;
                psoOptions.SinglePickInSpace = true;
                psoOptions.MessageForAdding = "Выберите внешние контуры";

                outResult = ed.GetSelection(psoOptions);
                if (outResult.Status != PromptStatus.OK)
                    return;
                SelectionSet outSS = outResult.Value;
                ObjectId[] outIds = outSS.GetObjectIds();
                foreach (ObjectId oId in outIds)
                {
                    if (!oId.ObjectClass.DxfName.Equals("LWPOLYLINE") && !oId.ObjectClass.DxfName.Equals("CIRCLE") &&
                        !oId.ObjectClass.DxfName.Equals("ELLIPSE"))
                    {
                        ed.WriteMessage("\nВыбраны недопустимые примитивы.");
                        return;
                    }
                }
                /*считываем внутренние контуры*/
                PromptSelectionResult inResult = null;
                psoOptions.MessageForAdding = "Выберите внутренние контуры";

                inResult = ed.GetSelection(psoOptions);
                if (inResult.Status != PromptStatus.OK)
                    return;
                SelectionSet inSS = inResult.Value;
                ObjectId[] inIds = inSS.GetObjectIds();
                foreach (ObjectId oId in inIds)
                {
                    if (!oId.ObjectClass.DxfName.Equals("LWPOLYLINE") && !oId.ObjectClass.DxfName.Equals("CIRCLE") &&
                        !oId.ObjectClass.DxfName.Equals("ELLIPSE"))
                    {
                        ed.WriteMessage("\nВыбраны недопустимые примитивы.");
                        return;
                    }
                }

                XmlDocument xDoc = new XmlDocument();
                xDoc.Load("profiles.xml");
                XmlElement xRoot = xDoc.DocumentElement;
                XmlElement profileElem = xDoc.CreateElement("profile");
                XmlAttribute nameAttr = xDoc.CreateAttribute("name");
                XmlText nameText = xDoc.CreateTextNode(Path.GetFileNameWithoutExtension(MyPlugin.doc.Name));
                nameAttr.AppendChild(nameText);
                profileElem.Attributes.Append(nameAttr);
                xRoot.AppendChild(profileElem);

                /*Записываю внешние контуры*/
                XmlElement outElem = xDoc.CreateElement("Out");
                profileElem.AppendChild(outElem);
                for (int i = 0; i < outIds.Length; i++)
                {
                    if (outIds[i].ObjectClass.DxfName.Equals("LWPOLYLINE"))
                    {
                        using (Transaction transaction = MyPlugin.doc.Database.TransactionManager.StartTransaction())
                        {
                            Polyline polyline;
                            polyline = transaction.GetObject(outIds[i], OpenMode.ForRead) as Polyline;
                            XmlElement lineElem = xDoc.CreateElement("LWPolyLine");
                            outElem.AppendChild(lineElem);
                            for (int j = 0; j < polyline.NumberOfVertices; j++)
                            {
                                XmlElement XElem = xDoc.CreateElement("X" + j);
                                XmlElement YElem = xDoc.CreateElement("Y" + j);
                                XmlElement ZElem = xDoc.CreateElement("Z" + j);
                                XmlElement BulgeElem = xDoc.CreateElement("Bulge" + j);
                                XmlText XText = xDoc.CreateTextNode(polyline.GetPoint3dAt(j).X.ToString());
                                XmlText YText = xDoc.CreateTextNode(polyline.GetPoint3dAt(j).Y.ToString());
                                XmlText ZText = xDoc.CreateTextNode(polyline.GetPoint3dAt(j).Z.ToString());
                                XmlText BulgeText = xDoc.CreateTextNode(polyline.GetBulgeAt(j).ToString());

                                XElem.AppendChild(XText);
                                YElem.AppendChild(YText);
                                ZElem.AppendChild(ZText);
                                BulgeElem.AppendChild(BulgeText);
                                lineElem.AppendChild(XElem);
                                lineElem.AppendChild(YElem);
                                lineElem.AppendChild(ZElem);
                                lineElem.AppendChild(BulgeElem);
                            }
                        }
                    }
                    if (outIds[i].ObjectClass.DxfName.Equals("CIRCLE"))
                    {
                        using (Transaction transaction = MyPlugin.doc.Database.TransactionManager.StartTransaction())
                        {
                            //Center
                            Circle circle;
                            circle = transaction.GetObject(outIds[i], OpenMode.ForRead) as Circle;
                            XmlElement lineElem = xDoc.CreateElement("Circle");
                            outElem.AppendChild(lineElem);
                            XmlElement XElem = xDoc.CreateElement("X0");
                            XmlElement YElem = xDoc.CreateElement("Y0");
                            XmlElement ZElem = xDoc.CreateElement("Z0");
                            XmlText XText = xDoc.CreateTextNode(circle.Center.X.ToString());
                            XmlText YText = xDoc.CreateTextNode(circle.Center.Y.ToString());
                            XmlText ZText = xDoc.CreateTextNode(circle.Center.Z.ToString());
                            XElem.AppendChild(XText);
                            YElem.AppendChild(YText);
                            ZElem.AppendChild(ZText);
                            lineElem.AppendChild(XElem);
                            lineElem.AppendChild(YElem);
                            lineElem.AppendChild(ZElem);
                            //Normal
                            XmlElement normalXElem = xDoc.CreateElement("NormalX");
                            XmlElement normalYElem = xDoc.CreateElement("NormalY");
                            XmlElement normalZElem = xDoc.CreateElement("NormalZ");
                            XmlText normalXText = xDoc.CreateTextNode(circle.Normal.X.ToString());
                            XmlText normalYText = xDoc.CreateTextNode(circle.Normal.Y.ToString());
                            XmlText normalZText = xDoc.CreateTextNode(circle.Normal.Z.ToString());
                            normalXElem.AppendChild(normalXText);
                            normalYElem.AppendChild(normalYText);
                            normalZElem.AppendChild(normalZText);
                            lineElem.AppendChild(normalXElem);
                            lineElem.AppendChild(normalYElem);
                            lineElem.AppendChild(normalZElem);
                            //Radius
                            XmlElement RElem = xDoc.CreateElement("R");
                            XmlText RText = xDoc.CreateTextNode(circle.Radius.ToString());
                            RElem.AppendChild(RText);
                            lineElem.AppendChild(RElem);
                        }
                    }
                    if (outIds[i].ObjectClass.DxfName.Equals("ELLIPSE"))
                    {
                        using (Transaction transaction = MyPlugin.doc.Database.TransactionManager.StartTransaction())
                        {
                            Ellipse ellipse;
                            ellipse = transaction.GetObject(outIds[i], OpenMode.ForRead) as Ellipse;
                            XmlElement lineElem = xDoc.CreateElement("Ellipse");
                            outElem.AppendChild(lineElem);
                            //center
                            XmlElement XElem = xDoc.CreateElement("X0");
                            XmlElement YElem = xDoc.CreateElement("Y0");
                            XmlElement ZElem = xDoc.CreateElement("Z0");
                            XmlText XText = xDoc.CreateTextNode(ellipse.Center.X.ToString());
                            XmlText YText = xDoc.CreateTextNode(ellipse.Center.Y.ToString());
                            XmlText ZText = xDoc.CreateTextNode(ellipse.Center.Z.ToString());
                            XElem.AppendChild(XText);
                            YElem.AppendChild(YText);
                            ZElem.AppendChild(ZText);
                            lineElem.AppendChild(XElem);
                            lineElem.AppendChild(YElem);
                            lineElem.AppendChild(ZElem);
                            //normal
                            XmlElement normalXElem = xDoc.CreateElement("NormalX");
                            XmlElement normalYElem = xDoc.CreateElement("NormalY");
                            XmlElement normalZElem = xDoc.CreateElement("NormalZ");
                            XmlText normalXText = xDoc.CreateTextNode(ellipse.Normal.X.ToString());
                            XmlText normalYText = xDoc.CreateTextNode(ellipse.Normal.Y.ToString());
                            XmlText normalZText = xDoc.CreateTextNode(ellipse.Normal.Z.ToString());
                            normalXElem.AppendChild(normalXText);
                            normalYElem.AppendChild(normalYText);
                            normalZElem.AppendChild(normalZText);
                            lineElem.AppendChild(normalXElem);
                            lineElem.AppendChild(normalYElem);
                            lineElem.AppendChild(normalZElem);
                            //MajorAxis
                            XmlElement mAxisXElem = xDoc.CreateElement("mAxisX");
                            XmlElement mAxisYElem = xDoc.CreateElement("mAxisY");
                            XmlElement mAxisZElem = xDoc.CreateElement("mAxisZ");
                            XmlText mAxisXText = xDoc.CreateTextNode(ellipse.MajorAxis.X.ToString());
                            XmlText mAxisYText = xDoc.CreateTextNode(ellipse.MajorAxis.Y.ToString());
                            XmlText mAxisZText = xDoc.CreateTextNode(ellipse.MajorAxis.Z.ToString());
                            mAxisXElem.AppendChild(mAxisXText);
                            mAxisYElem.AppendChild(mAxisYText);
                            mAxisZElem.AppendChild(mAxisZText);
                            lineElem.AppendChild(mAxisXElem);
                            lineElem.AppendChild(mAxisYElem);
                            lineElem.AppendChild(mAxisZElem);
                            //RadiusRatio
                            XmlElement radRatioElem = xDoc.CreateElement("RadiusRatio");
                            XmlText radRatioText = xDoc.CreateTextNode(ellipse.RadiusRatio.ToString());
                            radRatioElem.AppendChild(radRatioText);
                            lineElem.AppendChild(radRatioElem);
                            //StartAngle
                            XmlElement startAngElem = xDoc.CreateElement("StartAngle");
                            XmlText startAngText = xDoc.CreateTextNode(ellipse.StartAngle.ToString());
                            startAngElem.AppendChild(startAngText);
                            lineElem.AppendChild(startAngElem);
                            //EndAngle
                            XmlElement endAngElem = xDoc.CreateElement("EndAngle");
                            XmlText endAngText = xDoc.CreateTextNode(ellipse.EndAngle.ToString());
                            endAngElem.AppendChild(endAngText);
                            lineElem.AppendChild(endAngElem);
                        }
                    }
                }

                /*Записываю внутренние контуры*/
                XmlElement inElem = xDoc.CreateElement("In");
                profileElem.AppendChild(inElem);
                for (int i = 0; i < inIds.Length; i++)
                {
                    if (inIds[i].ObjectClass.DxfName.Equals("LWPOLYLINE"))
                    {
                        using (Transaction transaction = MyPlugin.doc.Database.TransactionManager.StartTransaction())
                        {
                            Polyline polyline;
                            polyline = transaction.GetObject(inIds[i], OpenMode.ForRead) as Polyline;
                            XmlElement lineElem = xDoc.CreateElement("LWPolyLine");
                            inElem.AppendChild(lineElem);
                            for (int j = 0; j < polyline.NumberOfVertices; j++)
                            {
                                XmlElement XElem = xDoc.CreateElement("X" + j);
                                XmlElement YElem = xDoc.CreateElement("Y" + j);
                                XmlElement ZElem = xDoc.CreateElement("Z" + j);
                                XmlElement BulgeElem = xDoc.CreateElement("Bulge" + j);
                                XmlText XText = xDoc.CreateTextNode(polyline.GetPoint3dAt(j).X.ToString());
                                XmlText YText = xDoc.CreateTextNode(polyline.GetPoint3dAt(j).Y.ToString());
                                XmlText ZText = xDoc.CreateTextNode(polyline.GetPoint3dAt(j).Z.ToString());
                                XmlText BulgeText = xDoc.CreateTextNode(polyline.GetBulgeAt(j).ToString());

                                XElem.AppendChild(XText);
                                YElem.AppendChild(YText);
                                ZElem.AppendChild(ZText);
                                BulgeElem.AppendChild(BulgeText);
                                lineElem.AppendChild(XElem);
                                lineElem.AppendChild(YElem);
                                lineElem.AppendChild(ZElem);
                                lineElem.AppendChild(BulgeElem);
                            }
                        }
                    }
                    if (inIds[i].ObjectClass.DxfName.Equals("CIRCLE"))
                    {
                        using (Transaction transaction = MyPlugin.doc.Database.TransactionManager.StartTransaction())
                        {
                            //Center
                            Circle circle;
                            circle = transaction.GetObject(inIds[i], OpenMode.ForRead) as Circle;
                            XmlElement lineElem = xDoc.CreateElement("Circle");
                            inElem.AppendChild(lineElem);
                            XmlElement XElem = xDoc.CreateElement("X0");
                            XmlElement YElem = xDoc.CreateElement("Y0");
                            XmlElement ZElem = xDoc.CreateElement("Z0");
                            XmlText XText = xDoc.CreateTextNode(circle.Center.X.ToString());
                            XmlText YText = xDoc.CreateTextNode(circle.Center.Y.ToString());
                            XmlText ZText = xDoc.CreateTextNode(circle.Center.Z.ToString());
                            XElem.AppendChild(XText);
                            YElem.AppendChild(YText);
                            ZElem.AppendChild(ZText);
                            lineElem.AppendChild(XElem);
                            lineElem.AppendChild(YElem);
                            lineElem.AppendChild(ZElem);
                            //Normal
                            XmlElement normalXElem = xDoc.CreateElement("NormalX");
                            XmlElement normalYElem = xDoc.CreateElement("NormalY");
                            XmlElement normalZElem = xDoc.CreateElement("NormalZ");
                            XmlText normalXText = xDoc.CreateTextNode(circle.Normal.X.ToString());
                            XmlText normalYText = xDoc.CreateTextNode(circle.Normal.Y.ToString());
                            XmlText normalZText = xDoc.CreateTextNode(circle.Normal.Z.ToString());
                            normalXElem.AppendChild(normalXText);
                            normalYElem.AppendChild(normalYText);
                            normalZElem.AppendChild(normalZText);
                            lineElem.AppendChild(normalXElem);
                            lineElem.AppendChild(normalYElem);
                            lineElem.AppendChild(normalZElem);
                            //Radius
                            XmlElement RElem = xDoc.CreateElement("R");
                            XmlText RText = xDoc.CreateTextNode(circle.Radius.ToString());
                            RElem.AppendChild(RText);
                            lineElem.AppendChild(RElem);
                        }
                    }
                    if (inIds[i].ObjectClass.DxfName.Equals("ELLIPSE"))
                    {
                        using (Transaction transaction = MyPlugin.doc.Database.TransactionManager.StartTransaction())
                        {
                            Ellipse ellipse;
                            ellipse = transaction.GetObject(inIds[i], OpenMode.ForRead) as Ellipse;
                            XmlElement lineElem = xDoc.CreateElement("Ellipse");
                            inElem.AppendChild(lineElem);
                            //center
                            XmlElement XElem = xDoc.CreateElement("X0");
                            XmlElement YElem = xDoc.CreateElement("Y0");
                            XmlElement ZElem = xDoc.CreateElement("Z0");
                            XmlText XText = xDoc.CreateTextNode(ellipse.Center.X.ToString());
                            XmlText YText = xDoc.CreateTextNode(ellipse.Center.Y.ToString());
                            XmlText ZText = xDoc.CreateTextNode(ellipse.Center.Z.ToString());
                            XElem.AppendChild(XText);
                            YElem.AppendChild(YText);
                            ZElem.AppendChild(ZText);
                            lineElem.AppendChild(XElem);
                            lineElem.AppendChild(YElem);
                            lineElem.AppendChild(ZElem);
                            //normal
                            XmlElement normalXElem = xDoc.CreateElement("NormalX");
                            XmlElement normalYElem = xDoc.CreateElement("NormalY");
                            XmlElement normalZElem = xDoc.CreateElement("NormalZ");
                            XmlText normalXText = xDoc.CreateTextNode(ellipse.Normal.X.ToString());
                            XmlText normalYText = xDoc.CreateTextNode(ellipse.Normal.Y.ToString());
                            XmlText normalZText = xDoc.CreateTextNode(ellipse.Normal.Z.ToString());
                            normalXElem.AppendChild(normalXText);
                            normalYElem.AppendChild(normalYText);
                            normalZElem.AppendChild(normalZText);
                            lineElem.AppendChild(normalXElem);
                            lineElem.AppendChild(normalYElem);
                            lineElem.AppendChild(normalZElem);
                            //MajorAxis
                            XmlElement mAxisXElem = xDoc.CreateElement("mAxisX");
                            XmlElement mAxisYElem = xDoc.CreateElement("mAxisY");
                            XmlElement mAxisZElem = xDoc.CreateElement("mAxisZ");
                            XmlText mAxisXText = xDoc.CreateTextNode(ellipse.MajorAxis.X.ToString());
                            XmlText mAxisYText = xDoc.CreateTextNode(ellipse.MajorAxis.Y.ToString());
                            XmlText mAxisZText = xDoc.CreateTextNode(ellipse.MajorAxis.Z.ToString());
                            mAxisXElem.AppendChild(mAxisXText);
                            mAxisYElem.AppendChild(mAxisYText);
                            mAxisZElem.AppendChild(mAxisZText);
                            lineElem.AppendChild(mAxisXElem);
                            lineElem.AppendChild(mAxisYElem);
                            lineElem.AppendChild(mAxisZElem);
                            //RadiusRatio
                            XmlElement radRatioElem = xDoc.CreateElement("RadiusRatio");
                            XmlText radRatioText = xDoc.CreateTextNode(ellipse.RadiusRatio.ToString());
                            radRatioElem.AppendChild(radRatioText);
                            lineElem.AppendChild(radRatioElem);
                            //StartAngle
                            XmlElement startAngElem = xDoc.CreateElement("StartAngle");
                            XmlText startAngText = xDoc.CreateTextNode(ellipse.StartAngle.ToString());
                            startAngElem.AppendChild(startAngText);
                            lineElem.AppendChild(startAngElem);
                            //EndAngle
                            XmlElement endAngElem = xDoc.CreateElement("EndAngle");
                            XmlText endAngText = xDoc.CreateTextNode(ellipse.EndAngle.ToString());
                            endAngElem.AppendChild(endAngText);
                            lineElem.AppendChild(endAngElem);
                        }
                    }
                }
                xDoc.Save("profiles.xml");
            }
        }

        /// <summary>
        /// Retrieve or create an Entry with the given name into the Extension Dictionary of the passed-in object.
        /// </summary>
        /// <param name="id">The object hosting the Extension Dictionary</param>
        /// <param name="entryName">The name of the dictionary entry to get or set</param>
        /// <returns>The ObjectId of the diction entry, old or new</returns>
        public static ObjectId GetSetExtensionDictionaryEntry(ObjectId id, string entryName)
        {
            ObjectId ret = ObjectId.Null;

            using (Transaction tr = MyPlugin.doc.Database.TransactionManager.StartTransaction())
            {
                DBObject obj = (DBObject)tr.GetObject(id, OpenMode.ForRead);
                if (obj.ExtensionDictionary == ObjectId.Null)
                {
                    obj.UpgradeOpen();
                    obj.CreateExtensionDictionary();
                    obj.DowngradeOpen();
                }

                DBDictionary dict = (DBDictionary)tr.GetObject(obj.ExtensionDictionary, OpenMode.ForRead);
                if (!dict.Contains(entryName))
                {
                    Xrecord xRecord = new Xrecord();
                    dict.UpgradeOpen();
                    dict.SetAt(entryName, xRecord);
                    tr.AddNewlyCreatedDBObject(xRecord, true);

                    ret = xRecord.ObjectId;
                }
                else
                    ret = dict.GetAt(entryName);

                tr.Commit();
            }

            return ret;
        }

        /// <summary>
        /// Remove the named Dictionary Entry from the passed-in object if applicable.
        /// </summary>
        /// <param name="id">The object hosting the Extension Dictionary</param>
        /// <param name="entryName">The name of the Dictionary Entry to remove</param>
        /// <returns>True if really removed, false if not there at all</returns>
        public static bool RemoveExtensionDictionaryEntry(ObjectId id, string entryName)
        {
            bool ret = false;

            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                DBObject obj = (DBObject)tr.GetObject(id, OpenMode.ForRead);
                if (obj.ExtensionDictionary != ObjectId.Null)
                {
                    DBDictionary dict = (DBDictionary)tr.GetObject(obj.ExtensionDictionary, OpenMode.ForRead);
                    if (dict.Contains(entryName))
                    {
                        dict.UpgradeOpen();
                        dict.Remove(entryName);
                        ret = true;
                    }
                }

                tr.Commit();
            }

            return ret;
        }
    }
}
