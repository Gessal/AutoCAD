using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCAD_PointsReader.Examples
{
    class Example1
    {
        /// <summary>
        /// Твердое тело по траектории
        /// </summary>
        [CommandMethod("SAP")]
        public void SweepAlongPath()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Ask the user to select a region to extrude
            PromptEntityOptions peo1 = new PromptEntityOptions("\nSelect profile or curve to sweep: ");
            peo1.SetRejectMessage("\nEntity must be a region, curve or planar surface.");
            peo1.AddAllowedClass(typeof(Region), false);
            peo1.AddAllowedClass(typeof(Curve), false);
            peo1.AddAllowedClass(typeof(PlaneSurface), false);
            PromptEntityResult per = ed.GetEntity(peo1);
            if (per.Status != PromptStatus.OK)
                return;
            ObjectId regId = per.ObjectId;

            // Ask the user to select an extrusion path
            PromptEntityOptions peo2 = new PromptEntityOptions("\nSelect path along which to sweep: ");
            peo2.SetRejectMessage("\nEntity must be a curve.");
            peo2.AddAllowedClass(typeof(Curve), false);
            per = ed.GetEntity(peo2);
            if (per.Status != PromptStatus.OK)
                return;
            ObjectId splId = per.ObjectId;
            PromptKeywordOptions pko = new PromptKeywordOptions("\nSweep a solid or a surface?");
            pko.AllowNone = true;
            pko.Keywords.Add("SOlid");
            pko.Keywords.Add("SUrface");
            pko.Keywords.Default = "SOlid";
            PromptResult pkr = ed.GetKeywords(pko);
            bool createSolid = (pkr.StringResult == "SOlid");
            if (pkr.Status != PromptStatus.OK)
                return;

            // Now let's create our swept surface
            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                try
                {
                    Entity sweepEnt = tr.GetObject(regId, OpenMode.ForRead) as Entity;
                    Curve pathEnt = tr.GetObject(splId, OpenMode.ForRead) as Curve;
                    if (sweepEnt == null || pathEnt == null)
                    {
                        ed.WriteMessage("\nProblem opening the selected entities.");
                        return;
                    }

                    // We use a builder object to create
                    // our SweepOptions
                    SweepOptionsBuilder sob = new SweepOptionsBuilder();

                    // Align the entity to sweep to the path
                    sob.Align = SweepOptionsAlignOption.AlignSweepEntityToPath;

                    // The base point is the start of the path
                    sob.BasePoint = pathEnt.StartPoint;

                    // The profile will rotate to follow the path
                    sob.Bank = true;

                    // Now generate the solid or surface...
                    Entity ent;
                    if (createSolid)
                    {
                        Solid3d sol = new Solid3d();
                        sol.CreateSweptSolid(sweepEnt, pathEnt, sob.ToSweepOptions());
                        ent = sol;
                    }
                    else
                    {
                        SweptSurface ss = new SweptSurface();
                        ss.CreateSweptSurface(sweepEnt, pathEnt, sob.ToSweepOptions());
                        ent = ss;
                    }

                    // ... and add it to the modelspace
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    ms.AppendEntity(ent);
                    tr.AddNewlyCreatedDBObject(ent, true);
                    tr.Commit();
                }
                catch
                {

                }
            }
        }

        /// <summary>
        /// Добавление записи объекту
        /// </summary>
        /// <param name="ename"></param> id объекта которому вы задаем запись
        /// <param name="key"></param> ключ записи
        /// <param name="value"></param> значение записи
        public static void SetExtDictionaryValueString(ObjectId ename, string key, string value)
        {
            if (ename == ObjectId.Null) throw new ArgumentNullException("ename");
            if (String.IsNullOrEmpty(key)) throw new ArgumentNullException("key");

            var doc = Application.DocumentManager.MdiActiveDocument;

            using (var transaction = doc.Database.TransactionManager.StartTransaction())
            {
                // Открыл объект по id для чтения
                var entity = transaction.GetObject(ename, OpenMode.ForWrite);
                if (entity == null)
                    throw new DataException("Ошибка при записи текстового значения в ExtensionDictionary: entity " +
                                            "с ObjectId=" + ename + " не найдена");
                //Получение или создание словаря extDictionary
                var extensionDictionaryId = entity.ExtensionDictionary;
                if (extensionDictionaryId == ObjectId.Null)
                {
                    entity.CreateExtensionDictionary();                     // если такого словаря не было
                    extensionDictionaryId = entity.ExtensionDictionary;     // то создаем его
                }
                var extDictionary = (DBDictionary)transaction.GetObject(extensionDictionaryId, OpenMode.ForWrite);
                // Запись значения в словарь
                if (String.IsNullOrEmpty(value))
                {
                    if (extDictionary.Contains(key))
                        extDictionary.Remove(key);
                    return;
                }
                var xrec = new Xrecord();                                   // создаем новую запись
                xrec.Data = new ResultBuffer(new TypedValue((int)DxfCode.ExtendedDataAsciiString, value)); // указываю тип значения записи (текст)
                extDictionary.SetAt(key, xrec); // записываю пару ключ-значение в словарь
                transaction.AddNewlyCreatedDBObject(xrec, true);
                Debug.WriteLine(entity.Handle + "['" + key + "'] = '" + value + "'");
                transaction.Commit();

            }
        }

        /// <summary>
        /// Получение значения записи (строка)
        /// </summary>
        /// <param name="ename"></param> объект, у которого нужно прочитать запись
        /// <param name="key"></param> ключ записи
        /// <returns></returns>
        public static string GetExtDictionaryValueString(ObjectId ename, string key)
        {
            if (ename == ObjectId.Null) throw new ArgumentNullException("ename");
            if (String.IsNullOrEmpty(key)) throw new ArgumentNullException("key");


            var doc = Application.DocumentManager.MdiActiveDocument;

            using (var transaction = doc.Database.TransactionManager.StartTransaction())
            {
                var entity = transaction.GetObject(ename, OpenMode.ForRead);
                if (entity == null)
                    throw new DataException("Ошибка при чтении текстового значения из ExtensionDictionary: полилиния с ObjectId=" + ename + " не найдена");

                var extDictionaryId = entity.ExtensionDictionary;
                if (extDictionaryId == ObjectId.Null)
                    throw new DataException("Ошибка при чтении текстового значения из ExtensionDictionary: словарь не найден");
                var extDic = (DBDictionary)transaction.GetObject(extDictionaryId, OpenMode.ForRead);

                if (!extDic.Contains(key))
                    return null;
                var myDataId = extDic.GetAt(key);
                var readBack = (Xrecord)transaction.GetObject(myDataId, OpenMode.ForRead);
                return (string)readBack.Data.AsArray()[0].Value;
            }
        }
    }
}
