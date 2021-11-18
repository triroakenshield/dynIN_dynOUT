using System.Linq;

using Db = Autodesk.AutoCAD.DatabaseServices;
using Gem = Autodesk.AutoCAD.Geometry;

namespace dynIN_dynOUT
{
    /// <summary>
    /// Вспомогательный класс для создания объектов 
    /// в базе чертежа
    /// </summary>
    public static class AddEntity
    {
        /// <summary>
        /// Проверяем и создаем заданный слой
        /// </summary>
        /// <param name="name">Имя слоя</param>
        /// <param name="create">true - если слоя с таким именем нет, то он будет создан</param>
        /// <returns></returns>
        public static bool CreateLayer(string name, bool create)
        {
            var outBool = false;
            var db = Db.HostApplicationServices.WorkingDatabase;

            using (var acTrans = db.TransactionManager.StartTransaction())
            {
                var layerTable = (Db.LayerTable)acTrans.GetObject(db.LayerTableId, Db.OpenMode.ForRead);

                if (!layerTable.Has(name))
                {
                    if (create)
                    {
                        using (var layerRecord = new Db.LayerTableRecord())
                        {
                            layerRecord.Name = name;
                            layerTable.UpgradeOpen();
                            layerTable.Add(layerRecord);
                            acTrans.AddNewlyCreatedDBObject(layerRecord, true);
                        }
                        outBool = true;
                    }
                }
                else
                {
                    outBool = true;
                }

                acTrans.Commit();
            }
            return outBool;
        }

        /// <summary>
        /// Проверяем и создаем заданный блок
        /// </summary>
        /// <param name="blockName">Имя блока</param>
        /// <returns></returns>
        public static Db.ObjectId CreateBlockReference(string blockName)
        {
            var newBtrId = Db.ObjectId.Null;
            var db = Db.HostApplicationServices.WorkingDatabase;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var acBlkTbl = tr.GetObject(db.BlockTableId, Db.OpenMode.ForRead) as Db.BlockTable;
                if (!(acBlkTbl is null) && acBlkTbl.Has(blockName))
                {
                    var acBlkTblRec = tr.GetObject(acBlkTbl[blockName], Db.OpenMode.ForWrite) as Db.BlockTableRecord;

                    //var BRId = Db.ObjectId.Null;

                    using (var ms = (Db.BlockTableRecord)tr.GetObject(acBlkTbl[Db.BlockTableRecord.ModelSpace], Db.OpenMode.ForWrite))
                    using (var br = new Db.BlockReference(Gem.Point3d.Origin, acBlkTblRec.ObjectId))
                    {

                        ms.AppendEntity(br);
                        tr.AddNewlyCreatedDBObject(br, true);

                        //Получаем все определения атрибутов из определения блока
                        var attDefs = acBlkTblRec.Cast<Db.ObjectId>()
                            .Where(n => n.ObjectClass.Name == "AcDbAttributeDefinition")
                            .Select(n => (Db.AttributeDefinition)tr.GetObject(n, Db.OpenMode.ForRead))
                            .Where(n => !n.Constant);//Исключаем константные атрибуты, т.к. для них AttributeReference не создаются.

                        foreach (var attRef in attDefs)
                        {
                            var ar = new Db.AttributeReference();
                            ar.SetAttributeFromBlock(attRef, br.BlockTransform);
                            br.AttributeCollection.AppendAttribute(ar);
                            tr.AddNewlyCreatedDBObject(ar, true);

                        }
                        newBtrId = br.ObjectId;
                    }
                }
                tr.Commit();
            }
            return newBtrId;
        }
    }
}