using System;
using System.Linq;

using Db = Autodesk.AutoCAD.DatabaseServices;

namespace dynIN_dynOUT
{
    /// <summary>
    /// https://sites.google.com/site/bushmansnetlaboratory/moi-zametki/attsynch
    /// Методы расширений для объектов класса Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
    /// </summary>
    public static class BlockReferenceExtensionMethods
    {
        public static string EffectiveName(this Db.BlockReference acBlockRef)
        {
            var blockName = acBlockRef.Name;

            if (acBlockRef.IsDynamicBlock)
            {
                var dynamicBlockTableRecordId = acBlockRef.DynamicBlockTableRecord;

                using (var blrNam = dynamicBlockTableRecordId.Open(Db.OpenMode.ForRead) as Db.BlockTableRecord)
                {
                    blockName = blrNam.Name;
                    blrNam.Close();
                }
            }

            return blockName;
        }

        /// <summary>
        /// Синхронизация вхождений блоков с их определением
        /// </summary>
        /// <param name="btr">Запись таблицы блоков, принятая за определение блока</param>
        /// <param name="directOnly">Следует ли искать только на верхнем уровне, или же нужно 
        /// анализировать и вложенные вхождения, т.е. следует ли рекурсивно обрабатывать блок в блоке:
        /// true - только верхний; false - рекурсивно проверять вложенные блоки.</param>
        /// <param name="removeSuperfluous">
        /// Следует ли во вхождениях блока удалять лишние атрибуты (те, которых нет в определении блока).</param>
        /// <param name="setAttDefValues">
        /// Следует ли всем атрибутам, во вхождениях блока, назначить текущим значением значение по умолчанию.</param>
        /// <param name="addSuperfluous">
        /// Следует ли добавлять отсутствующие в blockReference атрибуты: true - добавлять</param>
        public static void AttSync(this Db.BlockTableRecord btr, bool directOnly, bool removeSuperfluous, bool setAttDefValues)
        {
            var db = btr.Database;
            using (var wdb = new WorkingDatabaseSwitcher(db))
            {
                using (var t = db.TransactionManager.StartTransaction())
                {
                    var bt = (Db.BlockTable)t.GetObject(db.BlockTableId, Db.OpenMode.ForRead);

                    //Получаем все определения атрибутов из определения блока
                    var attdefs = btr.Cast<Db.ObjectId>()
                        .Where(n => n.ObjectClass.Name == "AcDbAttributeDefinition")
                        .Select(n => (Db.AttributeDefinition)t.GetObject(n, Db.OpenMode.ForRead))
                        .Where(n => !n.Constant);//Исключаем константные атрибуты, т.к. для них AttributeReference не создаются.

                    //В цикле перебираем все вхождения искомого определения блока
                    foreach (Db.ObjectId brId in btr.GetBlockReferenceIds(directOnly, false))
                    {
                        var br = (Db.BlockReference)t.GetObject(brId, Db.OpenMode.ForWrite);

                        //Проверяем имена на соответствие. В том случае, если вхождение блока "A" вложено в определение блока "B", 
                        //то вхождения блока "B" тоже попадут в выборку. Нам нужно их исключить из набора обрабатываемых объектов 
                        //- именно поэтому проверяем имена.
                        if (br.Name != btr.Name) continue;

                        //Получаем все атрибуты вхождения блока
                        var attrefs = br.AttributeCollection.Cast<Db.ObjectId>()
                            .Select(n => (Db.AttributeReference)t.GetObject(n, Db.OpenMode.ForWrite));

                        //Тэги существующих определений атрибутов
                        var dtags = attdefs.Select(n => n.Tag);
                        //Тэги существующих атрибутов во вхождении
                        var rtags = attrefs.Select(n => n.Tag);

                        //Если требуется - удаляем те атрибуты, для которых нет определения 
                        //в составе определения блока
                        if (removeSuperfluous)
                            foreach (var attref in attrefs.Where(n => rtags.Except(dtags).Contains(n.Tag)))
                                attref.Erase(true);

                        //Свойства существующих атрибутов синхронизируем со свойствами их определений
                        foreach (var attref in attrefs.Where(n => dtags
                            .Join(rtags, a => a, b => b, (a, b) => a).Contains(n.Tag)))
                        {
                            var ad = attdefs.First(n => n.Tag == attref.Tag);

                            //Метод SetAttributeFromBlock, используемый нами далее в коде, сбрасывает
                            //текущее значение многострочного атрибута. Поэтому запоминаем это значение,
                            //чтобы восстановить его сразу после вызова SetAttributeFromBlock.
                            var value = attref.TextString;
                            attref.SetAttributeFromBlock(ad, br.BlockTransform);
                            //Восстанавливаем значение атрибута
                            attref.TextString = value;

                            if (attref.IsMTextAttribute)
                            {
                            }

                            //Если требуется - устанавливаем для атрибута значение по умолчанию
                            if (setAttDefValues) attref.TextString = ad.TextString;

                            attref.AdjustAlignment(db);
                        }

                        //Если во вхождении блока отсутствуют нужные атрибуты - создаём их
                        var attdefsNew = attdefs.Where(n => dtags
                            .Except(rtags).Contains(n.Tag));

                        foreach (var ad in attdefsNew)
                        {
                            var attref = new Db.AttributeReference();
                            attref.SetAttributeFromBlock(ad, br.BlockTransform);
                            attref.AdjustAlignment(db);
                            br.AttributeCollection.AppendAttribute(attref);
                            t.AddNewlyCreatedDBObject(attref, true);
                        }
                    }
                    btr.UpdateAnonymousBlocks();
                    t.Commit();
                } // end Transaction


                //Если это динамический блок
                if (btr.IsDynamicBlock)
                {
                    using (var t = db.TransactionManager.StartTransaction())
                    {
                        foreach (Db.ObjectId id in btr.GetAnonymousBlockIds())
                        {
                            var _btr = (Db.BlockTableRecord)t.GetObject(id, Db.OpenMode.ForWrite);

                            //Получаем все определения атрибутов из оригинального определения блока
                            var attdefs = btr.Cast<Db.ObjectId>()
                                .Where(n => n.ObjectClass.Name == "AcDbAttributeDefinition")
                                .Select(n => (Db.AttributeDefinition)t.GetObject(n, Db.OpenMode.ForRead));

                            //Получаем все определения атрибутов из определения анонимного блока
                            var attdefs2 = _btr.Cast<Db.ObjectId>()
                                .Where(n => n.ObjectClass.Name == "AcDbAttributeDefinition")
                                .Select(n => (Db.AttributeDefinition)t.GetObject(n, Db.OpenMode.ForWrite));

                            //Определения атрибутов анонимных блоков следует синхронизировать 
                            //с определениями атрибутов основного блока

                            //Тэги существующих определений атрибутов
                            var dtags = attdefs.Select(n => n.Tag);
                            var dtags2 = attdefs2.Select(n => n.Tag);

                            //1. Удаляем лишние
                            foreach (var attdef in attdefs2.Where(n => !dtags.Contains(n.Tag)))
                            {
                                attdef.Erase(true);
                            }

                            //2. Синхронизируем существующие
                            foreach (var attdef in attdefs.Where(n => dtags
                               .Join(dtags2, a => a, b => b, (a, b) => a)
                               .Contains(n.Tag)))
                            {
                                var ad = attdefs2.First(n => n.Tag == attdef.Tag);
                                ad.Position = attdef.Position;
                                //ad.TextStyle = attdef.TextStyle;

                                //Если требуется - устанавливаем для атрибута значение по умолчанию
                                if (setAttDefValues) ad.TextString = attdef.TextString;

                                ad.Tag = attdef.Tag;
                                ad.Prompt = attdef.Prompt;
                                ad.LayerId = attdef.LayerId;
                                ad.Rotation = attdef.Rotation;
                                ad.LinetypeId = attdef.LinetypeId;
                                ad.LineWeight = attdef.LineWeight;
                                ad.LinetypeScale = attdef.LinetypeScale;
                                //ad.Annotative = attdef.Annotative;
                                ad.Color = attdef.Color;
                                ad.Height = attdef.Height;
                                ad.HorizontalMode = attdef.HorizontalMode;
                                ad.Invisible = attdef.Invisible;
                                ad.IsMirroredInX = attdef.IsMirroredInX;
                                ad.IsMirroredInY = attdef.IsMirroredInY;
                                ad.Justify = attdef.Justify;
                                ad.LockPositionInBlock = attdef.LockPositionInBlock;
                                ad.MaterialId = attdef.MaterialId;
                                ad.Oblique = attdef.Oblique;
                                ad.Thickness = attdef.Thickness;
                                ad.Transparency = attdef.Transparency;
                                ad.VerticalMode = attdef.VerticalMode;
                                ad.Visible = attdef.Visible;
                                ad.WidthFactor = attdef.WidthFactor;

                                ad.CastShadows = attdef.CastShadows;
                                ad.Constant = attdef.Constant;
                                ad.FieldLength = attdef.FieldLength;
                                ad.ForceAnnoAllVisible = attdef.ForceAnnoAllVisible;
                                ad.Preset = attdef.Preset;
                                ad.Prompt = attdef.Prompt;
                                ad.Verifiable = attdef.Verifiable;

                                ad.AdjustAlignment(db);
                            }

                            //3. Добавляем недостающие
                            foreach (var attdef in attdefs.Where(n => !dtags2.Contains(n.Tag)))
                            {
                                var ad = new Db.AttributeDefinition();
                                ad.SetDatabaseDefaults();
                                ad.Position = attdef.Position;
                                //ad.TextStyle = attdef.TextStyle;
                                ad.TextString = attdef.TextString;
                                ad.Tag = attdef.Tag;
                                ad.Prompt = attdef.Prompt;

                                ad.LayerId = attdef.LayerId;
                                ad.Rotation = attdef.Rotation;
                                ad.LinetypeId = attdef.LinetypeId;
                                ad.LineWeight = attdef.LineWeight;
                                ad.LinetypeScale = attdef.LinetypeScale;
                                //ad.Annotative = attdef.Annotative;
                                ad.Color = attdef.Color;
                                ad.Height = attdef.Height;
                                ad.HorizontalMode = attdef.HorizontalMode;
                                ad.Invisible = attdef.Invisible;
                                ad.IsMirroredInX = attdef.IsMirroredInX;
                                ad.IsMirroredInY = attdef.IsMirroredInY;
                                ad.Justify = attdef.Justify;
                                ad.LockPositionInBlock = attdef.LockPositionInBlock;
                                ad.MaterialId = attdef.MaterialId;
                                ad.Oblique = attdef.Oblique;
                                ad.Thickness = attdef.Thickness;
                                ad.Transparency = attdef.Transparency;
                                ad.VerticalMode = attdef.VerticalMode;
                                ad.Visible = attdef.Visible;
                                ad.WidthFactor = attdef.WidthFactor;

                                ad.CastShadows = attdef.CastShadows;
                                ad.Constant = attdef.Constant;
                                ad.FieldLength = attdef.FieldLength;
                                ad.ForceAnnoAllVisible = attdef.ForceAnnoAllVisible;
                                ad.Preset = attdef.Preset;
                                ad.Prompt = attdef.Prompt;
                                ad.Verifiable = attdef.Verifiable;

                                _btr.AppendEntity(ad);
                                t.AddNewlyCreatedDBObject(ad, true);
                                ad.AdjustAlignment(db);
                            }
                            //Синхронизируем все вхождения данного анонимного определения блока
                            _btr.AttSync(directOnly, removeSuperfluous, setAttDefValues);
                        }
                        //Обновляем геометрию определений анонимных блоков, полученных на основе 
                        //этого динамического блока
                        btr.UpdateAnonymousBlocks();
                        t.Commit();
                    }
                }
            }
        }
    }

    /// <summary>
    /// Изменяя базу данных чертежей, очень важно контролировать то, какая база данных является текущей. 
    /// Класс <c>WorkingDatabaseSwitcher</c>
    /// берёт на себя контроль над тем, чтобы текущей была именно та база данных, которая нужна.
    /// </summary>
    /// <example>
    /// Пример использования класса:
    /// <code>
    /// //db - объект Database
    /// using (WorkingDatabaseSwitcher hlp = new WorkingDatabaseSwitcher(db)) {
    ///     // тут наш код</code>
    /// }</example>
    public sealed class WorkingDatabaseSwitcher : IDisposable
    {
        private Db.Database prevDb;

        /// <summary>
        /// База данных, в контексте которой должна производиться работа. Эта база данных на время становится текущей.
        /// По завершению работы текущей станет та база, которая была ею до этого.
        /// </summary>
        /// <param name="db">База данных, которая должна быть установлена текущей</param>
        public WorkingDatabaseSwitcher(Db.Database db)
        {
            prevDb = Db.HostApplicationServices.WorkingDatabase;
            Db.HostApplicationServices.WorkingDatabase = db;
        }

        /// <summary>
        /// Возвращаем свойству <c>HostApplicationServices.WorkingDatabase</c> прежнее значение
        /// </summary>
        public void Dispose()
        {
            Db.HostApplicationServices.WorkingDatabase = prevDb;
        }
    }
}