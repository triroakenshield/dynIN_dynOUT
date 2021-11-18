using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using Db = Autodesk.AutoCAD.DatabaseServices;
using Rtm = Autodesk.AutoCAD.Runtime;

namespace dynIN_dynOUT
{
    /// <summary>Читаем данные из txt файла</summary>
    internal static class DynIN
    {
        internal static void IN()
        {
            // Получение текущего документа и базы данных
            var acDoc = Application.DocumentManager.MdiActiveDocument;
            if (acDoc == null) return;
            var acCurDb = acDoc.Database;
            var acEd = acDoc.Editor;

            //1. Читаем и парсим файл
            //OpenFileDialog openFileDialog = new OpenFileDialog("Выберите CSV файл",
            //                              "*.csv",
            //                              "csv",
            //                              "Выбор файла",
            //                              OpenFileDialog.OpenFileDialogFlags.NoUrls & OpenFileDialog.OpenFileDialogFlags.DefaultIsFolder );

            //if (openFileDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
            //string fileName = openFileDialog.Filename;

            var openFileDialog1 = new OpenFileDialog
            {
                Title = "Выберите CSV файл",
                Filter = "csv файлы (*.csv)|*.csv|Все файлы (*.*)|*.*",
                FileName = "",
                InitialDirectory = Settings.Data.Lastpath,
                RestoreDirectory = false
            };

            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            var fileName = openFileDialog1.FileName;

            Settings.Data.Lastpath = new FileInfo(fileName).DirectoryName;

            var fileLines = new List<string>();
            try
            {
                fileLines = File.ReadAllLines(fileName, Encoding.GetEncoding(1251)).ToList();
            }
            catch (Exception e)
            {
                acEd.WriteMessage($"\nDynIN.IN-Ошибка чтения файла: {e.Message}");
                return;
            }

            var propertyList = new List<Property>();
            //Парсим первую строку
            var rowHead = fileLines[0].Split(';').ToList();
            //Парсим основное тело
            for (var i = 1; i < fileLines.Count; i++)
            {
                var prop = new Property();
                //bool t = prop.Sets(rowHead, fileLines[i]);
                if (prop.Sets(rowHead, fileLines[i])) propertyList.Add(prop);
            }

            //Блокируем документ
            using (var docLock = acDoc.LockDocument())
            {
                //Прежде всего пройдемся по всем объектам 
                //и посмотрим все ли слои есть в базе
                foreach (var i in propertyList)
                {
                    try
                    {
                        // Validate the provided symbol table name
                        // И проверим имя слоя на плохие символы
                        Db.SymbolUtilityServices.ValidateSymbolName(i.Layer, false);
                        AddEntity.CreateLayer(i.Layer, Settings.Data.CreateLayer);
                    }
                    catch
                    {
                        // An exception has been thrown, indicating that
                        // the name is invalid
                        acEd.WriteMessage($"\n{i.Layer} is an invalid Layer name and it name replace to \"0\".");
                        i.Layer = "0";
                    }
                }

                // старт транзакции
                using (var acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    //прогресс бар
                    var pm = new Rtm.ProgressMeter();
                    pm.Start("Progress of processing BlockReference");
                    pm.SetLimit(propertyList.Count);

                    foreach (var prop in propertyList)
                    {
                        var id = Db.ObjectId.Null;
                        if (prop.BlockName == "")
                        {
                            try
                            {
                                id = acCurDb.GetObjectId(false, prop.Handle, 0);
                            }
                            catch (Exception e)
                            {
                                acEd.WriteMessage($"\nDynIN.IN-Ошибка поиска объекта: {e.Message}");
                            }
                        }
                        else
                        {
                            if (Settings.Data.CreateBlocReference && prop.BlockName != "")
                            {
                                id = Db.ObjectId.Null;
                                id = AddEntity.CreateBlockReference(prop.BlockName);
                            }
                        }

                        if (id.IsNull && !id.IsResident && !id.IsValid && id.IsErased) break;

                        //Полученный объект вообще блок, если нет то переходим к следующему
                        if (!id.ObjectClass.IsDerivedFrom(Rtm.RXObject.GetClass(typeof(Db.BlockReference)))) break;

                        var acBlRef = acTrans.GetObject(id, Db.OpenMode.ForWrite) as Db.BlockReference;
                        var blr = (Db.BlockTableRecord)acTrans.GetObject(acBlRef.DynamicBlockTableRecord,
                                                                    Db.OpenMode.ForRead);

                        prop.Handle = acBlRef.Handle;
                        prop.BlockName = acBlRef.EffectiveName();

                        var propsBlockRef = acBlRef.GetType().GetProperties();
                        var propElement = prop.GetType().GetProperties();

                        foreach (var propInfo in propElement)
                        {
                            try
                            {
                                //System.Reflection.PropertyInfo propBlock = propsBlockRef.Where(x => x.Name == propInfo.Name).FirstOrDefault();
                                //if (propBlock != null) propInfo.SetValue(prop, propBlock.GetValue(acBlRef, null));

                                var propBlock = propsBlockRef.Where(x => x.Name == propInfo.Name).FirstOrDefault();

                                var oo = propInfo.GetValue(prop, null);
                                if (propBlock != null)
                                {
                                    propBlock.SetValue(acBlRef, propInfo.GetValue(prop, null),null);
                                }

                            }
                            catch (Exception ex)
                            {
                                acEd.WriteMessage($"\nError: DynIN-IN -> {ex.Message}");
                            }
                        }

                        if (blr.HasAttributeDefinitions)
                        {
                            var attrCol = acBlRef.AttributeCollection;
                            if (attrCol.Count > 0)
                            {
                                foreach (Db.ObjectId AttID in attrCol)
                                {
                                    var acAttRef = acTrans.GetObject(AttID,
                                                            Db.OpenMode.ForRead) as Db.AttributeReference;

                                    foreach (var i in prop.Attribut)
                                    {
                                        //Обновляем только в том случае если были изменения
                                        if (acAttRef.Tag == i.Key && acAttRef.TextString != i.Value)
                                        {
                                            acAttRef.UpgradeOpen();
                                            acAttRef.TextString = i.Value;
                                            //acAttRef.RecordGraphicsModified(true);
                                            acAttRef.DowngradeOpen();
                                            break;
                                        }
                                    }
                                }
                            }   //Проверка что кол аттрибутов больше 0
                        }  //Проверка наличия атрибутов

                        var acBlockDynProp = acBlRef.DynamicBlockReferencePropertyCollection;
                        if (acBlockDynProp != null)
                        {
                            foreach (Db.DynamicBlockReferenceProperty obj in acBlockDynProp)
                            {
                                foreach (var i in prop.DynProp)
                                {
                                    //Дополнительно проверяем можно ли вообще обновить значение
                                    if (obj.PropertyName == i.Key && !obj.ReadOnly)
                                    {
                                        //Нужно проверить тип объекта
                                        if (obj.UnitsType == Db.DynamicBlockReferencePropertyUnitsType.Angular
                                            || obj.UnitsType == Db.DynamicBlockReferencePropertyUnitsType.Distance
                                            || obj.UnitsType == Db.DynamicBlockReferencePropertyUnitsType.Area)
                                        {
                                            var d = double.Parse(i.Value.ToString());
                                            //Обновляем только в том случае если были изменения
                                            //Как то коряво
                                            if (obj.Value != d as object)
                                                obj.Value = d;
                                        }

                                        //http://adn-cis.org/chtenie-tabliczyi-svojstv-bloka-dlya-dinamicheskogo-bloka.html
                                        //Тут можно посмотреть наименование свойств в таблице   

                                        //http://adn-cis.org/forum/index.php?topic=603.msg2033#msg2033
                                        if (obj.UnitsType == Db.DynamicBlockReferencePropertyUnitsType.NoUnits)
                                        {

                                            object d = null;// = new object();

                                            switch (obj.PropertyTypeCode)
                                            {
                                                case (short)DwgDataType.kDwgNull: //0
                                                    break;
                                                case (short)DwgDataType.kDwgReal: //1
                                                    d = double.Parse(i.Value.ToString());
                                                    break;                          //return true;
                                                case (short)DwgDataType.kDwgInt32:  //2
                                                    d = int.Parse(i.Value.ToString());
                                                    break;                          //return true;
                                                case (short)DwgDataType.kDwgInt16:  //3 
                                                                                    //Flip state
                                                                                    //Block Properties Table

                                                    //Отрицательное значение - вроде как выставленное по умолчанию...
                                                    // и отрицательное значение не присвоить Block Properties Table
                                                    //TODO , а как с Flip state еще не тестировал.
                                                    var j = short.Parse(i.Value.ToString());
                                                    if (j > 0) d = j;

                                                    break;
                                                case (short)DwgDataType.kDwgInt8: //4
                                                    //Возможно так... Int8
                                                    d = short.Parse(i.Value.ToString());
                                                    break;
                                                case (short)DwgDataType.kDwgText: //5  
                                                    //Lookup
                                                    d = i.Value.ToString();
                                                    break;
                                                case (short)DwgDataType.kDwgBChunk: //6
                                                    break;
                                                case (short)DwgDataType.kDwgHandle: //7
                                                    d = long.Parse(i.Value.ToString());
                                                    break;
                                                case (short)DwgDataType.kDwgHardOwnershipId: //8
                                                    break;
                                                case (short)DwgDataType.kDwgSoftOwnershipId: //9
                                                    break;
                                                case (short)DwgDataType.kDwgHardPointerId: //12 
                                                    //Origin (double[2])
                                                    break;
                                                case (short)DwgDataType.kDwgSoftPointerId: //11
                                                    break;
                                                case (short)DwgDataType.kDwg3Real: //12
                                                    break;
                                                case (short)DwgDataType.kDwgInt64: //13
                                                    d = Int64.Parse(i.Value.ToString());
                                                    break;

                                                case (short)DwgDataType.kDwgNotRecognized: //19
                                                    break;
                                                default:
                                                    throw new InvalidCastException("You can't cast a weird value!");
                                            }

                                            if (d != null && obj.Value != d && !obj.ReadOnly)
                                            {
                                                obj.Value = d;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        //обновляем атрибуты
                        //blr.AttSync(true, false, false);
                        //маркеруем блок, как блок с измененный графикой
                        acBlRef.RecordGraphicsModified(true);
                        pm.MeterProgress();
                    }
                    pm.Stop();
                    acTrans.Commit();
                }

                //обновляем атрибуты
                using (var tr = acCurDb.TransactionManager.StartTransaction())
                {
                    //прогресс бар
                    var pm = new Rtm.ProgressMeter();
                    pm.Start("Progress of processing update AttributeReference");
                    pm.SetLimit(propertyList.Count);

                    var acBlkTbl = tr.GetObject(acCurDb.BlockTableId, Db.OpenMode.ForRead) as Db.BlockTable;

                    var listBlockName = new List<string>();
                    foreach (var i in propertyList)
                        if (!listBlockName.Contains(i.BlockName)) listBlockName.Add(i.BlockName);

                    foreach (var i in listBlockName)
                    {
                        var acBlkTblRec = tr.GetObject(acBlkTbl[i], Db.OpenMode.ForRead) as Db.BlockTableRecord;
                        acBlkTblRec.AttSync(true, false, false);
                        pm.MeterProgress();
                    }
                    pm.Stop();
                    tr.Commit();
                }
            }

            //5. Оповещаем пользователя о завершении работы
            //Перерисовать графику
            //http://adn-cis.org/forum/index.php?topic=8361.0
            acDoc.TransactionManager.FlushGraphics();
            acEd.WriteMessage($"\nDone.");
        }
    }
}