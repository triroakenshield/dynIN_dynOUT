using dynINOUT_UI;

using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using cad = Autodesk.AutoCAD.ApplicationServices.Application;
using Db = Autodesk.AutoCAD.DatabaseServices;
using Ed = Autodesk.AutoCAD.EditorInput;
using Rtm = Autodesk.AutoCAD.Runtime;

namespace dynIN_dynOUT
{
    /// <summary>Сохраняем данные в txt файл</summary>
    internal static class DynOUT
    {
        //Публичная переменная в которой храниться список имен блоков
        internal static ObservableCollection<string> _blockNameList = new ObservableCollection<string>();

        internal static void OUT()
        {
            //string statrPath = Settings.Data.Lastpath ;

            //1. Подражая Attout сначала выбираем файл в который будет сохраняться информация
            // TODO по умолчанию имя нового файла должно соответствовать имени чертежа

            //SaveFileDialog openFileDialog = new SaveFileDialog("Выберите CSV файл",
            //                                          "*.csv",
            //                                          "csv",
            //                                          "Выбор файла",
            //                                          SaveFileDialog.SaveFileDialogFlags.NoUrls);
            //if (openFileDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
            //string fileName = openFileDialog.Filename;

            var openFileDialog1 = new SaveFileDialog
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

            // Получение текущего документа и базы данных
            var acDoc = Application.DocumentManager.MdiActiveDocument;
            if (acDoc == null) return;
            var acCurDb = acDoc.Database;
            var acEd = acDoc.Editor;

            //2. Запрашиваем у пользователя выбор блоков
            var tv = new Db.TypedValue[4];
            tv.SetValue(new Db.TypedValue((int)Db.DxfCode.Operator, "<AND"), 0);
            //Фильтр по типу объекта
            tv.SetValue(new Db.TypedValue((int)Db.DxfCode.Start, "INSERT"), 1);
            //Фильтр по имени слоя
            tv.SetValue(new Db.TypedValue((int)Db.DxfCode.LayerName, "*"), 2);
            tv.SetValue(new Db.TypedValue((int)Db.DxfCode.Operator, "AND>"), 3);

            var sf = new Ed.SelectionFilter(tv);

            var psr = acEd.GetSelection(sf);
            if (psr.Status != Ed.PromptStatus.OK) return;

            //Сюда нужно встроить фильтрацию по имени блока
            _blockNameList.Clear(); // чистим хранилище имен блоков
            using (var docloc = acDoc.LockDocument())
            {
                foreach (Ed.SelectedObject acSSObj in psr.Value)
                {
                    if (!acSSObj.ObjectId.IsNull && acSSObj.ObjectId.IsResident && acSSObj.ObjectId.IsValid && !acSSObj.ObjectId.IsErased)
                        if (acSSObj.ObjectId.ObjectClass.IsDerivedFrom(Rtm.RXObject.GetClass(typeof(Db.BlockReference))))
                        {
                            var blockName = "";
                            using (var acBlRef = acSSObj.ObjectId.Open(Db.OpenMode.ForRead) as Db.BlockReference)
                            {
                                //TODO По хорошему этот кусок кода надо бы вынести в экстеншен метод класса BlockReference
                                blockName = acBlRef.EffectiveName();
                                acBlRef.Close();
                            }
                            if (!_blockNameList.Contains(blockName)) _blockNameList.Add(blockName);
                        }
                }
            }

            //Тут показываем пользователю окошко с выбором блоков по именам
            if (_blockNameList.Count > 1)
            {
                var dlg = new MainWindow();
                dlg.AddBlockNameList(_blockNameList);
                cad.ShowModalWindow(dlg);

                _blockNameList.Clear();

                foreach (var i in dlg.BindingList)
                    if (i.Value) _blockNameList.Add(i.Key);
            }

            //3. Проходимся по выбранным блокам и собираем информацию
            var propertyList = new List<Property>();

            //Блокируем документ
            using (var docLock = acDoc.LockDocument())
            {
                using (Db.Transaction acTrans = acCurDb.TransactionManager.StartOpenCloseTransaction())
                {
                    var acSSet = psr.Value;

                    foreach (Ed.SelectedObject acSSObj in acSSet)
                    {
                        if (acSSObj != null)
                        {
                            var acEnt = acTrans.GetObject(acSSObj.ObjectId,
                                                     Db.OpenMode.ForRead) as Db.Entity;
                            if (acEnt != null)
                            {
                                if (acEnt is Db.BlockReference acBlRef)
                                {
                                    var blr = (Db.BlockTableRecord)acTrans.GetObject(acBlRef.DynamicBlockTableRecord,
                                                                                                    Db.OpenMode.ForRead);

                                    //Фильтр по именам блоков
                                    if (_blockNameList.Contains(acBlRef.EffectiveName()))
                                    {
                                        var prop = new Property(acBlRef.Handle);


                                        if (blr.HasAttributeDefinitions)
                                        {
                                            var attrCol = acBlRef.AttributeCollection;
                                            if (attrCol.Count > 0)
                                            {
                                                foreach (Db.ObjectId AttID in attrCol)
                                                {
                                                    var acAttRef = acTrans.GetObject(AttID,
                                                                            Db.OpenMode.ForRead) as Db.AttributeReference;

                                                    //TODO Необходимо проверить и учесть наличие полей
                                                    if (!prop.Attribut.ContainsKey(acAttRef.Tag))
                                                        prop.Attribut.Add(acAttRef.Tag, acAttRef.TextString);
                                                    else
                                                        acEd.WriteMessage($"\nВ блоке {blr.Name}->{acBlRef.Name} присутствуют атрибуты с одинаковыми тегами");

                                                }
                                            }   //Проверка что кол атрибутов больше 0
                                        }  //Проверка наличия атрибутов  

                                        var acBlockDynProp = acBlRef.DynamicBlockReferencePropertyCollection;
                                        if (acBlockDynProp != null)
                                        {
                                            foreach (Db.DynamicBlockReferenceProperty obj in acBlockDynProp)
                                            {
                                                //TODO а вот тут вопрос, нужно ли выводить значения, которые только ReadOnly
                                                if (obj.PropertyName != "Origin")
                                                {
                                                    if (!prop.DynProp.ContainsKey(obj.PropertyName))

                                                        prop.DynProp.Add(obj.PropertyName, obj.Value);
                                                    else
                                                        acEd.WriteMessage($"\nВ блоке {blr.Name}->{acBlRef.Name} присутствуют динамические свойства с одинаковыми именами");
                                                }
                                            }
                                        }

                                        //http://adndevblog.typepad.com/autocad/2012/05/comparing-properties-of-two-entities.html
                                        var propsBlockRef = acBlRef.GetType().GetProperties();
                                        var propElement = prop.GetType().GetProperties();

                                        foreach (var propInfo in propElement)
                                        {
                                            try
                                            {
                                                var propBlock = propsBlockRef
                                                    .Where(x => x.Name == propInfo.Name)
                                                    .FirstOrDefault();
                                                if (propBlock != null) 
                                                    propInfo.SetValue(prop, propBlock.GetValue(acBlRef, null), null);

                                            }
                                            catch (Exception ex)
                                            {

                                            }
                                        }
                                        propertyList.Add(prop);
                                    }
                                }   //Проверка, что объект это ссылка на блок
                            }
                        }
                    }
                    acTrans.Commit();
                }
            }

            //4. Формируем одну большую таблицу с данными
            //4.1 Считаем общее количество уникальны тегов атрибутов и уникальных названй динамических свойств
            var unicAttName = new List<string>();
            var unicDynName = new List<string>();
            var unicPropnName = new List<string>();

            foreach (var s in propertyList)
            {
                foreach (var i in s.Attribut)
                    if (!unicAttName.Contains("a_" + i.Key)) unicAttName.Add("a_" + i.Key);

                foreach (var i in s.DynProp)
                    if (!unicDynName.Contains("d_" + i.Key)) unicDynName.Add("d_" + i.Key);

                foreach (var i in s.GetType().GetProperties())
                    if (!unicPropnName.Contains("p_" + i.Name)) unicPropnName.Add("p_" + i.Name);
            }

            //4.2 Заполняем массив
            var rowList = new List<string[]>();

            var rowHead = new List<string>();
            rowHead.Add("Handle");
            rowHead.AddRange(unicAttName);
            rowHead.AddRange(unicDynName);
            rowHead.AddRange(unicPropnName);

            rowList.Add(rowHead.ToArray());

            foreach (var s in propertyList)
                rowList.Add(s.Gets(rowHead));

            //5. Выводим собранные данные в файл
            try
            {
                using (var sw = new StreamWriter(fileName, false, System.Text.Encoding.GetEncoding(1251)))
                {
                    foreach (var s in rowList)
                    {
                        sw.WriteLine(String.Join(";", s));
                    }
                }

                //using (StreamWriter sw = new StreamWriter(fileName, true, System.Text.Encoding.Default))
                //{
                //    sw.WriteLine("Дозапись");
                //    sw.Write(4.5);
                //}
            }
            catch (Exception e)
            {
                Console.WriteLine($"\nDynOUT.Out-Ошибка записи в файл: {e.Message}");
            }

            //6. Оповещаем пользователя о завершении работы
            acEd.WriteMessage($"\nЭкспорт завершен.");
        }
    }
}