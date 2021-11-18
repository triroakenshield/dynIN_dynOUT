using System;
using System.Collections.Specialized;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices.Core;
using App = Autodesk.AutoCAD.ApplicationServices;
using Db = Autodesk.AutoCAD.DatabaseServices;
using Ed = Autodesk.AutoCAD.EditorInput;
using Rtm = Autodesk.AutoCAD.Runtime;

[assembly: Rtm.CommandClass(typeof(dynIN_dynOUT.Commands))]

namespace dynIN_dynOUT
{
    public class Commands : Rtm.IExtensionApplication
    {
        /// <summary>
        /// Загрузка библиотеки
        /// http://through-the-interface.typepad.com/through_the_interface/2007/03/getting_the_lis.html
        /// </summary>
        #region 
        public void Initialize()
        {
            var assemblyFileFullName = GetType().Assembly.Location;
            var assemblyName = System.IO.Path.GetFileName(GetType().Assembly.Location);

            // Just get the commands for this assembly
            var dm = Application.DocumentManager;
            var asm = Assembly.GetExecutingAssembly();
            var acEd = dm.MdiActiveDocument.Editor;

            // Сообщаю о том, что произведена загрузка сборки 
            //и указываю полное имя файла,
            // дабы было видно, откуда она загружена
            acEd.WriteMessage(string.Format("\n{0} {1} {2}.\n{3}: {4}\n{5}\n",
                      "Assembly", assemblyName, "Loaded",
                      "Assembly File:", assemblyFileFullName,
                       "Copyright © Владимир Шульжицкий, 2018"));

            //Вывожу список комманд определенных в библиотеке
            acEd.WriteMessage("\nStart list of commands: \n\n");

            var cmds = GetCommands(asm, false);
            foreach (var cmd in cmds)
                acEd.WriteMessage(cmd + "\n");

            acEd.WriteMessage("\n\nEnd list of commands.\n");
        }

        public void Terminate()
        {
            Console.WriteLine("finish!");
        }

        /// <summary>
        /// Получение списка команд определенных в сборке
        /// </summary>
        /// <param name="asm"></param>
        /// <param name="markedOnly"></param>
        /// <returns></returns>
        private static string[] GetCommands(Assembly asm, bool markedOnly)
        {
            var sc = new StringCollection();
            var objs = asm.GetCustomAttributes(typeof(Rtm.CommandClassAttribute), true);
            Type[] tps;
            var numTypes = objs.Length;
            if (numTypes > 0)
            {
                tps = new Type[numTypes];
                for (var i = 0; i < numTypes; i++)
                {
                    if (objs[i] is Rtm.CommandClassAttribute cca)
                    {
                        tps[i] = cca.Type;
                    }
                }
            }
            else
            {
                // If we're only looking for specifically
                // marked CommandClasses, then use an
                // empty list
                if (markedOnly)
                    tps = new Type[0];
                else
                    tps = asm.GetExportedTypes();
            }
            foreach (var tp in tps)
            {
                var meths = tp.GetMethods();
                foreach (var meth in meths)
                {
                    objs = meth.GetCustomAttributes(typeof(Rtm.CommandMethodAttribute), true);
                    foreach (var obj in objs)
                    {
                        var attb = (Rtm.CommandMethodAttribute)obj;
                        sc.Add(attb.GlobalName);
                    }
                }
            }
            var ret = new string[sc.Count];
            sc.CopyTo(ret, 0);
            return ret;
        }
        #endregion

        [Rtm.CommandMethod("dynIN")]
        public static void dynIN()
        {
            //Читаем данные из txt файла
            DynIN.IN();
        }

        [Rtm.CommandMethod("dynOUT")]
        public static void dynOUT()
        {
            //Сохраняем данные в txt файл
            DynOUT.OUT();
        }

        [Rtm.CommandMethod("pp")]
        public static void dynSET()
        {
            //Сохраняем данные в txt файл
            //DynSET.OUT();
        }

        [Rtm.CommandMethod("GetAllDynamicBlockParameters")]
        public void GetAllDynamicBlockParameters()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var editor = doc.Editor;
            var option = new Ed.PromptEntityOptions("\n" + "Select a block");
            var result = editor.GetEntity(option);
            if (result.Status == Ed.PromptStatus.OK)
            {
                var id = result.ObjectId;
                using (var trans = db.TransactionManager.StartTransaction())
                {
                    var blockRef = (Db.BlockReference)trans.GetObject(id, Db.OpenMode.ForWrite);
                    var properties = blockRef.DynamicBlockReferencePropertyCollection;

                    for (var i = 0; i < properties.Count; i++)
                    {
                        var property = properties[i];

                        editor.WriteMessage($"\n{property.PropertyName} | {property.PropertyTypeCode} | {property.Value}");
                        //property.Value = (double)25;
                    }
                }
            }
        }
    }
}