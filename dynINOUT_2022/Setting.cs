using Autodesk.AutoCAD.ApplicationServices.Core;

using System;
using System.IO;
using System.Xml.Serialization;

namespace dynIN_dynOUT
{
    //тут нужно сделать синглтон!!!
    public class Settings
    {
        private static Sets _settings;

        public static Sets Data
        {
            //get { return _settings; }
            //set { _settings = value; }
            get
            {
                _settings = getParam();
                return _settings;
            }
            set
            {
                _settings = value;
                saveParam(_settings);
            }
        }

        private static Sets getParam()
        {
            // Set a variable to the My Documents path.
            var mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var path = mydocpath + Convert.ToString($"\\dynINOUTSetting.xml");

            Sets myObject;
            if (File.Exists(path))
            {
                try
                {
                    var mySerializer = new XmlSerializer(typeof(Sets));
                    using (var myFileStream = new FileStream(path, FileMode.Open))
                    {
                        myObject = (Sets)mySerializer.Deserialize(myFileStream);
                    }
                }
                catch (InvalidOperationException ex)
                {
                    File.Delete(path);//, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin);
                    myObject = new Sets(true, true, mydocpath);
                    saveParam(myObject);
                }
            }
            else
            {
                myObject = new Sets(true, true, mydocpath);
                saveParam(myObject);
            }
            return myObject;
        }

        private static void saveParam(Sets Setting)
        {
            try
            {
                var mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                var path = mydocpath + Convert.ToString($"\\dynINOUTSetting.xml");

                var ser = new XmlSerializer(typeof(Sets));
                using (TextWriter writer = new StreamWriter(path, false))
                {
                    ser.Serialize(writer, Setting);
                }
            }
            catch (FileNotFoundException ex)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nError: Settings-saveParam:{ex.Message}");
            }
        }
    }

    public class Sets
    {
        /// <summary>
        ///Создавать или нет слой, если его нет в базе чертежа
        /// </summary>
        public  bool CreateLayer { get; set; }

        /// <summary>
        /// ///Создавать или нет блок, если вместо хендла написано имя блока
        /// </summary>
        public  bool CreateBlocReference { get; set; }

        public string Lastpath { get; set; }

        public Sets() { }

        public Sets(bool createLayer, bool createBlocReference, string lastPath)
        {
            this.CreateLayer = createLayer;
            this.CreateBlocReference = createBlocReference;
            this.Lastpath = lastPath;
        }
    }
}