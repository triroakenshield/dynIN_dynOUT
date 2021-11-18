using System;
using System.Collections.Generic;
using System.Linq;

using Db = Autodesk.AutoCAD.DatabaseServices;
using Gem = Autodesk.AutoCAD.Geometry;

namespace dynIN_dynOUT
{
    enum DwgDataType
    {
        kDwgNull = 0,
        kDwgReal = 1,
        kDwgInt32 = 2,
        kDwgInt16 = 3,
        kDwgInt8 = 4,
        kDwgText = 5,
        kDwgBChunk = 6,
        kDwgHandle = 7,
        kDwgHardOwnershipId = 8,
        kDwgSoftOwnershipId = 9,
        kDwgHardPointerId = 10,
        kDwgSoftPointerId = 11,
        kDwg3Real = 12,
        kDwgInt64 = 13,
        kDwgNotRecognized = 19
    };

    public class Property
    {
        //Field
        public Db.Handle Handle;
        public Dictionary<string, string> Attribut = new Dictionary<string, string>();
        public Dictionary<string, object> DynProp = new Dictionary<string, object>();
        public string BlockName = "";

        //Property
        public Gem.Point3d Position { get; set; }
        public double Rotation { get; set; }
        public Gem.Scale3d ScaleFactors { get; set; }
        public string Layer { get; set; }
        public int ColorIndex { get; set; }

        public Property()
        {

        }

        public Property(Db.Handle handle)
        {
            Handle = handle;
        }

        public bool Sets(List<string> rowHead, string strLines)
        {
            var l = strLines.Split(';').ToList();

            //Попробуем спарсить хендл , если не получится, то запишем эту строку как имя блока
            long h=0;
            var tryConvertToLong = long.TryParse(l[0].Replace("\'", ""), out h);
            if (tryConvertToLong)
            {
                Handle = new Db.Handle(h);
            }
            else
            {
                BlockName = l[0].Replace("\'", "");
                Handle = new Db.Handle();
            }

            for (var j = 1; j < l.Count; j++)
            {
                if (l[j] != "")
                {
                    var NameProp = rowHead.ElementAt(j);
                    if (NameProp.Substring(0, 2) == "a_")
                        Attribut.Add(NameProp.Substring(2, NameProp.Length - 2), l[j]);

                    if (NameProp.Substring(0, 2) == "d_")
                        DynProp.Add(NameProp.Substring(2, NameProp.Length - 2), l[j]);

                    if (NameProp.Substring(0, 2) == "p_")
                    {
                        var name = NameProp.Substring(2, NameProp.Length - 2);

                        if (name == "Rotation") Rotation = double.Parse(l[j]);

                        if (name == "Layer")
                        {
                            //Проверяем, есть ли такой слой, если нет и создание слоев запрещено, то пишем "0"
                            if (!AddEntity.CreateLayer(l[j], false) && !Settings.Data.CreateLayer)
                            {
                                Layer = "0";
                            }
                            else
                            {
                                Layer = l[j];
                            }
                        }

                        if (name == "ColorIndex") ColorIndex = int.Parse(l[j]);

                        if (name == "Position")
                        {
                            //var str = (l[j].Substring(1, l[j].Length - 2)).Split(',').ToList();
                            //List<double> dbl = new List<double>();
                            //foreach (var t in str)
                            //    dbl.Add(double.Parse(t));

                            var dbl = (l[j].Substring(1, l[j].Length - 2)).Split(',').ToList()
                                .ConvertAll<double>(delegate (string i) { return double.Parse(i); });

                            Position =  new Gem.Point3d(dbl.ToArray());
                        }

                        if (name == "ScaleFactors")
                        {
                            //var str = (l[j].Substring(1, l[j].Length - 2)).Split(',').ToList();
                            //List<double> dbl = new List<double>();
                            //foreach (var t in str)
                            //    dbl.Add(double.Parse(t));

                            var dbl = (l[j].Substring(1, l[j].Length - 2)).Split(',').ToList()
                                .ConvertAll<double>(delegate (string i) { return double.Parse(i); });

                            ScaleFactors = new Gem.Scale3d(dbl.ToArray());
                        }
                    }
                }
            }
            return true;
        }

        public string[] Gets(List<string> rowHead)
        {
            //Создаем массив длинной , равной длинне заголовка
            var row = new string[rowHead.Count];
            //Все ячейки массива заполняем по умолчанию табуляциями
            for (var i = 0; i < row.Length; i++)
                row[i] = "";

            //В первую ячейку массива пишу хендл объекта
            row[0] = $"\'{Handle.Value.ToString()}";

            foreach (var i in Attribut)
            {
                var indxUnicAttName = rowHead.FindIndex(x => x == "a_" + i.Key);
                row[indxUnicAttName] = i.Value;
            }


            foreach (var i in DynProp)
            {
                var indxUnicDynName = rowHead.FindIndex(x => x == "d_" + i.Key);
                row[indxUnicDynName] = i.Value.ToString();
            }

            foreach (var prop in GetType().GetProperties())
            {
                var indxUnicPropName = rowHead.FindIndex(x => x == "p_" + prop.Name);

                var val1 = prop.GetValue(this, null).ToString();

                row[indxUnicPropName] = val1;
            }

            return row;
        }

        //public override string ToString()
        //{
        //    return string.Format("privateField: {0}, publicField: {1}");
        //}
    }

    public static class StringExtensions
    {
        public static bool ToBoolean(this string value)
        {
            switch (value.ToLower())
            {
                case "true":
                    return true;
                case "t":
                    return true;
                case "1":
                    return true;
                case "0":
                    return false;
                case "false":
                    return false;
                case "f":
                    return false;
                default:
                    throw new InvalidCastException("You can't cast a weird value to a bool!");
            }
        }
    }
}