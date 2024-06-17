using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SKToolsAddins.Utils
{
    public class ExcelRangeNameParseHelper
    {
        public enum ObjectType
        {
            Reset,
            IchijiKairyou,
            NijiKairyou,
            GaikouKairyou,
            Foundation,
            HashiraGata,
            Framing,
            Mashiuchi,
            Umemodoshi,
            Saiseki,
            Mabashira,
            DomaExca,
            Doma,
            DomaDansa,
            Pit
        }
        public abstract class RangeObject
        {
            public ObjectType ObjectType { get; protected set; }
            public abstract bool Parse(params string[] parameters);
        }
        public static RangeObject Parse(string rangeName)
        {
            if (rangeName.ToLower() == "reset")
            {
                RngObjReset rngObjReset = new RngObjReset();
                return rngObjReset;
            }
            var rangeNameNoSheetArr = rangeName.Split('!');
            string rangeNameNoSheet = "";
            if (rangeNameNoSheetArr.Count() == 2)
            {
                rangeNameNoSheet = rangeNameNoSheetArr[1];
            }
            else
            {
                rangeNameNoSheet = rangeNameNoSheetArr[0];
            }
            var arr = rangeNameNoSheet.Split('_');
            switch (arr[1].ToLower())
            {
                case "ichijikairyou":
                    RngObjIchijiKairyou rngObjIchijiKairyou = new RngObjIchijiKairyou();
                    rngObjIchijiKairyou.Parse(arr[2]);
                    return rngObjIchijiKairyou;
                case "nijikairyou":
                    RngObjNijiKairyou rngObjNijiKairyou = new RngObjNijiKairyou();
                    rngObjNijiKairyou.Parse(arr[2]);
                    return rngObjNijiKairyou;
                case "gaikoukairyou":
                    RngObjGaikouKairyou rngObjGaikouKairyou = new RngObjGaikouKairyou();
                    rngObjGaikouKairyou.Parse(arr[2]);
                    return rngObjGaikouKairyou;
                case "foundation":
                    RngObjFoundation rngObjFoundation = new RngObjFoundation();
                    rngObjFoundation.Parse(arr[2]);
                    return rngObjFoundation;
                case "hashiragata":
                    RngObjHashiraGata rngObjHashiraGata = new RngObjHashiraGata();
                    var newArr = arr.Where((source, index) => (index != 0) && (index != 1)).ToArray();
                    rngObjHashiraGata.Parse(newArr);
                    return rngObjHashiraGata;
                case "framing":
                    RngObjFraming rngObjFraming = new RngObjFraming();
                    rngObjFraming.Parse(arr[2]);
                    return rngObjFraming;
                case "mashiuchi":
                    RngObjMashiuchi rngObjMashiuchi = new RngObjMashiuchi();
                    rngObjMashiuchi.Parse(arr[2], arr[3]);
                    return rngObjMashiuchi;
                case "umemodoshi":
                    RngObjUmemodoshi rngObjUmemodoshi = new RngObjUmemodoshi();
                    rngObjUmemodoshi.Parse(arr[2], arr[3]);
                    return rngObjUmemodoshi;
                case "saiseki":
                    RngObjSaiseki rngObjSaiseki = new RngObjSaiseki();
                    rngObjSaiseki.Parse(arr[2]);
                    return rngObjSaiseki;
                case "mabashira":
                    RngObjMabashira rngObjMabashira = new RngObjMabashira();
                    rngObjMabashira.Parse(arr[2]);
                    return rngObjMabashira;
                case "domaexca":
                    RngObjDomaExca rngObjDomaExca = new RngObjDomaExca();
                    rngObjDomaExca.Parse(arr[2]);
                    return rngObjDomaExca;
                case "doma":
                    RngObjDoma rngObjDoma = new RngObjDoma();
                    rngObjDoma.Parse(arr[2]);
                    return rngObjDoma;
                case "domadansa":
                    RngObjDomaDansa rngObjDomaDansa = new RngObjDomaDansa();
                    rngObjDomaDansa.Parse(arr[2]);
                    return rngObjDomaDansa;
                case "pit":
                    RngObjPit rngObjPit = new RngObjPit();
                    rngObjPit.Parse(arr[2]);
                    return rngObjPit;
                default:
                    break;
            }
            return null;
        }
        public class RngObjReset : RangeObject
        {
            public RngObjReset()
            {
                ObjectType = ObjectType.Reset;
            }
            public override bool Parse(params string[] parameters)
            {
                throw new NotImplementedException();
            }
        }
        public class RngObjIchijiKairyou : RangeObject
        {
            public string Thickness { get; set; }
            public RngObjIchijiKairyou()
            {
                ObjectType = ObjectType.IchijiKairyou;
            }
            public override bool Parse(params string[] parameters)
            {
                Thickness = parameters[0].Trim();
                return true;
            }
        }
        public class RngObjNijiKairyou : RangeObject
        {
            public string GlOffset { get; set; }
            public RngObjNijiKairyou()
            {
                ObjectType = ObjectType.NijiKairyou;
            }
            public override bool Parse(params string[] parameters)
            {
                GlOffset = parameters[0].Trim();
                return true;
            }
        }
        public class RngObjGaikouKairyou : RangeObject
        {
            public string Volume { get; set; }
            public RngObjGaikouKairyou()
            {
                ObjectType = ObjectType.GaikouKairyou;
            }
            public override bool Parse(params string[] parameters)
            {
                Volume = parameters[0].Trim();
                return true;
            }
        }
        public class RngObjFoundation : RangeObject
        {
            public string TypeMark { get; set; }
            public RngObjFoundation()
            {
                ObjectType = ObjectType.Foundation;
            }
            public override bool Parse(params string[] parameters)
            {
                TypeMark = parameters[0].Trim();
                return true;
            }
        }
        public class RngObjHashiraGata : RangeObject
        {
            public List<string> TypeMarkList { get; set; }
            public RngObjHashiraGata()
            {
                ObjectType = ObjectType.HashiraGata;
            }
            public override bool Parse(params string[] parameters)
            {
                TypeMarkList = parameters.ToList();
                return true;
            }
        }
        public class RngObjFraming : RangeObject
        {
            public string TypeMark { get; set; }
            public RngObjFraming()
            {
                ObjectType = ObjectType.Framing;
            }
            public override bool Parse(params string[] parameters)
            {
                TypeMark = parameters[0].Trim();
                return true;
            }
        }
        public class RngObjMashiuchi : RangeObject
        {
            public string Host { get; set; }
            public string Size { get; set; }
            public RngObjMashiuchi()
            {
                ObjectType = ObjectType.Mashiuchi;
            }
            public override bool Parse(params string[] parameters)
            {
                Host = parameters[0].Trim();
                Size = parameters[1].Trim();
                return true;
            }
        }
        public class RngObjUmemodoshi : RangeObject
        {
            public string Host { get; set; }
            public string Size { get; set; }
            public RngObjUmemodoshi()
            {
                ObjectType = ObjectType.Umemodoshi;
            }
            public override bool Parse(params string[] parameters)
            {
                Host = parameters[0].Trim();
                Size = parameters[1].Trim();
                return true;
            }
        }
        public class RngObjSaiseki : RangeObject
        {
            public string Volume { get; set; }
            public RngObjSaiseki()
            {
                ObjectType = ObjectType.Saiseki;
            }
            public override bool Parse(params string[] parameters)
            {
                Volume = parameters[0].Trim();
                return true;
            }
        }
        public class RngObjMabashira : RangeObject
        {
            public string TypeName { get; set; }
            public RngObjMabashira()
            {
                ObjectType = ObjectType.Mabashira;
            }
            public override bool Parse(params string[] parameters)
            {
                TypeName = parameters[0].Trim();
                return true;
            }
        }
        public class RngObjDomaExca : RangeObject
        {
            public string Size { get; set; }
            public RngObjDomaExca()
            {
                ObjectType = ObjectType.DomaExca;
            }
            public override bool Parse(params string[] parameters)
            {
                Size = parameters[0].Trim();
                return true;
            }
        }
        public class RngObjDoma : RangeObject
        {
            public string Size { get; set; }
            public RngObjDoma()
            {
                ObjectType = ObjectType.Doma;
            }
            public override bool Parse(params string[] parameters)
            {
                Size = parameters[0].Trim();
                return true;
            }
        }
        public class RngObjDomaDansa : RangeObject
        {
            public string Size { get; set; }
            public RngObjDomaDansa()
            {
                ObjectType = ObjectType.DomaDansa;
            }
            public override bool Parse(params string[] parameters)
            {
                Size = parameters[0].Trim();
                return true;
            }
        }
        public class RngObjPit : RangeObject
        {
            public string Name { get; set; }
            public RngObjPit()
            {
                ObjectType = ObjectType.Pit;
            }
            public override bool Parse(params string[] parameters)
            {
                Name = parameters[0].Trim();
                return true;
            }
        }
    }
}
