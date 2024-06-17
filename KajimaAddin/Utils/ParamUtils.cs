﻿using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SKToolsAddins.Utils
{
    public static class ParamUtils
    {
        public static dynamic GetParameterValue(this Parameter parameter)
        {
            if (parameter == null) return null;
            StorageType storageType = parameter.StorageType;
            switch (storageType)
            {
                case StorageType.None:
                    return parameter.AsString();
                case StorageType.Integer:
                    return parameter.AsInteger();
                case StorageType.Double:
                    return Math.Round(parameter.AsDouble(), 3);
                case StorageType.String:
                    return parameter.AsString();
                case StorageType.ElementId:
                    return parameter.AsElementId();
                default:
                    return parameter.AsString();
            }
        }

    }
}