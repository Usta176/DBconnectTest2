using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSqlServerDb
{
    public class CDefinition
    {
        public class CTool
        {
            public Int32 Id;
            public string Name;
        }

        public class CToolCharacteristic
        {
            public Int32 Id;
            public string ToolName;
            public string Key;
            public string Value;
        }


    }
}
