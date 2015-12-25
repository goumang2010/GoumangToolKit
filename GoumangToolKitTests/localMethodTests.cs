using Microsoft.VisualStudio.TestTools.UnitTesting;
using mysqlsolution;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace mysqlsolution.Tests
{
    [TestClass()]
    public class localMethodTests
    {
        [TestMethod()]
        public void GetConfigValueTest()
        {
          dynamic pp = (localMethod.GetConfigValue("MONGO_URI", "DBCfg.py"));
            Assert.AreEqual(pp, "mongodb://autorivet:222222@127.0.0.1/SACI");
        }
    }
}