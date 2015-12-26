using Microsoft.VisualStudio.TestTools.UnitTesting;
using GoumangToolKit.NET4._6;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoumangToolKit.NET4._6.Tests
{
    [TestClass()]
    public class PartDBTests
    {

        [TestMethod()]
        public void UpdatePartDBTest()
        {
            PartDB pd = new PartDB();
            var dd= pd.UpdatePartDB(new List<string>() { @"F:\temp\AM1110_PVR" });
            //var dd = pd.UpdatePartDB(new List<string>() { @"\\192.168.2.10\Cseries-data1" });
            Assert.AreEqual(dd.Result, true);
        }
    }
}