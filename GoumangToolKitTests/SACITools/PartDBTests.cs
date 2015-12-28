using GoumangToolKit;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using GoumangToolKit.NET4._6;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoumangToolKit.Tests
{
    [TestClass()]
    public class PartDBTests
    {
        [TestMethod()]
        public void UpdateFTPPartDBTest()
        {

            PartDB pd = new PartDB("FTPfilePath");
            var dd = pd.UpdateFTPPartDB("/");
            //var dd = pd.UpdatePartDB(new List<string>() { @"\\192.168.2.10\Cseries-data1" });
            Assert.AreEqual(dd.Result, true);
        }
    }
}

namespace GoumangToolKit.NET4._6.Tests
{
    [TestClass()]
    public class PartDBTests
    {

        [TestMethod()]
        public void UpdatePartDBTest()
        {
            PartDB pd = new PartDB("filePath");
            //var dd= pd.UpdatePartDB(@"");
            var dd = pd.UpdatePartDB( @"\\192.168.2.10\Cseries-data1" );
            Assert.AreEqual(dd.Result, true);
        }
    }
}