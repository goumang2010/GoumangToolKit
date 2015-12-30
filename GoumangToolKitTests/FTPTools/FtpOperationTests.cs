using Microsoft.VisualStudio.TestTools.UnitTesting;
using GoumangToolKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoumangToolKit.Tests
{
    [TestClass()]
    public class FtpOperationTests
    {
        [TestMethod()]
        public void DownloadTest()
        {
            FtpOperation cc = new FtpOperation();
            string err = "";
           var flag= cc.Download(@"E:\\C01317361-gg.CATPart", "ftp://61.161.226.2//MID-FUSELAGE/ENG DATA SET-UPDATED/C01301038-003 --.CATPart", out err);


            Assert.AreEqual(flag,true);
        }
    }
}