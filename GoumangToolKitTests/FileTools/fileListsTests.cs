using Microsoft.VisualStudio.TestTools.UnitTesting;
using GoumangToolKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace GoumangToolKit.Tests
{
    [TestClass()]
    public class fileListsTests
    {
        [TestMethod()]
        public void WalkTreeTest()
        {
            var ff = new List<FileInfo>();
             ff.WalkTreeAsync(@"E:\Autorivet_program\question");
            var vv = new List<FileInfo>();
            vv.WalkTreeAsync(@"E:\Autorivet_program\question");
            Assert.AreEqual(ff.Count, vv.Count);

        }
    }
}