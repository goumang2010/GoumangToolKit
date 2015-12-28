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
    public class historyJsonModelTests
    {
        [TestMethod()]
        public void AppendToFileTest()
        {
            historyJsonModel hm = new historyJsonModel()
            {
                writer = "SB",
                date = DateTime.Now.ToShortDateString(),
                operation = "Done"

            };
            var dd = hm.AppendToFile(@"F:\Prepare\SOFTWARE_PUB\Autorivet_team_manage\settings\PartDBRoutine");
            Assert.AreEqual(dd, true);
        }

        [TestMethod()]
        public void ReadFromFileTest()
        {
            var dd = jsonMethod.ReadFromFile(@"F:\Prepare\SOFTWARE_PUB\Autorivet_team_manage\settings\PartDBRoutine");
            var count = dd.Count();
            Assert.AreEqual(count, 2);
        }
    }
}