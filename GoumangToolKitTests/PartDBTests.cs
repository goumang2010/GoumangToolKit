using Microsoft.VisualStudio.TestTools.UnitTesting;
using AutorivetService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GoumangToolKit.NET4._6;
using GoumangToolKit;

namespace AutorivetService.Tests
{
    [TestClass()]
    public class PartDBTests
    {
        [TestMethod()]
        public void RunWorkTest()
        {
            string routinepath = @"F:\Prepare\SOFTWARE_PUB\Autorivet_team_manage\settings\PartDBRoutine";
            //检查配置文件
            var PartRoutine = jsonMethod.ReadFromFile(routinepath);
            var lastrocord = PartRoutine.Last();
            var lastDate = DateTime.Parse(lastrocord.date);
            var span = DateTime.Now.Date - lastDate;
            //If the span is so long, then update the database
            if (span.Days > 7)
            {
                //Get the update path
                var up = GoumangToolKit.localMethod.GetConfigValue("UPDATE_PATH", "PartDBCfg.py");
                PartDB pd = new PartDB("filepath");
                foreach (dynamic pp in up)
                {
                    var kk= pd.UpdatePartDB((string)pp);
                   
                }
            }

            PartRoutine.Add(new historyJsonModel()
            {
                writer="System",
                date=DateTime.Now.ToShortDateString(),
                operation="Done"

            });
           var rr= PartRoutine.WriteToFile(routinepath);
            Assert.AreEqual(rr, true);
        }
    }
}