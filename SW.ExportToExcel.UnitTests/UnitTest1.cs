using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SW.ExportToExcel.UnitTests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        async public Task TestMethod1()
        {

            var mylist = new List<Employee>()
            {
                new Employee
                {
                    Name = "samer",
                    DoB = DateTime.UtcNow,
                    Id = 12
                },
                new Employee
                {
                    Name = "wael",
                    DoB = DateTime.UtcNow,
                    Id = 13
                }

            };

            var bytes = await mylist.ExportToExcel(); 

        }

        private class Employee
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public DateTime DoB { get; set; }
        }
    }
}
