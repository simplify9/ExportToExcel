using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace SW.ExportToExcel.UnitTests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
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

            var bytes = mylist.ExportToExcel(); 

        }

        private class Employee
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public DateTime DoB { get; set; }
        }
    }
}
