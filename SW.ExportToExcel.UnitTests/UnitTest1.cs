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

            var mylist = new List<Employee>();

            // mylist.Ex

        }

        private class Employee
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public DateTime DoB { get; set; }
        }
    }
}
