using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Buffers.Text;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace SW.ExportToExcel.UnitTests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public async Task TestMethod1()
        {
            var mylist = new List<Employee>
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

        [TestMethod]
        public async Task TestMethod2()
        {
            var dict = new List<Dictionary<string, string>>
            {
                new Dictionary<string, string>
                {
                    ["name"] = "Wheeb",
                    ["age"] = "20"
                },
                new Dictionary<string, string>
                {
                    ["name"] = "Joe",
                    ["age"] = "42"
                }
            };
            var stream = new MemoryStream();
            await dict.WriteExcel(stream);
            Assert.IsNotNull(stream);
            //  var b = Convert.ToBase64String(stream.ToArray());
            // Console.WriteLine(b);
        }

        private class Employee
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public DateTime DoB { get; set; }
        }
    }
}