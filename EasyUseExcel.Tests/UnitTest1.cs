using EasyUseExcel.Attribute;
using IgnoreAttribute = EasyUseExcel.Attribute.IgnoreAttribute;
using OrderAttribute = EasyUseExcel.Attribute.OrderAttribute;

namespace EasyUseExcel.Tests
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void Test1()
        {
            var testCase = new List<TestModel>();
            for (var i = 1; i < 6; i++) 
            {
                testCase.Add(CreateTestModel(i));
            }

            var stream = ExcelWriter.Excute(testCase);

            Assert.IsTrue(stream.Length > 0);

            var result = ExcelReader.Excute<TestModel>(stream, 1, 2);

            Assert.IsTrue(result.Count > 0);

            Assert.Pass();
        }

        [Test]
        public void Test2()
        {
            var testCase = new List<TestModel>();
            for (var i = 1; i < 6; i++)
            {
                testCase.Add(CreateTestModel(i));
            }

            var stream = ExcelWriter.Excute<TestModel, TestModel>(testCase, testCase);

            Assert.IsTrue(stream.Length > 0);

            var result = ExcelReader.Excute<TestModel>(stream, 2, 2);

            Assert.IsTrue(result.Count > 0);

            Assert.Pass();
        }

        public TestModel CreateTestModel(int seq) 
        {
            return new TestModel()
            {
                Seq = seq,
                Name = $"User{seq}",
                Age = seq + 18,
                Phone = $"9999999{seq}",
                Remark = "TestData"
            };
        }
    }

    public class TestModel 
    {
        [Order(1)]
        public int Seq { get; set; }

        [Order(2)]
        [Display("UserName")]
        public string Name { get; set; }

        [Order(4)]
        public int Age { get; set; }

        [Order(3)]
        public string Phone { get; set; }

        [Ignore]
        public string Remark { get; set; }
    }
  
}