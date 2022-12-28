# EasyUseExcel
use ClosedXML<br>

C# excel helper<br>

Attribute List
<table>
    <thead>
        <tr>
            <th>
                class name       
            </th>
            <th>
                description
            </th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>
                Order
            </td>
            <td>
                排序用
            </td>
        </tr>
        <tr>
            <td>
                Display
            </td>
            <td>
                自訂表頭
            </td>
        </tr>
        <tr>
            <td>
                Ignore
            </td>
            <td>
                忽略
            </td>
        </tr>
        <tr>
            <td>
                RowSpan
            </td>
            <td>
                同HTML RowSpan
            </td>
        </tr>
    </tbody>
</table>

## ExcelWriter   

Model to Excel   
example:
```
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

    public class test()
    {
         var list = new List<TestModel>();
         list.add(new TestModel()
            {
                Seq = 1,
                Name = "User1",
                Age = 18,
                Phone = "99999999999",
                Remark = "TestData"
        });
        var stream = ExcelWriter.Excute(testCase);
    }
```
## ExcelReader
Excel to Model  
example:
```
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

    public class test()
    {
        var stream = // your file
        var result = ExcelReader.Excute<TestModel>(stream, BeginSheet: 1, BeginRow: 2);
    }
```
