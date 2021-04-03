using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_dynamic
{
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }

        public Person(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }
    class Program
    {
        static List<Person> persons = new List<Person>();
        static Program()
        {
            persons.Add(new Person("Frank", 25));
            persons.Add(new Person("Joe", 24));
        }

        static void Main(string[] args)
        {
            dynamic excelType = Type.GetTypeFromProgID("Excel.Application");
            var excelObj = Activator.CreateInstance(excelType);
            excelObj.Visible = true;

            excelObj.Workbooks.Add();
            dynamic workSheet = excelObj.ActiveSheet;

            workSheet.Cells[1, 1] = "Names";
            workSheet.Cells[1, 2] = "Age";

            int rowIndex = 1;
            foreach (var person in persons)
            {
                rowIndex++;
                workSheet.Cells[rowIndex, 1] = person.Name;
                workSheet.Cells[rowIndex, 2] = person.Age;
            }
        }
    }
}
