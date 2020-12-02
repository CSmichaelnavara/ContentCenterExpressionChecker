using System;
using System.Runtime.InteropServices;
using Inventor;

namespace ContentCenterExpressionChecker
{
    class Program
    {
        static void Main(string[] args)
        {

            var inventor = (Application)Marshal.GetActiveObject("Inventor.Application");

            Console.WriteLine($"{"familyName".PadRight(40)}{"columnName".PadRight(25)}expression");
            Console.WriteLine(new string('-', 100));

            new ContentCenterBrowser().Browse(inventor.ContentCenter.TreeViewTopNode);

            Console.WriteLine(new string('-', 100));
            Console.WriteLine("Press any key to quit...");
            Console.ReadKey();
        }
    }

    class ContentCenterBrowser
    {
        private string columnName;
        private string familyName;

        public void Browse(ContentTreeViewNode node)
        {


            foreach (ContentFamily family in node.Families)
            {
                CheckFamily(family);
            }

            foreach (ContentTreeViewNode childNode in node.ChildNodes)
            {
                Browse(childNode);
            }
        }

        private void CheckExpression(object expression)
        {
            if (expression == null)
                return;
            if (expression is string stringExpression)
            {
                bool badExpression = false;

                var expressionParts = stringExpression.Split('&');
                if (expressionParts.Length < 2) return; //Expression neobsahuje '&'

                foreach (string expressionPart in expressionParts)
                {
                    string part = expressionPart.Trim();
                    //Kazda cast vyrazu musi byt uzavrena bud v uvozovkach ("...") nebo ve slozenych zavorkach ({...})
                    badExpression |= !(part.StartsWith("\"") && part.EndsWith("\"") || part.StartsWith("{") && part.EndsWith("}"));

                }

                if (badExpression)
                    Console.WriteLine($"{familyName.PadRight(40)}{columnName.PadRight(25)}{expression}");
            }
        }

        private void CheckFamily(ContentFamily family)
        {
            if (!family.IsModifiable)
                return;
            familyName = family.DisplayName;

            foreach (ContentTableColumn column in family.TableColumns)
            {
                if (column.DataType != ValueTypeEnum.kStringType) continue;

                if (column.InternalName == "MATERIAL") continue;
                if (column.InternalName == "THREADCLASSEXT") continue;
                if (column.InternalName == "THREADTYPE") continue;
                columnName = column.InternalName;

                CheckExpression(column.Expression);
            }
        }
    }
}