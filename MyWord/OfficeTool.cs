using Spire.Doc;
using Spire.Presentation;
using Spire.Xls;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace OfficeTool
{
    class OfficeTool
    {
        static Regex reg = new Regex(@"[\u4e00-\u9fa5]");//正则表达式

        static void Main(string[] args)
        { 
            Boolean linux = false;
            string inputFileName = "";
            string outputFileName = "";
            string type = "";
            if (linux)
            {
                inputFileName = args[0];
                outputFileName = args[1];
                type = args[2];
            }
            else 
            {
                inputFileName = "d:/code/test.doc";
                outputFileName = "d:/code/test.html";
                type = "docx";
            }

            if ("pptx".Equals(type))
            {
                PptxToHtml(inputFileName, outputFileName, linux);
            }
            else if ("xlsx".Equals(type))
            {
                ExcelToHtml(inputFileName, outputFileName);
            }
            else 
            {
                DocxToHtml(inputFileName, outputFileName);
            }
        }

        static string getNewLine(string line)
        {
            if (line.Contains("<text") && reg.IsMatch(line) && line.Contains("<tspan")) 
            {
                string[] arr = line.Split("<tspan");
                string height = "0";
                double position = 0.00;
                for (int i = 1; i < arr.Length; i++) 
                {
                    if (arr[i].Contains("textLength"))
                    {
                        string[] items = arr[i].Split("\"");
                        int index = Array.IndexOf(items, " textLength=");
                        string len = items[index+1];
                        string itemHeight = items[3];
                        double length = Convert.ToDouble(len);
                        double width = length;
                        if (reg.IsMatch(arr[i]))
                        {
                            width = Math.Round(length * 1.6, 2);
                        }
                        if (i > 1 && height.Equals(itemHeight))
                        {
                            items[1] = position + "";
                            position += width;
                        }
                        else
                        {
                            string pos = items[1];
                            position = Convert.ToDouble(pos) + width;
                        }
                        height = itemHeight;
                        items[items.Length - 2] = width + "";
                        arr[i] = string.Join("\"", items);
                    }
                }
                line = string.Join("<tspan", arr);
            }
            return line;
        }

        static void PptxToHtml(string inputFileName, string outputFileName, Boolean linux) 
        {
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(inputFileName);
            presentation.SaveToFile(outputFileName, Spire.Presentation.FileFormat.Html);
            string output = "";
            if (linux)
            {
                string[] arr = File.ReadAllLines(outputFileName);
                for (int i = 0; i < arr.Length; i++)
                {
                    string line = arr[i].Replace("Evaluation Warning : The document was created with  Spire.Presentation for .NET", "");
                    arr[i] = getNewLine(line);
                }
                output = string.Join("\r\n", arr);
            }
            else 
            {
                string line = File.ReadAllText(outputFileName);
                output = line.Replace("Evaluation Warning : The document was created with  Spire.Presentation for .NET", "");
            }
            File.WriteAllText(outputFileName, output, System.Text.Encoding.UTF8);//保存结果
        }

        static void DocxToHtml(string inputFileName, string outputFileName)
        {
            Document doc = new Document();
            doc.LoadFromFile(inputFileName);
            doc.SaveToFile(outputFileName, Spire.Doc.FileFormat.Html);
            string output = File.ReadAllText(outputFileName);
            string text = output.Replace("Evaluation Warning: The document was created with Spire.Doc for .NET.", "");
            File.WriteAllText(outputFileName, text, System.Text.Encoding.UTF8);//保存结果
        }


        static void ExcelToHtml(string inputFileName, string outputFileName)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(inputFileName);
            //convertExceltoHTML
            var list = workbook.Worksheets;
            int count = 1;
            int index = outputFileName.LastIndexOf(".");
            string head = outputFileName.Substring(0, index);
            string footer = outputFileName.Substring(index);
            foreach (Worksheet sheet in list)
            {
                string outputName = head + "-" + count + footer;
                sheet.SaveToHtml(outputName);
                string sss = File.ReadAllText(outputName);
                string output = sss.Replace("Evaluation&nbsp;Warning&nbsp;:&nbsp;The&nbsp;document&nbsp;was&nbsp;created&nbsp;with&nbsp;&nbsp;Spire.XLS&nbsp;for&nbsp;.NET", "");
                File.WriteAllText(outputName, output, System.Text.Encoding.UTF8);//保存结果
                count++;
            }
        }


    }
}
