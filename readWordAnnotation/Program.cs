using System;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.IO;
using System.Text;
using System.Collections;

namespace readWordAnnotation
{
    class Program
    {
        string dirPath = "D:\\WORKSPACE\\助教工作\\作业批改\\作业数据\\实验0-1";
        ArrayList filePath = new ArrayList();

        static void test()
        {
            Document doc = new Document();

            doc.LoadFromFile("D:\\WORKSPACE\\助教工作\\作业批改\\作业数据\\实验0-1\\20190918-马姝媛\\56518797-029班马姝媛20190918 实验0-1.docx");


            StringBuilder sb = new StringBuilder();

            foreach (Comment comment in doc.Comments)
            {
                foreach (Paragraph paragraph in comment.Body.Paragraphs)
                {
                    sb.AppendLine(paragraph.Text);
                    Console.WriteLine(paragraph.Text);
                }
            }
            File.WriteAllText("aaa.txt", sb.ToString()); // 写入文件


        }


        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            test();
        }
    }
}
