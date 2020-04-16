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

        public  void getAllFilePath()
        {
            DirectoryInfo theFolder = new DirectoryInfo(dirPath);
            foreach (var dir in theFolder.GetDirectories())
            {
                string newPath = Path.Combine(dirPath, dir.ToString());
                Console.WriteLine(newPath); // 获得当前学生文件夹

                var files = Directory.GetFiles(newPath);
                if (files.Length > 1)
                    Console.WriteLine(files.Length+" ============");
                foreach (var c_file in files)
                {
                    //Console.WriteLine(c_file);
                    filePath.Add(c_file);
                }


            }
            Console.WriteLine(filePath.Count);
        }
        static void Main(string[] args)
        {
            Program f_you = new Program();
            Console.WriteLine("Hello World!");
            //test();
            f_you.getAllFilePath();
            Console.ReadLine();
        }
    }
}
