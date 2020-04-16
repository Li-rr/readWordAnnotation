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
        string outputDir = "D:\\WORKSPACE\\助教工作\\作业批改\\作业数据\\实验0-1改\\";
        ArrayList filePath = new ArrayList();
        double score_per_problem = 3.7;

        public void readFileToTxt()
        {
            string t_var = "";
            string t_temp = "";
            string stu_name = "";
            string stu_no = "";
            double score = 0;
            foreach (var stu_path in filePath)
            {
                //if (!stu_path.ToString().Contains("潘天宇"))
                //    continue;
                t_var = stu_path.ToString();
                t_temp = t_var.Split('\\')[6];
                stu_name = t_temp.Split('-')[0];
                stu_no = t_temp.Split('-')[1];
                Console.WriteLine(stu_no+" "+stu_name);
                
                Document doc = new Document();
                doc.LoadFromFile(t_var);
                StringBuilder sb = new StringBuilder();
                score = 0;
                foreach (Comment comment in doc.Comments)
                {
                    foreach (Paragraph paragraph in comment.Body.Paragraphs)
                    {

                        if(paragraph.Text.Contains("错") && !paragraph.Text.Contains("半")){
                            score += score_per_problem;
                        }
                        if (paragraph.Text.Contains("半错"))
                            score += score_per_problem / 2;
                        sb.AppendLine(paragraph.Text);
                        Console.WriteLine(paragraph.Text);
                    }
                    Console.WriteLine(score);
                }
                sb.AppendLine((100-score).ToString());
                File.WriteAllText(outputDir+ stu_name + "-"+stu_no+".txt", sb.ToString()); // 写入文件

            }
            //Document doc = new Document();

            //doc.LoadFromFile("D:\\WORKSPACE\\助教工作\\作业批改\\作业数据\\实验0-1\\20190918-马姝媛\\56518797-029班马姝媛20190918 实验0-1.docx");


            //StringBuilder sb = new StringBuilder();

            //foreach (Comment comment in doc.Comments)
            //{
            //    foreach (Paragraph paragraph in comment.Body.Paragraphs)
            //    {
            //        sb.AppendLine(paragraph.Text);
            //        Console.WriteLine(paragraph.Text);
            //    }
            //}
            //File.WriteAllText("aaa.txt", sb.ToString()); // 写入文件


        }

        public  void getAllFilePath()
        {
            DirectoryInfo theFolder = new DirectoryInfo(dirPath);
            foreach (var dir in theFolder.GetDirectories())
            {
                string newPath = Path.Combine(dirPath, dir.ToString());
                //Console.WriteLine(newPath); // 获得当前学生文件夹

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
            f_you.readFileToTxt();
            Console.ReadLine();
        }
    }
}
