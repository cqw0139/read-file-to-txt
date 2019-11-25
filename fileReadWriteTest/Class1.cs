using System;
using System.Text;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace fileReadWriteTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string JobToken = string.Empty;
            string line = string.Empty;

            try
            {

                StringBuilder text = new StringBuilder();
                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                object miss = System.Reflection.Missing.Value;
                object path = @"D:/visual/projects/ConsoleApp3/bin/Debug/test.txt";
                object readOnly = true;
                Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

                for (int i = 0; i < docs.Paragraphs.Count; i++)
                {
                    text.Append(" \r\n " + docs.Paragraphs[i + 1].Range.Text.ToString());
                }


                using (StringReader reader = new StringReader(text.ToString()))
                {
                    line = reader.ReadLine();

                    if (line.ToString().Contains("JobToken"))
                    {
                        JobToken = line.ToString().Replace("JobToken", "").Trim();
                    }

                }

            }
            catch (Exception e)
            {
                string startupPath = Environment.CurrentDirectory;
                File.WriteAllText(startupPath + "/test.txt", "JobToken:mvsdnvweoinvewr");
            }


            Console.WriteLine("Token: ");
        }
    }
}
