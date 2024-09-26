using System;
using MySql.Data.MySqlClient; 
using Microsoft.Office.Interop.Word;

namespace MailMergeExample
{
    class Program
    {
        static void Main(string[] args)
        {
            string templatePath = @"C:\Users\Asani\Downloads\mail_merge_template.docx";
            string outputDirectory = @"C:\Users\Asani\Downloads\";

            string connectionString = "Server=localhost;Database=db_mail_marge;User=root;Password=;";
            string query = "SELECT * FROM Debtor";

            Application wordApp = new Application();

            try
            {
                using (MySqlConnection conn = new MySqlConnection(connectionString))
                {
                    conn.Open();
                    using (MySqlCommand cmd = new MySqlCommand(query, conn))
                    {
                        using (MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            int fileIndex = 1;

                            while (reader.Read())
                            {
                                Document wordDoc = wordApp.Documents.Open(templatePath);

                                var data = new
                                {
                                    Today = Convert.ToDateTime(reader["Today"]).ToString("MMMM dd, yyyy"),
                                    DebtorName = reader["DebtorName"].ToString(),
                                    Jabatan = reader["Jabatan"].ToString(),
                                    DebtorEmail = reader["DebtorEmail"].ToString(),
                                    Dear = reader["Dear"].ToString(),
                                    PlanNo = reader["PlanNo"].ToString(),
                                    BuildAddress = reader["BuildAddress"].ToString(),
                                    LotNo = reader["LotNo"].ToString()
                                };

                                ReplacePlaceholder(wordDoc, "<Today>", data.Today);
                                ReplacePlaceholder(wordDoc, "<DebtorName>", data.DebtorName);
                                ReplacePlaceholder(wordDoc, "<Jabatan>", data.Jabatan);
                                ReplacePlaceholder(wordDoc, "<DebtorEmail>", data.DebtorEmail);
                                ReplacePlaceholder(wordDoc, "<Dear>", data.Dear);
                                ReplacePlaceholder(wordDoc, "<PlanNo>", data.PlanNo);
                                ReplacePlaceholder(wordDoc, "<BuildAddress>", data.BuildAddress);
                                ReplacePlaceholder(wordDoc, "<LotNo>", data.LotNo);

                                string outputPath = $"{outputDirectory}output_{fileIndex}.docx";
                                wordDoc.SaveAs2(outputPath);
                                Console.WriteLine($"Mail merge selesai untuk record {fileIndex}. Hasil disimpan di {outputPath}");

                                wordDoc.Close();
                                fileIndex++;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                wordApp.Quit();
            }
        }

        static void ReplacePlaceholder(Document doc, string placeholder, string value)
        {
            foreach (Microsoft.Office.Interop.Word.Range range in doc.StoryRanges)
            {
                range.Find.Execute(FindText: placeholder, ReplaceWith: value);
            }
        }
    }
}
