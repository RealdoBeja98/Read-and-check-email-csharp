using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using System.Net.Http;


namespace IC.MailCheck

{
/// <summary>
/// ///
/// </summary>
    class Program
    {
        static void openFile()
        {
            Excel.Application application = new Excel.Application();
            string path_of_file = "C:\\Users\\RealdoB\\source\\repos\\IC.EmailCheck2\\IC.MailCheck\\final_email_list.xlsx";
            Excel.Workbook workbook = application.Workbooks.Open(path_of_file);
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Excel.Range range = worksheet.UsedRange;
            int column_of_final_email = 4;
            transform_excelsheet_to_matrix(range, column_of_final_email, worksheet, workbook);
        }
        static void transform_excelsheet_to_matrix(Excel.Range range, int column_of_final_email, Excel.Worksheet worksheet, Excel.Workbook workbook)
        {
            string[,] arr = new string[range.Rows.Count, range.Columns.Count];
            for (int row = 95; row <= range.Rows.Count; row++)
            {
                List<string> list_of_information_in_excel = new List<string>();
                for (int coll = 1; coll <= 3; coll++)
                {
                    arr[row - 1, coll - 1] = ((Excel.Range)range.Cells[row, coll]).Text;
                    list_of_information_in_excel.Add(arr[row - 1, coll - 1]);
                }
                Console.WriteLine();
                var full_name_in_first_col_in_excel = list_of_information_in_excel[0];//.Replace(" ", "");
                var website_domain_in_excel = list_of_information_in_excel[2];
                printFullname_domain_and_clear_array(full_name_in_first_col_in_excel, website_domain_in_excel, list_of_information_in_excel);
                string checked_email = do_checks_of_email(full_name_in_first_col_in_excel, website_domain_in_excel);
                worksheet.Cells[row, column_of_final_email] = checked_email;

                
                workbook.Save();

            }
            workbook.Close();
        }

        static string do_checks_of_email(string full_name_in_first_col_in_excel, string website_domain_in_excel)
        {

             string domain_without_protocoll = clear_protocoll_from_website(website_domain_in_excel);
             //nomecognome@domain.it
             string deleted_white_space_from_full_name = full_name_in_first_col_in_excel;
             string str1 = deleted_white_space_from_full_name.Replace(" ", "") + "@" + domain_without_protocoll;
             Console.WriteLine(str1);
             var result1 = call_api(str1.ToLower()).Result;
             if (result1) 
                 return str1.ToLower();

            //nome.cognome@domain.it
             string full_name_with_point = full_name_in_first_col_in_excel;
             string str2 = full_name_with_point.Replace(" ", ".") + "@" + domain_without_protocoll;
             Console.WriteLine(str2);
             var result2 = call_api(str2.ToLower()).Result;
             if (result2) 
                 return str2.ToLower();

            //nome_cognome@domain.it
             string full_name_with_underscore = full_name_in_first_col_in_excel;
             string str3 = full_name_with_underscore.Replace(" ", "_") + "@" + domain_without_protocoll;
             Console.WriteLine(str3);

             var result3 = call_api(str3.ToLower()).Result;
             if (result3) 
                 return str3.ToLower();
            //nome@domain.it
             string name_with_domain = full_name_in_first_col_in_excel;
             string str4 = name_with_domain.Substring(0, name_with_domain.IndexOf(" ") + 0) + "@" + domain_without_protocoll;
             Console.WriteLine(str4);

             var result4 = call_api(str4.ToLower()).Result;
             if (result4) 
                 return str4.ToLower();
            //cognome@domain.it
             string surname_with_domain = full_name_in_first_col_in_excel;
             string str5 = surname_with_domain.Substring(surname_with_domain.LastIndexOf(" ") + 1) + "@" + domain_without_protocoll;
             Console.WriteLine(str5);

             var result5 = call_api(str5.ToLower()).Result;
             if (result5) 
                 return str5.ToLower();
            //n.cognome@domain.it
             string first_letter_of_name_with_surname = full_name_in_first_col_in_excel;
             string str6 = first_letter_of_name_with_surname.Remove(1, first_letter_of_name_with_surname.IndexOf(" ")).Insert(1, ".") + "@" + domain_without_protocoll;
             Console.WriteLine(str6);

             var result6 = call_api(str6.ToLower()).Result;
             if (result6) 
                 return str6.ToLower();

             return "Not found";

        }

        static async Task<bool> call_api(string email)
        {
            var httpClient = HttpClientFactory.Create();
            var access_key = "3ce1f192713f9cab33da3de0b0fb242b";
            var url = "http://apilayer.net/api/check?access_key=" + access_key + "&email=" + email + "&format=1";
            var data = await httpClient.GetStringAsync(url);
            JObject json = JObject.Parse(data);
            var api_result = json["smtp_check"].Value<bool>();
            return api_result;

        }
        static void printFullname_domain_and_clear_array(string full_name_in_first_col_in_excel, string website_domain_in_excel, List<string> list_of_information_in_excel)
        {
            Console.WriteLine("=========");
            Console.WriteLine(full_name_in_first_col_in_excel);
            string domain_without_protocoll = clear_protocoll_from_website(website_domain_in_excel);
            Console.WriteLine(domain_without_protocoll);
            Console.WriteLine();
            list_of_information_in_excel.Clear();

        }

        static string clear_protocoll_from_website(string website_domain_in_excel)
        {
            string pattern = @"http:\/\/(?:www\.)?";

            return Regex.Replace(website_domain_in_excel, pattern, String.Empty);
        }
        static async Task Main(string[] args)
        {
            
            openFile();
        }
    }
}
