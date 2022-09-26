using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace convertToExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {

            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("api");
            //string[] columns = {"Stt", "Method", "Name", "Url", "Parameter", "Validate", "Note"};
            Columns[] columns =
            {
                 new Columns("Stt", 10),
                 new Columns("Name", 40),
                 new Columns("Method", 20),
                 new Columns("Url", 60),
                 new Columns("Paramater", 40),
                 new Columns("Validate", 40),
                 new Columns("Note", 80)
             };

            #region title
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            for (int i = 0; i < columns.Length; i++)
            {
                workSheet.Cells[1, i + 1].Value = columns[i].name;
                workSheet.Column(i + 1).Width = columns[i].width;
            }
            #endregion



            dynamic jsonFile = JsonConvert.DeserializeObject(File.ReadAllText(@"C:\Users\son\Desktop\testapi.json"));
            dynamic typeApps = jsonFile["item"];
            int indexTypeApp = 1;
            int indexRow = 2;

            for (int i = 0; i < typeApps.Count; i++)
            {
                workSheet.Cells[indexRow, 1].Value = indexTypeApp.ToString();
                workSheet.Cells[indexRow, 2].Value = typeApps[i]["name"].ToString();

                configStyleHeader(ref workSheet, indexRow, columns.Length, System.Drawing.Color.Aqua);

                indexRow++;
                int indexGroup = 1;
                JArray groups = ((JObject)typeApps[i]).ContainsKey("item") ? (JArray)typeApps[i]["item"] : new JArray();
                for (int j = 0; j < groups.Count; j++)
                {
                    JObject group = (JObject)groups[j];
                   
                    JArray apis = group.ContainsKey("item") ? (JArray)group["item"] : new JArray();
                    if (apis.Count > 0)
                    {
                        workSheet.Cells[indexRow, 1].Value = indexTypeApp.ToString() + "." + indexGroup.ToString();
                        workSheet.Cells[indexRow, 2].Value = group["name"].ToString();

                        configStyleHeader(ref workSheet, indexRow, columns.Length, System.Drawing.Color.AntiqueWhite);

                        indexRow++;
                        for (int k = 0; k < apis.Count; k++)
                        {
                            JObject api = (JObject)apis[k];
                            //set column stt, name
                            workSheet.Cells[indexRow, 1].Value = indexTypeApp.ToString() + "." + indexGroup.ToString() + "." + (k + 1).ToString();


                            string name = api.ContainsKey("name") ? (string)api.GetValue("name") : "";
                            JObject request = api.ContainsKey("request") ? (JObject)api.GetValue("request") : new JObject();
                            JObject url = request.ContainsKey("url") ? (JObject)request.GetValue("url") : new JObject();
                            string urlName = url.ContainsKey("raw") ? (string)url.GetValue("raw") : "";
                            string method = request.ContainsKey("method") ? (string)request.GetValue("method") : "";
                            string description = request.ContainsKey("description") ? (string)request.GetValue("description") : "";

                            workSheet.Cells[indexRow, 2].Value = name;
                            workSheet.Cells[indexRow, 3].Value = method;
                            workSheet.Cells[indexRow, 4].Value = urlName;
                            workSheet.Cells[indexRow, 7].Value = description;

                            if (method == "GET" || method == "DELETE")
                            {
                                JArray queriesTemp = url.ContainsKey("query") ? (JArray)url["query"] : new JArray();
                                List<JObject> queries = new List<JObject>();
                                foreach (JObject obj in queriesTemp)
                                {
                                    if (obj.ContainsKey("description"))
                                    {
                                        queries.Add(obj);
                                    }
                                }

                                if (queries.Count > 0)
                                {

                                    workSheet.Cells[indexRow, 1, indexRow + queries.Count - 1, 1].Merge = true;
                                    workSheet.Cells[indexRow, 2, indexRow + queries.Count - 1, 2].Merge = true;
                                    workSheet.Cells[indexRow, 3, indexRow + queries.Count - 1, 3].Merge = true;
                                    workSheet.Cells[indexRow, 4, indexRow + queries.Count - 1, 4].Merge = true;
                                    workSheet.Cells[indexRow, 7, indexRow + queries.Count - 1, 7].Merge = true;
                                    foreach (JObject query in queries)
                                    {
                                        string key = query.ContainsKey("key") ? (string)query.GetValue("key") : "";
                                        string descriptionKey = query.ContainsKey("description") ? (string)query.GetValue("description") : "";
                                        workSheet.Cells[indexRow, 5].Value = key;
                                        workSheet.Cells[indexRow, 6].Value = descriptionKey;
                                        indexRow++;
                                    }
                                }
                                else
                                {
                                    indexRow++;
                                }
                                //Console.WriteLine(queries);
                            }
                            else
                            {
                                Body body = getBody(request.ContainsKey("body") ? (JObject)request.GetValue("body") : new JObject());
                                List<Parameter> parameters = handValidation(body);
                                if (parameters.Count > 0)
                                {
                                    workSheet.Cells[indexRow, 1, indexRow + parameters.Count - 1, 1].Merge = true;
                                    workSheet.Cells[indexRow, 2, indexRow + parameters.Count - 1, 2].Merge = true;
                                    workSheet.Cells[indexRow, 3, indexRow + parameters.Count - 1, 3].Merge = true;
                                    workSheet.Cells[indexRow, 4, indexRow + parameters.Count - 1, 4].Merge = true;
                                    workSheet.Cells[indexRow, 7, indexRow + parameters.Count - 1, 7].Merge = true;
                                    foreach (Parameter param in parameters)
                                    {
                                        string key = param.name;
                                        string descriptionKey = param.description;
                                        workSheet.Cells[indexRow, 5].Value = key;
                                        workSheet.Cells[indexRow, 6].Value = descriptionKey;
                                        indexRow++;
                                    }
                                }
                                else
                                {
                                    indexRow++;
                                }
                            }



                        }
                        indexGroup++;
                    }


                }
                indexTypeApp++;
            }


            // file name with .xlsx extension 
            string p_strPath = "result.xlsx";

            if (File.Exists(p_strPath))
                File.Delete(p_strPath);

            // Create excel file on physical disk 
            FileStream objFileStrm = File.Create(p_strPath);
            objFileStrm.Close();

            // Write content to excel file 
            File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
            //Close Excel package
            excel.Dispose();
            Console.ReadKey();











           /* List<Parameter> result = new List<Parameter>();
            string test = "{\r\n    \"device\": \"son1999tmgl@gmail.com\",//M\r\n    \"persionActionEmail\": \"son1999tmgl@gmail.com\",//M\r\n    \"signType\": 6,//M\r\n    \"employeeId\": 43,//M\r\n    \"employeeSignId\": 77,//M\r\n    \"documentIds\": [\r\n        56\r\n    ],//M\r\n    \"employeeName\": \"CÔNG TY ABCD\"//M\r\n}";
            JObject jsontest = JObject.Parse(test);
            string pattern = "\\\"([\\w_-]+)\\\":[^(//\\n)]{1,100}//([^\\n]+)";
            Regex myRegex = new Regex(pattern);
            MatchCollection demo = myRegex.Matches(test);
           Console.WriteLine(demo.Count);
            foreach (Match m in demo)
            {
                Console.WriteLine(m.Groups[1].Value + ":" + m.Groups[2].Value);
            }*/
        }


        static public void configStyleHeader(ref ExcelWorksheet workSheet, int indexRow, int numberColumns, Color color)
        {
            workSheet.Cells[indexRow, 2, indexRow, numberColumns].Merge = true;
            workSheet.Cells[indexRow, 2, indexRow, numberColumns].Style.Font.Bold = true;
            workSheet.Cells[indexRow, 1, indexRow, numberColumns].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells[indexRow, 1, indexRow, numberColumns].Style.Fill.BackgroundColor.SetColor(color);
        }

        static public Body getBody(JObject objectBody)
        {
            Body body = new Body();
            body.mode = objectBody.ContainsKey("mode") ? (string)objectBody.GetValue("mode") : "";
            if(body.mode == "raw")
            {
                body.value = objectBody.ContainsKey(body.mode) ? (string)objectBody.GetValue(body.mode) : ""; 
            }else if(body.mode == "formdata")
            {
                body.value = objectBody.ContainsKey(body.mode) ? (JArray)objectBody.GetValue(body.mode) : new JArray();
            }
            
            return body;
        }

        static public List<Parameter> handValidation(Body body)
        {
            List<Parameter> result = new List<Parameter>();
            if(body.mode == "raw")
            {
                string raw = (string)body.value;
                if(raw != "")
                {
                    string pattern = "\\\"([\\w_-]+)\\\":[^\\n]+//([M|O][^\\n]{0,200}\\n)";
                    Regex regex = new Regex(pattern);
                    MatchCollection listRegex = regex.Matches(raw);
                    Console.WriteLine("-----------------");
                    Console.WriteLine(raw);
                    foreach(Match match in listRegex)
                    {
                        Console.WriteLine(match.Groups[1].Value + ": " + match.Groups[2].Value);
                        result.Add(new Parameter(match.Groups[1].Value, match.Groups[2].Value));
                    }
                }
            }
            else if(body.mode == "formdata")
            {
                foreach(JObject param in body.value)
                {
                    string name = param.ContainsKey("key") ? (string)param.GetValue("key") : "";
                    string description = param.ContainsKey("description") ? (string)param.GetValue("description") : "";
                    result.Add(new Parameter(name, description));
                }
            }
            return result;
        }

    }

    class Columns
    {
        public Columns(string name, int width)
        {
            this.name = name;
            this.width = width;
        }
        public string name { get; set; }
        public int width { get; set; }
    }

    class Body
    {
        public string mode { get; set; }
        public dynamic value { get; set; }
    }

    class Parameter
    {
        public Parameter(string name, string description)
        {
            this.name = name;
            this.description = description;
        }

        public string name { get; set; }
        public string description { get; set; }

        public override string ToString()
        {
            return this.name +"---"+ this.description;
        }
    }
}
