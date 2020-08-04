using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlusTest
{
    public class Startup
    {
        // This method gets called by the runtime. Use this method to add services to the container.
        // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940
        public void ConfigureServices(IServiceCollection services)
        {
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseRouting();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapGet("/", async context =>
                {
                    //            var Articles = new[]
                    //{
                    //        new {
                    //            Id = "101", Name = "C++"
                    //        },
                    //        new {
                    //            Id = "102", Name = "Python"
                    //        },
                    //        new {
                    //            Id = "103", Name = "Java Script"
                    //        },
                    //        new {
                    //            Id = "104", Name = "GO"
                    //        },
                    //        new {
                    //            Id = "105", Name = "Java"
                    //        },
                    //        new {
                    //            Id = "106", Name = "C#"
                    //        }
                    //    };

                    //            // Creating an instance 
                    //            // of ExcelPackage 
                    //            ExcelPackage excel = new ExcelPackage();

                    //            // name of the sheet 
                    //            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");

                    //            // setting the properties 
                    //            // of the work sheet  
                    //            workSheet.TabColor = System.Drawing.Color.Black;
                    //            workSheet.DefaultRowHeight = 12;

                    //            // Setting the properties 
                    //            // of the first row 
                    //            workSheet.Row(1).Height = 20;
                    //            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //            workSheet.Row(1).Style.Font.Bold = true;

                    //            // Header of the Excel sheet 
                    //            workSheet.Cells[1, 1].Value = "S.No";
                    //            workSheet.Cells[1, 2].Value = "Id";
                    //            workSheet.Cells[1, 3].Value = "Name";

                    //            // Inserting the article data into excel 
                    //            // sheet by using the for each loop 
                    //            // As we have values to the first row  
                    //            // we will start with second row 
                    //            int recordIndex = 2;

                    //            foreach (var article in Articles)
                    //            {
                    //                workSheet.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                    //                workSheet.Cells[recordIndex, 2].Value = article.Id;
                    //                workSheet.Cells[recordIndex, 3].Value = article.Name;
                    //                recordIndex++;
                    //            }

                    //            // By default, the column width is not  
                    //            // set to auto fit for the content 
                    //            // of the range, so we are using 
                    //            // AutoFit() method here.  
                    //            workSheet.Column(1).AutoFit();
                    //            workSheet.Column(2).AutoFit();
                    //            workSheet.Column(3).AutoFit();

                    //            // file name with .xlsx extension  
                    //            string p_strPath = "geeksforgeeks.xlsx";

                    //            if (File.Exists(p_strPath))
                    //                File.Delete(p_strPath);

                    //            // Create excel file on physical disk  
                    //            FileStream objFileStrm = File.Create(p_strPath);
                    //            objFileStrm.Close();

                    //            // Write content to excel file  
                    //            File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
                    //            //Close Excel package 
                    //            excel.Dispose();
                    //            Console.ReadKey();


                    //--------------------------------------
                    var fileinfo = new FileInfo("template.xlsx");
                    if (fileinfo.Exists)
                    {
                        using (ExcelPackage p = new ExcelPackage(fileinfo))
                        {
                            //using (stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
                            {
                                //p.Load(stream);
                                //ExcelWorksheet ws = p.Workbook.Worksheets.Add(wsName + wsNumber.ToString());
                                //ws.Cells[1, 1].Value = wsName;
                                //ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                //ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                                //ws.Cells[1, 1].Style.Font.Bold = true;


                                ExcelWorksheet ws = p.Workbook.Worksheets[0];
                                ws.Cells[9, 3].Value = 2;
                                ws.Cells[10, 3].Value = 2;
                                var fileinforesult = new FileInfo("1.xlsx");
                                p.SaveAs(fileinforesult);
                            }

                        }

                    }

                    await context.Response.WriteAsync("Hello World!");

                });
            });
        }
    }
}
