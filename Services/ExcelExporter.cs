using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using bluesolutions.Models;

namespace bluesolutions.Services
{
    public class ExcelExporter
    {
        public void ExportToExcel(List<Post> posts, Stream stream)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Posts");

                worksheet.Cell(1, 1).Value = "공고명";
                worksheet.Cell(1, 2).Value = "검색 키워드";
                worksheet.Cell(1, 3).Value = "공고기관";
                worksheet.Cell(1, 4).Value = "수요기관";
                worksheet.Cell(1, 5).Value = "입력일시";
                worksheet.Cell(1, 6).Value = "검색일시";

                for (int i = 0; i < posts.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = posts[i].Title;
                    worksheet.Cell(i + 2, 2).Value = posts[i].Keyword;
                    worksheet.Cell(i + 2, 3).Value = posts[i].PostingInstitution;
                    worksheet.Cell(i + 2, 4).Value = posts[i].RequestingInstitution;
                    worksheet.Cell(i + 2, 5).Value = posts[i].InputDate;
                    worksheet.Cell(i + 2, 6).Value = posts[i].SearchDate;
                }

                workbook.SaveAs(stream);
            }
        }
    }
}
