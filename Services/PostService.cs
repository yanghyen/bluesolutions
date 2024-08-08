using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using ClosedXML.Excel;
using bluesolutions.Models;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;

namespace bluesolutions.Services
{
    public class PostService
    {
        private readonly HttpClient _httpClient;

        public PostService(HttpClient httpClient)
        {
            _httpClient = httpClient;
        }

        public async Task<List<Post>> GetPostsAsync(string keyword, DateTime startDate, DateTime endDate)
        {
            var posts = new List<Post>();

            new DriverManager().SetUpDriver(new ChromeConfig());

            var options = new ChromeOptions();
            options.AddArgument("--no-sandbox");
            options.AddArgument("--disable-dev-shm-usage");

            // Chromedriver 경로 설정
            using var driver = new ChromeDriver(options);

            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5)); // 30초 대기 설정

            try
            {
                Console.WriteLine("Navigating to URL...");
                driver.Navigate().GoToUrl("http://www.g2b.go.kr/index.jsp");

                Console.WriteLine("Filling out the form...");
                var keywordInput = wait.Until(drv => drv.FindElement(By.XPath("//*[@id='bidNm']")));
                keywordInput.SendKeys(keyword);

                // JavaScript로 날짜 입력
                var js = (IJavaScriptExecutor)driver;
                js.ExecuteScript("document.getElementById('fromBidDt').value = arguments[0];", startDate.ToString("yyyy/MM/dd"));
                js.ExecuteScript("document.getElementById('toBidDt').value = arguments[0];", endDate.ToString("yyyy/MM/dd"));

                Console.WriteLine("Clicking the search button...");
                var searchButton = wait.Until(drv => drv.FindElement(By.XPath("//*[@id='searchForm']/div/fieldset[1]/ul/li[4]/dl/dd[3]/a")));
                searchButton.Click();

                Console.WriteLine("Waiting for the search results to load...");
                await Task.Delay(3000); 
                // 결과 테이블이 로드될 때까지 기다림
                try
                {
                    wait.Until(drv => drv.FindElement(By.XPath("//*[@id='resultForm']/div[2]/table/tbody")));
                }
                catch (WebDriverTimeoutException)
                {
                    Console.WriteLine("Timed out waiting for the search results to load.");
                }
                // 검색 결과 가져오기
                var rows = wait.Until(drv => drv.FindElements(By.XPath("//*[@id='resultForm']/div[2]/table/tbody")));
                Console.WriteLine($"{rows.Count} rows found.");

                foreach (var row in rows)
                {
                    try
                    {
                        var titleNode = row.FindElement(By.XPath(".//*[@id='resultForm']/div[2]/table/tbody/tr/td[4]/div/a"));
                        var institutionNode = row.FindElement(By.XPath(".//*[@id='resultForm']/div[2]/table/tbody/tr/td[5]/div"));
                        var requestInstitutionNode = row.FindElement(By.XPath(".//*[@id='resultForm']/div[2]/table/tbody/tr/td[6]/div"));
                        var inputDateNode = row.FindElement(By.XPath(".//*[@id='resultForm']/div[2]/table/tbody/tr/td[8]/div"));

                        var post = new Post
                        {
                            Title = titleNode?.Text.Trim() ?? "No Title",
                            Keyword = keyword,
                            PostingInstitution = institutionNode?.Text.Trim() ?? "No Institution",
                            RequestingInstitution = requestInstitutionNode?.Text.Trim() ?? "No Request Institution",
                            InputDate = DateTime.TryParse(inputDateNode?.Text.Trim(), out var date) ? date : DateTime.MinValue,
                            SearchDate = DateTime.Now
                        };

                        if (post.InputDate >= startDate && post.InputDate <= endDate)
                        {
                            posts.Add(post);
                        }
                    }
                    catch (NoSuchElementException ex)
                    {
                        Console.WriteLine("Element not found: " + ex.Message);
                    }
                }
            }
            catch (WebDriverException ex)
            {
                Console.WriteLine("WebDriver exception: " + ex.Message);
            }
            finally
            {
                driver.Quit();
            }

            return posts;
        }

        public void SavePostsToExcel(List<Post> posts, string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                filePath = "장표.xlsx";
            }
            else if (!filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                filePath += ".xlsx";
            }
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Posts");

                worksheet.Cell(1, 1).Value = "공고명";
                worksheet.Cell(1, 2).Value = "공고기관";
                worksheet.Cell(1, 3).Value = "수요기관";
                worksheet.Cell(1, 4).Value = "입력일시";
                worksheet.Cell(1, 5).Value = "검색일시";

                for (int i = 0; i < posts.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = posts[i].Title;
                    worksheet.Cell(i + 2, 2).Value = posts[i].PostingInstitution;
                    worksheet.Cell(i + 2, 3).Value = posts[i].RequestingInstitution;
                    worksheet.Cell(i + 2, 4).Value = posts[i].InputDate.ToString("yyyy-MM-dd HH:mm");
                    worksheet.Cell(i + 2, 5).Value = posts[i].SearchDate.ToString("yyyy-MM-dd HH:mm");
                }

                workbook.SaveAs(filePath);
            }
        }
    }
}
