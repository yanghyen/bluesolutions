using Microsoft.AspNetCore.Mvc;
using System;
using System.Threading.Tasks;
using bluesolutions.Services;
using bluesolutions.Models;
using System.Collections.Generic;


namespace bluesolutions.Controllers
{
    public class PostsController : Controller
    {
        private readonly PostService _postService;

        public PostsController(PostService postService)
        {
            _postService = postService;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Index(DateTime startDate, DateTime endDate, string filePath)
        {
            const string keyword = "RPA";

            if (string.IsNullOrEmpty(filePath))
            {
                ViewBag.Message = "파일 경로를 입력해 주세요.";
                return View();
            }

            if (!filePath.EndsWith("장표.xlsx", StringComparison.OrdinalIgnoreCase))
            {
                filePath += "장표.xlsx";
            }

            try
            {
                var posts = await _postService.GetPostsAsync(keyword, startDate, endDate);
                _postService.SavePostsToExcel(posts, filePath);

                ViewBag.Message = "엑셀 파일이 성공적으로 저장되었습니다.";
            }
            catch (Exception ex)
            {
                ViewBag.Message = $"오류 발생: {ex.Message}";
            }

            return View();
        }
    }
}
