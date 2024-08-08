// Models/Post.cs
using System;

namespace bluesolutions.Models
{
    public class Post
    {
        public string Title { get; set; } = string.Empty;
        public string Keyword { get; set; } = string.Empty;
        public string PostingInstitution { get; set; } = string.Empty;
        public string RequestingInstitution { get; set; } = string.Empty;
        public DateTime InputDate { get; set; }
        public DateTime SearchDate { get; set; }
    }
}
