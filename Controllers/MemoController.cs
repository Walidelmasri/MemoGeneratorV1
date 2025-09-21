using System.IO;
using System.Threading.Tasks;
using MemoGenerator.Models;
using MemoGenerator.Services.Abstractions;
using MemoGenerator.ViewModels;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;

namespace MemoGenerator.Controllers
{
    public class MemoController : Controller
    {
        private readonly IMemoPdfService _pdf;
        private readonly IWebHostEnvironment _env;

        public MemoController(IMemoPdfService pdf, IWebHostEnvironment env)
        {
            _pdf = pdf;
            _env = env;
        }

        [HttpGet]
        public IActionResult Create() => View(new MemoCreateVm());

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(MemoCreateVm vm)
        {
            if (!ModelState.IsValid)
                return View(vm);

            // Load default images if present (optional)
            var bannerPath = Path.Combine(_env.WebRootPath, "images", "memo-banner.png");
            var footerPath = Path.Combine(_env.WebRootPath, "images", "memo-footer.png");
            byte[]? banner = System.IO.File.Exists(bannerPath) ? await System.IO.File.ReadAllBytesAsync(bannerPath) : null;
            byte[]? footer = System.IO.File.Exists(footerPath) ? await System.IO.File.ReadAllBytesAsync(footerPath) : null;

            // Always auto-generate memo number + date (no user input)
            string AutoMemoNumber() => $"M-{System.DateTime.UtcNow:yyyyMMdd-HHmmss}";
            string OrdinalDate(System.DateTime dt)
            {
                int d = dt.Day;
                string suffix = (d % 10 == 1 && d != 11) ? "st"
                              : (d % 10 == 2 && d != 12) ? "nd"
                              : (d % 10 == 3 && d != 13) ? "rd" : "th";
                return $"{dt:MMMM} {d}{suffix} {dt:yyyy}";
            }

            var doc = new MemoDocument
            {
                To = vm.To,
                Through = vm.Through,
                From = vm.From,
                Subject = vm.Subject,
                Body = vm.Body,
                Classification = vm.Classification,

                MemoNumber = AutoMemoNumber(),
                DateText   = OrdinalDate(System.DateTime.UtcNow),

                BannerImage = vm.UseDefaultImages ? banner : null,
                FooterImage = vm.UseDefaultImages ? footer : null
            };

            var pdf = await _pdf.GenerateAsync(doc);
            var fileName = $"Memo_{System.DateTime.UtcNow:yyyyMMdd_HHmm}.pdf";
            return File(pdf, "application/pdf", fileName);
        }
    }
}
