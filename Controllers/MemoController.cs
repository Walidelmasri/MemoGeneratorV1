using System;
using System.IO;
using System.Threading.Tasks;
using HtmlAgilityPack;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using MemoGenerator.Models;
using MemoGenerator.Services.Abstractions;
using MemoGenerator.ViewModels;

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

        // --- minimal sanitizer using HtmlAgilityPack (keeps text-align only) ---
        private static string SanitizeBody(string? html)
        {
            html ??= string.Empty;

            // Wrap so we always have a single root to work with
            var doc = new HtmlDocument();
            doc.LoadHtml($"<div id='root'>{html}</div>");
            var root = doc.GetElementbyId("root");

            // Allowed
            var allowedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "p","br","strong","b","em","i","u","ul","ol","li","div","span" };
            var allowedAlignValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "left","right","center","justify" };

            void CleanNode(HtmlNode node)
            {
                // Work on a copy since we'll modify the children
                var children = node.ChildNodes.ToArray();
                foreach (var child in children)
                {
                    if (child.NodeType == HtmlNodeType.Element)
                    {
                        if (!allowedTags.Contains(child.Name))
                        {
                            // unwrap: keep inner content, drop the tag
                            foreach (var gc in child.ChildNodes.ToArray())
                                node.InsertBefore(gc, child);
                            node.RemoveChild(child);
                            continue;
                        }

                        // strip all attributes except style:text-align
                        var style = child.GetAttributeValue("style", null);
                        child.Attributes.RemoveAll();

                        if (!string.IsNullOrEmpty(style))
                        {
                            // keep only text-align
                            string? alignValue = null;
                            var parts = style.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                            foreach (var p in parts)
                            {
                                var kv = p.Split(':', 2, StringSplitOptions.TrimEntries);
                                if (kv.Length == 2 && kv[0].Equals("text-align", StringComparison.OrdinalIgnoreCase))
                                {
                                    var val = kv[1].Trim().ToLowerInvariant();
                                    if (allowedAlignValues.Contains(val))
                                        alignValue = val;
                                }
                            }
                            if (!string.IsNullOrEmpty(alignValue))
                                child.SetAttributeValue("style", $"text-align:{alignValue}");
                        }

                        CleanNode(child); // recurse
                    }
                    else if (child.NodeType == HtmlNodeType.Comment)
                    {
                        node.RemoveChild(child);
                    }
                }
            }

            CleanNode(root);
            return root.InnerHtml;
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(MemoCreateVm vm)
        {
            if (!ModelState.IsValid)
                return View(vm);

            // Sanitize posted HTML body
            var safeBodyHtml = SanitizeBody(vm.Body);

            // Optional header/footer art
            byte[]? banner = null;
            byte[]? footer = null;

            var root = _env.WebRootPath ?? "wwwroot";
            var bannerPath = Path.Combine(root, "images", "memo-banner.png");
            var footerPath = Path.Combine(root, "images", "memo-footer.png");

            if (System.IO.File.Exists(bannerPath))
                banner = await System.IO.File.ReadAllBytesAsync(bannerPath);
            if (System.IO.File.Exists(footerPath))
                footer = await System.IO.File.ReadAllBytesAsync(footerPath);

            var doc = new MemoDocument
            {
                To = vm.To,
                Through = string.IsNullOrWhiteSpace(vm.Through) ? null : vm.Through,
                From = vm.From,
                Subject = vm.Subject,
                Body = safeBodyHtml,      // sanitized HTML
                Classification = vm.Classification,
                BannerImage = banner,
                FooterImage = footer
            };

            var pdf = await _pdf.GenerateAsync(doc);
            var fileName = $"Memo_{DateTime.UtcNow:yyyyMMdd_HHmm}.pdf";
            return File(pdf, "application/pdf", fileName);
        }
    }
}
