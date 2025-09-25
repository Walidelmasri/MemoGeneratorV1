using System;
using System.Threading.Tasks;
using HtmlAgilityPack;
using MemoGenerator.Models;
using MemoGenerator.Services.Abstractions;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace MemoGenerator.Services
{
    public sealed class QuestPdfMemoService : IMemoPdfService
    {
        // ---- RTL detector ----
        static bool ContainsArabic(string? s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            foreach (var ch in s)
            {
                int code = ch;
                if ((code >= 0x0600 && code <= 0x06FF) ||
                    (code >= 0x0750 && code <= 0x077F) ||
                    (code >= 0x08A0 && code <= 0x08FF))
                    return true;
            }
            return false;
        }

        // Explicit RTL on node or fallback to content detection
        static bool IsRtlNode(HtmlNode node)
        {
            var dir = node.GetAttributeValue("dir", "").Trim().ToLowerInvariant();
            if (dir == "rtl") return true;

            var style = node.GetAttributeValue("style", "");
            if (!string.IsNullOrEmpty(style) &&
                style.IndexOf("direction:rtl", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;

            return ContainsArabic(node.InnerText);
        }

        // ---- Read text-align from tag style/attribute ----
        static string? ReadTextAlign(HtmlNode node)
        {
            var style = node.GetAttributeValue("style", "");
            if (!string.IsNullOrEmpty(style))
            {
                var parts = style.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                foreach (var p in parts)
                {
                    var kv = p.Split(':', 2, StringSplitOptions.TrimEntries);
                    if (kv.Length == 2 && kv[0].Equals("text-align", StringComparison.OrdinalIgnoreCase))
                        return kv[1].ToLowerInvariant();
                }
            }
            var alignAttr = node.GetAttributeValue("align", null);
            return alignAttr?.ToLowerInvariant();
        }

        // ---- NEW: detect a blank paragraph/div (<p><br></p>, only &nbsp;, or whitespace) ----
        static bool IsBlankBlock(HtmlNode block)
        {
            if (!(block.Name.Equals("p", StringComparison.OrdinalIgnoreCase) ||
                  block.Name.Equals("div", StringComparison.OrdinalIgnoreCase)))
                return false;

            // Remove zero-width and NBSP then check inner text
            var text = HtmlEntity.DeEntitize(block.InnerText)
                                  .Replace("\u200B", "")  // zero-width space
                                  .Replace("\u00A0", " ") // NBSP
                                  .Trim();

            if (text.Length > 0)
                return false;

            // Also consider inner HTML variants of a blank para
            var html = block.InnerHtml.Trim().ToLowerInvariant();
            return html == "" || html == "<br>" || html == "<br/>" || html == "<br />";
        }

        // ---- Inline renderer: <b>, <strong>, <i>, <em>, <u>, and <br> ----
        static void RenderInline(HtmlNode node, TextDescriptor t, TextStyle baseStyle, TextStyle arabicStyle)
        {
            if (node.NodeType == HtmlNodeType.Text)
            {
                var text = HtmlEntity.DeEntitize(node.InnerText);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    var style = ContainsArabic(text) ? arabicStyle : baseStyle;
                    t.Span(text).Style(style);
                }
                return;
            }

            var tag = node.Name.ToLowerInvariant();

            foreach (var child in node.ChildNodes)
            {
                // <br> -> blank line in the same paragraph
                if (child.NodeType == HtmlNodeType.Element &&
                    child.Name.Equals("br", StringComparison.OrdinalIgnoreCase))
                {
                    t.EmptyLine();
                    continue;
                }

                var style = baseStyle;

                if (tag is "strong" or "b")
                    style = style.SemiBold();

                if (tag is "em" or "i")
                    style = style.Italic();

                if (child.NodeType == HtmlNodeType.Text)
                {
                    var txt = HtmlEntity.DeEntitize(child.InnerText);
                    if (!string.IsNullOrWhiteSpace(txt))
                    {
                        var s = ContainsArabic(txt) ? arabicStyle : style;
                        var span = t.Span(txt).Style(s);
                        if (tag == "u") span.Underline();
                    }
                }
                else
                {
                    RenderInline(child, t, style, arabicStyle);
                }
            }
        }

        static void RenderParagraph(IContainer parent, HtmlNode block, float widthPt, TextStyle baseStyle, TextStyle arabicStyle)
        {
            // --- NEW: treat completely blank <p>/<div> as vertical space between paragraphs ---
            if (IsBlankBlock(block))
            {
                parent.Height(12);   // tweak spacing for one empty paragraph
                return;
            }

            var align = ReadTextAlign(block);
            var rtl = IsRtlNode(block);

            var alignToUse = string.IsNullOrWhiteSpace(align)
                ? (rtl ? "right" : "left")
                : align;

            var container = parent.AlignCenter().Width(widthPt);
            if (rtl) container = container.ContentFromRightToLeft();

            container.Text(t =>
            {
                t.DefaultTextStyle(ds => ds.LineHeight(1.35f));

                switch (alignToUse)
                {
                    case "justify":
                        t.Justify();       // if your QuestPDF version doesn’t support this, it’s ignored
                        break;
                    case "center":
                        t.AlignCenter();
                        break;
                    case "right":
                        t.AlignRight();
                        break;
                    case "left":
                    default:
                        t.AlignLeft();
                        break;
                }

                foreach (var child in block.ChildNodes)
                    RenderInline(child, t, baseStyle, arabicStyle);
            });
        }

        static void RenderList(IContainer parent, HtmlNode list, bool ordered, float widthPt, TextStyle baseStyle, TextStyle arabicStyle)
        {
            var items = list.SelectNodes("./li");
            if (items == null) return;

            int index = 1;
            foreach (var li in items)
            {
                var bullet = ordered ? $"{index}. " : "• ";
                var rtl = IsRtlNode(li);
                var align = ReadTextAlign(li) ?? (rtl ? "right" : "left");

                var rowContainer = parent.AlignCenter().Width(widthPt);
                if (rtl) rowContainer = rowContainer.ContentFromRightToLeft();

                rowContainer.Row(row =>
                {
                    row.ConstantItem(18).Text(b =>
                    {
                        switch (align)
                        {
                            case "center": b.AlignCenter(); break;
                            case "right":  b.AlignRight();  break;
                            default:       b.AlignLeft();   break;
                        }
                        b.Span(bullet);
                    });

                    row.RelativeItem().Text(t =>
                    {
                        t.DefaultTextStyle(ds => ds.LineHeight(1.35f));
                        switch (align)
                        {
                            case "justify": t.Justify();     break;
                            case "center":  t.AlignCenter(); break;
                            case "right":   t.AlignRight();  break;
                            default:        t.AlignLeft();   break;
                        }
                        foreach (var child in li.ChildNodes)
                            RenderInline(child, t, baseStyle, arabicStyle);
                    });
                });

                index++;
            }
        }

        static void RenderRichBody(IContainer container, string html, float widthPt, TextStyle baseStyle, TextStyle arabicStyle)
        {
            // Plain text fallback
            if (string.IsNullOrWhiteSpace(html) || html.IndexOf('<') < 0)
            {
                var text = (html ?? "").Replace("\r\n", "\n");
                container.Text(t =>
                {
                    t.DefaultTextStyle(ds => ds.LineHeight(1.35f));
                    foreach (var line in text.Split('\n'))
                    {
                        if (string.IsNullOrWhiteSpace(line))
                        {
                            t.EmptyLine();
                            continue;
                        }

                        var rtl = ContainsArabic(line);
                        if (rtl)
                        {
                            t.AlignRight();
                            t.Span(line).Style(arabicStyle);
                        }
                        else
                        {
                            t.AlignLeft();
                            t.Span(line).Style(baseStyle);
                        }
                    }
                });
                return;
            }

            // HTML path (Quill: <p>…</p>, blank lines as <p><br></p>)
            var doc = new HtmlDocument();
            doc.LoadHtml(html);

            var body = doc.DocumentNode.SelectSingleNode("//body") ?? doc.DocumentNode;
            var blocks = body.ChildNodes;

            container.Column(col =>
            {
                foreach (var node in blocks)
                {
                    if (node.NodeType == HtmlNodeType.Text && string.IsNullOrWhiteSpace(node.InnerText))
                        continue;
                    if (node.NodeType != HtmlNodeType.Element)
                        continue;

                    switch (node.Name.ToLowerInvariant())
                    {
                        case "p":
                        case "div":
                            RenderParagraph(col.Item(), node, widthPt, baseStyle, arabicStyle);
                            break;

                        case "br":
                            col.Item().Height(12); // explicit extra blank line between blocks
                            break;

                        case "ul":
                            RenderList(col.Item(), node, ordered: false, widthPt, baseStyle, arabicStyle);
                            break;

                        case "ol":
                            RenderList(col.Item(), node, ordered: true, widthPt, baseStyle, arabicStyle);
                            break;

                        default:
                            RenderParagraph(col.Item(), node, widthPt, baseStyle, arabicStyle);
                            break;
                    }
                }
            });
        }

        // ---- Main ----
        public Task<byte[]> GenerateAsync(MemoDocument m)
        {
            QuestPDF.Settings.License = LicenseType.Community;

            static string HexOr(string? hex, string fallback)
                => string.IsNullOrWhiteSpace(hex) ? fallback : hex!.Trim();

            static string OrdinalDate(DateTime dt)
            {
                int d = dt.Day;
                string suffix = (d % 10 == 1 && d != 11) ? "st"
                              : (d % 10 == 2 && d != 12) ? "nd"
                              : (d % 10 == 3 && d != 13) ? "rd" : "th";
                return $"{dt:MMMM} {d}{suffix} {dt:yyyy}";
            }

            // Colors (English lighter, Arabic darker)
            var labelHexEn   = "#1CA76A";
            var labelHexAr   = "#137B3C";
            var underlineHex = HexOr(m.UnderlineColorHex, "#137B3C");
            var bodyHex      = "#000000";

            // Layout
            const float FieldWidthPt      = 420f;
            const float ThroughWidthPt    = FieldWidthPt - 40f; // Through narrower, no underline
            const float MemoWidthPt       = 220f;               // left memo number block
            const float DateWidthPt       = 260f;               // centered date block
            const float LineThickness     = 0.8f;
            const float FooterReservePt   = 84f;

            // Spacing
            const float TopClusterSpacing = 2f;
            const float LabelBlockSpacing = 0f;
            const float UnderlinePadTop   = 0f;
            const float UnderlinePadBot   = 0f;
            const float FieldLineHeight   = 1.28f;
            const float GapAfterSubjectPt = 20f;

            // Text styles (Tajawal everywhere)
            var baseText   = TextStyle.Default.FontSize(11).FontFamily("Tajawal").FontColor(bodyHex);
            if (!string.IsNullOrWhiteSpace(m.FontFamilyLatin))
                baseText = baseText.FontFamily(m.FontFamilyLatin);

            var arabicText = TextStyle.Default.FontSize(11).FontFamily("Tajawal").FontColor(bodyHex);
            if (!string.IsNullOrWhiteSpace(m.FontFamilyArabic))
                arabicText = arabicText.FontFamily(m.FontFamilyArabic);

            // Assets
            byte[]? bannerBytes = (m.BannerImage is { Length: > 0 }) ? m.BannerImage : null;
            byte[]? footerBytes = (m.FooterImage is { Length: > 0 }) ? m.FooterImage : null;

            // Values
            string classificationValue = (m.Classification ?? "").Trim();
            string memoNumber = string.IsNullOrWhiteSpace(m.MemoNumber) ? $"M-{DateTime.UtcNow:yyyyMMdd-HHmmss}" : m.MemoNumber!.Trim();
            string dateText   = string.IsNullOrWhiteSpace(m.DateText)   ? OrdinalDate(DateTime.UtcNow)          : m.DateText!.Trim();

            var pdf = Document.Create(doc =>
            {
                doc.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(m.PageMarginPt);
                    page.DefaultTextStyle(baseText);

                    // Footer background art
                    page.Background().Element(bg =>
                    {
                        if (footerBytes != null)
                            bg.AlignBottom().Image(footerBytes).FitWidth();
                    });

                    // Header banner (first page)
                    page.Header().ShowOnce().PaddingTop(3).Element(h =>
                    {
                        if (bannerBytes != null)
                            h.Image(bannerBytes).FitWidth();
                    });

                    // Content (reserve footer space)
                    page.Content()
                        .PaddingBottom(FooterReservePt)
                        .Column(col =>
                    {
                        col.Spacing(TopClusterSpacing);

                        // Memo No. (left; no underline)
                        col.Item().AlignLeft().Width(MemoWidthPt).Column(c =>
                        {
                            c.Spacing(LabelBlockSpacing);
                            c.Item().Text(t => t.Span("Memo No.").SemiBold().FontColor(labelHexEn).FontSize(10));
                            c.Item().Text(t =>
                            {
                                t.AlignLeft();
                                t.Span(memoNumber).FontSize(10);
                            });
                        });

                        // Date (centered; underline)
                        col.Item().AlignCenter().Width(DateWidthPt).Column(c =>
                        {
                            c.Spacing(LabelBlockSpacing);
                            c.Item().Row(rr =>
                            {
                                rr.RelativeItem().Text(t => t.Span("Date").SemiBold().FontColor(labelHexEn));
                                rr.RelativeItem().AlignRight().Text(t => t.Span("التاريخ").Style(arabicText).SemiBold().FontColor(labelHexAr));
                            });
                            c.Item()
                             .PaddingTop(UnderlinePadTop)
                             .BorderBottom(LineThickness).BorderColor(underlineHex)
                             .PaddingBottom(UnderlinePadBot)
                             .Text(t =>
                             {
                                 t.AlignCenter();
                                 t.Span(dateText);
                             });
                        });

                        // Field helpers
                        void FieldBlock(string en, string ar, string? value)
                        {
                            col.Item().AlignCenter().Width(FieldWidthPt).Column(b =>
                            {
                                b.Spacing(LabelBlockSpacing);

                                b.Item().Row(r =>
                                {
                                    r.RelativeItem().Text(t => t.Span(en).SemiBold().FontColor(labelHexEn));
                                    r.RelativeItem().AlignRight().Text(t => t.Span(ar).Style(arabicText).SemiBold().FontColor(labelHexAr));
                                });

                                b.Item()
                                 .BorderBottom(LineThickness).BorderColor(underlineHex)
                                 .Text(t =>
                                 {
                                     t.DefaultTextStyle(ds => ds.LineHeight(FieldLineHeight));
                                     t.AlignCenter();
                                     t.Span(value ?? "");
                                 });
                            });
                        }

                        void ThroughBlock(string en, string ar, string? value)
                        {
                            col.Item().AlignCenter().Width(ThroughWidthPt).Column(b =>
                            {
                                b.Spacing(LabelBlockSpacing);

                                b.Item().Row(r =>
                                {
                                    r.RelativeItem().Text(t => t.Span(en).SemiBold().FontColor(labelHexEn));
                                    r.RelativeItem().AlignRight().Text(t => t.Span(ar).Style(arabicText).SemiBold().FontColor(labelHexAr));
                                });

                                // no underline
                                b.Item().Text(t =>
                                {
                                    t.DefaultTextStyle(ds => ds.LineHeight(FieldLineHeight));
                                    t.AlignCenter();
                                    t.Span(value ?? "");
                                });
                            });
                        }

                        // Fields
                        FieldBlock("To",       "إلى",     m.To);
                        ThroughBlock("Through","بواسطة", m.Through);
                        FieldBlock("From",     "من",      m.From);
                        FieldBlock("Subject",  "الموضوع", m.Subject);

                        // Gap before body
                        col.Item().Height(GapAfterSubjectPt);

                        // Body: HTML (preserve blank lines + RTL + justify)
                        col.Item()
                           .AlignCenter()
                           .Width(FieldWidthPt)
                           .Element(body =>
                        {
                            RenderRichBody(body, m.Body ?? "", FieldWidthPt, baseText, arabicText);
                        });
                    });

                    // Footer: classification line ALWAYS visible at top of footer area
                    page.Footer()
                        .PaddingTop(2)
                        .Column(f =>
                    {
                        var classText = string.IsNullOrWhiteSpace(classificationValue) ? "—" : classificationValue;
                        f.Item().Text(t =>
                        {
                            t.AlignCenter();
                            t.Span("(");
                            t.Span("Classification").SemiBold().FontColor(labelHexEn);
                            t.Span(": ");
                            t.Span(classText);
                            t.Span(")");
                        });
                    });
                });
            }).GeneratePdf();

            return Task.FromResult(pdf);
        }
    }
}
