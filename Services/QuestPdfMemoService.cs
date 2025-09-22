using System;
using System.Threading.Tasks;
using MemoGenerator.Models;
using MemoGenerator.Services.Abstractions;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using QuestPDF.Helpers;

namespace MemoGenerator.Services
{
    public sealed class QuestPdfMemoService : IMemoPdfService
    {
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

            // colors
            var labelHex     = HexOr(m.LabelColorHex,     "#137B3C");
            var underlineHex = HexOr(m.UnderlineColorHex, "#137B3C");
            var bodyHex      = "#000000";

            // layout
            const float FieldWidthPt      = 420f;
            const float ThroughWidthPt    = FieldWidthPt - 40f; // Through a bit narrower (no underline)
            const float MemoWidthPt       = 220f;               // left memo block width
            const float DateWidthPt       = 260f;               // centered date block width
            const float LineThickness     = 0.8f;
            const float FooterReservePt   = 84f;

            // NEW: tighter knobs
            const float TopClusterSpacing = 2f;   // was 4 → tighter
            const float LabelBlockSpacing = 0f;   // internal label/value block spacing
            const float UnderlinePadTop   = 0f;   // was 1–2
            const float UnderlinePadBot   = 0f;   // was 1–2
            const float FieldLineHeight   = 1.28f;
            const float GapAfterSubjectPt = 20f;  // explicit space between Subject and Body

            // text styles
            var baseText   = TextStyle.Default.FontSize(11).FontFamily("Tajawal").FontColor(bodyHex);
            if (!string.IsNullOrWhiteSpace(m.FontFamilyLatin))
                baseText = baseText.FontFamily(m.FontFamilyLatin);

            var arabicText = TextStyle.Default.FontSize(11).FontFamily("Tajawal").FontColor(bodyHex);
            if (!string.IsNullOrWhiteSpace(m.FontFamilyArabic))
                arabicText = arabicText.FontFamily(m.FontFamilyArabic);

            // assets
            byte[]? bannerBytes = (m.BannerImage is { Length: > 0 }) ? m.BannerImage : null;
            byte[]? footerBytes = (m.FooterImage is { Length: > 0 }) ? m.FooterImage : null;

            // values
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

                    // background footer art
                    page.Background().Element(bg =>
                    {
                        if (footerBytes != null)
                            bg.AlignBottom().Image(footerBytes).FitWidth();
                    });

                    // header banner (first page)
                    page.Header().ShowOnce().PaddingTop(3).Element(h =>
                    {
                        if (bannerBytes != null)
                            h.Image(bannerBytes).FitWidth();
                    });

                    // content
                    page.Content()
                        .PaddingBottom(FooterReservePt)
                        .Column(col =>
                    {
                        col.Spacing(TopClusterSpacing); // tighter overall spacing

                        // --- Memo No. (left-aligned under header art, NO underline, smaller) ---
                        col.Item().AlignLeft().Width(MemoWidthPt).Column(c =>
                        {
                            c.Spacing(LabelBlockSpacing);
                            // English label only
                            c.Item().Text(t => t.Span("Memo No.").SemiBold().FontColor(labelHex).FontSize(10));
                            // value
                            c.Item().Text(t =>
                            {
                                t.AlignLeft();
                                t.Span(memoNumber).FontSize(10);
                            });
                        });

                        // --- Date (centered; same style with underline; minimal padding) ---
                        col.Item().AlignCenter().Width(DateWidthPt).Column(c =>
                        {
                            c.Spacing(LabelBlockSpacing);
                            c.Item().Row(rr =>
                            {
                                rr.RelativeItem().Text(t => t.Span("Date").SemiBold().FontColor(labelHex));
                                rr.RelativeItem().AlignRight().Text(t => t.Span("التاريخ").SemiBold().FontColor(labelHex).Style(arabicText));
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

                        // helpers
                        void FieldBlock(string en, string ar, string? value)
                        {
                            col.Item().AlignCenter().Width(FieldWidthPt).Column(b =>
                            {
                                b.Spacing(LabelBlockSpacing);

                                b.Item().Row(r =>
                                {
                                    r.RelativeItem().Text(t => t.Span(en).SemiBold().FontColor(labelHex));
                                    r.RelativeItem().AlignRight().Text(t => t.Span(ar).SemiBold().FontColor(labelHex).Style(arabicText));
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
                                    r.RelativeItem().Text(t => t.Span(en).SemiBold().FontColor(labelHex));
                                    r.RelativeItem().AlignRight().Text(t => t.Span(ar).SemiBold().FontColor(labelHex).Style(arabicText));
                                });

                                // no underline for Through
                                b.Item().Text(t =>
                                {
                                    t.DefaultTextStyle(ds => ds.LineHeight(FieldLineHeight));
                                    t.AlignCenter();
                                    t.Span(value ?? "");
                                });
                            });
                        }

                        // fields (tighter)
                        FieldBlock("To",       "إلى",     m.To);
                        ThroughBlock("Through","بواسطة", m.Through);
                        FieldBlock("From",     "من",      m.From);
                        FieldBlock("Subject",  "الموضوع", m.Subject);

                        // --- explicit breathing room between last field and body ---
                        col.Item().Height(GapAfterSubjectPt);

                        // body (per-line LTR/RTL)
                        col.Item()
                           .AlignCenter()
                           .Width(FieldWidthPt)
                           .Element(body =>
                        {
                            var text = (m.Body ?? "").Replace("\r\n", "\n");
                            body.Column(c2 =>
                            {
                                foreach (var raw in text.Split('\n'))
                                {
                                    var line = raw.TrimEnd();
                                    if (string.IsNullOrWhiteSpace(line))
                                    {
                                        c2.Item().Height(10);
                                        continue;
                                    }

                                    if (ContainsArabic(line))
                                    {
                                        c2.Item().Text(t => { t.DefaultTextStyle(ds => ds.LineHeight(1.35f)); t.AlignRight(); t.Span(line).Style(arabicText); });
                                    }
                                    else
                                    {
                                        c2.Item().Text(t => { t.DefaultTextStyle(ds => ds.LineHeight(1.35f)); t.AlignLeft(); t.Span(line); });
                                    }
                                }
                            });
                        });
                    });

                    // footer: compact classification line
                    page.Footer().Column(f =>
                    {
                        if (!string.IsNullOrWhiteSpace(classificationValue))
                        {
                            f.Item().PaddingBottom(6).Text(t =>
                            {
                                t.AlignCenter();
                                t.Span("(");
                                t.Span("Classification").SemiBold().FontColor(labelHex);
                                t.Span(": ");
                                t.Span(classificationValue);
                                t.Span(")");
                            });
                        }
                    });
                });
            }).GeneratePdf();

            return Task.FromResult(pdf);
        }
    }
}
