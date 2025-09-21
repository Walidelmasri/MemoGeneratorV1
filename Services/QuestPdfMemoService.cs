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

            var labelHex = HexOr(m.LabelColorHex, "#137B3C");
            var underlineHex = HexOr(m.UnderlineColorHex, "#137B3C");
            var bodyHex = "#000000";

            const float FieldWidthPt   = 420f;
            const float ThroughWidthPt = FieldWidthPt - 40f; // slightly narrower Through
            const float LineThickness  = 0.8f;
            const float FooterReservePt = 84f;

var baseText   = TextStyle.Default.FontSize(11).FontFamily("Tajawal").FontColor(bodyHex);
            if (!string.IsNullOrWhiteSpace(m.FontFamilyLatin))
                baseText = baseText.FontFamily(m.FontFamilyLatin);

var arabicText = TextStyle.Default.FontSize(11).FontFamily("Tajawal").FontColor(bodyHex);
            if (!string.IsNullOrWhiteSpace(m.FontFamilyArabic))
                arabicText = arabicText.FontFamily(m.FontFamilyArabic);

            byte[]? bannerBytes = (m.BannerImage is { Length: > 0 }) ? m.BannerImage : null;
            byte[]? footerBytes = (m.FooterImage is { Length: > 0 }) ? m.FooterImage : null;

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

                    page.Background().Element(bg =>
                    {
                        if (footerBytes != null)
                            bg.AlignBottom().Image(footerBytes).FitWidth();
                    });

                    page.Header().ShowOnce().PaddingTop(6).Element(h =>
                    {
                        if (bannerBytes != null)
                            h.Image(bannerBytes).FitWidth();
                    });

                    page.Content().PaddingBottom(FooterReservePt).Column(col =>
                    {
                        col.Spacing(8);

                        // Memo No. (left, smaller, no underline) + Date (right, with underline)
                        col.Item().AlignCenter().Width(FieldWidthPt).Row(r =>
                        {
                            r.RelativeItem().Column(c =>
                            {
                                c.Spacing(1);
                                c.Item().Row(rr =>
                                {
                                    rr.RelativeItem().Text(t => t.Span("Memo No.").SemiBold().FontColor(labelHex).FontSize(10));
                                    rr.RelativeItem().AlignRight().Text(t => t.Span("رقم المذكرة").Style(arabicText).SemiBold().FontColor(labelHex).FontSize(10));
                                });
                                c.Item().Text(t =>
                                {
                                    t.AlignLeft();
                                    t.Span(memoNumber).FontSize(10);
                                });
                            });

                            r.ConstantItem(24).Text("");

                            r.RelativeItem().Column(c =>
                            {
                                c.Spacing(2);
                                c.Item().Row(rr =>
                                {
                                    rr.RelativeItem().Text(t => t.Span("Date").SemiBold().FontColor(labelHex));
                                    rr.RelativeItem().AlignRight().Text(t => t.Span("التاريخ").SemiBold().FontColor(labelHex).Style(arabicText));
                                });
                                c.Item().PaddingTop(2).BorderBottom(LineThickness).BorderColor(underlineHex).PaddingBottom(2)
                                 .Text(t => { t.AlignCenter(); t.Span(dateText); });
                            });
                        });

                        col.Item().Height(2);

                        void FieldBlock(string en, string ar, string? value)
                        {
                            col.Item().AlignCenter().Width(FieldWidthPt).Column(b =>
                            {
                                b.Spacing(0);
                                b.Item().Row(r =>
                                {
                                    r.RelativeItem().Text(t => t.Span(en).SemiBold().FontColor(labelHex));
                                    r.RelativeItem().AlignRight().Text(t => t.Span(ar).SemiBold().FontColor(labelHex).Style(arabicText));
                                });
                                b.Item().BorderBottom(LineThickness).BorderColor(underlineHex)
                                    .Text(t => { t.DefaultTextStyle(ds => ds.LineHeight(1.35f)); t.AlignCenter(); t.Span(value ?? ""); });
                            });
                        }

                        void ThroughBlock(string en, string ar, string? value)
                        {
                            col.Item().AlignCenter().Width(ThroughWidthPt).Column(b =>
                            {
                                b.Spacing(0);
                                b.Item().Row(r =>
                                {
                                    r.RelativeItem().Text(t => t.Span(en).SemiBold().FontColor(labelHex));
                                    r.RelativeItem().AlignRight().Text(t => t.Span(ar).SemiBold().FontColor(labelHex).Style(arabicText));
                                });
                                b.Item().Text(t => { t.DefaultTextStyle(ds => ds.LineHeight(1.35f)); t.AlignCenter(); t.Span(value ?? ""); });
                            });
                        }

                        FieldBlock("To", "إلى", m.To);
                        ThroughBlock("Through", "بواسطة", m.Through);
                        FieldBlock("From", "من", m.From);
                        FieldBlock("Subject", "الموضوع", m.Subject);

                        // Body: per-line alignment
                        col.Item().AlignCenter().Width(FieldWidthPt).PaddingTop(10).Element(body =>
                        {
                            var text = (m.Body ?? "").Replace("\r\n", "\n");
                            body.Column(c =>
                            {
                                foreach (var raw in text.Split('\n'))
                                {
                                    var line = raw.TrimEnd();
                                    if (string.IsNullOrWhiteSpace(line)) { c.Item().Height(12); continue; }

                                    if (ContainsArabic(line))
                                    {
                                        c.Item().Text(t => { t.DefaultTextStyle(ds => ds.LineHeight(1.35f)); t.AlignRight(); t.Span(line).Style(arabicText); });
                                    }
                                    else
                                    {
                                        c.Item().Text(t => { t.DefaultTextStyle(ds => ds.LineHeight(1.35f)); t.AlignLeft(); t.Span(line); });
                                    }
                                }
                            });
                        });
                    });

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
