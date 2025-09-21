using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.IO;
using MemoGenerator.Services;
using MemoGenerator.Services.Abstractions;
using QuestPDF.Drawing;   // <-- IMPORTANT

var builder = WebApplication.CreateBuilder(args);

builder.Services
    .AddControllersWithViews()
    .AddMvcOptions(o => o.SuppressImplicitRequiredAttributeForNonNullableReferenceTypes = true);

builder.Services.AddSingleton<IMemoPdfService, QuestPdfMemoService>();

var app = builder.Build();

// Register Tajawal fonts for QuestPDF (reads from wwwroot/fonts)
var fontsDir = Path.Combine(app.Environment.WebRootPath ?? "wwwroot", "fonts");
var reg = Path.Combine(fontsDir, "Tajawal-Regular.ttf");
var bold = Path.Combine(fontsDir, "Tajawal-Bold.ttf");
if (File.Exists(reg))  FontManager.RegisterFont(File.OpenRead(reg));
if (File.Exists(bold)) FontManager.RegisterFont(File.OpenRead(bold));

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Memo}/{action=Create}/{id?}");

app.Run();
