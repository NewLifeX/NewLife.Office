using System;
using System.IO;
using System.Text;
using NewLife.Office;

var ms = new MemoryStream();
var w = new PdfWriter();
w.BeginPage();
w.DrawText("Test", 56, 780, 12);
w.EmbedFile("test.txt", Encoding.UTF8.GetBytes("Hello"));
w.Save(ms);
var text = Encoding.Latin1.GetString(ms.ToArray());
Console.WriteLine(text.Substring(0, Math.Min(3000, text.Length)));
