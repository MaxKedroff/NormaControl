using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NormaControl.Services
{
    public class DocxFontChecker
    {
        public List<(int page, int line, string invalidFont, string paragraphText, string runText)> CheckFont(Stream stream)
        {
            var errors = new List<(int page, int line, string invalidFont, string paragraphText, string runText)>();

            try
            {
                using var wordDoc = WordprocessingDocument.Open(stream, false);
                var body = wordDoc.MainDocumentPart.Document.Body;
                int page = 1, line = 1;

                foreach (var paragraph in body.Elements<Paragraph>())
                {
                    // Получаем текст параграфа
                    string paragraphText = paragraph.InnerText;

                    foreach (var run in paragraph.Elements<Run>())
                    {
                        // Получаем текст из текущего Run
                        string runText = run.InnerText;

                        // Проверяем шрифт
                        var runProperties = run.RunProperties;
                        if (runProperties?.RunFonts?.Ascii != null)
                        {
                            var fontName = runProperties.RunFonts.Ascii.Value;
                            if (fontName != "Times New Roman")
                            {
                                errors.Add((page, line, fontName, paragraphText, runText));
                            }
                        }

                        // Проверка на разрыв страницы
                        if (run.Elements<Break>().Any(b => b.Type == BreakValues.Page))
                        {
                            page++;
                            line = 0; // Начинаем новую страницу
                        }

                        line++;
                    }
                }

                return errors;
            }
            catch (Exception e)
            {
                throw new Exception($"Ошибка чтения файла: {e.Message}");
            }
        }
    }
}
