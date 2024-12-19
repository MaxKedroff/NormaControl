using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using NormaControl.Services; // Импорт вашего сервиса

namespace NormaControl.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocumentCheckController : ControllerBase
    {
        private readonly DocxFontChecker _fontChecker;

        public DocumentCheckController()
        {
            _fontChecker = new DocxFontChecker();
        }

        [HttpPost("check-font")]
        public IActionResult CheckFont([FromForm] IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("Файл не был загружен или он пуст.");
            }

            try
            {
                using var stream = file.OpenReadStream();
                var errors = _fontChecker.CheckFont(stream);

                if (errors.Any())
                {
                    return Ok(new
                    {
                        Message = "Обнаружены ошибки в документе.",
                        Errors = errors.Select(e => new
                        {
                            Page = e.page,
                            Line = e.line,
                            InvalidFont = e.invalidFont,
                            ParagraphText = e.paragraphText,
                            RunText = e.runText
                        })
                    });
                }
                else
                {
                    return Ok(new { Message = "Документ соответствует требованиям: все шрифты Times New Roman." });
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { Message = "Ошибка при обработке файла.", Error = ex.Message });
            }
        }

    }
}
