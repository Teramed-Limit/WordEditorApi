using Microsoft.AspNetCore.Mvc;
using WordEditorApi.DTOs;
using WordEditorApi.Services;

namespace WordEditorApi.Controller;

[Route("api/[controller]")]
[ApiController]
public class DocumentController : ControllerBase
{
    private readonly IDocumentService _documentService;

    public DocumentController(IDocumentService documentService)
    {
        _documentService = documentService;
    }

    [HttpGet]
    public IActionResult GetDocuments()
    {
        var documents = _documentService.GetDocuments();
        return Ok(documents);
    }

    [HttpPost]
    public async Task<IActionResult> CreateDocument([FromBody] CreateDocumentRequest request)
    {
        var document = await _documentService.CreateDocument(request);
        return Ok(document);
    }

    [HttpGet("{id}")]
    public async Task<IActionResult> GetEditorConfig(string id, string fileType, string mode)
    {
        try
        {
            if (!Enum.TryParse<DocumentMode>(mode, true, out var documentMode))
                return BadRequest($"無效的mode參數: {mode}");

            var config = await _documentService.GetEditorConfig(id, fileType, documentMode);
            return Ok(config);
        }
        catch (FileNotFoundException)
        {
            return NotFound();
        }
    }

    [HttpPost("{id}/callback")]
    public async Task<IActionResult> SaveCallback(string id)
    {
        try
        {
            await _documentService.ProcessCallback(id, Request.Body);
            return Ok(new { error = 0 });
        }
        catch (FileNotFoundException)
        {
            return NotFound(new { error = 1, message = "找不到指定的文檔" });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = 1, message = "處理回調時發生內部錯誤" });
        }
    }
}