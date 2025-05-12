using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using WordEditorApi.Models;

namespace WordEditorApi.Controller;

[Route("api/[controller]")]
[ApiController]
public class DocumentsController : ControllerBase
{
    private readonly string storagePath = Path.Combine(Directory.GetCurrentDirectory(), "Documents");
    private readonly string serverUrl = "http://localhost:7144"; // 替換為你的實際 public 網址

    public DocumentsController()
    {
        if (!Directory.Exists(storagePath))
            Directory.CreateDirectory(storagePath);
    }

    // 取得檔案列表
    [HttpGet]
    public IActionResult GetDocuments()
    {
        var files = Directory.GetFiles(storagePath)
            .Select(f => new DocumentInfo
            {
                Id = Path.GetFileNameWithoutExtension(f),
                FileName = Path.GetFileName(f),
                Url = $"{serverUrl}/Documents/{Path.GetFileName(f)}"
            }).ToList();

        return Ok(files);
    }

    // 新增檔案
    [HttpPost]
    public IActionResult CreateDocument()
    {
        var newId = Guid.NewGuid().ToString();
        var fileName = $"{newId}.docx";
        var filePath = Path.Combine(storagePath, fileName);

        // 複製空白模板或創建空白檔案
        System.IO.File.Copy("Empty.docx", filePath); // 你需要準備一個空的 Word 檔案

        return Ok(new { Id = newId });
    }

    // 取得 OnlyOffice 編輯用設定
    [HttpGet("{id}/editor-config")]
    public IActionResult GetEditorConfig(string id)
    {
        var fileName = $"{id}.docx";
        var fileUrl = $"{serverUrl}/Documents/{fileName}";

        var config = new
        {
            document = new
            {
                fileType = "docx",
                key = id,
                title = fileName,
                url = fileUrl
            },
            editorConfig = new
            {
                mode = "edit",
                callbackUrl = $"{serverUrl}/api/docs/{id}/callback"
            }
        };

        return Ok(config);
    }

    // 接收 OnlyOffice 編輯結果（儲存）
    [HttpPost("{id}/callback")]
    public async Task<IActionResult> SaveCallback(string id)
    {
        using var reader = new StreamReader(Request.Body);
        var body = await reader.ReadToEndAsync();
        var json = JsonDocument.Parse(body);
        var status = json.RootElement.GetProperty("status").GetInt32();

        if (status == 2) // 已完成儲存
        {
            var downloadUrl = json.RootElement.GetProperty("url").GetString();
            var targetPath = Path.Combine(storagePath, $"{id}.docx");

            using var httpClient = new HttpClient();
            var fileBytes = await httpClient.GetByteArrayAsync(downloadUrl);
            await System.IO.File.WriteAllBytesAsync(targetPath, fileBytes);
        }

        return Ok(new { error = 0 });
    }
}