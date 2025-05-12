using System.Text.Json;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.AcroForms;
using PdfSharpCore.Pdf.IO;
using WordEditorApi.DTOs;
using WordEditorApi.Utils;


namespace WordEditorApi.Services;

public interface IDocumentService
{
    List<DocumentResponse> GetDocuments();
    Task<DocumentResponse> CreateDocument(CreateDocumentRequest request);
    Task<object> GetEditorConfig(string id, string fileType, DocumentMode mode);
    Task ProcessCallback(string id, Stream requestBody);
}

public class DocumentService : IDocumentService
{
    private readonly string _storagePath;
    private readonly string _serverUrl;
    private readonly ILogger<DocumentService> _logger;

    public DocumentService(IConfiguration configuration, ILogger<DocumentService> logger)
    {
        _storagePath = configuration.GetSection("DocStoragePath").Value;
        _serverUrl = configuration.GetSection("DocStorageUrl").Value;
        _logger = logger;

        if (!Directory.Exists(_storagePath))
            Directory.CreateDirectory(_storagePath);
    }

    public List<DocumentResponse> GetDocuments()
    {
        return Directory.GetFiles(_storagePath)
            .Select(f =>
            {
                var fileInfo = new FileInfo(f);
                return new DocumentResponse
                {
                    Id = Path.GetFileNameWithoutExtension(f),
                    FileName = Path.GetFileName(f),
                    FileType = Path.GetExtension(f).TrimStart('.'),
                    Url = $"{_serverUrl}/{Path.GetFileName(f)}",
                    CreatedAt = fileInfo.CreationTime.ToString("yyyy-MM-ddTHH:mm:ss"),
                    UpdatedAt = fileInfo.LastWriteTime.ToString("yyyy-MM-ddTHH:mm:ss")
                };
            }).ToList();
    }

    public async Task<DocumentResponse> CreateDocument(CreateDocumentRequest request)
    {
        var newId = Guid.NewGuid().ToString();
        var fileName = string.IsNullOrEmpty(request?.Title)
            ? $"{newId}.docx"
            : $"{newId}_{DocumentUtils.NormalizeFileName(request.Title)}.docx";

        var filePath = Path.Combine(_storagePath, fileName);
        System.IO.File.Copy("Empty.docx", filePath);

        var documentKey = DocumentUtils.GenerateSimpleDocumentKey(fileName);
        var fileUrl = $"{_serverUrl}/{fileName}";

        return new DocumentResponse
        {
            Id = newId,
            FileName = fileName,
            FileType = "docx",
            Url = fileUrl,
            CreatedAt = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
            UpdatedAt = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss")
        };
    }

    public async Task<object> GetEditorConfig(string id, string fileType, DocumentMode mode)
    {
        var files = Directory.GetFiles(_storagePath)
            .Where(f => Path.GetFileNameWithoutExtension(f).StartsWith(id))
            .ToList();

        if (!files.Any())
            throw new FileNotFoundException($"找不到ID為 {id} 的文檔");

        var filePath = files.First();
        var fileNameExt = Path.GetFileName(filePath);
        var fileName = Path.GetFileNameWithoutExtension(filePath);
        var fileUrl = $"{_serverUrl}/{fileNameExt}";

        var fileInfo = new FileInfo(filePath);
        var lastModified = fileInfo.LastWriteTime;
        var version = 1;

        var documentKey = DocumentUtils.GenerateDocumentKey(fileNameExt, id, lastModified, version);


        FillForm(new Dictionary<string, string>
        {
            { "Text2", "預設值1" },
            { "CheckBox1", "Yes" },
            { "text2", "test" }
        }, filePath);

        return new
        {
            Document = new
            {
                FileType = fileType,
                Key = documentKey,
                Title = fileNameExt,
                Url = fileUrl,
                Permissions = new
                {
                    Download = true,
                    Print = true,
                    Copy = true,
                    Edit = mode == DocumentMode.Edit,
                    Review = mode == DocumentMode.Edit,
                    Comment = mode == DocumentMode.Edit,
                    FillForms = mode == DocumentMode.FillForms
                }
            },
            DocumentType = DocumentUtils.GetDocumentTypeByFileType(fileType),
            Type = "desktop",
            EditorConfig = new
            {
                CallbackUrl = $"http://localhost:5292/api/Document/{fileName}/callback",
                Mode = "edit",
                Lang = "zh-TW",
                Region = "zh-TW",
                User = new User { Id = "userId", Name = "userName" },
                Customization = new
                {
                    Forcesave = true,
                    SubmitForm = mode == DocumentMode.FillForms,
                },
            }
        };
    }

    public async Task ProcessCallback(string id, Stream requestBody)
    {
        using var reader = new StreamReader(requestBody);
        var body = await reader.ReadToEndAsync();
        var callbackJson = JsonDocument.Parse(body);
        var root = callbackJson.RootElement;

        var docKey = root.GetProperty("key").GetString();
        var status = root.GetProperty("status").GetInt32();

        var files = Directory.GetFiles(_storagePath)
            .Where(f => Path.GetFileNameWithoutExtension(f).StartsWith(id))
            .ToList();

        if (!files.Any())
            throw new FileNotFoundException($"找不到ID為 {id} 的文檔");

        var filePath = files.First();
        var fileName = Path.GetFileName(filePath);

        await ProcessCallbackStatus(status, root, filePath, fileName);
    }

    private async Task ProcessCallbackStatus(int status, JsonElement root, string filePath, string fileName)
    {
        switch (status)
        {
            case 1: // 用戶連接或斷開文檔協作
                await ProcessUserConnection(root, fileName);
                break;
            case 2: // 文檔準備保存
                // await ProcessDocumentSave(root, filePath, fileName);
                break;
            case 3: // 文檔保存錯誤
                _logger.LogError($"文檔 {fileName} 保存失敗");
                break;
            case 4: // 文檔關閉無更改
                _logger.LogInformation($"文檔 {fileName} 已關閉，無修改");
                break;
            case 6: // 強制保存
                await ProcessForceSave(root, filePath, fileName);
                break;
            case 7: // 強制保存錯誤
                _logger.LogError($"文檔 {fileName} 強制保存失敗");
                break;
            default:
                _logger.LogWarning($"未知狀態: {status}");
                break;
        }
    }

    private async Task ProcessUserConnection(JsonElement root, string fileName)
    {
        if (root.TryGetProperty("actions", out var actionsElement))
        {
            foreach (var action in actionsElement.EnumerateArray())
            {
                var actionType = action.GetProperty("type").GetInt32();
                var userId = action.GetProperty("userid").GetString();
                _logger.LogInformation($"文檔 {fileName}: 用戶 {userId} 執行了動作類型 {actionType}");
            }
        }
    }

    private async Task ProcessDocumentSave(JsonElement root, string filePath, string fileName)
    {
        await SaveDocumentFromUrl(root, filePath);

        if (root.TryGetProperty("history", out var historyElement))
        {
            await SaveDocumentHistory(historyElement, fileName);
        }

        if (root.TryGetProperty("users", out var usersElement))
        {
            var lastEditUser = usersElement.EnumerateArray().FirstOrDefault().GetString();
            _logger.LogInformation($"文檔 {fileName} 由用戶 {lastEditUser} 最後編輯");
        }
    }

    private async Task ProcessForceSave(JsonElement root, string filePath, string fileName)
    {
        if (root.TryGetProperty("forcesavetype", out var forceSaveTypeElement))
        {
            var forceSaveType = forceSaveTypeElement.GetInt32();
            await ProcessForceSaveType(forceSaveType, root);
        }

        // 表單提交，要更換名字
        if (forceSaveTypeElement.GetInt32() == 3)
        {
            var extension = Path.GetExtension(filePath);
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
            var newFilePath = Path.Combine(_storagePath, $"{fileNameWithoutExtension}_form{extension}");
            await SaveDocumentFromUrl(root, newFilePath);
            return;
        }

        await SaveDocumentFromUrl(root, filePath);
    }

    private async Task ProcessForceSaveType(int forceSaveType, JsonElement root)
    {
        switch (forceSaveType)
        {
            case 3: // 表單提交
                if (root.TryGetProperty("formsdataurl", out var formsDataUrlElement))
                {
                    var formsDataUrl = formsDataUrlElement.GetString();
                    await ProcessFormData(formsDataUrl);
                }

                break;
        }
    }

    private async Task SaveDocumentFromUrl(JsonElement root, string filePath)
    {
        if (root.TryGetProperty("url", out var urlElement))
        {
            var downloadUrl = urlElement.GetString();
            using var httpClient = new HttpClient();
            var fileBytes = await httpClient.GetByteArrayAsync(downloadUrl);
            await System.IO.File.WriteAllBytesAsync(filePath, fileBytes);

            if (root.TryGetProperty("changesurl", out var changesUrlElement))
            {
                var changesUrl = changesUrlElement.GetString();
                await SaveChangesHistory(changesUrl, filePath);
            }
        }
    }

    private async Task SaveDocumentHistory(JsonElement historyElement, string fileName)
    {
        // TODO: 實現歷史記錄保存邏輯
        _logger.LogInformation($"保存文檔 {fileName} 的歷史記錄");
    }

    private async Task SaveChangesHistory(string changesUrl, string filePath)
    {
        try
        {
            var historyDir = Path.Combine(Path.GetDirectoryName(filePath), "history");
            if (!Directory.Exists(historyDir))
            {
                Directory.CreateDirectory(historyDir);
            }

            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            var historyFilePath = Path.Combine(historyDir, $"{fileName}_{timestamp}.zip");

            using var httpClient = new HttpClient();
            var fileBytes = await httpClient.GetByteArrayAsync(changesUrl);
            await System.IO.File.WriteAllBytesAsync(historyFilePath, fileBytes);

            _logger.LogInformation($"已保存文檔變更歷史至 {historyFilePath}");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "保存變更歷史時發生錯誤");
        }
    }

    private async Task ProcessFormData(string formsDataUrl)
    {
        try
        {
            using var httpClient = new HttpClient();
            var jsonString = await httpClient.GetStringAsync(formsDataUrl);
            var formData = JsonDocument.Parse(jsonString);

            // foreach (var formField in formData.RootElement.EnumerateArray())
            // {
            //     var key = formField.GetProperty("key").GetString();
            //     var tag = formField.GetProperty("tag").GetString();
            //     var value = formField.GetProperty("value").GetString();
            //     var type = formField.GetProperty("type").GetString();
            //
            //     await ProcessFormField(type, key, value, tag);
            // }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "處理表單數據時發生錯誤");
            throw;
        }
    }

    private async Task ProcessFormField(string type, string key, string value, string tag)
    {
        switch (type?.ToLower())
        {
            case "text":
                _logger.LogInformation($"處理文本表單: 鍵={key}, 值={value}");
                break;
            case "checkbox":
                var isChecked = value?.ToLower() == "true";
                _logger.LogInformation($"處理複選框: 鍵={key}, 已選中={isChecked}");
                break;
            case "picture":
                if (!string.IsNullOrEmpty(value))
                {
                    _logger.LogInformation($"處理圖片: 鍵={key}");
                }

                break;
            case "combobox":
            case "dropdownlist":
                _logger.LogInformation($"處理下拉列表: 鍵={key}, 選中值={value}");
                break;
            case "datetime":
                if (DateTime.TryParse(value, out DateTime dateValue))
                {
                    _logger.LogInformation($"處理日期時間: 鍵={key}, 日期={dateValue:yyyy-MM-dd HH:mm:ss}");
                }

                break;
            case "radio":
                _logger.LogInformation($"處理單選按鈕組: 鍵={key}, 選中值={value}");
                break;
            default:
                _logger.LogWarning($"未處理的表單類型: {type}, 鍵: {key}, 值: {value}");
                break;
        }

        await LogFormSubmission(key, type, value, tag);
    }

    private async Task LogFormSubmission(string key, string type, string value, string tag)
    {
        var logMessage = new
        {
            Timestamp = DateTime.Now,
            FormKey = key,
            FormType = type,
            FormValue = value,
            FormTag = tag
        };

        var logJson = JsonSerializer.Serialize(logMessage);
        _logger.LogInformation($"表單提交日誌: {logJson}");
    }

    private string FillForm(Dictionary<string, string> data, string templatePath)
    {
        // 檢查檔案是否存在
        if (!File.Exists(templatePath))
        {
            throw new FileNotFoundException($"找不到範本檔案: {templatePath}");
        }

        using var inputDocument = PdfReader.Open(templatePath, PdfDocumentOpenMode.Modify);
        if (inputDocument.AcroForm == null)
            return templatePath;

        inputDocument.AcroForm.Elements.SetBoolean("/NeedAppearances", true);
        var form = inputDocument.AcroForm.Fields;

        foreach (var item in form.Names)
        {
            Console.WriteLine($"欄位名稱: {item} ");
        }

        // 遍歷所有表單欄位並填寫
        foreach (var field in data)
        {
            try
            {
                SetField(form, field.Key, field.Value);
            }
            catch (Exception ex)
            {
                _logger.LogError($"填寫欄位 {field.Key} 時發生錯誤: {ex.Message}");
            }
        }

        // 取得檔案名稱和路徑
        string fileName = Path.GetFileNameWithoutExtension(templatePath);
        string directory = Path.GetDirectoryName(templatePath);
        string newFileName = $"{fileName}_auto.pdf";
        string newFilePath = Path.Combine(directory, newFileName);

        // 儲存填寫後的PDF
        inputDocument.Save(newFilePath);

        return newFilePath;
    }

    private void SetField(PdfAcroField.PdfAcroFieldCollection fields, string name, string value)
    {
        if (fields[name] is PdfTextField textField)
        {
            textField.Value = new PdfString(value);
        }
    }
}