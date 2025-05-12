namespace WordEditorApi.DTOs;
using System.ComponentModel;

public class CreateDocumentRequest
{
    public string? Title { get; set; }
}

public class DocumentResponse
{
    public string Id { get; set; }
    public string FileName { get; set; }
    public string FileType { get; set; }
    public string Url { get; set; }
    public string CreatedAt { get; set; }
    public string UpdatedAt { get; set; }
}

public class DocumentConfig
{
    public string DocumentType { get; set; }
    public string Type { get; set; }
    public DocumentInfo Document { get; set; }
    public EditorConfig EditorConfig { get; set; }
    public string Height { get; set; }
    public string Width { get; set; }
    public string Token { get; set; }
}

public class DocumentInfo
{
    public string FileType { get; set; }
    public string Key { get; set; }
    public ReferenceData ReferenceData { get; set; }
    public string Title { get; set; }
    public string Url { get; set; }
    public Info Info { get; set; }
    public Permissions Permissions { get; set; }
}

public class ReferenceData
{
    public string FileKey { get; set; }
    public string InstanceId { get; set; }
    public string Key { get; set; }
}

public class Info
{
    public string Owner { get; set; }
    public string Uploaded { get; set; }
    public bool? Favorite { get; set; }
    public string Folder { get; set; }
    public List<object> SharingSettings { get; set; }
}

public class Permissions
{
    public bool? ChangeHistory { get; set; }
    public bool? Chat { get; set; }
    public bool? Comment { get; set; }
    public object CommentGroups { get; set; }
    public bool? Copy { get; set; }
    public bool? DeleteCommentAuthorOnly { get; set; }
    public bool? Download { get; set; }
    public bool? Edit { get; set; }
    public bool? EditCommentAuthorOnly { get; set; }
    public bool? FillForms { get; set; }
    public bool? ModifyContentControl { get; set; }
    public bool? ModifyFilter { get; set; }
    public bool? Print { get; set; }
    public bool? Protect { get; set; }
    public bool? Rename { get; set; }
    public bool? Review { get; set; }
    public List<string> ReviewGroups { get; set; }
    public List<string> UserInfoGroups { get; set; }
}

public class EditorConfig
{
    public string CallbackUrl { get; set; }
    public string Mode { get; set; }
    public string Lang { get; set; }
    public string Region { get; set; }
    public User User { get; set; }
    public Customization Customization { get; set; }
    public Embedded Embedded { get; set; }
    public Plugins Plugins { get; set; }
}

public class User
{
    public string Id { get; set; }
    public string Name { get; set; }
    public string Group { get; set; }
    public string Image { get; set; }
    public string Firstname { get; set; }
    public string Lastname { get; set; }
}

public class Customization
{
    // 依照 IConfig.customization 結構補齊
}

public class Embedded
{
    // 依照 IConfig.editorConfig.embedded 結構補齊
}

public class Plugins
{
    // 依照 IConfig.editorConfig.plugins 結構補齊
}

public enum DocumentMode
{
    [Description("Edit")]
    Edit,
    [Description("FillForms")]
    FillForms,
}

