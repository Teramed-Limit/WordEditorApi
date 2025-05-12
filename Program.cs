using WordEditorApi.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Add services to the container.
builder.Services.AddScoped<IDocumentService, DocumentService>();

// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
// Add services to the container.
builder.Services.AddControllers();

// 配置 CORS 以允許所有來源的請求
builder.Services.AddCors(options =>
{
    // var allowedOrigins = configuration.GetSection("AllowedOrigins").Get<string[]>();
    options.AddPolicy(
        "AllowAll",
        policy =>
        {
            policy.WithOrigins(["http://localhost:5173", "https://localhost:5173"]).AllowAnyMethod().AllowAnyHeader().AllowCredentials();
        }
    );
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

if (!Directory.Exists(Path.Combine(Directory.GetCurrentDirectory(), "Documents")))
    Directory.CreateDirectory(Path.Combine(Directory.GetCurrentDirectory(), "Documents"));

// 允許前端與 OnlyOffice 存取文件
app.UseStaticFiles();

app.UseHttpsRedirection();

// 使用 CORS
app.UseCors("AllowAll");


app.MapControllers();

app.Run();

