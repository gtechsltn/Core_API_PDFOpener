

using Core_API_DocOpener;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.AddCors(options =>
{
    options.AddPolicy("cors", policy =>
    {
        policy.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader();
    });
});

var app = builder.Build();

// Configure the HTTP request pipeline.
 
    app.UseSwagger();
    app.UseSwaggerUI();


app.UseHttpsRedirection();
app.UseCors("cors");


app.MapGet("/documents", async (HttpContext context, IWebHostEnvironment env) => {

  List<string> files = new List<string>();

    // 1. Get the directory path
    string dirPath = Path.Combine(env.ContentRootPath, "Files");
    // 2. Read Files
    var documents = Directory.GetFiles(dirPath).ToList();

    // 3. Extract file names
    files = documents.Select(f => Path.GetFileName(f)).ToList();

    return Results.Ok(files);

});

// The Endpoint

app.MapGet("/document/{file}", async (HttpContext context, IWebHostEnvironment env, string file) => {

    // 1. Get the directory path
    string dirPath = Path.Combine(env.ContentRootPath, "Files");
    // 2. Read Files
    var files = Directory.GetFiles(dirPath);

    var fileName = files.FirstOrDefault(f => f.Contains(file));

    // 3. If file is not found return NotFound
    if (fileName == null)
    {
        return Results.NotFound($"File {file}.docx is not available");
    }

    // 4. Define a memory Stream
    var memory = new MemoryStream();

    // 5. COpy the File into the MemoryStream
    using (var stream = new FileStream(fileName, FileMode.Open))
    {
        await stream.CopyToAsync(memory);
    }
    
    memory.Position = 0;
    // 6. The File Response
    return Results.File(memory, "application/pdf", fileName);
    // return Results.File(fileName, "application/octet-stream", file);    
});

app.Run();
 
