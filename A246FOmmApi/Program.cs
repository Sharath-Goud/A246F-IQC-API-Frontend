using A246FOmmApi.Data;

var builder = WebApplication.CreateBuilder(args);

// Register controllers.
builder.Services.AddControllers();

// Register Swagger/OpenAPI.
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Register SQL Server connection factory.
builder.Services.AddSingleton<SqlConnectionFactory>();

var app = builder.Build();

// Enable Swagger for IIS testing.
app.UseSwagger();
app.UseSwaggerUI();

app.UseAuthorization();

app.MapControllers();

app.Run();