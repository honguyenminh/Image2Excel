using Cocona;
using Image2Excel.CommandLine;
using Microsoft.Extensions.Logging;

var builder = CoconaApp.CreateBuilder();
builder.Logging.AddDebug();
builder.Logging.AddConsole();

var app = builder.Build();
app.AddCommands<MainCommand>();

app.Run();