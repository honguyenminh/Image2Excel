using Cocona;
using Image2Excel.CommandLine;
using Microsoft.Extensions.DependencyInjection;

var builder = CoconaApp.CreateBuilder();
builder.Services.AddSingleton<Image2Excel.Version>();
builder.Services.AddSingleton<ConsoleLogger>();

var app = builder.Build();
app.AddCommands<MainCommand>();

app.Run();