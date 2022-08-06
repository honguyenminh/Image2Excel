using Cocona;
using Image2Excel;
using Image2Excel.CommandLine;
using Microsoft.Extensions.DependencyInjection;
using Version = Image2Excel.Core.Metadata.Version;

var version = new Version(Major: 0, Minor: 2, IsPreRelease: true);

var builder = CoconaApp.CreateBuilder();
builder.Services.AddSingleton<ConsoleLogger>();
builder.Services.AddSingleton(version);

var app = builder.Build();
app.AddCommands<MainCommand>();

app.Run();