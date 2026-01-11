using Spectre.Console;
using Spectre.Console.Cli;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using UglyToad.PdfPig;

namespace PdfSuite.CLI.Commands;

public sealed class ExtractTextCommand : Command<ExtractTextCommand.Settings>
{
    public sealed class Settings : CommandSettings
    {
        [Description("Input PDF to extract text from")]
        [CommandArgument(0, "<input>")]
        public string Input { get; init; } = string.Empty;

        [Description("Optional output text file; if omitted, prints to console")]
        [CommandOption("-o|--output <FILE>")]
        public string? Output { get; init; }
    }

    public override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        try
        {
            var inputPath = Path.GetFullPath(settings.Input);
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"File not found: {inputPath}");

            var text = new System.Text.StringBuilder();
            using (var doc = PdfDocument.Open(inputPath))
            {
                foreach (var page in doc.GetPages())
                {
                    text.AppendLine(page.Text);
                }
            }

            if (!string.IsNullOrWhiteSpace(settings.Output))
            {
                var outPath = Path.GetFullPath(settings.Output);
                Directory.CreateDirectory(Path.GetDirectoryName(outPath) ?? ".");
                File.WriteAllText(outPath, text.ToString());
                AnsiConsole.MarkupLine("[green]Extracted text →[/] [yellow]{0}[/]", Markup.Escape(outPath));
            }
            else
            {
                AnsiConsole.WriteLine(text.ToString());
            }

            return 0;
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine("[red]Extract failed:[/] {0}", Markup.Escape(ex.Message));
            return -1;
        }
    }
}
