using Spectre.Console;
using Spectre.Console.Cli;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PdfSuite.CLI.Commands;

public sealed class SplitCommand : Command<SplitCommand.Settings>
{
    public sealed class Settings : CommandSettings
    {
        [Description("Input PDF to split")]
        [CommandArgument(0, "<input>")]
        public string Input { get; init; } = string.Empty;

        [Description("Output directory for single-page PDFs")]
        [CommandOption("-o|--outdir <DIR>")]
        public string? OutputDirectory { get; init; }

        public override ValidationResult Validate()
        {
            if (string.IsNullOrWhiteSpace(Input))
                return ValidationResult.Error("Input file is required");
            return ValidationResult.Success();
        }
    }

    public override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        try
        {
            var inputPath = Path.GetFullPath(settings.Input);
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"File not found: {inputPath}");

            var outDir = settings.OutputDirectory;
            if (string.IsNullOrWhiteSpace(outDir))
            {
                var name = Path.GetFileNameWithoutExtension(inputPath);
                outDir = Path.Combine(Path.GetDirectoryName(inputPath) ?? ".", $"{name}_pages");
            }

            Directory.CreateDirectory(outDir);

            using var input = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);
            for (int i = 0; i < input.PageCount; i++)
            {
                using var doc = new PdfDocument();
                doc.AddPage(input.Pages[i]);
                var outPath = Path.Combine(outDir, $"page-{i + 1}.pdf");
                doc.Save(outPath);
            }

            AnsiConsole.MarkupLine("[green]Split[/] {0} pages → [yellow]{1}[/]", input.PageCount, Markup.Escape(outDir));
            return 0;
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine("[red]Split failed:[/] {0}", Markup.Escape(ex.Message));
            return -1;
        }
    }
}
