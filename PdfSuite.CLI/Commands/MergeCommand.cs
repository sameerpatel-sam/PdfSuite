using Spectre.Console;
using Spectre.Console.Cli;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PdfSuite.CLI.Commands;

public sealed class MergeCommand : Command<MergeCommand.Settings>
{
    public sealed class Settings : CommandSettings
    {
        [Description("Output PDF file path")] 
        [CommandOption("-o|--output <FILE>")]
        public string? Output { get; init; }

        [Description("Input PDF files to merge (in order)")]
        [CommandArgument(0, "<inputs>")]
        public string[] Inputs { get; init; } = Array.Empty<string>();

        public override ValidationResult Validate()
        {
            if (string.IsNullOrWhiteSpace(Output))
                return ValidationResult.Error("--output is required");
            if (Inputs.Length < 2)
                return ValidationResult.Error("Provide at least two input PDFs to merge");
            return ValidationResult.Success();
        }
    }

    public override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        try
        {
            var outPath = Path.GetFullPath(settings.Output!);
            using var output = new PdfDocument();

            foreach (var file in settings.Inputs)
            {
                var path = Path.GetFullPath(file);
                if (!File.Exists(path))
                {
                    throw new FileNotFoundException($"Input not found: {path}");
                }

                using var input = PdfReader.Open(path, PdfDocumentOpenMode.Import);
                for (int i = 0; i < input.PageCount; i++)
                {
                    output.AddPage(input.Pages[i]);
                }
            }

            Directory.CreateDirectory(Path.GetDirectoryName(outPath) ?? ".");
            output.Save(outPath);
            AnsiConsole.MarkupLine("[green]Merged[/] {0} files → [yellow]{1}[/]", settings.Inputs.Length, Markup.Escape(outPath));
            return 0;
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine("[red]Merge failed:[/] {0}", Markup.Escape(ex.Message));
            return -1;
        }
    }
}
