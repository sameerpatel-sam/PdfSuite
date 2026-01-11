using Spectre.Console;
using Spectre.Console.Cli;
using System.Diagnostics.CodeAnalysis;

namespace PdfSuite.CLI;

public static class Program
{
	public static int Main(string[] args)
	{
		var app = new CommandApp();

		app.Configure(config =>
		{
			config.SetApplicationName("pdfsuite");
			config.ValidateExamples();

			config.AddCommand<Commands.MergeCommand>("merge")
				.WithDescription("Merge multiple PDFs into one.")
				.WithExample(new[] { "merge", "-o", "out.pdf", "a.pdf", "b.pdf" });

			config.AddCommand<Commands.SplitCommand>("split")
				.WithDescription("Split a PDF into single-page PDFs.")
				.WithExample(new[] { "split", "-o", "outdir", "input.pdf" });

			config.AddCommand<Commands.ExtractTextCommand>("extract-text")
				.WithDescription("Extract text from a PDF (all pages by default).")
				.WithExample(new[] { "extract-text", "-o", "text.txt", "input.pdf" });

			config.SetInterceptor(new RootInterceptor());
		});

		try
		{
			return app.Run(args);
		}
		catch (Exception ex)
		{
			AnsiConsole.MarkupLine("[red]Error:[/] {0}", Markup.Escape(ex.Message));
			return -1;
		}
	}

	private sealed class RootInterceptor : ICommandInterceptor
	{
		public void Intercept([NotNull] CommandContext context, [NotNull] CommandSettings settings)
		{
			AnsiConsole.Write(new FigletText("PdfSuite").Centered().Color(Color.Blue));
		}
	}
}
