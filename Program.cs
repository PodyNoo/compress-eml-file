using MimeKit;
using Multipart = MimeKit.Multipart;
using ICSharpCode.SharpZipLib.Zip;
using ZipFile = ICSharpCode.SharpZipLib.Zip.ZipFile;
using SkiaSharp;
using QPdfNet;
using QPdfNet.Enums;

class Program
{
	private static readonly int _jpegQuality = 75; // TODO: add as parameter
	private static readonly List<Task> _runningTasks = [];
	private static readonly object _lock = new();

	static void Main(string[] args)
	{
		if (args.Length < 1)
		{
			Console.WriteLine("Usage: compress-eml-file.exe <path to folder>");
			return;
		}

		string path = args[0];

		if (!Directory.Exists(path))
		{
			Console.WriteLine($"Folder '{path}' does not exist.");
			return;
		}

		string[] emlFiles = Directory.GetFiles(path, "*.eml");
		if (emlFiles.Length == 0)
		{
			Console.WriteLine("No .eml files found in the specified folder.");
			return;
		}

		Console.WriteLine("start of compress-eml-file.exe");

		string outputPath = Path.Combine(Path.GetDirectoryName(path)!, "output"); // TODO: add as parameter
		Directory.CreateDirectory(outputPath);

		foreach (string? emlFile in emlFiles)
		{
			if (emlFile == null) continue;
			// REWORK: structure this, don't do on the fly + separate logs of each tasks and wait end of task work to ouput logs
			Task task = Task.Run(() =>
			{
				lock (_lock)
				{
					Console.BackgroundColor = ConsoleColor.DarkBlue;
					Console.ForegroundColor = ConsoleColor.White;
					Console.WriteLine($"Processing: {emlFile}");
					Console.ResetColor();
				}
				string emlFileOutput = Path.Combine(outputPath, $"{Path.GetFileNameWithoutExtension(emlFile)}_compressed.eml");

				MimeMessage message;

				try
				{
					message = MimeMessage.Load(emlFile);
				}
				catch (Exception ex)
				{
					Console.WriteLine("Failed to load email\n" + ex.ToString());
					return;
				}

				List<Task> tasks = [];

				CompressMimeEntity(message.Body, tasks);

				Task.WaitAll(tasks.ToArray());

				try
				{
					using (FileStream fs = File.Create(emlFileOutput))
						message.WriteTo(fs);
					Console.WriteLine($"Processed email saved to: {Path.GetFullPath(emlFileOutput)}");
				}
				catch (Exception ex)
				{
					Console.WriteLine("Failed to save processed email\n" + ex.ToString());
				}
			});
			_runningTasks.Add(task);
		}
		Task.WaitAll(_runningTasks.ToArray());
		Console.WriteLine("end of compress-eml-file.exe");
	}

	static void CompressMimeEntity(MimeEntity? entity, List<Task> tasks)
	{
		if (entity == null) return;
		if (entity is Multipart multipart)
		{
			foreach (MimeEntity? part in multipart)
			{
				CompressMimeEntity(part, tasks);
			}
		}
		else if (entity is MimePart part)
		{
			Action? action = null;
			if (part.ContentType.MediaType.StartsWith("image", StringComparison.OrdinalIgnoreCase))
			{
				action = () => CompressImagePart(part);
			}
			else if (part.ContentType.MediaType.Equals("application", StringComparison.OrdinalIgnoreCase))
			{
				if (part.ContentType.MediaSubtype.Equals("pdf", StringComparison.OrdinalIgnoreCase))
				{
					action = () => CompressPdfPart(part);
				}
				else if (part.ContentType.MediaSubtype.Equals("x-zip-compressed", StringComparison.OrdinalIgnoreCase) || part.ContentType.MediaSubtype.Equals("zip", StringComparison.OrdinalIgnoreCase))
				{
					action = () => CompressZipPart(part);
				}
			}
			if (action != null)
			{
				tasks.Add(Task.Run(action));
			}
		}
	}

	static void CompressImagePart(MimePart part)
	{
		string partFilename = part.FileName ?? "inline image";
		try
		{
			using MemoryStream originalStream = new();
			part.Content.DecodeTo(originalStream);
			originalStream.Position = 0;
			Stream? compressedImageStream = CompressImage(originalStream, _jpegQuality);
			if (compressedImageStream == null) return;
			part.Content = new MimeContent(compressedImageStream, ContentEncoding.Default);
			part.ContentType.MediaSubtype = "jpeg";
			if (!string.IsNullOrEmpty(part.FileName))
			{
				part.FileName = Path.GetFileNameWithoutExtension(part.FileName) + ".jpg";
			}
			Console.WriteLine($"Compressed image: {partFilename}");
		}
		catch (Exception ex)
		{
			Console.WriteLine($"Error processing image: {partFilename}\n" + ex.ToString());
		}
	}

	static Stream? CompressImage(MemoryStream imageStream, int quality)
	{
		//TODO: needs to check if format is supported by SkiaSharp before, or maybe not needed if it is only returning null instead of throwing exception.
		imageStream.Position = 0;
		SKImage image = SKImage.FromEncodedData(imageStream);
		SKData? encoded = image?.Encode(SKEncodedImageFormat.Jpeg, quality);
		return encoded?.AsStream();
	}

	static void CompressPdfPart(MimePart part)
	{
		string partFilename = part.FileName ?? "inline .pdf";
		try
		{
			MemoryStream originalStream = new();
			part.Content.DecodeTo(originalStream);
			originalStream.Position = 0;

			MemoryStream optimizedPdf = OptimizePdf(originalStream, part.FileName);
			part.Content = new MimeContent(optimizedPdf, ContentEncoding.Default);

			Console.WriteLine($"Compressed .pdf: {partFilename}");
		}
		catch (Exception ex)
		{
			Console.WriteLine($"Error processing .pdf: {partFilename}\n" + ex.ToString());
		}
	}

	static MemoryStream OptimizePdf(MemoryStream pdfStream, string? filename = null)
	{
		string tempDirectory = Path.Combine(Path.GetTempPath(), "compress-eml-file");
		Directory.CreateDirectory(tempDirectory);
		string tempFilename = Path.GetFileNameWithoutExtension(SanitizeFilename(filename, "temp_pdf")) + "_" + DateTime.Now.ToFileTimeUtc();
		string inputPath = Path.Combine(tempDirectory, tempFilename + ".pdf");
		string outputPath = Path.Combine(tempDirectory, tempFilename + "_optimized.pdf");
		File.WriteAllBytes(inputPath, pdfStream.ToArray());

		Job qpdfJob = new();
		ExitCode exitCode = qpdfJob
			.InputFile(inputPath)
			.OptimizeImages()
			.KeepInlineImages()
			.CompressionLevel(9)
			.RecompressFlate()
			.CompressStreams(true)
			.Linearize()
			.NoWarn()
			.RemoveUnreferencedResources()
			.OutputFile(outputPath)
			.Run(out string? logs);

		if (exitCode == ExitCode.ErrorsFoundFileNotProcessed)
		{
			Console.WriteLine($"Error happened while using QPdfNet to optimize pdf: {inputPath}");
			if (!string.IsNullOrWhiteSpace(logs))
			{
				Console.WriteLine(logs);
			}
		}

		MemoryStream result = new(File.ReadAllBytes(outputPath));
		File.Delete(inputPath);
		File.Delete(outputPath);
		return result;
	}

	static string SanitizeFilename(string? filename, string fallback)
	{
		filename ??= fallback;
		char[] invalids = Path.GetInvalidFileNameChars();
		return string.Join("_", filename.Split(invalids, StringSplitOptions.RemoveEmptyEntries)).TrimEnd('.');
	}

	static void CompressZipPart(MimePart part)
	{
		string partFilename = part.FileName ?? "inline .zip";
		using MemoryStream memoryStream = new();
		part.Content.DecodeTo(memoryStream);
		memoryStream.Position = 0;

		using ZipFile zipFile = new(memoryStream);
		MemoryStream finalMemoryStream = new();
		ZipOutputStream zipOutputStream = new(finalMemoryStream);
		zipOutputStream.SetLevel(9);

		foreach (ZipEntry zipEntry in zipFile)
		{
			if (zipEntry.IsDirectory) continue;
			using Stream zipStream = zipFile.GetInputStream(zipEntry);
			// Check if this is the entry we want to modify (an image)
			// TODO: we could also had pdf and re - use OptimizePdf()
			ZipEntry? newZipEntry = null;
			if (MimeTypes.GetMimeType(zipEntry.Name).StartsWith("image"))
			{
				using MemoryStream fileMemoryStream = new();
				zipStream.CopyTo(fileMemoryStream);
				fileMemoryStream.Position = 0;

				Stream? imageStream = CompressImage(fileMemoryStream, _jpegQuality);
				if (imageStream != null)
				{
					if (imageStream is not MemoryStream ms)
					{
						ms = new MemoryStream();
						imageStream.CopyTo(ms);
					}

					byte[] imageBytes = ms.GetBuffer();

					newZipEntry = new(Path.GetFileNameWithoutExtension(zipEntry.Name) + ".jpg")
					{
						DateTime = zipEntry.DateTime,
						Size = imageBytes.Length
					};

					zipOutputStream.PutNextEntry(newZipEntry);
					zipOutputStream.Write(imageBytes, 0, imageBytes.Length);
					Console.WriteLine($"Compressed image inside .zip ({partFilename}) : {zipEntry.Name}");
				}
			}
			if (zipEntry == null)
			{
				// If not the modified entry or can't compress image, copy as-is
				zipOutputStream.PutNextEntry(zipEntry);
				zipStream.CopyTo(zipOutputStream);
			}
		}
		zipOutputStream.Finish();
		part.Content = new MimeContent(finalMemoryStream, ContentEncoding.Default);
	}
}