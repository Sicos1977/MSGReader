# Copilot Instructions

## Project Guidelines
even if the user communicates in Dutch.
- Release process: bump `Version`, `AssemblyVersion` (x.y.z.0) and `FileVersion` (x.y.z.0) in `MsgReader\MsgReader.csproj` and replace `PackageReleaseNotes` in the same file. The NuGet notes are a short bullet list using `•` characters, starting with a "Version x.y.z:" line (or a headline for security releases), and prefixed with `⚠ Breaking:` for breaking changes. GitHub release notes use Markdown with Changes, Installation (`dotnet add package MsgReader --version x.y.z`) and a Full changelog compare link.

## Coding Style
- Targets: net462, netstandard2.0, netstandard2.1, net8.0, net9.0, net10.0 with `LangVersion` latest. Use `#if` / `NETSTANDARD2_0_OR_GREATER` for framework-specific code and keep code compatible with the oldest target.
- Every .cs file starts with the standard header: file name, `Author: Kees van Spelde <sicos2002@hotmail.com>`, `Copyright (c) 2013-2026 Kees van Spelde. (www.magic-sessions.com)` and the MIT permission text. Copy it from an existing file.
- Use file-scoped namespaces (`namespace MsgReader.Outlook;`), with `using` directives above the namespace.
- Indentation is 4 spaces and braces are on their own line (Allman style).
- Use `var` for locals and modern C# patterns (e.g. `is not { CanWrite: true }`). Single-statement guard clauses may be written on one line: `if (x) return;`.
- Group members in `#region Fields`, `#region Properties`, and one region per method/feature (`#region WriteToLog` ... `#endregion`). The `#endregion` sits directly after the closing brace.
- Private fields use an underscore prefix (`_data`). Use `internal` for anything that is not public API; use `sealed` where possible.
- All public and internal members get XML documentation: `/// <summary>` with the text indented by 5 spaces after `///` (`///     Text`), plus `<param>`, `<remarks>`, `<see cref="..."/>` and `<c>true</c>` where relevant. Doc text is in English and starts with a capital.
- Add ReSharper suppressions as `// ReSharper disable ...` comments between the usings and the namespace.
- Custom exceptions live in `MsgReader\Exceptions`, are prefixed with `MR` (e.g. `MRMaxEmbeddedDepthReached`), derive from `Exception`, have the three standard constructors (parameterless, message, message + inner) which are `internal`, and carry an XML summary.
- Logging goes through `Logger.WriteToLog(...)`, which is internal and writes to a thread-static stream. Placeholders in log messages are wrapped in single quotes: `'{…}'`.
- Treat input files as untrusted: guard against deep nesting, oversized allocations, and truncated data, and throw typed exceptions.
- Localized strings are in `MsgReader\Localization\LanguageConsts*.resx`.
