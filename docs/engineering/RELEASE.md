# Release and Packaging Engineering

## Status

Authoritative for YADG V1 release engineering after the M0011 late-bound-COM correction.

This supersedes the earlier M0011 requirement to preserve MSBuild `COMReference` / `ResolveComReference` / `tlbimp` Office interop generation.

## V1 release target

```text
YADG 1.0.0
```

M0011 produces a validated release candidate. It does not publish externally, create a GitHub Release/tag, or add GitHub workflows.

## Supported distribution

```text
Package ID:   Yadg
Package type: .NET tool
Command:      yadg
Target:       net10.0-windows
Architecture: x64
Version:      1.0.0
```

The package is framework-dependent. Only the CLI project is packable; source libraries are implementation details.

## Version authority

Use one central MSBuild version:

```text
VersionPrefix = 1.0.0
VersionSuffix = empty
```

Package version, `yadg --version`, documentation, and release evidence must agree.

## Microsoft Word automation

### Required binding

V1 uses built-in .NET/Windows late-bound COM only.

Activate Word through:

```text
Type.GetTypeFromProgID("Word.Application")
-> Activator.CreateInstance
-> late-bound Word object model
```

Use `dynamic` or equivalent reflection-based late binding.

Do not use:

- Office `COMReference`;
- `ResolveComReference`;
- `tlbimp`;
- `Microsoft.Office.Interop.Word` or `Microsoft.Office.Core` NuGet packages;
- generated `Interop.Microsoft.Office.Interop.Word.dll`;
- generated `Interop.Microsoft.Office.Core.dll`;
- handwritten/source-generated Office COM interfaces.

### Behavioral preservation

Only the binding changes. Preserve the existing Word renderer contract:

- desktop Word installed and COM-registered;
- interactive logged-on user session;
- STA execution;
- owned Word instance;
- non-visible automation/alert suppression;
- macro suppression where practical;
- bounded timeout behavior;
- field/index/pagination finalization;
- finalized DOCX output;
- clean close/quit best-effort;
- actionable diagnostics;
- no unattended/service/server support claim.

Do not introduce a helper process or redesign finalization semantics merely for this correction unless late binding cannot satisfy existing real-Word tests.

### Package/licensing consequence

The YADG package must contain no Microsoft Office binary or generated Office interop wrapper.

Microsoft Word is an external, separately licensed runtime prerequisite.

YADG remains MIT licensed. Third-party dependencies remain subject to their own licenses.

## Build and pack

The full-Visual-Studio-MSBuild interop-generation requirement is removed.

Authoritative release construction is ordinary .NET SDK tooling:

```powershell
dotnet restore Yadg.slnx
dotnet build Yadg.slnx --configuration Release
dotnet pack src/Yadg.Cli/Yadg.Cli.csproj --configuration Release --no-build --output <path>
```

Canonical wrapper:

```powershell
./eng/pack.ps1
./eng/pack.ps1 -OutputPath <path>
```

Default output:

```text
artifacts/package
```

`eng/pack.ps1` must validate version, remove stale expected RC output, restore/build normally, pack only the CLI project, and produce exactly one `Yadg.1.0.0.nupkg`.

Do not retain custom package ZIP rewriting/TFM normalization unless an actual .NET tool packaging defect remains after interop removal and is documented in the execution ledger.

## Package metadata

Required metadata remains:

```text
PackageId               Yadg
PackAsTool              true
ToolCommandName         yadg
Authors                 carlrabbit
RepositoryType          git
RepositoryUrl           https://github.com/carlrabbit/yadg
PackageProjectUrl       https://github.com/carlrabbit/yadg
PackageReadmeFile       README.md
PackageLicenseExpression MIT
```

README/license wording must not imply Microsoft Word or third-party components are relicensed under MIT.

## Package-content validation

Inspect the `.nupkg`. Verify expected YADG/runtime dependencies and reject tests, fixtures, review evidence, secrets, unrelated build output, and Office interop binaries.

Specifically forbidden:

```text
Interop.Microsoft.Office.Interop.Word.dll
Interop.Microsoft.Office.Core.dll
Microsoft.Office.Interop.Word.dll
Microsoft.Office.Core.dll
```

No equivalent generated Office wrapper may be substituted under another name.

## NuGet push tooling

Retain:

```powershell
./eng/publish-nuget.ps1 -Source <source>
./eng/publish-nuget.ps1 -Source <source> -PackagePath <path>
```

`-Source` is mandatory. There is no implicit NuGet.org source. No credential is committed. `NUGET_API_KEY` may come from the environment. Do not use `--skip-duplicate`.

M0011 validates this only against a local/non-external destination.

## Tier 4 consumer validation

Canonical command:

```powershell
./eng/test-m0011-tier4.ps1
```

Install the local `Yadg 1.0.0` package into an isolated tool directory and invoke the installed command.

Validate:

1. package metadata/content and absence of Office interop assemblies;
2. isolated `dotnet tool install`;
3. installed `yadg --version`;
4. root and command-specific help;
5. representative `check`;
6. representative `build`;
7. real Microsoft Word render through late-bound `Word.Application`;
8. real LibreOffice render;
9. `publish`;
10. no dependency on repository `bin`/`obj`;
11. release evidence.

The installed package must successfully render through real Word while containing no Office interop wrapper.

## CLI help release surface

Validate actual help, not merely command-name presence:

```text
yadg --help
yadg check --help
yadg build --help
yadg render --help
yadg publish --help
```

Help must accurately communicate command purpose, workspace default, renderer IDs/default, LibreOffice-only renderer path, publish-path precedence, and supported version flag behavior.

README syntax must agree with actual help.

## System.CommandLine

Evaluate upgrading the existing prerelease `System.CommandLine` dependency to the current stable 2.x version available at implementation time.

Upgrade unless focused tests demonstrate a material incompatibility with YADG's CLI contract. If incompatible, record evidence and escalate rather than silently retaining the prerelease package.

## README contract

README is V1 user documentation. It must include:

- product/template ownership model;
- supported platform and runtime prerequisites;
- installation;
- a minimal complete workspace example;
- minimal Markdown example;
- minimal visible prepared-DOCX template example;
- end-to-end `check -> build -> render -> publish`;
- all public commands/options;
- values/producer/publish configuration;
- a concrete Mermaid producer configuration example;
- template vocabulary;
- relative heading composition;
- supported value stories;
- references/captions;
- Word/LibreOffice differences;
- PDF status;
- publish semantics;
- trust/limitations/troubleshooting.

Normal users must not be told they need Visual Studio MSBuild, `ResolveComReference`, or `tlbimp`.

Do not mention DMS.

## CHANGELOG contract

Until actual external release/tag publication, use:

```text
## [Unreleased]
```

for the pending V1 entry.

A later explicitly authorized release action may convert it to:

```text
## 1.0.0 — <actual release date>
```

The changelog summarizes user-visible capabilities and limitations, not milestone implementation history.

## Release evidence

Earlier PR11 evidence/package hashes are invalid after this correction.

Fresh evidence must record:

- corrected repository revision;
- package path/hash;
- package inspection;
- explicit absence of Office interop assemblies;
- .NET/Windows provenance;
- Word/LibreOffice versions;
- late-bound Word smoke;
- M0010 Tier-3 evidence on the corrected revision;
- installed help/version/check/build/render/publish;
- local publish-script validation;
- README/changelog audit;
- external publication = none.

## Human release review

Retain one narrow blocking release review only for subjective judgment that automation cannot decide.

The human reviews:

1. whether the V1 README is understandable and materially accurate;
2. whether the pending V1 changelog/release description is acceptable;
3. whether the exact validated RC is acceptable to mark release-ready for a later publication action.

Package hash, repository revision, runtime versions, and automated results are derived automatically by review tooling. The human does not transcribe them.

## External publication

Still excluded:

- no external NuGet push;
- no GitHub Release;
- no release tag;
- no GitHub workflow.
