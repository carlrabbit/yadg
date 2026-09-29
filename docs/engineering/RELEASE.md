# Release and Packaging Engineering

## Status

Authoritative for YADG 1.1 release engineering after M0013.

Historical release milestones remain historical evidence; this document describes the current release target and reusable packaging contract.

## Current release target

```text
YADG 1.1.0
```

M0013 produces a validated release candidate and does not publish externally.

## Distribution

```text
Package ID:        Yadg
Package type:      .NET tool
Command:           yadg
Target framework:  net10.0
Platform target:   x64
Version:           1.1.0
```

The package is framework-dependent.

Only the CLI project is packable.

## Platform policy

The package intentionally uses:

```text
net10.0
```

rather than `net10.0-windows`.

The authoritative supported/validated full-product locus remains Windows x64.

The portable TFM does not imply that every renderer is supported on every .NET 10 platform.

Office-independent CLI/authoring code must not become compile-time Windows-only solely because the product includes a Microsoft Word renderer.

### Microsoft Word runtime boundary

Word finalization is Windows-only.

A Word render request on a non-Windows runtime must fail through a normal actionable YADG diagnostic before calling Windows-only APIs such as:

- `Thread.SetApartmentState`;
- COM activation;
- Word object-model automation.

Do not rely on an uncaught `PlatformNotSupportedException` as the platform contract.

YOLO may attempt the alternate renderer after the normal Word platform failure according to the existing resilience specification.

M0013 does not certify full non-Windows product support.

## Central version authority

Use the central MSBuild product version:

```text
VersionPrefix = 1.1.0
VersionSuffix = empty
```

The following must agree:

- package version;
- packed nuspec;
- installed `yadg --version`;
- README release version;
- Tier-4 evidence;
- release review evidence.

Release scripts should derive the current version from central authority instead of duplicating literal version strings where practical.

## Microsoft Word binding

YADG uses built-in .NET/Windows late-bound COM:

```text
Type.GetTypeFromProgID("Word.Application")
-> Activator.CreateInstance
-> late-bound Word object model
```

Do not use:

- Office `COMReference`;
- `ResolveComReference`;
- `tlbimp`;
- `Microsoft.Office.Interop.Word` / `Microsoft.Office.Core` NuGet packages;
- generated Office interop wrappers;
- handwritten/source-generated Office COM interfaces.

The package contains no Microsoft Office binary or generated wrapper.

Microsoft Word is a separately licensed external runtime prerequisite.

## Word runtime behavior

Preserve:

- interactive logged-on Windows user session;
- installed/activated/COM-registered desktop Word;
- STA execution on Windows;
- YADG-owned Word instance;
- non-visible automation/alert suppression;
- macro suppression where practical;
- bounded timeout behavior;
- field/index/pagination finalization;
- finalized DOCX output;
- clean close/quit best effort;
- actionable diagnostics;
- no unattended service/server automation claim.

## Build and pack

Canonical release construction:

```powershell
./eng/pack.ps1
```

Equivalent underlying shape:

```powershell
dotnet restore Yadg.slnx
dotnet build Yadg.slnx --configuration Release
dotnet pack src/Yadg.Cli/Yadg.Cli.csproj --configuration Release --no-build --output <path>
```

Default package output:

```text
artifacts/package/Yadg.1.1.0.nupkg
```

`eng/pack.ps1` must:

1. read/validate the central product version;
2. remove stale package outputs that would make selection ambiguous;
3. restore/build through ordinary .NET SDK tooling;
4. pack only the CLI project;
5. produce exactly one current `Yadg.<version>.nupkg`.

Do not reintroduce custom package rewriting or Office type-library generation.

## Required package metadata

```text
PackageId                Yadg
PackAsTool               true
ToolCommandName          yadg
Authors                  carlrabbit
RepositoryType           git
RepositoryUrl            https://github.com/carlrabbit/yadg
PackageProjectUrl        https://github.com/carlrabbit/yadg
PackageReadmeFile        README.md
PackageLicenseExpression MIT
```

README/license wording must not imply Microsoft Word, LibreOffice, Mermaid tooling, or other third-party components are relicensed under MIT.

## Package-content validation

Inspect the `.nupkg`.

Require:

- DotnetTool package type;
- current ID/version;
- packaged `README.md`;
- packaged `LICENSE`;
- expected YADG assemblies/runtime content.

Reject:

- Office interop/generated wrappers;
- tests/fixtures;
- `.review`/`.execution`;
- release/test evidence;
- `.git`;
- repository `bin`/`obj`;
- obvious secrets/credentials;
- unrelated build output.

## NuGet push tooling

Retain:

```powershell
./eng/publish-nuget.ps1 -Source <source>
./eng/publish-nuget.ps1 -Source <source> -PackagePath <path>
```

`-Source` is mandatory.

There is no implicit NuGet.org source.

No credential is committed.

`NUGET_API_KEY` may be supplied from the environment when the selected source requires it.

Before push, inspect the nuspec and require the expected package ID/current central version.

M0013 validates this script only against a local/non-external NuGet destination.

## Tier 2

Canonical repository validation:

```powershell
./eng/validate.ps1
```

This remains the ordinary source-tree validation layer.

## Tier 3

Current authoritative real Office integration for the 1.1 feature line:

```powershell
./eng/test-m0012-tier3.ps1
```

Required locus:

```text
Windows x64
interactive user session
Microsoft Word installed/activated/COM-registered
LibreOffice Writer installed
```

Tier-3 evidence must belong to the same repository revision used for the release candidate.

## Tier 4 — packaged consumer validation

Canonical command:

```powershell
./eng/test-m0013-tier4.ps1
```

Install the freshly packed `Yadg 1.1.0` into an isolated tool directory and invoke the installed command.

Validate:

1. package metadata/content;
2. absence of Office interop wrappers;
3. isolated .NET tool installation;
4. installed `yadg --version`;
5. root and command-specific help;
6. `init`;
7. strict `check --list`;
8. `inspect styles`;
9. `inspect template`;
10. strict `build`;
11. representative YOLO `check/build` recovery;
12. real Microsoft Word render;
13. real LibreOffice render;
14. representative YOLO renderer fallback;
15. downstream-only `publish`;
16. no dependency on repository `bin/obj`;
17. exact release evidence.

Tier 4 proves the installed consumer surface. It does not replace Tier 3.

## Help release surface

Validate actual installed help:

```text
yadg --help
yadg init --help
yadg check --help
yadg inspect --help
yadg inspect styles --help
yadg inspect template --help
yadg build --help
yadg render --help
yadg publish --help
```

Help must accurately communicate:

- workspace defaults;
- `check --list`;
- `--yolo` scope;
- renderer IDs/default;
- LibreOffice `--renderer-path`;
- publication destination option;
- inspection commands.

README/examples must agree with actual help.

## Documentation release surface

The M0013 planning package supplies the intended documentation content.

Release implementation verifies rather than rewrites it.

Verify:

- command examples execute/parse;
- YAML examples match schema;
- Markdown examples match grammar;
- template examples match tag/front-matter/prototype contracts;
- version/platform/runtime statements are accurate;
- YOLO behavior/examples match implementation.

Only narrow factual corrections discovered by validation are allowed without returning to planning.

## Changelog

Before external publication, keep the prepared 1.1 content under:

```text
## [Unreleased]
```

Do not fabricate a release date.

An explicitly authorized publication operation later converts it to:

```text
## 1.1.0 — <actual release date>
```

## Release evidence

Write ignored evidence under:

```text
artifacts/release/evidence/M0013/
```

Record at least:

- exact repository revision;
- working tree status;
- release version;
- package path/hash;
- package inspection result;
- package content summary;
- .NET SDK;
- Windows version;
- Word version;
- LibreOffice version;
- M0012 Tier-3 evidence reference/revision;
- isolated tool-install result;
- installed version/help;
- bootstrap/strict authoring/inspection/build;
- YOLO authoring;
- Word/LibreOffice render;
- YOLO renderer fallback;
- publish;
- documentation/example audit;
- external publication = none.

If production code/package/docs change, regenerate affected evidence.

## Human release review

M0013 owns:

```text
HR-M0013-01
```

The human judges:

- documentation usability/accuracy for 1.1;
- prepared changelog/release description;
- acceptability of the exact validated RC.

Hashes, revisions, runtime versions and automated pass/fail data are bound automatically.

Canonical gate:

```powershell
./eng/review-check.ps1 --milestone M0013
```

## External publication

Excluded from M0013:

- NuGet.org push;
- GitHub Release;
- Git tag;
- publication workflow changes.

A later explicitly authorized release operation consumes the validated candidate.
