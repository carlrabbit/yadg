# Release and Packaging Engineering

## Status

Authoritative for YADG V1 package construction, consumer validation, and NuGet publication tooling introduced by M0011.

This document defines release engineering. It does not change YADG document-authoring semantics.

## V1 release target

The M0011 release candidate version is:

```text
1.0.0
```

The release candidate is considered ready only after the M0011 validation and human-review gates pass.

M0011 does not push `1.0.0` to an external NuGet feed.

## Supported distribution surface

YADG V1 has one supported distributable package:

```text
NuGet package ID: Yadg
Package type:      .NET tool
Tool command:      yadg
Target:            net10.0-windows
Architecture:      x64
Version:           1.0.0
```

The package is framework-dependent.

The supported consumer platform for V1 is Windows x64 with a compatible .NET 10 SDK/runtime environment.

The package does not promise Linux or macOS execution.

## Public API surface

The supported V1 product interface is:

```text
yadg command-line interface
workspace/file contracts documented by project specs
prepared DOCX template vocabulary
```

The C# projects/assemblies under `src/` are implementation details.

M0011 does not publish `Yadg.Core`, `Yadg.Word`, `Yadg.Renderer`, or `Yadg.WordRenderer` as separately supported NuGet libraries.

Every non-CLI project must be explicitly non-packable so an ordinary solution/repository pack cannot accidentally create public library packages.

## Version authority

Repository version authority for M0011 is a single central MSBuild property:

```text
VersionPrefix = 1.0.0
VersionSuffix = empty
```

The central property belongs in `Directory.Build.props` unless the live repository has an equivalent central MSBuild authority at implementation time.

The effective package version, assembly informational version presented by the CLI, release evidence, and documentation must agree on:

```text
1.0.0
```

M0011 adds a public:

```text
yadg --version
```

surface that prints the effective product/package semantic version.

It must not print a stale separately hard-coded value.

## Package metadata

The CLI project is configured as the single .NET tool package with at least:

```text
PackageId           Yadg
PackAsTool          true
ToolCommandName     yadg
Authors             carlrabbit
Description         Template-first document authoring from Markdown and prepared DOCX templates.
RepositoryType      git
RepositoryUrl       https://github.com/carlrabbit/yadg
PackageProjectUrl   https://github.com/carlrabbit/yadg
PackageReadmeFile   README.md
```

The packed README is the repository `README.md`.

Package tags and other ordinary metadata are implementation mechanics provided they do not make unsupported product claims.

### License metadata

The repository entering M0011 has no project license file or license expression.

M0011 must not invent or select a software license.

Therefore the release candidate must not claim a `PackageLicenseExpression` or `PackageLicenseFile` unless an explicit project-owner licensing decision has become repository authority before packaging implementation reaches that point.

The absence of a license claim must be surfaced in the release evidence and human release review.

Actual external publication remains out of scope for M0011.

## Word interop packaging

M0011 preserves the M0009 build decision:

```text
Windows + installed Word
-> full Visual Studio MSBuild
-> COMReference / ResolveComReference / tlbimp
-> generated managed interop assemblies
```

The packaged tool must contain every generated Word/Office interop wrapper assembly required to execute the installed tool.

The installed tool must not depend on the build machine's generated `obj`/`bin` locations.

Do not replace the generated-interoperability approach with an Office interop NuGet package, `dynamic`, handwritten COM interfaces, source generation, or another COM binding strategy as part of release packaging.

If `PackAsTool`/MSBuild packaging cannot carry the required generated interop assemblies into the installed tool, M0011 is blocked and must return to planning.

## Pack script

M0011 defines:

```powershell
./eng/pack.ps1
./eng/pack.ps1 -OutputPath <path>
```

Default output:

```text
artifacts/package
```

The script:

1. validates/uses the central release version;
2. restores required NuGet dependencies;
3. locates a supported full Visual Studio MSBuild capable of the M0009 `ResolveComReference` build;
4. fails clearly if that build capability is unavailable;
5. packs the CLI project as the `Yadg` .NET tool in Release configuration;
6. writes the package only to the requested/default package output directory;
7. verifies exactly one expected `Yadg.1.0.0.nupkg` release-candidate package is produced.

Do not use `dotnet pack` as an authoritative fallback when full MSBuild/COM-reference tooling is unavailable.

M0011 does not require or emit a symbols package.

The pack script must be usable non-interactively from a future CI/CD environment on an appropriately provisioned Windows build agent.

## Package-content verification

Release validation inspects the `.nupkg` rather than treating successful packing as sufficient evidence.

At minimum verify:

- NuGet package ID `Yadg`;
- package version `1.0.0`;
- package type is a .NET tool;
- command is `yadg`;
- target framework is the expected Windows .NET 10 tool target;
- repository/readme/author/description metadata is present and correct;
- repository README is packaged;
- CLI and all required YADG implementation assemblies are present;
- generated Word and Office Core interop wrapper assemblies required by the installed CLI are present;
- no test binaries, test fixtures, review evidence, repository secrets, source-tree `obj`/`bin` paths, or unrelated package artifacts are included;
- no unsupported package license claim is present unless project licensing authority exists.

## NuGet publish script

M0011 defines:

```powershell
./eng/publish-nuget.ps1 -Source <source>
./eng/publish-nuget.ps1 -Source <source> -PackagePath <path-to-nupkg>
```

`-Source` is mandatory.

There is no implicit/default NuGet.org source.

When `-PackagePath` is omitted, the script resolves the exact M0011 package from the default package output directory and fails on absence or ambiguity.

Before pushing, the script verifies that the package is exactly the expected package ID/version:

```text
Yadg 1.0.0
```

Credentials are never stored in the repository or command source.

For feeds requiring an API key, the script may use:

```text
NUGET_API_KEY
```

from the environment.

If that variable is absent, the script may rely on normal NuGet source authentication/configuration and must not invent a credential.

The script must not use `--skip-duplicate` for a stable release. An already-existing package/version is a publication condition that must be surfaced.

### M0011 publish-script validation

M0011 validates the publish script only against a temporary/local filesystem NuGet destination or equivalent non-external test target.

It must not push to NuGet.org or any real external/private production feed.

## Consumer validation

M0011 defines an authoritative Tier-4 release/consumer command:

```powershell
./eng/test-m0011-tier4.ps1
```

It runs on the same class of interactive Windows release workstation used for real Word validation and additionally requires LibreOffice for the representative packaged-tool renderer smoke.

The Tier-4 test:

1. runs/consumes the M0011 pack output;
2. creates isolated `DOTNET_CLI_HOME`, NuGet package/cache, tool-install, workspace, and delivery directories;
3. installs `Yadg` version `1.0.0` from the local package output using `dotnet tool install --tool-path ...`;
4. invokes the installed `yadg`, not `dotnet run` or repository build outputs;
5. verifies `yadg --version` reports `1.0.0`;
6. verifies `yadg --help` exposes the supported top-level commands;
7. runs representative `check` and `build` behavior on a realistic M0010-derived workspace;
8. finalizes representative authored DOCX through the installed package with the Microsoft Word renderer;
9. finalizes a representative authored DOCX through the installed package with the LibreOffice renderer;
10. runs `publish` through the installed package and verifies the delivered finalized DOCX;
11. proves the installed tool resolves the packaged YADG and generated Office interop assemblies without using repository `bin`/`obj`;
12. uninstalls/removes the isolated tool state best-effort;
13. writes release evidence.

Tier 4 is package/consumer evidence. It does not replace the complete M0010 four-path compatibility matrix.

## Lower-tier release prerequisites

On the M0011 release-candidate revision, authoritative release evidence includes successful:

```powershell
./eng/validate.ps1
./eng/test-m0010-tier3.ps1
./eng/test-m0011-tier4.ps1
```

`eng/test-m0010-tier3.ps1` re-proves the realistic four-path Word/LibreOffice matrix on the actual release-candidate revision.

The M0011 Tier-4 test proves the installed package wiring and consumer path.

## Release evidence

Implementation creates repository-local M0011 release evidence under a location such as:

```text
artifacts/release/evidence/M0011/
```

Exact filenames are implementation-owned.

Evidence must include, in machine-readable or plainly auditable form:

- repository revision;
- release version;
- package ID/version/path;
- package SHA-256;
- package metadata/content inspection result;
- Windows version;
- .NET SDK version;
- full MSBuild identity/version used to pack;
- observable Word version;
- observable LibreOffice version;
- M0010 Tier-3 result/evidence reference;
- isolated tool-install result;
- installed `yadg --version` output;
- installed `yadg --help` smoke result;
- installed-tool check/build/Word-render/LibreOffice-render/publish results;
- publish-script local-target test result;
- documentation audit result;
- explicit confirmation that no external NuGet publication occurred;
- explicit license-metadata state.

The `.nupkg` itself is a generated release candidate and is not required to be committed to source control.

The evidence must identify the exact package hash reviewed by the human gate.

## Public documentation contract

M0011 performs the V1 public documentation audit rather than deferring it.

### README

`README.md` becomes V1 user documentation and must be understandable without milestone history.

It covers, at minimum:

- what YADG is and the template-first ownership model;
- supported V1 platform;
- installation as the `Yadg` .NET tool;
- local-package installation example for development/release validation;
- intended feed installation form after publication;
- prerequisites and which capabilities require Word, LibreOffice, or an external Mermaid producer;
- workspace layout;
- minimal end-to-end `check -> build -> render -> publish` workflow;
- all public commands/options;
- `YADG.md` values/producer/publish configuration;
- `section`, `content`, direct figure/table, prepared table, and value template vocabulary;
- relative heading composition;
- supported visible value stories;
- figures/tables/captions/references;
- Word and LibreOffice renderer differences;
- PDF's actual status;
- publication semantics;
- security/trust boundaries for external producers and Office automation;
- important V1 limitations;
- troubleshooting guidance for missing Word/interop build tooling, Word runtime, LibreOffice, and producer failures.

The normal user-facing README must not require readers to understand milestone IDs such as M0003/M0006/M0010 to understand current behavior.

Historical milestone references may remain only where genuinely historical.

### CHANGELOG

M0011 adds a repository-root:

```text
CHANGELOG.md
```

with a `1.0.0` release entry summarizing the supported product at a user-relevant level.

It is not a copy of milestone history.

### CLI help/version

`yadg --help`, command-specific help, README command syntax, and current product behavior must agree.

`yadg --version` is a release surface and must agree with the package version.

### Samples, website, library API

M0011 does not create a separate sample application, website, or public C# API documentation set.

The realistic test fixtures remain compatibility fixtures, not public samples.

The README/workspace examples are the V1 onboarding surface.

## Diagnostic surface

Existing diagnostic codes remain user-visible troubleshooting identifiers.

M0011 documentation may reference important codes where useful, but exact diagnostic message prose is not a separately versioned public API.

M0011 does not redesign the diagnostic system.

## Release review

The release candidate is subject to blocking human review `HR-M0011-01`.

The review is performed against the exact package hash and release-candidate repository revision represented by the release evidence.

A changed package or changed release-candidate revision before approval requires refreshed evidence and another decision in the same active review.

After M0011 completes, the review is historical evidence and does not approve future releases.

## External publication

M0011 ends with a validated, reviewed `1.0.0` release candidate and working NuGet publication tooling.

It does not:

- push to NuGet.org;
- push to another production/private feed;
- create a GitHub Release;
- create a Git tag;
- create GitHub Actions or another CI/CD workflow;
- reserve/transfer a NuGet package prefix/ownership;
- choose a project software license.

Those actions require an explicit later release/publication decision.
