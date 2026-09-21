# Milestone — M0011 V1.0 Release Readiness

## Release target

```text
YADG 1.0.0
```

M0011 produces a validated and human-approved release candidate.

M0011 does not publish that release candidate to an external NuGet feed, create a GitHub Release, create a tag, or add GitHub workflows.

## Execution profile

| Field | Value |
|---|---|
| Lifecycle state | ready |
| Milestone type | release-readiness |
| Mode | ai-executed-human-reviewed |
| Baseline implementation model | GPT-5.6 Luna |
| Repository role | product-tool |
| Repository maturity metadata | initial-implementation |
| Applicable profiles | repository-wide `base` + `artifact-first-runtime` |
| Release version | `1.0.0` |
| Distribution artifact | one `Yadg` .NET tool package |
| Supported consumer platform | Windows x64 / .NET 10 |
| Tier-2 command | `./eng/validate.ps1` |
| Tier-3 dependency | `./eng/test-m0010-tier3.ps1` |
| Tier-4 command | `./eng/test-m0011-tier4.ps1` |
| Human review | blocking `HR-M0011-01` |
| External publication | excluded |

## Goal

Turn the completed YADG feature set into a defensible V1 release candidate by closing the remaining release-engineering and public-documentation gap.

M0011 must:

1. establish one supported NuGet/.NET tool distribution artifact;
2. establish authoritative `1.0.0` versioning;
3. add deterministic Windows/MSBuild pack tooling compatible with the M0009 generated Office interop build;
4. add safe parameterized NuGet push tooling without performing an external release;
5. validate the package through the actual consumer installation mechanism;
6. audit and rewrite public documentation for V1 rather than milestone history;
7. create release evidence and pass a blocking human release review.

## Target state

A release operator on the authoritative Windows release workstation can execute:

```powershell
./eng/validate.ps1
./eng/test-m0010-tier3.ps1
./eng/pack.ps1
./eng/test-m0011-tier4.ps1
```

and obtain exactly one validated release-candidate package:

```text
Yadg.1.0.0.nupkg
```

The package installs through the .NET tool mechanism and exposes:

```text
yadg
```

The installed tool, rather than repository build output, successfully exercises representative:

```text
--version
--help
check
build
render --renderer word
render --renderer libreoffice
publish
```

behavior.

A separate push script exists for later authorized publication:

```powershell
./eng/publish-nuget.ps1 -Source <explicit-source>
```

but M0011 does not invoke it against an external/production feed.

## Scope

### Package and versioning

- one `Yadg` package;
- `PackAsTool=true`;
- tool command `yadg`;
- version `1.0.0`;
- Windows x64 / `net10.0-windows`;
- framework-dependent .NET tool;
- central MSBuild version authority;
- public `yadg --version`;
- repository/package metadata and README inclusion;
- non-CLI projects explicitly non-packable;
- generated Word/Office interop wrappers included in the package.

### Release scripts

- `eng/pack.ps1`;
- configurable package output path;
- full Visual Studio MSBuild/COMReference build/pack path;
- explicit failure when required MSBuild interop tooling is absent;
- `eng/publish-nuget.ps1`;
- explicit NuGet source;
- no hardcoded credentials;
- no default external feed;
- local/non-external push-script validation.

### Consumer validation

- `eng/test-m0011-tier4.ps1`;
- isolated .NET tool installation from the generated local package;
- package metadata/content inspection;
- installed command/version/help;
- installed-tool representative workspace lifecycle;
- installed Word and LibreOffice renderer smoke;
- installed publish smoke;
- release evidence.

### Public documentation

- complete README audit/rewrite for current V1;
- installation/prerequisites;
- end-to-end workflow;
- public CLI/options;
- current workspace/template contracts;
- relative heading composition;
- document-wide value substitution;
- structured content/references;
- Mermaid producer configuration/trust;
- renderer differences and prerequisites;
- PDF's actual status;
- publishing semantics;
- supported platform/limitations/troubleshooting;
- removal of milestone-dependent user documentation;
- root `CHANGELOG.md` with a `1.0.0` entry.

### Release review

- release-candidate evidence;
- package review;
- documentation review;
- blocking `HR-M0011-01`;
- no waiver.

## Non-goals

M0011 does not:

- add new document-authoring features;
- change M0010 heading/value semantics;
- replace the M0009 generated Office interop approach;
- publish separate YADG library NuGet packages;
- make YADG cross-platform;
- make a self-contained/native executable distribution;
- add MSI/MSIX/ZIP installers;
- publish PDF artifacts as a product distribution format;
- publish to NuGet.org or another production feed;
- create a GitHub Release or Git tag;
- create GitHub Actions or other CI/CD workflows;
- choose or invent a software license;
- create a public website;
- create a separate sample application;
- redesign diagnostic codes/messages.

## Resolved decisions and constraints

### One distributable

V1 supports exactly one distributable package:

```text
Package ID: Yadg
Package type: .NET tool
Command: yadg
Version: 1.0.0
```

The source projects are not supported public libraries.

The implementation must make non-CLI projects non-packable.

### Platform/runtime

V1 package target:

```text
net10.0-windows
x64
framework-dependent
```

The release/consumer locus is Windows x64 with the .NET 10 environment required to install/run the tool.

Word and LibreOffice are runtime capability dependencies only for their corresponding renderer operations.

External Mermaid tooling remains workspace-configured and is not bundled.

### Version authority

Use one central MSBuild release version:

```text
VersionPrefix = 1.0.0
VersionSuffix = empty
```

Package version and `yadg --version` must derive from that version rather than separate hard-coded public values.

### Pack path

Packing must use the existing Windows/full-MSBuild generated-interoperability build path.

`dotnet pack` is not a permitted fallback when it cannot execute the required `ResolveComReference` tooling.

The package must carry every generated interop wrapper required by the installed tool.

Failure to package/install those generated wrappers is an escalation boundary, not permission to switch interop strategies.

### Package metadata

Required package metadata is governed by:

```text
docs/engineering/RELEASE.md
```

The repository entering M0011 has no license authority. Implementation must not invent one.

### Publish script

The NuGet push script always requires:

```text
-Source <source>
```

There is no implicit NuGet.org endpoint.

No API key or credential is committed.

External publication is prohibited during M0011.

### Release validation topology

M0011 deliberately separates:

```text
Tier 2 repository validation
Tier 3 real document/runtime compatibility
Tier 4 actual packaged-tool consumer validation
human release review
external publication (later, excluded)
```

Passing repository tests or `dotnet/msbuild pack` alone is not release readiness.

### M0010 evidence on release revision

The complete M0010 real Word/LibreOffice compatibility suite must pass again on the M0011 release-candidate revision.

Historical M0010 success on an earlier commit is insufficient release evidence.

### Public documentation

README is a user surface, not a milestone diary.

Current behavior must be documented in product terms.

Milestone identifiers may remain only where history itself is being discussed.

`CHANGELOG.md` provides release history.

### Licensing

M0011 does not select a project license.

Absence of package license metadata is allowed for the release candidate because no project licensing authority exists; that state must be explicit in evidence and human review.

This does not constitute a recommendation to publish publicly without a license.

### Publication boundary

M0011 stops after approved release-candidate readiness.

The actual decision to push the package, choose the target feed, create a tag/release, or establish CI/CD is separate.

## Required authority

Implementation starts with this milestone and reads:

- `docs/engineering/RELEASE.md`
- `docs/ENGINEERING.md`
- `docs/SPECS.md`
- `docs/ARCHITECTURE.md`
- `docs/TERMINOLOGY.md`
- `docs/specs/DOCUMENT-COMPOSITION.md`
- `docs/specs/RENDERING.md`
- `docs/specs/WORD-RENDERER.md`
- `docs/specs/PUBLISHING.md`
- `docs/specs/CONTENT-PRODUCERS.md`
- `docs/specs/WORKSPACE-VALUES.md`
- `README.md`
- `AGENTS.md`
- `.review/pending/HR-M0011-01.md`

Implementation may inspect live project/source/test/engineering files needed to execute the contract.

Ordinary implementation must not read the external guide repository or planning conversation to reconstruct release policy.

## Focus areas

### 1. Package/version surface

Establish central V1 versioning, package metadata, one packable CLI project, tool command/version output, README inclusion, and interop-wrapper package completeness.

### 2. Release engineering scripts

Add pack, local package inspection, local consumer validation, and explicit-source NuGet push tooling using the authoritative Windows/full-MSBuild build path.

### 3. V1 public documentation audit

Rewrite public documentation against the actual M0010/M0009 product rather than accumulated milestone-era prose, and add the 1.0.0 changelog.

### 4. Consumer/release evidence

Install the generated package in isolation, exercise representative real functionality through the installed command, preserve current Tier-3 runtime evidence, and create auditable release evidence.

### 5. Release review and closure

Prepare exact-hash review evidence, obtain human approval, rerun the milestone review check, reread authority, and perform the release-readiness completion audit.

## Acceptance criteria

### Versioning

- central repository version authority resolves to `1.0.0`;
- no competing/stale package version authority exists;
- built package version is `1.0.0`;
- installed `yadg --version` prints `1.0.0`;
- version output is derived from build/package version metadata, not a separate stale constant.

### Package identity/surface

- only the CLI project produces a supported NuGet package;
- non-CLI source projects are explicitly non-packable;
- package ID is `Yadg`;
- package type is .NET tool;
- tool command is `yadg`;
- package target/runtime assumptions match Windows x64/.NET 10;
- package README is the audited repository README;
- required package metadata from `docs/engineering/RELEASE.md` is present;
- package contains all YADG runtime assemblies required by the command;
- package contains the generated Microsoft Word and Office Core interop wrapper assemblies required by M0009;
- installed tool does not load required assemblies from repository `bin`/`obj`.

### Package hygiene

- exactly one expected `Yadg.1.0.0.nupkg` is produced by the default pack command;
- package has no test binaries/test fixtures/review artifacts;
- package has no credentials/secrets;
- package has no unrelated repository files;
- package has no accidental separate YADG library packages;
- package has no unsupported license metadata claim;
- package hash is recorded in release evidence.

### Pack script

- `./eng/pack.ps1` succeeds on the authoritative Windows build locus;
- `-OutputPath` overrides the default package output;
- full Visual Studio MSBuild is used for the COMReference-capable pack path;
- absence of required full MSBuild/interop tooling fails clearly;
- no `dotnet pack` fallback silently weakens the M0009 build contract;
- script is non-interactive and suitable for future CI execution;
- failed pack does not present a stale package as current success.

### NuGet publish script

- `./eng/publish-nuget.ps1` requires explicit `-Source`;
- no production/default source is embedded;
- exact package ID/version is checked before push;
- default package-path resolution fails on missing/ambiguous package;
- repository contains no API key;
- `NUGET_API_KEY` may be consumed from environment when needed;
- script can operate with normal NuGet-configured authentication when an API key is not supplied;
- stable release push does not use `--skip-duplicate`;
- local/non-external validation proves the script pushes the intended package;
- M0011 never invokes it against an external production feed.

### Installed-tool Tier 4

- package installs from local package output with `dotnet tool install --tool-path`;
- installation uses isolated CLI/cache/tool directories;
- installed tool is invoked directly;
- installed `--version` and `--help` succeed;
- installed `check` succeeds on a representative realistic workspace;
- installed `build` produces authored DOCX;
- installed Word renderer finalizes representative authored DOCX through real Microsoft Word;
- installed LibreOffice renderer finalizes representative authored DOCX through real LibreOffice;
- installed `publish` delivers the expected finalized DOCX;
- installed tool resolves packaged generated interop wrappers;
- temporary tool/workspace state is cleaned best-effort.

### Existing validation on RC revision

- `./eng/validate.ps1` passes;
- `./eng/test-m0010-tier3.ps1` passes the complete four-path matrix on the M0011 revision;
- M0009 Word build/interop contract is preserved;
- M0010 realistic template/heading/value semantics are preserved;
- publish behavior remains unchanged.

### README V1 audit

README accurately and coherently covers the V1 contract required by `docs/engineering/RELEASE.md`, including:

- purpose/ownership model;
- platform/install prerequisites;
- NuGet tool installation shape;
- workspace layout and first workflow;
- every public top-level command and relevant options;
- template/Markdown/workspace concepts;
- relative headings;
- visible value stories;
- structured content/references;
- external Mermaid producer;
- Word/LibreOffice behavior/prerequisites;
- PDF limitations;
- publish behavior;
- security/trust;
- V1 limitations and troubleshooting.

The main user path contains no stale statements contradicted by M0009/M0010.

Current functionality is not described primarily in terms of milestone IDs.

### CHANGELOG

- root `CHANGELOG.md` exists;
- has a `1.0.0` entry;
- summarizes user-visible V1 capabilities/limitations rather than copying milestone task lists;
- version/date information agrees with the release candidate.

### CLI/docs consistency

- root `--help` lists `check`, `build`, `render`, `publish`;
- command help and README use the same option names/default behavior;
- `--version` agrees with package version;
- renderer prerequisite wording clearly separates runtime use from build/package-generation requirements;
- README does not claim PDF publication;
- README does not claim unsupported unattended/server Word automation.

### Release evidence

Release evidence contains every item required by `docs/engineering/RELEASE.md`.

Evidence identifies the exact repository revision and package SHA reviewed.

Evidence states that no external package publication, GitHub Release, or tag occurred.

### Human review

- implementation prepares `HR-M0011-01` evidence;
- review tooling recognizes M0011 without weakening previous review checks;
- review check fails while pending/negative;
- implementation does not fabricate approval;
- waiver is impossible;
- after actual approval, `./eng/review-check.ps1 --milestone M0011` passes for the exact release evidence/package hash.

## Validation

### Tier 0 — edit/package sanity

On the authoritative Windows build locus:

- restore/build changed release/package surfaces;
- inspect version/package properties;
- validate scripts parse/argument handling;
- verify package project selection/non-packable library projects.

Use the repository's existing full-MSBuild requirements rather than assuming plain `dotnet build` can satisfy Word COM reference generation.

### Tier 1 — focused

Cover at least:

- `--version`;
- package metadata/version mapping;
- pack output-path behavior;
- publish-script mandatory-source and package-selection behavior;
- local publish-script destination;
- package content/hygiene assertions;
- README/changelog consistency checks that can be automated reasonably.

### Tier 2 — repository

```powershell
./eng/validate.ps1
```

Run on Windows with the M0009 full-MSBuild/Office development prerequisites.

### Tier 3 — current real-runtime compatibility

```powershell
./eng/test-m0010-tier3.ps1
```

Run on an interactive Windows session with real Microsoft Word and LibreOffice.

All four M0010 origin/renderer paths remain mandatory.

### Pack

```powershell
./eng/pack.ps1
```

Expected default artifact:

```text
artifacts/package/Yadg.1.0.0.nupkg
```

### Tier 4 — consumer/release

```powershell
./eng/test-m0011-tier4.ps1
```

Required locus:

```text
interactive Windows x64
.NET 10 SDK
full Visual Studio MSBuild/ResolveComReference tooling
desktop Microsoft Word installed/activated
LibreOffice installed
```

Expected evidence:

- package inspection;
- isolated tool installation;
- installed command/version/help;
- installed check/build;
- installed Word render;
- installed LibreOffice render;
- installed publish;
- package/interops resolution;
- release manifest/hashes.

### Publish-script test

The implementation tests:

```powershell
./eng/publish-nuget.ps1 -Source <temporary-local-source>
```

against a non-external local destination.

Do not point it at NuGet.org or a production/private feed during M0011.

### Tier 5 — release review

Review ID:

```text
HR-M0011-01
```

Completion command:

```powershell
./eng/review-check.ps1 --milestone M0011
```

Approval must correspond to the exact M0011 release evidence and package hash.

## Constrained execution

The implementation environment may lack one or more release capabilities.

If full Visual Studio MSBuild/ResolveComReference tooling is unavailable:

- portable/documentation work may proceed;
- do not claim pack success;
- do not substitute `dotnet pack`;
- record the release milestone blocked on the authoritative pack locus.

If Microsoft Word or LibreOffice is unavailable:

- lower-tier/package work may proceed where possible;
- do not weaken Tier 3/Tier 4 renderer evidence;
- leave the release milestone blocked.

If network access to ordinary NuGet dependency sources is unavailable:

- do not publish externally;
- use already restored/cached packages where valid;
- otherwise record the affected pack/install validation blocked.

Human approval can only be supplied by the reviewer.

## Release evidence and review flow

Implementation maintains its normal execution ledger:

```text
.execution/M0011-v1-release-readiness.md
```

Release closure is:

```text
implement package/version/scripts/docs
-> Tier 0/1
-> ./eng/validate.ps1
-> ./eng/test-m0010-tier3.ps1
-> ./eng/pack.ps1
-> package inspection
-> ./eng/test-m0011-tier4.ps1
-> local publish-script validation
-> assemble exact-hash release evidence
-> human HR-M0011-01
-> review-check
-> fresh authority reread
-> milestone/evidence/ledger reconciliation
-> release-readiness completion audit
```

Do not infer release readiness from partial child output.

## Direct documentation impact

M0011 directly changes public/release documentation, including:

- `README.md`;
- new `CHANGELOG.md`;
- package metadata/readme surface;
- direct release/packaging engineering documentation where necessary;
- CLI help/version surface.

This documentation work is part of M0011, not a deferred documentation-sync pass.

## Deferred documentation synchronization

None is required for the V1 public surfaces in scope.

Guide/profile metadata remains planning metadata and is not changed merely to record release success unless a separate guide-synchronization task requires it.

## Publication operations explicitly excluded

M0011 must not:

```text
dotnet nuget push <package> --source https://api.nuget.org/...
create GitHub Release
create/push release tag
modify GitHub Actions
publish to a production/private NuGet feed
```

The publish script itself is in scope; external publication is not.

## Baseline-executability audit

Planning confirms:

- release version and package identity are fixed;
- supported package/platform/API surface is fixed;
- only CLI is packable;
- M0009 interop build/pack strategy is preserved;
- package-content obligations are fixed;
- version authority is fixed;
- pack and NuGet push script behavior is fixed;
- external publication boundary is fixed;
- Tier 2/3/4 targets and loci are fixed;
- public documentation obligations are fixed;
- license behavior is fixed without inventing a license;
- release evidence is fixed;
- blocking human review is fixed;
- GitHub workflows are explicitly excluded.

Implementation choices remaining are local mechanics, not release policy.

M0011 is ready for the configured baseline implementation model.

## Escalation boundary

Return to planning rather than inventing policy if implementation would require:

- changing package ID `Yadg`;
- publishing more than one supported package;
- exposing C# libraries as supported public APIs;
- changing V1 platform/architecture/framework;
- changing version `1.0.0`;
- changing the M0009 generated Office interop strategy;
- omitting required interop wrappers from the installed package;
- changing the pack build locus away from full Visual Studio MSBuild;
- selecting a project software license;
- choosing/defaulting an external NuGet feed;
- adding GitHub workflow/release/tag publication;
- weakening the M0010 Tier-3 matrix;
- weakening the installed-package Tier-4 target;
- removing/block-bypassing the human release review.

If the .NET tool package cannot contain and resolve the generated Office interop assemblies using the mandated M0009 build approach, stop and report M0011 blocked.
