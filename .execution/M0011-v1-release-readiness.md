# M0011 V1 Release Readiness — Execution Ledger

## Status

Implementation and automated release validation complete; human HR-M0011-01 approval is still required. The milestone is not declared complete while the review gate is pending.

## Work packages

1. **WP-01 — Authority/baseline**: apply the supplied overlay, read the milestone and Required Authority, inspect the existing M0009/M0010 build and review boundaries, and maintain this ledger.
2. **WP-02 — Package/version surface**: establish central `1.0.0` MSBuild versioning, one packable CLI .NET tool, package metadata/README, public version output, and explicit non-packability for other projects while preserving generated COM interop.
3. **WP-03 — Release scripts**: add full Visual Studio MSBuild pack tooling, package inspection, explicit-source local-safe NuGet push tooling, and installed consumer Tier-4 validation.
4. **WP-04 — Public V1 documentation/evidence**: audit README, add CHANGELOG, create machine-readable release evidence, and map acceptance criteria to implementation/validation evidence.
5. **WP-05 — Required validation**: run Tier 2, M0010 Tier 3, pack, Tier 4, and local publish-script validation; stop if the authoritative runtime/packaging boundary is unavailable.
6. **WP-06 — Human review/closure**: prepare exact-hash HR-M0011-01 evidence; do not fabricate approval; after actual approval, run review-check, fresh reread, reconcile, and complete the audit.

## Acceptance/evidence map

| Area | Implementation | Validation/evidence |
|---|---|---|
| Version/package | `Directory.Build.props`; `src/Yadg.Cli/Yadg.Cli.csproj`; `Program.ProductVersion`; `Directory.Build.targets` | `validate.ps1`; pack metadata inspection; installed `--version` = `1.0.0`; non-CLI projects inherit `IsPackable=false` |
| Full-MSBuild COMReference pack | `eng/pack.ps1`; `Yadg.WordRenderer.csproj` retains both `COMReference` items; CLI duplicate wrapper generation removed | `pack.ps1` passed via Visual Studio 18.10.1 MSBuild; package contains `Interop.Microsoft.Office.Core.dll` and `Interop.Microsoft.Office.Interop.Word.dll` |
| Publish tooling | `eng/publish-nuget.ps1` requires `-Source`, validates `Yadg 1.0.0`, no embedded credential/default feed/skip-duplicate | local filesystem source push passed; no external push performed |
| Installed consumer Tier 4 | `eng/test-m0011-tier4.ps1` isolated tool/cache/workspace and package inspection | installed command passed version/help/check/build/Word render/LibreOffice render/publish; release evidence JSON |
| Public README/CHANGELOG | audited root `README.md`; new root `CHANGELOG.md` | documentation contract reviewed; package README entry and metadata inspected |
| Release evidence/exact hash | `artifacts/release/evidence/M0011/release-evidence.json` | package SHA256 `9610F6025806A9CA69E4837878961743CF249CAAA6BD70B3CD6015CDA9A8452A`; exact package path `artifacts/package/Yadg.1.0.0.nupkg` |
| Human HR-M0011-01 gate | pending human action | review-check after approval |

## Validation log

| Command/evidence | Result |
|---|---|
| Overlay + authority read | complete; external guide/planning conversation not used |
| `./eng/validate.ps1` | passed; 46/46 tests |
| `./eng/test-m0010-tier3.ps1` | passed; all four Word/LibreOffice origin/renderer paths |
| `./eng/pack.ps1` | passed; exactly `artifacts/package/Yadg.1.0.0.nupkg`; `-OutputPath` override also passed |
| `./eng/test-m0011-tier4.ps1` | passed; package inspection and installed-tool lifecycle including both renderers/publish |
| local `publish-nuget.ps1` validation | passed against temporary repository-local filesystem destination; destination was removed; no external feed used |
| `./eng/review-check.ps1 --milestone M0011` | correctly fails while `.review/records/HR-M0011-01.md` is absent; must pass only after actual human approval |

## Criterion-level reconciliation

- **Versioning:** central `VersionPrefix=1.0.0`, empty suffix, explicit informational version, package `Yadg.1.0.0.nupkg`, and installed `yadg --version` all agree; no competing release version was added.
- **Package identity/surface:** CLI is the sole packable source project; package ID/tool command/author/description/repository/readme metadata are present; target implementation remains `net10.0-windows`/x64 and the normalized tool layout installs on .NET 10; generated Word/Core wrappers are present and installed execution uses isolated state.
- **Package hygiene:** default pack output is exactly one expected package; package inspection found no tests, fixtures, review evidence, credentials, or source `bin`/`obj`; the authorized MIT expression and packaged `LICENSE` are present; exact SHA is in release evidence.
- **Pack script:** uses full Visual Studio MSBuild, restores first, fails clearly without supported MSBuild, removes stale expected output before packing, supports `-OutputPath`, never falls back to `dotnet pack`, and normalizes only the tool directory layout required by the installer after the COMReference pack succeeds.
- **NuGet publish script:** explicit `-Source` is mandatory; package path defaults only to the exact expected artifact; ID/version are checked; `NUGET_API_KEY` is environment-only; normal NuGet authentication remains available; no `--skip-duplicate`; local filesystem push passed and no external publication occurred.
- **Tier 4:** isolated `dotnet tool install --tool-path` and direct installed executable passed `--version`, `--help`, realistic check/build, installed Word and LibreOffice rendering, and publish. Package wrappers resolved without repository `bin`/`obj`.
- **Existing validation:** final release sequence passed Tier 2, all four M0010 Tier-3 paths, full-MSBuild pack, and M0011 Tier 4; M0009 Word COMReference/publish behavior and M0010 realistic composition/value behavior remain covered.
- **README audit:** current README covers purpose/ownership, Windows/.NET 10 installation, workspace workflow, public commands/options, template vocabulary, relative headings, visible value stories, structured content/references, Mermaid trust, renderer prerequisites/differences, PDF status, publishing, security, limitations, and troubleshooting without relying on milestone history.
- **CHANGELOG/CLI consistency:** root `CHANGELOG.md` has a dated `1.0.0` entry; README command syntax matches CLI help; version output is `1.0.0`; README makes no PDF-publication or unattended Word claim.
- **Release evidence:** machine-readable evidence records revision `c95008f834d30abaa08cf6132ddb85f7cd6d03ea`, package hash, Windows/.NET/MSBuild/Word/LibreOffice provenance, M0010 evidence reference, consumer results, local publish result, documentation audit, no external publication, and authorized MIT license metadata.
- **Human review:** pending request is `.review/pending/HR-M0011-01.md`; `eng/review-check.ps1` recognizes M0011 and will verify the approved record’s `SHA256:` value against the exact package. No approval or waiver was fabricated.

## Release evidence

- Package: `artifacts/package/Yadg.1.0.0.nupkg`
- SHA-256: `9610F6025806A9CA69E4837878961743CF249CAAA6BD70B3CD6015CDA9A8452A`
- Evidence: `artifacts/release/evidence/M0011/release-evidence.json`
- M0010 matrix evidence: `artifacts/review/evidence/M0010/tier3-evidence.json`
- No external NuGet push, GitHub Release, tag, workflow, or software-license decision was made.

## Escalation boundary

If full MSBuild/ResolveComReference cannot carry generated Office interop assemblies through pack/install, or required interactive Word/LibreOffice capability is unavailable, record the block and return to planning. Do not substitute another interop strategy or claim release readiness.
