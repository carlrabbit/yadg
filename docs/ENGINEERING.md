# Engineering

## Scope

Repository-wide engineering and validation policy for YADG.

## Stack/boundaries

YADG is C#/.NET with System.CommandLine and Open XML SDK.

Office-independent authoring remains free of Microsoft Office/LibreOffice dependencies.

M0008 adds external producer processes but no in-process JS/runtime dependency.

## External producer validation

Portable Tier 0-2 tests may use deterministic local fake producer executables/process fixtures where useful for process/error-path coverage.

They must not claim real Mermaid compatibility.

Real M0008 compatibility is Tier 3 and requires network access for test-tool acquisition unless already satisfied by a verified local test cache.

## Test-tool version manifest

Authoritative M0008 external-test versions are stored in:

```text
eng/test-tools.json
```

The manifest uses exact versions.

At M0008 planning time it pins:

- Bun `1.4.2`;
- `@mermaid-js/mermaid-cli` `11.17.0`.

The Bun pin represents the current latest stable release at planning time.

Tests must not dynamically resolve a moving `latest` selector at execution time.

Version upgrades are explicit edits to the JSON manifest.

## Tier-3 Bun acquisition

M0008 Tier-3 test infrastructure must:

1. read `eng/test-tools.json`;
2. download the official Bun release matching the configured exact version for the current supported test platform into a test-owned temporary/cache location;
3. avoid requiring or modifying a machine-global Bun installation;
4. verify the downloaded executable reports the configured Bun version;
5. use that executable for the real Mermaid integration test.

The exact official release URL construction/platform archive extraction mechanics are implementation-owned, but provenance must identify the resulting Bun version and platform.

Tests should reuse an already verified local downloaded archive/executable within one test run/cache rather than redownloading unnecessarily.

## Mermaid one-shot execution

The real Tier-3 test must invoke the configured Mermaid package/version through Bun one-time package execution rather than `npm install`, a repo `node_modules`, or a global Mermaid install.

The required logical invocation is equivalent to:

```text
<downloaded-bun> x --bun --package @mermaid-js/mermaid-cli@<version> mmdc -i <input> -o <output>
```

The package name/version come from `eng/test-tools.json`.

The test must not require system `node`, `npm`, `npx`, or a preinstalled `mmdc`.

A failure of the pinned Bun/Mermaid combination to render the required scenario blocks M0008 completion; do not silently substitute npm/Node in the authoritative test.

## Network distinction

Canonical repository validation:

```powershell
./eng/validate.ps1
```

must remain suitable for ordinary development without mandatory external downloads.

The real Bun/Mermaid Tier-3 target is invoked explicitly and is a separate network-capable validation target.

If the locus cannot access the required official Bun release/npm registry/Chromium dependencies, Tier 0-2 may pass but M0008 Tier-3 success must not be claimed.

## DOCX integration

Real M0008 Tier-3 validation continues through a real filesystem workspace and real DOCX package after Mermaid output is produced.

Validate semantic placement/embedded PNG/caption/reference structures structurally; byte equality is not required.

## Process diagnostics

Capture bounded external producer exit status and stderr/stdout where practical.

Do not expose credentials in diagnostics.

## Fixtures/hygiene

Mermaid test source and DOCX templates must be synthetic and redistribution-safe.

No confidential bank diagrams/content.

## Existing renderer

M0005 real LibreOffice validation policy remains unchanged. M0008 does not require LibreOffice to validate Mermaid authoring itself.
