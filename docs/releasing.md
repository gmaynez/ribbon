# Releasing Ribbon

Ribbon releases are built by `.github/workflows/release.yml`. The workflow runs when a tag matching `v1.*` is pushed, then applies stricter validation before publishing anything.

## Release contract

- Tags use semantic versioning in the form `v1.MINOR.PATCH`, such as `v1.2.0`. A suffix such as `v1.2.0-alpha.1`, `v1.2.0-beta.1`, or `v1.2.0-rc.1` creates a GitHub prerelease. When production signing secrets are absent, prereleases use a temporary self-signed manifest certificate and carry an untrusted-test-build warning.
- The tagged commit must be contained in `origin/main`. A matching tag on another branch fails before the build.
- The workflow uses the Windows Server 2025 Visual Studio 2026 image so one MSBuild 18 invocation can build both the .NET 10 broker and the .NET Framework 4.8 VSTO projects. It installs Microsoft's signed VSTO runtime redistributable and verifies its ClickOnce hosting assembly before the build.
- The workflow performs a complete Release build, runs the broker tests, verifies that each VSTO application manifest includes the broker runtime files, and refuses to publish unsigned manifests. Stable releases additionally require the configured release-signing certificate.
- It then publishes a self-contained `win-x64` broker and compiles the per-user Inno Setup installer.
- The GitHub Release contains `Ribbon-Setup-vVERSION.exe`, separate Grid, Quill, Deck, and Post ZIP files, and `SHA256SUMS.txt`.

## One-time GitHub configuration

Create a GitHub Actions environment named `release`. Add protection rules or required reviewers if releases should require approval. Stable releases require these environment secrets; prereleases use them when present and otherwise fall back to an ephemeral self-signed certificate:

- `RIBBON_SIGNING_CERTIFICATE_BASE64` — a Base64-encoded, publicly trusted code-signing PFX containing its private key.
- `RIBBON_SIGNING_CERTIFICATE_PASSWORD` — the PFX password.

To encode the certificate without writing its contents to the console:

```powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes('C:\secure\ribbon-release.pfx')) |
    Set-Clipboard
```

The certificate is imported into the ephemeral runner's current-user certificate store and removed in an `always()` cleanup step. The PFX and password are never passed on the MSBuild command line.

Repository settings must allow GitHub Actions to create releases with the repository `GITHUB_TOKEN`. The workflow declares only `contents: write` permission and pins its checkout action to an immutable commit SHA.

## Creating a release

Update `main`, create an annotated tag, and push only after the tagged commit is present on `origin/main`:

```powershell
git switch main
git pull --ff-only origin main
git tag -a v1.2.0 -m 'Ribbon v1.2.0'
git push origin v1.2.0
```

Do not move a published release tag. If a release build fails, fix the cause on `main` and create a new version tag.

For a local packaging check, use the explicit development-only switch. This relies on each VSTO project's existing local signing settings and development certificate:

```powershell
.\eng\Build-Release.ps1 -Version 1.2.0 -UseProjectSigningSettings
```

Inno Setup 7 must be installed to compile `eng\Ribbon.iss`. The 64-bit compiler is preferred (`%ProgramFiles%\Inno Setup 7\ISCC.exe`). Pass `-SkipInstaller` to produce only the host ZIP files.

Output signed with a local or workflow-generated development certificate is for testing only. Windows and Office will identify its publisher as untrusted, and it must not be represented as a stable Ribbon release.

## End-user installer

`Ribbon-Setup-vVERSION.exe` is a 64-bit per-user Inno Setup 7 package (`SetupArchitecture=x64`, `PrivilegesRequired=lowest`). The wizard presents the Apache 2.0 `LICENSE` for acceptance before component selection. It installs to `%LOCALAPPDATA%\Ribbon` without administrator rights:

| Path | Contents |
| --- | --- |
| `Grid\`, `Quill\`, `Deck\`, `Post\` | Signed VSTO payloads, including the four broker files listed in each application manifest |
| `Broker\` | Self-contained `Ribbon.Broker.exe` used by every host |

Prerequisites: the VSTO add-ins need Microsoft .NET Framework 4.8 (included with Windows 10 since the May 2019 Update) and the Microsoft Visual Studio 2010 Tools for Office runtime. Setup checks both and warns with a continue-or-cancel prompt when either is missing. A separate .NET 10 installation is not required because the installed broker is self-contained.

The installer offers component selection. The broker is a fixed, required component; Grid, Quill, Deck, and Post are individually selectable and checked by default. Post is only useful with classic desktop Outlook: the new Outlook is a WebView-based client that cannot load VSTO add-ins. During setup, Ribbon resolves `outlook.exe` through App Paths in both registry hives and views and accepts it only when the resolved executable really is `OUTLOOK.EXE`. If the user has switched to the new Outlook (`HKCU\Software\Microsoft\Office\16.0\Outlook\Preferences\NewOutlook = 1`) or an administrator forced that migration through policy, Post is left unselected with an explanatory message, and manually re-selecting it asks for confirmation.

The installer writes HKCU VSTO keys (`LoadBehavior=3`) whose `Manifest` values use `file:///{app}/<host>/<host>.vsto|vstolocal`. Keys and folders are written and removed only for the components selected during install. Uninstall does not delete conversations, downloaded agents, sessions, cache, logs, or checkpoints.

Close Excel, Word, PowerPoint, classic Outlook, and `Ribbon.Broker.exe` before an upgrade. The setup stops the broker process itself when it can.

A later MSI can wrap the same payload for Intune or GPO. VSTO add-ins still cannot be submitted to Microsoft Marketplace.

## Microsoft Marketplace

Ribbon's current Excel, Word, PowerPoint, and Outlook add-ins are VSTO add-ins. Microsoft states that Office VSTO and COM add-ins cannot be submitted to Microsoft Marketplace, so this repository cannot add a Store-publication stage for the current products. Distribute the signed VSTO packages directly or through an organization-managed Windows deployment channel.

Publishing inside Office would require a separate Office Web Add-in with a supported manifest and cross-platform implementation. That product could be submitted through Partner Center, but Microsoft currently states that Microsoft 365 Office add-ins do not have submission API support, so certification submission cannot be automated by this GitHub workflow.

References:

- [Make solutions available in Microsoft Marketplace and within Office](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/submit-to-appsource-via-partner-center)
- [Marketplace submission API support](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/submission-api-overview)
- [GitHub Actions workflow tag filters](https://docs.github.com/en/actions/reference/workflows-and-actions/workflow-syntax#onpushbranchestagsbranches-ignoretags-ignore)
- [Secure use of GitHub Actions](https://docs.github.com/en/actions/reference/security/secure-use)
