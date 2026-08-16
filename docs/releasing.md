# Releasing Ribbon

Ribbon releases are built by `.github/workflows/release.yml`. The workflow runs when a tag matching `v1.*` is pushed, then applies stricter validation before publishing anything.

## Release contract

- Tags use semantic versioning in the form `v1.MINOR.PATCH`, such as `v1.2.0`. A prerelease suffix such as `v1.2.0-rc.1` is supported and creates a GitHub prerelease.
- The tagged commit must be contained in `origin/main`. A matching tag on another branch fails before the build.
- The workflow performs a complete Release build, runs the broker tests, verifies that each VSTO application manifest includes the broker runtime files, and refuses to publish unsigned manifests.
- It then publishes a self-contained `win-x64` broker and compiles the per-user Inno Setup installer.
- The GitHub Release contains `Ribbon-Setup-vVERSION.exe`, separate Grid, Quill, and Deck ZIP files, and `SHA256SUMS.txt`.

## One-time GitHub configuration

Create a GitHub Actions environment named `release`. Add protection rules or required reviewers if releases should require approval, then add these environment secrets:

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

Inno Setup 6 must be installed to compile `eng\Ribbon.iss`. Pass `-SkipInstaller` to produce only the host ZIP files.

Output signed with a local development certificate is for pipeline development only and must not be distributed as a Ribbon release.

## End-user installer

`Ribbon-Setup-vVERSION.exe` is a per-user Inno Setup package (`PrivilegesRequired=lowest`). It installs to `%LOCALAPPDATA%\Ribbon` without administrator rights:

| Path | Contents |
| --- | --- |
| `Grid\`, `Quill\`, `Deck\` | Signed VSTO payloads, including the four broker files listed in each application manifest |
| `Broker\` | Self-contained `Ribbon.Broker.exe` used by every host |

The installer writes HKCU VSTO keys (`LoadBehavior=3`) whose `Manifest` values use `file:///{app}/<host>/<host>.vsto|vstolocal`. Uninstall removes those keys and the payload folders. It does not delete conversations, downloaded agents, sessions, cache, logs, or checkpoints.

Close Excel, Word, PowerPoint, and `Ribbon.Broker.exe` before an upgrade. The setup stops the broker process itself when it can.

A later MSI can wrap the same payload for Intune or GPO. VSTO add-ins still cannot be submitted to Microsoft Marketplace.

## Microsoft Marketplace

Ribbon's current Excel, Word, and PowerPoint add-ins are VSTO add-ins. Microsoft states that Office VSTO and COM add-ins cannot be submitted to Microsoft Marketplace, so this repository cannot add a Store-publication stage for the current products. Distribute the signed VSTO packages directly or through an organization-managed Windows deployment channel.

Publishing inside Office would require a separate Office Web Add-in with a supported manifest and cross-platform implementation. That product could be submitted through Partner Center, but Microsoft currently states that Microsoft 365 Office add-ins do not have submission API support, so certification submission cannot be automated by this GitHub workflow.

References:

- [Make solutions available in Microsoft Marketplace and within Office](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/submit-to-appsource-via-partner-center)
- [Marketplace submission API support](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/submission-api-overview)
- [GitHub Actions workflow tag filters](https://docs.github.com/en/actions/reference/workflows-and-actions/workflow-syntax#onpushbranchestagsbranches-ignoretags-ignore)
- [Secure use of GitHub Actions](https://docs.github.com/en/actions/reference/security/secure-use)
