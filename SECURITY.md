# Security policy

## Supported versions

Only the latest release receives security fixes. Check your version under
**Help → About XactCopy** and update from
[Releases](https://github.com/Wimukthi/XactCopy/releases/latest) before
reporting anything.

| Version | Supported |
| --- | --- |
| 2.0.x (latest release) | Yes |
| Older 2.0.x | No — upgrade first |
| 1.x (the retired VB.NET build) | No |

## Reporting a vulnerability

Report privately through GitHub: open the repository's **Security** tab and
choose **Report a vulnerability**. Please do not open a public issue for
something exploitable.

A useful report includes:

- The XactCopy version and your Windows version.
- What an attacker would gain, and what access they need to start with.
- Concrete steps to reproduce, ideally with a sample source tree or media.
- Any relevant part of the operations log or journal.

You can expect an acknowledgement within a few days. Fixes ship in the next
release, and the advisory credits the reporter unless you ask otherwise.

## Scope

XactCopy is a local desktop application. The parts worth scrutinising:

- **Worker IPC.** `XactCopy.exe` and `XactCopyExecutive.exe` talk over a named
  pipe created as a single instance, in byte mode, with a security descriptor
  limiting it to the current user. Frames are length-prefixed and capped, and the
  protocol version is checked rather than inferred.
- **Journal and bad-range-map integrity.** Journals are hash-chained and
  HMAC-signed; bad-range maps are signed and bound to a fingerprint of their
  source path, so a map can't be applied to media it wasn't made for. The journal
  HMAC key is stored as raw bytes and the map key is DPAPI-protected for the
  current user, both under `%LOCALAPPDATA%\XactCopy\security\`.
- **The updater.** It reads GitHub's `releases/latest` over HTTPS, downloads the
  package, and verifies it against the SHA-256 published with the release —
  either the asset digest reported by the API or a sibling `.sha256` asset — before
  installing. A release whose package cannot be checksummed is refused rather
  than installed unverified.
- **Explorer integration.** All shell entries are written per-user under
  `HKCU\Software\Classes` and removed when the option is turned off.

Out of scope: anything that requires the attacker to already have the user's
Windows account, physically damaged media producing wrong data (that is the
problem XactCopy exists to bound, not a vulnerability), and the retired 1.x
VB.NET implementation on the `vbnet-legacy` branch.

## Verifying a download

Every release asset ships with a `.sha256` file. Check it before running:

```powershell
Get-FileHash .\XactCopySetup-v2.0.0.6-win-x64.exe -Algorithm SHA256
```

The result must match the contents of
`XactCopySetup-v2.0.0.6-win-x64.exe.sha256`. Releases are not code-signed;
downloading only from the GitHub Releases page and checking the hash is the
intended verification path.
