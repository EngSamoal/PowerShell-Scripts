# Instructions for Claude Code sessions working in this repository

## Always commit and push automatically

Whenever you modify a script in this repo at the user's request, **commit and push the change
to GitHub yourself as part of finishing the task** — don't wait to be asked, and don't just
describe the change and leave it sitting locally. This applies by default to every script
request in this repo, not just ones where the user explicitly says "push it."

- Commit with a clear message describing what changed and why.
- Push to whichever branch your session is instructed to develop on. If a session has no
  specific branch instruction, ask the user which branch to use before pushing — do not push
  directly to `main` without the user's explicit go-ahead first.
- After pushing, tell the user where to find it (repo + branch + file), and consider sending
  the updated file directly as an attachment too — this user has repeatedly had trouble
  locating files on GitHub, so a direct file makes the handoff much more reliable.
- If a later session finds unrelated changes already on the target branch (e.g. edits made
  directly through the GitHub UI, not through a Claude session), stop and confirm with the user
  before overwriting anything — don't assume your version should win.

## Testing PowerShell scripts without Windows/PowerCLI

This environment is Linux and has no Windows, PowerCLI, or Word installed, but a portable
PowerShell 7 binary can be downloaded for real syntax/logic verification instead of relying on
visual inspection alone:

```bash
curl -sSL -o /tmp/pwsh.tar.gz "https://github.com/PowerShell/PowerShell/releases/download/v7.4.6/powershell-7.4.6-linux-x64.tar.gz"
mkdir -p /tmp/pwsh && tar -xzf /tmp/pwsh.tar.gz -C /tmp/pwsh && chmod +x /tmp/pwsh/pwsh
```

Before pushing any change to `VMware_Weekly_HealthCheck.ps1` (or similar scripts in this repo):

1. **Parse-check** with `[System.Management.Automation.Language.Parser]::ParseFile()` to catch
   syntax errors.
2. **Smoke-test the report-generation logic** against synthetic data with stubbed Word COM
   objects (Word/PowerCLI cmdlets don't exist here, but the pure PowerShell logic — filtering,
   array handling, string building — can and should be exercised for real). Several real bugs in
   this script (a single-row table silently corrupting into individual characters, a VLAN count
   inflated by counting switch uplink port groups as VLANs, `.Count` returning `$null` instead of
   0/1 for certain pipeline results) were only caught this way, not by reading the code.
3. When the user sends a photo of an error or an actual generated `.docx`/`.log` file, prefer
   inspecting the real file over guessing from the image: `.docx` files are zip archives —
   `unzip` them and read `word/document.xml` directly for exact table contents, using
   `libreoffice --headless --convert-to pdf` to render if a visual check is needed.
4. The user runs scripts in **Windows PowerShell 5.1 / PowerShell ISE**, not PowerShell 7. Code that
   passes on pwsh 7 can still fail on 5.1 (e.g. `Export-WklyReportDashboard.ps1` build r6 crashed on 5.1
   with "Argument types do not match"). Root cause, reproduced on pwsh 7 too: wrapping a
   `System.Collections.Generic.List[object]` in the array operator — `@($list)`, `@($obj.ListProp).Count`,
   `foreach ($x in @($list))` — throws that error. Iterate lists directly (`foreach ($x in $list)`, which is
   also safe for `$null`) and use `$list.Count`; `@(... | Where-Object ...)` pipelines are fine. Test helper
   code with real `List[object]` data, not hand-built arrays, and make optional or diagnostic features
   unable to abort the main output (per-item try/catch).
5. **Save scripts as UTF-8 *with* BOM.** Windows PowerShell 5.1 reads a BOM-less file as ANSI (Windows-1252),
   so any non-ASCII character (`→ ▲ – — ·`, emoji) is misread: string comparisons silently stop matching, and a
   character such as `→` inside quotes breaks parsing (`Export-WklyReportDashboardB.ps1` build B2 failed this way;
   pwsh 7 assumes UTF-8 and hid it). Check with `head -c3 file | od -An -tx1` (must be `ef bb bf`), and
   parse-check the file as 5.1 would by decoding it as code page 1252 when there is no BOM.
6. Don't guess twice at the same class of bug. If a fix doesn't hold up against the next report
   the user sends, get the actual error/log data before proposing another fix, rather than
   patching blind again.

## VMware_Weekly_HealthCheck.ps1 conventions

- Bump `$ScriptBuild` (near the top of the script) on every change and mention the new build
  string when handing the file back — the user runs this script from multiple machines/sessions
  and has repeatedly hit confusion from stale cached copies. The build string prints as the very
  first line of output, so it's always possible to confirm which version actually ran.
- Never fabricate data. Anything the script can't determine from vCenter/PowerCLI (e.g. physical
  storage-array hardware, backup-product status) must be surfaced as `Manual/External Required`
  or `Unable to Check` with an honest explanation in `Notes` — not guessed or left silently blank.
- When a check can fail (VAMI/CIS calls in particular have separate permissions from the regular
  vSphere API and fail often), catch it per-item and still emit a finding explaining *why* it
  failed, not just a blank/generic value — the report is the only place the user reliably looks,
  or a log file next to it.
