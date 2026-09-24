# securecrt-excel-runner

Drive network devices from a spreadsheet — with nothing but **SecureCRT** and **Excel** (or LibreOffice).

Each row of the workbook is one step: the command to send, the prompt that ends its output, the text
that means success. `rdwr.vbs` runs the rows in order, waits for each prompt before sending the next
command, and writes a verdict per row — with the **complete device output as a cell note** and a
timestamp. Change the job by editing the sheet, not the script.

<!-- VIDEO: on GitHub, edit this README and drag media/demo_cisco_devnet.mp4 onto this line -
     GitHub uploads it and inserts a link that shows as a video player. -->

**Demo** (2:23, no sound): the Cisco sample running against the DevNet IOS-XE sandbox, results filling
in live — [`media/demo_cisco_devnet.mp4`](media/demo_cisco_devnet.mp4). Camera recording of the screen.

---

## Why

Many engineers reach production devices only through **jump hosts where nothing can be installed** —
no Python, no pip, no Ansible, pyATS, Netmiko or Nornir. What *is* usually there: SecureCRT and MS
Office. This script needs exactly those two — SecureCRT runs VBScript, and VBScript drives Excel
over COM with no extra modules.

It also fixes the classic problem of **pasting a batch of commands into a CLI**: Cisco IOS drops
whatever you type while a `ping` runs, so a pasted list of four pings runs only the first. Junos queues
the rest and runs it unseen. One row per command, each sent only after the prompt is back, avoids both —
and every row gets its own verdict.

## What you need

- **SecureCRT** on Windows (tested: 8.7.2, 64-bit) with VBScript available (see
  [VBScript's future](#vbscripts-future)).
- **MS Excel** (tested: Excel 16 / Office 2016+) **or LibreOffice Calc** (tested: 25.8).
- A session in SecureCRT, already logged in and sitting at the prompt.

## How to run

1. Put the workbook somewhere and set its full path in `rdwr.vbs`, line 19:
   `g_strSpreadSheetPath = "C:\securecrt-excel-runner\samples\sample_cisco_devnet.xlsx"`
2. **Excel:** open the workbook in Excel **first**. (LibreOffice: open it, or let the script open it.)
   `g_BACKEND = "auto"` (line 24) uses Excel if it is running, else LibreOffice.
3. In SecureCRT, connect and log in, then **Script → Run… → `rdwr.vbs`**.
4. Watch column H fill in. Hover a cell in H to see the full output of that step.
5. **Save the workbook yourself** — the script never saves or closes it.

## The sheet

The script works on the **second sheet** (`PromptResp` in the samples); row 1 is a header, steps start
at row 2. The first sheet (`HostList`) is not used — it's left over from the VanDyke sample.

| Column | Meaning |
|---|---|
| B | empty = run the row · any text = skip it (`=SKiPPED=`) · `<< END >>` (any spacing / case) = stop |
| C | the **prompt** that ends the command's output, as literal text (e.g. `Router#`) |
| D | the **command** — one per cell (empty = just press Enter) |
| F | **expected tokens**, space separated: the **1st means success**, any other means failure. Empty = Ok unless an error marker appears |
| G | optional **timeout** in seconds for this row |
| H | written by the script: the verdict, with the full output as a **cell note** |
| I | written by the script: date / time |
| J, K | free — the samples use them for "expected H" and an explanation. The script never reads them |

The run stops at `<< END >>` **or at the first row with B, C and D all empty** — so mark blank
separator rows with `x` in B.

**Verdicts in H:**

| H | Means |
|---|---|
| `Ok` | 1st F token seen (or F empty, prompt back, no error marker) |
| `FAiL: got <token>` | a later F token was seen first |
| `FAiL: expected text not seen` | the prompt came back without any F token |
| `FAiL: error text '<marker>'` | F empty and an error marker was in the output |
| `FAiL: timeout Ns` | the prompt didn't come back in time — the note is a snapshot of the screen |
| `FAiL: not sent - no prompt Ns after row N timed out` | see [Timeouts](#timeouts) — the run stops here |
| `=SKiPPED=` | B had text; nothing sent |

## Settings (optional third sheet)

A sheet named **`Settings`**: key in column A, value in column B, from row 2 down to the first empty key.
Without it, the built-in defaults apply.

| Key | Value | Default |
|---|---|---|
| `TIMEOUT` | seconds (per read; column G overrides per row) | `30` |
| `PAGER` | one row per pager prompt, matched by its start | `---(more`, `--More--`, `<--- More --->`, `-- More --`, `---- More ----` |
| `PAGER_KEY` | `SPACE`, `ENTER` or literal text | `SPACE` |
| `ERROR` | one row per error marker, any case | `syntax error`, `unknown command`, `error:`, `% Invalid`, `% Incomplete`, `% Ambiguous`, `% Unknown` |

Any `PAGER` or `ERROR` row **replaces that whole default list** — list every marker you want for that
device. Type a value that starts with `-` with a leading apostrophe (`'---(more`), or the spreadsheet
takes it for a formula. The markers are checked top to bottom and the first one found is reported,
so put generic ones like `error:` last.

## How it behaves

- **Pagers** (`--More--`, `---(more NN%)---`, …) are answered with the pager key, and the pager prompt
  and the device's erase sequence are removed — a paged note is identical to an unpaged one.
  Better still, turn paging off in the first row (`terminal length 0`, `set cli screen-length 0`).
- **Error markers** catch commands that fail without an F token (`% Invalid input`, `syntax error`).
- <a name="timeouts"></a>**Timeouts:** a row that times out is marked and the run carries on — but the
  command may still be running (a long ping) or waiting at a question (`[confirm]`). So before the
  next row the script presses Enter and **waits for that row's prompt**; if it doesn't come within the
  row's timeout, the row is marked `not sent` and the run **stops** instead of typing into an unknown state.
- **One command per cell.** A cell with several lines is sent like a paste: IOS runs only the first line,
  Junos queues and runs the rest — either way only the first is checked. The samples show this on purpose.

## Samples — for demonstration only

Both samples run against **public** devices, so anyone can try them. They are **demonstrations, not
reference results**: public devices change, and your run will differ in places. Columns J and K explain
what each row is meant to show.

| File | Device | Notes |
|---|---|---|
| `samples/sample_cisco_devnet.xlsx` | Cisco DevNet **Catalyst 8000 Always-On** sandbox (IOS-XE 17.15), prompt `Cat8kv#` | Free Cisco account; launch the sandbox for per-reservation credentials, no VPN. It's **shared and reset regularly**: the `192.168.1.x` / `10.10.20.x` ping targets are lab addresses that may or may not answer on the day (in the recorded run `192.168.1.1/.2` didn't). |
| `samples/sample_cisco_devnet_DONE.xlsx` + `samples/logs/…_DONE.log` | same | The run from the video, with results and notes, and the SecureCRT session log of it. |
| `samples/sample_junos_att.xlsx` | AT&T public route server `route-server.ip.att.net` (Junos 23.2), prompt `rviews@route-server.ip.att.net>` | Login `rviews` / `rviews` (published in its banner). A **restricted** account: `show version` and `show system …` are refused — those rows show how a refused command is caught. It is a production box offered as a public service: **run it once, not in a loop.** |
| `samples/sample_junos_att_DONE.xlsx` + `samples/logs/…_DONE.log` | same | A real run, results and log (one third-party IP in the log's login line blanked). |

Change column C to your device's exact prompt before using a sample elsewhere.

## Tested on

| Device | OS | Via |
|---|---|---|
| DevNet Catalyst 8000v (sandbox) | IOS-XE 17.15 | Excel and LibreOffice |
| RouteViews ASR1004 (public, production) | IOS 15.5 / IOS-XE 3.16, user mode | LibreOffice |
| Juniper SRX320 (lab) | Junos 23.4 | LibreOffice |
| AT&T route server (public, production) | Junos 23.2 | Excel and LibreOffice |

## Tests (no device needed)

`tests/fake_crt.vbs` is a fake SecureCRT with a simulated Junos / IOS device (pagers, error text, slow
pings that drop or queue typed input). The tests run the real `Main` from `rdwr.vbs` against it:

```
cscript //nologo tests\test_logic.vbs calc      68 checks, LibreOffice
cscript //nologo tests\test_logic.vbs excel     68 checks, MS Excel (close Excel first)
cscript //nologo tests\test_backend.vbs         37 checks, LibreOffice backend
cscript //nologo tests\test_excel_backend.vbs   60 checks, Excel backend (close Excel first)
```

`tools/inspect_workbook.ps1 -Path <book.xlsx>` prints the rows and notes of a saved workbook
without opening Excel.

## Limitations

- The timeout applies **per read**, not per row: a command paged N times can take up to N × timeout.
- F tokens can't contain spaces (space is the separator). The prompt in C is literal text, not a pattern.
- A row with a prompt but no command still presses Enter — handy for manual steps.
- Excel keeps at most 32,767 characters in a note; longer output is cut with a marker saying so.
- SecureCRT 8.7.2's embedded Python 2.7 can't reach Excel (no pywin32, no `_ctypes`), which is why
  this is VBScript. Newer SecureCRT versions may differ.

## VBScript's future

Microsoft is phasing VBScript out of Windows: an optional feature, on by default, since Windows 11 24H2
/ Server 2025; **off by default around 2027**; removed later
([Microsoft's timeline](https://techcommunity.microsoft.com/blog/windows-itpro-blog/vbscript-deprecation-timelines-and-next-steps/4148301)).
Older Windows keeps it. Where it's gone, SecureCRT also runs **JScript**, which can drive Excel the same
way (`new ActiveXObject("Excel.Application")`) — the natural successor for the same no-install setup.

## Safety

The script sends **exactly what column D says** to whatever session is active. Review a workbook before
you point it at production, start with read-only `show` commands, and be polite to public servers.
No warranty — see the licence.

## History

The first version was written for **ProComm Plus** more than 20 years ago, to program Nortel PABXes and
other devices of the day. It was rebuilt on VanDyke's SecureCRT sample
`ExcelSpreadsheets-ReadingAndWriting.vbs`
([VanDyke scripting examples](https://www.vandyke.com/support/scripting/scripting-examples/)) — line 1
of `rdwr.vbs` keeps that attribution — and extended in 2026 with expected-output checks, pager
handling, timeouts with re-sync, error markers, a Settings sheet, a LibreOffice backend and offline tests.

The 2026 rework (fixes, tests, README) was written with Claude (Anthropic) as a coding assistant;
design, requirements and all live-device testing by the author.

## License

[MIT](LICENSE) — free to use, change and share; no warranty.
