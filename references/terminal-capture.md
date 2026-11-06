# Terminal Capture

Use `scripts/capture_terminal.py` for commands whose meaningful evidence is terminal text.

## Capture a Command

```powershell
python scripts/capture_terminal.py `
  --command "ip -4 addr show eth1" `
  --shell wsl `
  --title "确认实验接口" `
  --output screenshots/01-interface.png
```

Supported shells: `powershell`, `cmd`, `bash`, and `wsl`.

The PNG includes:

- the executed command;
- stdout and stderr;
- exit status;
- capture time.

The tool strips ANSI terminal escape sequences and limits very long output with an explicit
omission marker.

## Redact Sensitive Values

Repeat `--redact` with regular expressions:

```powershell
python scripts/capture_terminal.py `
  --command "printenv" `
  --redact "(?i)token=.*" `
  --redact "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}" `
  --title "环境变量检查" `
  --output screenshots/02-env.png
```

Redaction changes only the displayed command and output. The original command still runs.

## Capture Existing Text

```powershell
python scripts/capture_terminal.py `
  --input captures/command-output.txt `
  --title "命令结果" `
  --output screenshots/03-result.png
```

## Handoff Rule

The Agent may capture its own clean experiment output. When the user must reproduce the
experiment, leave `<!-- SHOT: ... -->` plus the image line in Markdown, and stop before DOCX
conversion until the user supplies the required PNG files.
